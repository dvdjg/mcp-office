import { listStyles } from '../../../src/tools/word/styles.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';

// Mock de las dependencias externas (simulando la interacción con Office)
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
}));
jest.mock('../../../src/utils/logger');

const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;
const mockLogger = logger as jest.Mocked<typeof logger>;

describe('word/styles integration tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockStyles: any;
    let mockOfficeAppInstance: any;

    beforeEach(() => {
        jest.clearAllMocks();

        // Configurar mocks para simular objetos COM de Word y la interacción con officeInterop
        mockStyles = {
            Count: 3,
            Item: jest.fn((index: number) => {
                if (index === 1) return { NameLocal: 'Normal' };
                if (index === 2) return { NameLocal: 'Heading 1' };
                if (index === 3) return { NameLocal: 'Heading 2' };
                return undefined;
            }),
        };
        mockDoc = {
            Styles: mockStyles,
            Close: jest.fn(),
        };
        mockWordApp = {
            Documents: {
                Open: jest.fn().mockReturnValue(mockDoc),
            },
        };

        // Mock de la estructura de OfficeAppInstance devuelta por getOfficeApplication
        mockOfficeAppInstance = {
            app: mockWordApp,
            openDocument: jest.fn().mockReturnValue(mockDoc), // Simula openDocument de officeInterop
            release: jest.fn(), // Simula release de officeInterop
        };

        mockGetOfficeApplication.mockResolvedValue(mockOfficeAppInstance);
        mockValidateFilePath.mockReturnValue(true); // Asumir que la validación de ruta es exitosa
    });

    // Prueba de integración para listStyles
    test('listStyles debería interactuar correctamente con officeInterop para listar estilos', async () => {
        const params = { filePath: 'C:/test/document.docx' };

        // Ejecutar la herramienta
        const result = await listStyles(params);

        // Verificar interacciones con officeInterop y mocks de COM
        expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
        expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
        // En una prueba de integración, verificamos que openDocument de officeInterop es llamado
        expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
        // Verificamos que se accede a la colección de estilos del documento mockeado
        expect(mockDoc.Styles).toBe(mockStyles);
        // Verificamos que se itera sobre los estilos mockeados
        expect(mockStyles.Item).toHaveBeenCalledWith(1);
        expect(mockStyles.Item).toHaveBeenCalledWith(2);
        expect(mockStyles.Item).toHaveBeenCalledWith(3);

        // Verificamos que los objetos COM mockeados son liberados
        expect(mockReleaseObject).toHaveBeenCalledWith({ NameLocal: 'Normal' });
        expect(mockReleaseObject).toHaveBeenCalledWith({ NameLocal: 'Heading 1' });
        expect(mockReleaseObject).toHaveBeenCalledWith({ NameLocal: 'Heading 2' });
        expect(mockDoc.Close).toHaveBeenCalledWith(false); // Documento cerrado
        expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc); // Documento liberado
        expect(mockOfficeAppInstance.release).toHaveBeenCalled(); // Instancia de OfficeApp liberada

        // Verificar el resultado de la herramienta
        expect(result).toEqual({ success: true, data: ['Normal', 'Heading 1', 'Heading 2'] });
        expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Attempting to list styles for document'));
        expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully listed'));
    });

    test('listStyles debería manejar errores al abrir el documento', async () => {
        // Simular que openDocument falla
        mockOfficeAppInstance.openDocument.mockReturnValue(null);
        const params = { filePath: 'C:/test/nonexistent.docx' };

        const result = await listStyles(params);

        expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
        expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
        expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);

        // No debería intentar acceder a Styles si el documento no se abrió
        expect(mockDoc.Styles).toBeUndefined(); // O verificar que no se accede a propiedades de mockDoc

        // Debería liberar la instancia de OfficeApp
        expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        // No debería intentar cerrar o liberar mockDoc si es null
        expect(mockDoc.Close).not.toHaveBeenCalled();
        expect(mockReleaseObject).not.toHaveBeenCalledWith(mockDoc);

        // Verificar el resultado de error
        expect(result.success).toBe(false);
        expect(result.error?.message).toContain('Failed to open document');
        expect(mockLogger.error).toHaveBeenCalledWith(expect.stringContaining('Error listing styles via COM'));
    });

     test('listStyles debería manejar errores al acceder a estilos individuales', async () => {
         // Simular un error al acceder al segundo estilo
        mockStyles.Item.mockImplementation((index: number) => {
            if (index === 1) return { NameLocal: 'Normal' };
            if (index === 2) throw new Error('COM Error accessing style');
            if (index === 3) return { NameLocal: 'Heading 2' };
            return undefined;
        });

        const params = { filePath: 'C:/test/document.docx' };
        const result = await listStyles(params);

        expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
        expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
        expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
        expect(mockDoc.Styles).toBe(mockStyles);
        expect(mockStyles.Item).toHaveBeenCalledWith(1);
        expect(mockStyles.Item).toHaveBeenCalledWith(2);
        expect(mockStyles.Item).toHaveBeenCalledWith(3);

        // Debería intentar liberar el primer y tercer estilo, pero no el segundo que falló
        expect(mockReleaseObject).toHaveBeenCalledWith({ NameLocal: 'Normal' });
        expect(mockReleaseObject).not.toHaveBeenCalledWith({ NameLocal: 'Heading 1' }); // Este falló
        expect(mockReleaseObject).toHaveBeenCalledWith({ NameLocal: 'Heading 2' });
        expect(mockDoc.Close).toHaveBeenCalledWith(false);
        expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
        expect(mockOfficeAppInstance.release).toHaveBeenCalled();

        // Debería devolver los estilos que pudo obtener y loggear el error
        expect(result).toEqual({ success: true, data: ['Normal', 'Heading 2'] });
        expect(mockLogger.error).toHaveBeenCalledWith(expect.stringContaining('Error accessing style at index 2'));
        expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully listed'));
    });
});