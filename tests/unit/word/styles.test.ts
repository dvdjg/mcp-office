import { applyStyle, listStyles } from '../../../src/tools/word/styles.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';

// Mock de las dependencias
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

describe('word/styles unit tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockSelection: any;
    let mockRange: any;
    let mockParagraphs: any;
    let mockStyles: any;

    beforeEach(() => {
        // Reset mocks antes de cada prueba
        jest.clearAllMocks();

        // Configurar mocks básicos para Word COM objects
        mockRange = { Style: '' };
        mockSelection = { Range: mockRange };
        mockParagraphs = {
            Count: 5,
            Item: jest.fn((index: number) => {
                if (index > 0 && index <= mockParagraphs.Count) {
                    return { Range: mockRange };
                }
                return undefined; // Simular COM object no encontrado
            }),
        };
        mockStyles = {
            Count: 3,
            Item: jest.fn((index: number) => {
                if (index === 1) return { NameLocal: 'Normal' };
                if (index === 2) return { NameLocal: 'Heading 1' };
                if (index === 3) return { NameLocal: 'Heading 2' };
                return undefined; // Simular COM object no encontrado
            }),
        };
        mockDoc = {
            Content: mockRange,
            Paragraphs: mockParagraphs,
            Styles: mockStyles,
            Close: jest.fn(),
        };
        mockWordApp = {
            Documents: {
                Open: jest.fn().mockResolvedValue(mockDoc),
            },
            Selection: mockSelection,
            Quit: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockWordApp);
        mockValidateFilePath.mockReturnValue(true); // Asumir que la validación de ruta es exitosa por defecto
    });

    // --- Pruebas Unitarias para applyStyle ---

    describe('applyStyle', () => {
        const baseParams = { filePath: 'C:/test/document.docx', style: 'Heading 1' };

        test('debería aplicar un estilo a la selección', async () => {
            const params = { ...baseParams, range: 'selection' };
            const result = await applyStyle(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath);
            expect(mockWordApp.Selection.Range).toBe(mockRange);
            expect(mockRange.Style).toBe(params.style);
            expect(mockDoc.Close).toHaveBeenCalledWith(false); // Cerrar sin guardar
            expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully applied style'));
        });

        test('debería aplicar un estilo al documento completo', async () => {
            const params = { ...baseParams, range: 'document' };
            const result = await applyStyle(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath);
            expect(mockDoc.Content).toBe(mockRange);
            expect(mockRange.Style).toBe(params.style);
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully applied style'));
        });

        test('debería aplicar un estilo a un párrafo específico', async () => {
            const params = { ...baseParams, range: 'paragraph:3' };
            const result = await applyStyle(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(3);
            expect(mockRange.Style).toBe(params.style);
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully applied style'));
        });

        test('debería manejar un filePath inválido', async () => {
            mockValidateFilePath.mockReturnValue(false);
            const params = { ...baseParams, range: 'document' };
            const result = await applyStyle(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            // No debería llamar a getOfficeApplication si la validación falla
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Invalid or potentially unsafe file path');
        });

        test('debería manejar un rango no soportado', async () => {
            const params = { ...baseParams, range: 'invalidRange' };
            const result = await applyStyle(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Unsupported range format');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
        });

        test('debería manejar un índice de párrafo inválido (no numérico)', async () => {
            const params = { ...baseParams, range: 'paragraph:abc' };
            const result = await applyStyle(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Invalid paragraph index format');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
        });

        test('debería manejar un índice de párrafo fuera de límites', async () => {
            const params = { ...baseParams, range: 'paragraph:10' }; // Solo hay 5 párrafos mockeados
            const result = await applyStyle(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(10);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Paragraph index 10 is out of bounds');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
        });

        test('debería manejar un error al abrir el documento', async () => {
            mockWordApp.Documents.Open.mockResolvedValue(null); // Simular fallo al abrir
            const params = { ...baseParams, range: 'document' };
            const result = await applyStyle(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Failed to open document');
            // No debería intentar cerrar doc si es null
            expect(mockDoc.Close).not.toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
        });

         test('debería manejar un error al obtener la selección', async () => {
            mockSelection.Range = null; // Simular fallo al obtener la selección
            const params = { ...baseParams, range: 'selection' };
            const result = await applyStyle(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Could not get range from selection');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
        });
    });

    // --- Pruebas Unitarias para listStyles ---

    describe('listStyles', () => {
        const baseParams = { filePath: 'C:/test/document.docx' };

        test('debería listar los estilos del documento', async () => {
            const result = await listStyles(baseParams);

            expect(mockValidateFilePath).toHaveBeenCalledWith(baseParams.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(baseParams.filePath);
            expect(mockDoc.Styles).toBe(mockStyles);
            expect(mockStyles.Count).toBe(3);
            expect(mockStyles.Item).toHaveBeenCalledWith(1);
            expect(mockStyles.Item).toHaveBeenCalledWith(2);
            expect(mockStyles.Item).toHaveBeenCalledWith(3);
            expect(mockReleaseObject).toHaveBeenCalledWith({ NameLocal: 'Normal' });
            expect(mockReleaseObject).toHaveBeenCalledWith({ NameLocal: 'Heading 1' });
            expect(mockReleaseObject).toHaveBeenCalledWith({ NameLocal: 'Heading 2' });
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            expect(result).toEqual({ success: true, data: ['Normal', 'Heading 1', 'Heading 2'] });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully listed'));
        });

        test('debería manejar un filePath inválido', async () => {
            mockValidateFilePath.mockReturnValue(false);
            const result = await listStyles(baseParams);

            expect(mockValidateFilePath).toHaveBeenCalledWith(baseParams.filePath);
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Invalid or potentially unsafe file path');
        });

        test('debería manejar un error al abrir el documento', async () => {
            mockWordApp.Documents.Open.mockResolvedValue(null); // Simular fallo al abrir
            const result = await listStyles(baseParams);

            expect(mockValidateFilePath).toHaveBeenCalledWith(baseParams.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(baseParams.filePath);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Failed to open document');
            // No debería intentar cerrar doc si es null
            expect(mockDoc.Close).not.toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
        });

        test('debería manejar errores al acceder a estilos individuales', async () => {
             // Simular un error al acceder al segundo estilo
            mockStyles.Item.mockImplementation((index: number) => {
                if (index === 1) return { NameLocal: 'Normal' };
                if (index === 2) throw new Error('COM Error accessing style');
                if (index === 3) return { NameLocal: 'Heading 2' };
                return undefined;
            });

            const result = await listStyles(baseParams);

            expect(mockValidateFilePath).toHaveBeenCalledWith(baseParams.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(baseParams.filePath);
            expect(mockDoc.Styles).toBe(mockStyles);
            expect(mockStyles.Count).toBe(3);
            expect(mockStyles.Item).toHaveBeenCalledWith(1);
            expect(mockStyles.Item).toHaveBeenCalledWith(2);
            expect(mockStyles.Item).toHaveBeenCalledWith(3);
            // Debería haber intentado liberar el primer y tercer estilo, pero no el segundo que falló
            expect(mockReleaseObject).toHaveBeenCalledWith({ NameLocal: 'Normal' });
            expect(mockReleaseObject).not.toHaveBeenCalledWith({ NameLocal: 'Heading 1' }); // Este falló
            expect(mockReleaseObject).toHaveBeenCalledWith({ NameLocal: 'Heading 2' });
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            // Debería devolver los estilos que pudo obtener
            expect(result).toEqual({ success: true, data: ['Normal', 'Heading 2'] });
            expect(mockLogger.error).toHaveBeenCalledWith(expect.stringContaining('Error accessing style at index 2'));
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully listed'));
        });
    });
});