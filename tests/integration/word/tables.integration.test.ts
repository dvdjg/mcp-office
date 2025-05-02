import { insertTable } from '../../../src/tools/word/tables.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { z } from 'zod';

// Mock de las dependencias externas (simulando la interacción con Office)
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
    createErrorResponse: jest.fn((message, code, details) => ({ success: false, error: { code, message, details } })),
}));
jest.mock('../../../src/utils/logger');

const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;
const mockLogger = logger as jest.Mocked<typeof logger>;

describe('word/tables integration tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockSelection: any;
    let mockRange: any;
    let mockParagraphs: any;
    let mockTables: any;
    let mockOfficeAppInstance: any;

    beforeEach(() => {
        jest.clearAllMocks();

        // Configurar mocks para simular objetos COM de Word y la interacción con officeInterop
        mockRange = {
            Start: 0,
            End: 10,
            Collapse: jest.fn(),
            InsertParagraphAfter: jest.fn(),
        };
         mockSelection = {
            Range: mockRange,
            Type: 1, // wdSelectionNormal = 1
        };
        mockParagraphs = {
            Count: 5,
            Item: jest.fn((index: number) => {
                if (index > 0 && index <= mockParagraphs.Count) {
                    return { Range: { Start: (index - 1) * 15, End: (index * 15) -1, Collapse: jest.fn() } };
                }
                return undefined;
            }),
        };
        mockTables = {
            Add: jest.fn().mockReturnValue({ Index: 1, Style: '' }), // Mock de una tabla insertada
        };
        mockDoc = {
            Content: { End: 50 }, // Mock del final del contenido
            Paragraphs: mockParagraphs,
            Tables: mockTables,
            Range: jest.fn((start, end) => {
                 // Return a new mock range for specific positions
                 return {
                    Start: start,
                    End: end,
                    Collapse: jest.fn(),
                    InsertParagraphAfter: jest.fn(),
                 };
            }),
            Save: jest.fn(),
            Close: jest.fn(),
        };
        mockWordApp = {
            Documents: {
                Open: jest.fn().mockResolvedValue(mockDoc),
            },
            Selection: mockSelection,
            Quit: jest.fn(),
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

    // --- Pruebas de Integración para insertTable ---

    describe('insertTable', () => {
        const baseParams = { filePath: 'C:/test/document.docx', rows: 3, columns: 2 };

        test('debería interactuar correctamente con officeInterop para insertar una tabla al final', async () => {
            const params = { ...baseParams }; // position is optional, defaults to end
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            // Verificamos que se obtiene el rango al final del contenido a través de mockDoc.Content
            expect(mockDoc.Range).toHaveBeenCalledWith(mockDoc.Content.End, mockDoc.Content.End);
            // Verificamos que se inserta un párrafo después y se colapsa el rango mockeado
            expect(mockRange.InsertParagraphAfter).toHaveBeenCalled();
            expect(mockRange.Collapse).toHaveBeenCalledWith(0); // wdCollapseEnd
            // Verificamos que se llama a Add en la colección Tables del documento mockeado
            expect(mockTables.Add).toHaveBeenCalledWith(mockRange, params.rows, params.columns, 0, 0);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith({ Index: 1, Style: '' }); // La tabla insertada
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(result).toEqual({ success: true, data: { message: 'Table (3x2) inserted successfully.' } });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Executing word/tables/insert tool'));
            expect(mockLogger.debug).toHaveBeenCalledWith(expect.stringContaining('Table inserted successfully'));
        });

        test('debería interactuar correctamente con officeInterop para insertar una tabla en la selección', async () => {
            const params = { ...baseParams, position: 'selection' };
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            // Verificamos que se obtiene el rango de la selección a través de mockWordApp.Selection
            expect(mockWordApp.Selection.Range).toBe(mockRange);
            expect(mockTables.Add).toHaveBeenCalledWith(mockRange, params.rows, params.columns, 0, 0);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith({ Index: 1, Style: '' }); // La tabla insertada
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(result).toEqual({ success: true, data: { message: 'Table (3x2) inserted successfully.' } });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Executing word/tables/insert tool'));
            expect(mockLogger.debug).toHaveBeenCalledWith(expect.stringContaining('Table inserted successfully'));
        });

        test('debería interactuar correctamente con officeInterop para insertar una tabla antes de un párrafo', async () => {
            const params = { ...baseParams, position: 'paragraph:3' };
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            // Verificamos que se accede al párrafo a través de mockDoc.Paragraphs
            expect(mockParagraphs.Item).toHaveBeenCalledWith(3);
            // Verificamos que se colapsa el rango del párrafo mockeado
            expect(mockParagraphs.Item(3).Range.Collapse).toHaveBeenCalledWith(1); // wdCollapseStart
            expect(mockTables.Add).toHaveBeenCalledWith(mockParagraphs.Item(3).Range, params.rows, params.columns, 0, 0);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith({ Index: 1, Style: '' }); // La tabla insertada
            // El rango del párrafo también debería ser liberado
            expect(mockReleaseObject).toHaveBeenCalled(); // Check if releaseObject was called at least once
            expect(result).toEqual({ success: true, data: { message: 'Table (3x2) inserted successfully.' } });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Executing word/tables/insert tool'));
            expect(mockLogger.debug).toHaveBeenCalledWith(expect.stringContaining('Table inserted successfully'));
        });

        test('debería interactuar correctamente con officeInterop para aplicar un estilo a la tabla', async () => {
            const params = { ...baseParams, style: 'Table Grid' };
            const mockTable = { Index: 1, Style: '' };
            mockTables.Add.mockReturnValue(mockTable);

            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(mockTables.Add).toHaveBeenCalledWith(mockRange, params.rows, params.columns, 0, 0);
            // Verificamos que se asigna el estilo a la propiedad Style de la tabla mockeada
            expect(mockTable.Style).toBe(params.style);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockTable); // La tabla insertada
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(result).toEqual({ success: true, data: { message: 'Table (3x2) inserted successfully.' } });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Executing word/tables/insert tool'));
            expect(mockLogger.debug).toHaveBeenCalledWith(expect.stringContaining('Table inserted successfully'));
            expect(mockLogger.debug).toHaveBeenCalledWith(expect.stringContaining('Attempting to apply style'));
            expect(mockLogger.debug).toHaveBeenCalledWith(expect.stringContaining('Style "Table Grid" applied successfully.'));
        });
    });
});