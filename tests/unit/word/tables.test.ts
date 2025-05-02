import { insertTable } from '../../../src/tools/word/tables.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { z } from 'zod';

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

describe('word/tables unit tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockSelection: any;
    let mockRange: any;
    let mockParagraphs: any;
    let mockTables: any;
    let mockOfficeAppInstance: any;

    beforeEach(() => {
        // Reset mocks antes de cada prueba
        jest.clearAllMocks();

        // Configurar mocks básicos para Word COM objects
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
                    return { Range: { Start: (index - 1) * 15, End: (index * 15) -1, Collapse: jest.fn() } }; // Mock simple paragraph range
                }
                return undefined; // Simular COM object no encontrado
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

        // Mock the OfficeAppInstance structure
        mockOfficeAppInstance = {
            app: mockWordApp,
            openDocument: jest.fn().mockReturnValue(mockDoc),
            release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockOfficeAppInstance);
        mockValidateFilePath.mockReturnValue(true); // Asumir que la validación de ruta es exitosa por defecto
    });

    // --- Pruebas Unitarias para insertTable ---

    describe('insertTable', () => {
        const baseParams = { filePath: 'C:/test/document.docx', rows: 3, columns: 2 };

        test('debería insertar una tabla al final del documento por defecto', async () => {
            const params = { ...baseParams }; // position is optional, defaults to end
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            // Verifica que se obtiene el rango al final del contenido
            expect(mockDoc.Range).toHaveBeenCalledWith(mockDoc.Content.End, mockDoc.Content.End);
            // Verifica que se inserta un párrafo después y se colapsa el rango
            expect(mockRange.InsertParagraphAfter).toHaveBeenCalled();
            expect(mockRange.Collapse).toHaveBeenCalledWith(0); // wdCollapseEnd
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

         test('debería insertar una tabla al final de un documento vacío', async () => {
            mockDoc.Content.End = 0; // Simular documento vacío
            const params = { ...baseParams };
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            // Verifica que se obtiene el rango al inicio (0, 0)
            expect(mockDoc.Range).toHaveBeenCalledWith(0, 0);
            // No debería insertar párrafo si el documento está vacío
            expect(mockRange.InsertParagraphAfter).not.toHaveBeenCalled();
            // Debería colapsar al inicio (que es 0,0)
            expect(mockRange.Collapse).toHaveBeenCalledWith(0); // wdCollapseEnd
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


        test('debería insertar una tabla en la selección', async () => {
            const params = { ...baseParams, position: 'selection' };
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
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

        test('debería insertar una tabla antes de un párrafo específico', async () => {
            const params = { ...baseParams, position: 'paragraph:3' };
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(3);
            // Verifica que se obtiene el rango del párrafo y se colapsa al inicio
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

        test('debería aplicar un estilo a la tabla si se especifica', async () => {
            const params = { ...baseParams, style: 'Table Grid' };
            const mockTable = { Index: 1, Style: '' };
            mockTables.Add.mockReturnValue(mockTable); // Ensure mockTable is returned

            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(mockTables.Add).toHaveBeenCalledWith(mockRange, params.rows, params.columns, 0, 0);
            expect(mockTable.Style).toBe(params.style); // Verifica que el estilo fue aplicado
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockTable); // La tabla insertada
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(result).toEqual({ success: true, data: { message: 'Table (3x2) inserted successfully.' } });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Executing word/tables/insert tool'));
            expect(mockLogger.debug).toHaveBeenCalledWith(expect.stringContaining('Table inserted successfully'));
            expect(mockLogger.debug).toHaveBeenCalledWith(expect.stringContaining('Style "Table Grid" applied successfully.'));
        });

        test('debería manejar un filePath inválido', async () => {
            mockValidateFilePath.mockReturnValue(false);
            const params = { ...baseParams };
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Invalid or potentially unsafe file path');
        });

        test('debería manejar un número de filas inválido (cero)', async () => {
            const params = { ...baseParams, rows: 0 };
            const result = await insertTable(params);

            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Number of rows must be a positive integer.');
            expect(mockGetOfficeApplication).not.toHaveBeenCalled(); // No debería proceder si la validación falla
        });

         test('debería manejar un número de columnas inválido (negativo)', async () => {
            const params = { ...baseParams, columns: -1 };
            const result = await insertTable(params);

            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Number of columns must be a positive integer.');
            expect(mockGetOfficeApplication).not.toHaveBeenCalled(); // No debería proceder si la validación falla
        });


        test('debería manejar una posición de párrafo inválida (no numérica)', async () => {
            const params = { ...baseParams, position: 'paragraph:abc' };
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Invalid paragraph index');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('debería manejar una posición de párrafo fuera de límites', async () => {
            const params = { ...baseParams, position: 'paragraph:10' }; // Solo hay 5 párrafos mockeados
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(10);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Paragraph index 10 out of bounds');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('debería manejar un error al abrir el documento', async () => {
            mockOfficeAppInstance.openDocument.mockReturnValue(null); // Simular fallo al abrir
            const params = { ...baseParams };
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Failed to open document');
            // No debería intentar cerrar doc si es null
            expect(mockDoc.Close).not.toHaveBeenCalled();
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

         test('debería manejar un error al obtener la selección', async () => {
            mockSelection.Range = null; // Simular fallo al obtener la selección
            const params = { ...baseParams, position: 'selection' };
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Cannot insert at selection: No selection found.');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('debería manejar un error al insertar la tabla', async () => {
            mockTables.Add.mockReturnValue(null); // Simular fallo al insertar la tabla
            const params = { ...baseParams };
            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(mockTables.Add).toHaveBeenCalledWith(mockRange, params.rows, params.columns, 0, 0);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Failed to insert table.');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(null); // table should be null
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange should still be released
        });

         test('debería manejar un error al aplicar el estilo', async () => {
            const params = { ...baseParams, style: 'Invalid Style' };
            const mockTable = { Index: 1, Style: '' };
            mockTables.Add.mockReturnValue(mockTable);
            // Simular un error al intentar aplicar el estilo
            Object.defineProperty(mockTable, 'Style', {
                set: jest.fn(() => { throw new Error('COM Error applying style'); }),
                get: jest.fn(() => ''),
                configurable: true,
            });

            const result = await insertTable(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(mockTables.Add).toHaveBeenCalledWith(mockRange, params.rows, params.columns, 0, 0);
            expect(mockTable.Style).toBe(params.style); // The setter was called
            expect(mockDoc.Save).toHaveBeenCalled(); // Debería guardar a pesar del error de estilo
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockTable); // La tabla insertada
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(result).toEqual({ success: true, data: { message: 'Table (3x2) inserted successfully.' } }); // Debería reportar éxito ya que la tabla se insertó
            expect(mockLogger.warn).toHaveBeenCalledWith(expect.stringContaining('Failed to apply style')); // Debería loggear una advertencia
        });
    });
});