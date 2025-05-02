import { insertTable } from '../../../src/tools/word/tables.tool';
import { modifyText } from '../../../src/tools/word/text.tool';
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

describe('word/tables and word/text integration tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockSelection: any;
    let mockRange: any;
    let mockTables: any;
    let mockTable: any;
    let mockCell: any;
    let mockCellRange: any;
    let mockOfficeAppInstance: any;

    beforeEach(() => {
        jest.clearAllMocks();

        // Configurar mocks para simular objetos COM de Word y la interacción con officeInterop
        mockCellRange = {
            Text: 'Initial Cell Text', // Texto inicial de la celda
            Delete: jest.fn(),
            Collapse: jest.fn(),
        };
        // Mock setter para Text en mockCellRange
        Object.defineProperty(mockCellRange, 'Text', {
            set: jest.fn((value) => { mockCellRange._text = value; }),
            get: jest.fn(() => mockCellRange._text || 'Initial Cell Text'), // Devolver valor asignado o inicial
            configurable: true,
        });


        mockCell = {
            Range: mockCellRange,
        };
        mockTable = {
            Index: 1,
            Style: '',
            Cell: jest.fn((row: number, col: number) => {
                // Devolver la celda mockeada para una celda específica (ej. 1, 1)
                if (row === 1 && col === 1) {
                    return mockCell;
                }
                // Devolver un mock genérico o undefined para otras celdas si es necesario
                return { Range: { Text: `Cell ${row},${col}`, Delete: jest.fn(), Collapse: jest.fn() } };
            }),
        };
        mockTables = {
            Add: jest.fn().mockReturnValue(mockTable), // Devolver la tabla mockeada al añadir
            Item: jest.fn((index: number) => {
                 if (index === 1) return mockTable; // Devolver la tabla mockeada por índice
                 return undefined;
            }),
            Count: 1, // Simular que hay una tabla después de insertarla
        };
        mockRange = { // Rango para inserción inicial de tabla
            Start: 0,
            End: 10,
            Collapse: jest.fn(),
            InsertParagraphAfter: jest.fn(),
        };
        mockSelection = { // Mock de selección si fuera necesario
            Range: mockRange,
            Type: 1, // wdSelectionNormal
        };
        mockDoc = {
            Content: { End: 50 },
            Tables: mockTables,
            Range: jest.fn().mockReturnValue(mockRange), // Rango para inserción
            Save: jest.fn(),
            Close: jest.fn(),
        };
        mockWordApp = {
            Documents: {
                Open: jest.fn().mockResolvedValue(mockDoc),
            },
            Selection: mockSelection, // Añadir mock de selección
        };

        // Mock de la estructura de OfficeAppInstance
        mockOfficeAppInstance = {
            app: mockWordApp,
            openDocument: jest.fn().mockReturnValue(mockDoc),
            release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockOfficeAppInstance);
        mockValidateFilePath.mockReturnValue(true);
    });

    test('debería insertar una tabla y luego modificar el texto de una celda', async () => {
        const filePath = 'C:/test/integration_doc.docx';
        const insertParams = { filePath, rows: 2, columns: 2, position: 'end' };
        const modifyParams = { filePath, range: 'table:1:cell:1:1', newText: 'Modified Cell Text' }; // Modificar celda (1,1) de la tabla 1

        // --- Paso 1: Insertar la tabla ---
        const insertResult = await insertTable(insertParams);

        // Verificar la inserción (simplificado, asumimos que funciona basado en pruebas unitarias)
        expect(insertResult.success).toBe(true);
        expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
        expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(filePath);
        expect(mockTables.Add).toHaveBeenCalledWith(mockRange, insertParams.rows, insertParams.columns, 0, 0);
        expect(mockDoc.Save).toHaveBeenCalledTimes(1); // Guardado después de insertar
        expect(mockDoc.Close).toHaveBeenCalledTimes(1); // Cerrado después de insertar
        expect(mockOfficeAppInstance.release).toHaveBeenCalledTimes(1); // Liberado después de insertar
        expect(mockReleaseObject).toHaveBeenCalledWith(mockTable);
        expect(mockReleaseObject).toHaveBeenCalledWith(mockRange);

        // Reset mocks para la segunda llamada (simulando reapertura del documento)
        // Es importante resetear mocks específicos que se esperan ser llamados de nuevo
        mockOfficeAppInstance.openDocument.mockClear();
        mockDoc.Save.mockClear();
        mockDoc.Close.mockClear();
        mockOfficeAppInstance.release.mockClear();
        mockReleaseObject.mockClear();
        // Asegurarse de que getOfficeApplication devuelva la misma instancia mockeada
        mockGetOfficeApplication.mockResolvedValue(mockOfficeAppInstance);
        // Asegurarse de que openDocument devuelva el mismo documento mockeado
        mockOfficeAppInstance.openDocument.mockReturnValue(mockDoc);


        // --- Paso 2: Modificar texto en la celda ---
        const modifyResult = await modifyText(modifyParams);

        // Verificar la modificación
        expect(modifyResult.success).toBe(true);
        // Verificar que se volvió a obtener la aplicación y abrir el documento
        expect(mockGetOfficeApplication).toHaveBeenCalledTimes(2); // Llamado de nuevo
        expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(filePath, false, false); // Llamado de nuevo para modificar

        // Verificar que se accedió a la tabla y celda correctas
        expect(mockTables.Item).toHaveBeenCalledWith(1); // Acceder a la tabla 1
        expect(mockTable.Cell).toHaveBeenCalledWith(1, 1); // Acceder a la celda (1,1)

        // Verificar que se modificó el texto del rango de la celda
        expect(mockCell.Range).toBe(mockCellRange);
        expect(mockCellRange.Text).toBe(modifyParams.newText); // Verificar que el setter fue llamado con el nuevo texto

        // Verificar guardado, cierre y liberación final
        expect(mockDoc.Save).toHaveBeenCalledTimes(1); // Guardado después de modificar
        expect(mockDoc.Close).toHaveBeenCalledTimes(1); // Cerrado después de modificar
        expect(mockOfficeAppInstance.release).toHaveBeenCalledTimes(1); // Liberado después de modificar
        expect(mockReleaseObject).toHaveBeenCalledWith(mockCellRange); // Rango de la celda liberado
        expect(mockReleaseObject).toHaveBeenCalledWith(mockCell); // Celda liberada
        expect(mockReleaseObject).toHaveBeenCalledWith(mockTable); // Tabla liberada

        // Verificar logs
        expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Executing word/tables/insert tool'));
        expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Executing word/text/modify tool'));
        expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully modified text in table 1, cell (1, 1)'));
    });

     test('debería manejar un error si la tabla especificada para modificar no existe', async () => {
        const filePath = 'C:/test/integration_doc.docx';
        // Intentar modificar una tabla que no existe (mockTables.Item devolverá undefined)
        const modifyParams = { filePath, range: 'table:2:cell:1:1', newText: 'Modified Cell Text' };

        // Simular que no hay tabla 2
        mockTables.Item.mockImplementation((index: number) => {
            if (index === 1) return mockTable;
            return undefined; // Tabla 2 no encontrada
        });

        const modifyResult = await modifyText(modifyParams);

        expect(modifyResult.success).toBe(false);
        expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
        expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(filePath, false, false);
        expect(mockTables.Item).toHaveBeenCalledWith(2); // Intento de acceder a tabla 2
        expect(mockTable.Cell).not.toHaveBeenCalled(); // No debería intentar acceder a la celda
        if (!modifyResult.success) {
            expect(modifyResult.error.message).toContain('Table index 2 is out of bounds');
        } else {
            fail('Expected modifyResult to be an error response');
        }

        // Verificar cierre y liberación
        expect(mockDoc.Save).not.toHaveBeenCalled(); // No debería guardar si falla
        expect(mockDoc.Close).toHaveBeenCalledWith(false);
        expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        expect(mockReleaseObject).not.toHaveBeenCalledWith(mockCellRange); // No se llegó a obtener el rango de la celda
    });

     test('debería manejar un error si la celda especificada para modificar no existe', async () => {
        const filePath = 'C:/test/integration_doc.docx';
        // Intentar modificar una celda que no existe (mockTable.Cell devolverá error o undefined)
        const modifyParams = { filePath, range: 'table:1:cell:9:9', newText: 'Modified Cell Text' };

        // Simular que la celda (9,9) no existe lanzando un error
        mockTable.Cell.mockImplementation((row: number, col: number) => {
            if (row === 1 && col === 1) return mockCell;
            throw new Error('Mock COM Error: Cell not found'); // Simular error COM
        });


        const modifyResult = await modifyText(modifyParams);

        expect(modifyResult.success).toBe(false);
        expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
        expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(filePath, false, false);
        expect(mockTables.Item).toHaveBeenCalledWith(1); // Acceder a tabla 1
        expect(mockTable.Cell).toHaveBeenCalledWith(9, 9); // Intento de acceder a celda (9,9)
        if (!modifyResult.success) {
            expect(modifyResult.error.message).toContain('Failed to get cell (9, 9) from table 1');
        } else {
            fail('Expected modifyResult to be an error response');
        }

        // Verificar cierre y liberación
        expect(mockDoc.Save).not.toHaveBeenCalled();
        expect(mockDoc.Close).toHaveBeenCalledWith(false);
        expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        expect(mockReleaseObject).toHaveBeenCalledWith(mockTable); // La tabla se obtuvo y debe liberarse
        expect(mockReleaseObject).not.toHaveBeenCalledWith(mockCell); // No se obtuvo la celda
    });

});