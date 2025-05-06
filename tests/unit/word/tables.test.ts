import { insertTable, insertTableFromArray, extractTableData, insertTableSchema, insertTableFromArraySchema, extractTableDataSchema } from '../../../src/tools/word/tables.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { z } from 'zod';
import * as fs from 'fs-extra';
import { Packer, Table, TableRow, TableCell, Paragraph, TextRun, Document as DocxDocument } from 'docx'; // Renamed Document to DocxDocument, Added TextRun
import { createErrorResponse as mockCreateErrorResponseUtil, handleToolError as mockHandleToolErrorUtil } from '../../../src/utils/errorHandler';

// Mock dependencies
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
    createErrorResponse: jest.fn((message, code, details) => ({ success: false, error: { code, message, details } })),
}));
jest.mock('../../../src/utils/logger');
jest.mock('fs-extra', () => ({
    pathExists: jest.fn(),
    writeFile: jest.fn(),
}));
jest.mock('docx', () => {
    const originalDocx = jest.requireActual('docx');
    return {
        ...originalDocx,
        Packer: { toBuffer: jest.fn() },
        Table: jest.fn().mockImplementation(props => ({ props })),
        TableRow: jest.fn().mockImplementation(props => ({ props })),
        TableCell: jest.fn().mockImplementation(props => ({ props })),
        Paragraph: jest.fn().mockImplementation(props => ({ props })),
        Document: jest.fn().mockImplementation(props => ({ props })), // Mock Document constructor
    };
});

const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;
const mockLoggerInfo = jest.spyOn(logger, 'info');
const mockLoggerWarn = jest.spyOn(logger, 'warn');
const mockLoggerError = jest.spyOn(logger, 'error');
const mockFsPathExists = fs.pathExists as jest.Mock;
const mockFsWriteFile = fs.writeFile as jest.MockedFunction<typeof fs.writeFile>;
const mockPackerToBuffer = Packer.toBuffer as jest.Mock;
const mockCreateErrorResponse = mockCreateErrorResponseUtil as jest.Mock;
const mockHandleToolError = mockHandleToolErrorUtil as jest.Mock;


describe('word/tables unit tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockSelection: any;
    let mockRange: any;
    let mockParagraphs: any;
    let mockTablesCollection: any;
    let mockCell: any;
    let mockBookmarkRange: any;

    beforeEach(() => {
        jest.clearAllMocks();

        mockCell = {
            Range: { Text: 'Cell Text', _textValue: 'Cell Text', release: jest.fn() },
            RowIndex: 1, ColumnIndex: 1, release: jest.fn()
        };
        mockRange = { Start: 0, End: 10, Collapse: jest.fn(), InsertParagraphAfter: jest.fn(), release: jest.fn() };
        mockSelection = { Range: mockRange, Type: 1, release: jest.fn() };
        mockParagraphs = {
            Count: 5,
            Item: jest.fn(idx => ({
                Range: { ...mockRange, Start: (idx - 1) * 15, End: (idx * 15) - 1, Collapse: jest.fn(), release: jest.fn() },
                release: jest.fn()
            })),
            release: jest.fn(),
        };
        const mockCreatedTableInstance = { // Renamed to avoid conflict
            Index: 1, Style: '', Cell: jest.fn(() => mockCell),
            Rows: { Count: 2, release: jest.fn() }, Columns: { Count: 2, release: jest.fn() },
            ApplyStyleHeadingRows: true, ApplyStyleFirstColumn: false, ApplyStyleBandedRows: false,
            ApplyStyleBandedColumns: false, AllowFormatting: true, release: jest.fn()
        };
        mockTablesCollection = {
            Add: jest.fn().mockReturnValue(mockCreatedTableInstance),
            Count: 1,
            Item: jest.fn(idx => (idx === 1 ? mockCreatedTableInstance : undefined)),
            release: jest.fn(),
        };
        mockBookmarkRange = { ...mockRange, Text: `Bookmark text`, Collapse: jest.fn(), release: jest.fn() };
        mockDoc = {
            Content: { End: 50, Start: 0, Text: 'Doc content', _textValue: 'Doc content', release: jest.fn() },
            Paragraphs: mockParagraphs, Tables: mockTablesCollection,
            Range: jest.fn((start, end) => ({ ...mockRange, Start: start, End: end, release: jest.fn() })),
            Bookmarks: jest.fn(name => ({ Range: mockBookmarkRange, release: jest.fn() })),
            Save: jest.fn(), Close: jest.fn(), release: jest.fn(),
        };
        mockWordApp = {
            Documents: { Open: jest.fn().mockReturnValue(mockDoc) }, // mockReturnValue for sync COM calls
            Selection: mockSelection, Quit: jest.fn(), release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockWordApp);
        mockValidateFilePath.mockImplementation(fp => fp);
        mockFsPathExists.mockResolvedValue(true);
        mockPackerToBuffer.mockResolvedValue(Buffer.from("docx-buffer"));
        mockFsWriteFile.mockImplementation(() => Promise.resolve(undefined));
        mockHandleToolError.mockImplementation((err, code) => ({ success: false, error: { code, message: err.message } }));
    });

    const testModes = [
        { mode: 'COM', useComInterop: true, default: false },
        { mode: 'Library', useComInterop: false, default: true },
    ];

    describe.each(testModes)('insertTable (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParams = { filePath: 'C:/test/document.docx', rows: 3, columns: 2 };
        const getParams = (options?: Partial<z.infer<typeof insertTableSchema>>) => ({ ...baseParams, ...options, useComInterop });

        if (useComInterop) {
            test('COM: should insert a table at the end by default', async () => {
                const params = getParams(); // position defaults to end
                const result = await insertTable(params);
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath);
                expect(mockTablesCollection.Add).toHaveBeenCalled();
                expect(mockDoc.Save).toHaveBeenCalled();
                expect(result.success).toBe(true);
                if (result.success) expect((result.data as { message: string }).message).toContain('COM: Table (3x2) inserted.');
                expect(mockReleaseObject).toHaveBeenCalledWith(mockTablesCollection.Add()); // Table
                expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            });

            test('COM: should insert table at selection', async () => {
                const params = getParams({ position: 'selection' });
                await insertTable(params);
                expect(mockTablesCollection.Add).toHaveBeenCalledWith(mockSelection.Range, params.rows, params.columns, 0, 0);
            });

            test('COM: should insert table at paragraph:N', async () => {
                const params = getParams({ position: 'paragraph:2' });
                await insertTable(params);
                expect(mockParagraphs.Item).toHaveBeenCalledWith(2);
                expect(mockTablesCollection.Add).toHaveBeenCalledWith(mockParagraphs.Item(2).Range, params.rows, params.columns, 0, 0);
            });
            
            test('COM: should apply style if provided', async () => {
                const params = getParams({ style: 'Table Grid' });
                await insertTable(params);
                const addedTable = mockTablesCollection.Add.mock.results[0].value;
                expect(addedTable.Style).toBe('Table Grid');
            });
            
            test('COM: should warn if style application fails', async () => {
                const addedTable = mockTablesCollection.Add.mock.results[0].value;
                addedTable.Style = jest.fn(() => { throw new Error("Style error"); }); // Make Style a setter that throws
                 const params = getParams({ style: 'InvalidStyle' });
                await insertTable(params);
                expect(mockLoggerWarn).toHaveBeenCalledWith(expect.stringContaining("Failed to apply style \"InvalidStyle\""));
            });


        } else { // Library path
            test('Library: should create a new file with a table (default useComInterop:false)', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = { ...baseParams, filePath: 'C:/test/new_doc_lib.docx' }; // No useComInterop
                await insertTable(params);
                expect(mockFsPathExists).toHaveBeenCalledWith(params.filePath);
                expect(mockPackerToBuffer).toHaveBeenCalled();
                expect(mockFsWriteFile).toHaveBeenCalledWith(params.filePath, Buffer.from("docx-buffer"));
                expect(Table).toHaveBeenCalled();
                expect(TableRow).toHaveBeenCalledTimes(params.rows);
                expect(TableCell).toHaveBeenCalledTimes(params.rows * params.columns);
                expect(Paragraph).toHaveBeenCalledTimes(params.rows * params.columns);
                expect(DocxDocument).toHaveBeenCalled();
            });

            test('Library: should fail if file exists', async () => {
                mockFsPathExists.mockResolvedValue(true);
                const params = getParams();
                const result = await insertTable(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('LIB_INSERT_EXISTING_FILE_NOT_SUPPORTED');
            });
            
            test('Library: should log warning if style is provided', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = getParams({ style: 'SomeStyle', filePath: 'C:/test/new_styled_lib.docx' });
                await insertTable(params);
                expect(mockLoggerWarn).toHaveBeenCalledWith(expect.stringContaining("Style parameter 'SomeStyle' received."));
            });
        }
    });

    describe.each(testModes)('insertTableFromArray (mode: $mode, default: $default)', ({ useComInterop }) => {
        const data = [['A1', 'B1'], ['A2', 'B2']];
        const baseParamsArray = { filePath: 'C:/test/document_array.docx', data };
        const getParams = (options?: Partial<z.infer<typeof insertTableFromArraySchema>>) => ({ ...baseParamsArray, ...options, useComInterop });


        if (useComInterop) {
            test('COM: should insert table from array and populate cells', async () => {
                const params = getParams();
                const result = await insertTableFromArray(params);
                expect(mockTablesCollection.Add).toHaveBeenCalled();
                const addedTable = mockTablesCollection.Add.mock.results[0].value;
                expect(addedTable.Cell).toHaveBeenCalledTimes(data.length * data[0].length);
                expect(addedTable.Cell(1,1).Range.Text).toBe('A1');
                expect(addedTable.Cell(2,2).Range.Text).toBe('B2');
                expect(result.success).toBe(true);
            });
            
            test('COM: should apply styleName and styleOptions', async () => {
                const params = getParams({
                    styleName: 'Grid Table 1 Light',
                    styleOptions: {
                        headerRow: true,
                        bandedRows: true,
                        firstColumn: false, // Added default
                        bandedColumns: false // Added default
                    }
                });
                await insertTableFromArray(params);
                const addedTable = mockTablesCollection.Add.mock.results[0].value;
                expect(addedTable.Style).toBe('Grid Table 1 Light');
                expect(addedTable.ApplyStyleHeadingRows).toBe(true);
                expect(addedTable.ApplyStyleBandedRows).toBe(true);
            });
            
            test('COM: should insert at bookmark', async () => {
                const params = getParams({ position: 'bookmark:TestBookmark' });
                await insertTableFromArray(params);
                expect(mockDoc.Bookmarks).toHaveBeenCalledWith('TestBookmark');
                expect(mockTablesCollection.Add).toHaveBeenCalledWith(mockBookmarkRange, data.length, data[0].length, 0, 0);
            });

        } else { // Library path
            test('Library: should create new file with table from array (default useComInterop:false)', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = { ...baseParamsArray, filePath: 'C:/test/new_doc_array_lib.docx' }; // No useComInterop
                await insertTableFromArray(params);
                expect(mockPackerToBuffer).toHaveBeenCalled();
                expect(mockFsWriteFile).toHaveBeenCalled();
                expect(Table).toHaveBeenCalled();
                expect(TableRow).toHaveBeenCalledTimes(data.length);
                expect(TableCell).toHaveBeenCalledTimes(data.length * data[0].length);
                expect(TextRun).toHaveBeenCalledWith('A1'); // Check TextRun usage
            });

            test('Library: should fail if file exists for insertTableFromArray', async () => {
                mockFsPathExists.mockResolvedValue(true);
                const params = getParams();
                const result = await insertTableFromArray(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('LIB_INSERT_ARRAY_EXISTING_FILE_NOT_SUPPORTED');
            });
            
            test('Library: should log warnings for styleName and styleOptions', async () => {
                 mockFsPathExists.mockResolvedValue(false);
                const params = getParams({
                    styleName: 'SomeStyle',
                    styleOptions: {
                        headerRow: true,
                        firstColumn: false, // Added default
                        bandedRows: false, // Added default
                        bandedColumns: false // Added default
                    },
                    filePath: 'C:/test/new_styled_array_lib.docx'
                });
                await insertTableFromArray(params);
                expect(mockLoggerWarn).toHaveBeenCalledWith(expect.stringContaining("StyleName 'SomeStyle' is COM-specific"));
                expect(mockLoggerWarn).toHaveBeenCalledWith(expect.stringContaining("Setting 'headerRow' directly on TableRow properties is complex"));
            });
        }
    });

    describe.each(testModes)('extractTableData (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParamsExtract = { filePath: 'C:/test/doc_with_table.docx', tableIndex: 1 };
        const getParams = (options?: Partial<z.infer<typeof extractTableDataSchema>>) => ({ ...baseParamsExtract, ...options, useComInterop });


        if (useComInterop) {
            test('COM: should extract table data correctly', async () => {
                const params = getParams();
                // Setup mockCell to return different text for different cells for a more robust test
                mockTablesCollection.Item(1).Cell.mockImplementation((r: number, c: number) => ({
                    Range: { Text: `R${r}C${c}\r\n`, release: jest.fn() }, // Add newlines to test trimming
                    RowIndex: r, ColumnIndex: c, release: jest.fn()
                }));

                const result = await extractTableData(params);
                expect(mockDoc.Tables.Item).toHaveBeenCalledWith(1);
                const tableToExtract = mockDoc.Tables.Item(1);
                expect(tableToExtract.Cell).toHaveBeenCalledTimes(tableToExtract.Rows.Count * tableToExtract.Columns.Count);
                expect(result.success).toBe(true);
                if (result.success) expect(result.data).toEqual([['R1C1', 'R1C2'], ['R2C1', 'R2C2']]);
            });

            test('COM: should handle table index out of bounds', async () => {
                mockTablesCollection.Count = 0;
                const params = getParams();
                const result = await extractTableData(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.message).toContain("Table index 1 out of bounds. Document has 0 tables.");
            });
            
            test('COM: should handle error if doc.Tables.Item fails', async () => {
                mockDoc.Tables.Item.mockImplementation(() => { throw new Error("Cannot access table"); });
                const params = getParams();
                const result = await extractTableData(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('WORD_TABLE_EXTRACT_DATA_FAILED_COM');
            });

        } else { // Library path
            test('Library: should return NOT_IMPLEMENTED_LIB_TABLE_EXTRACTION (default useComInterop:false)', async () => {
                const params = { ...baseParamsExtract }; // No useComInterop
                const result = await extractTableData(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_TABLE_EXTRACTION');
            });

            test('Library: should handle file not found for extractTableData', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = getParams();
                const result = await extractTableData(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('FILE_NOT_FOUND_LIB');
            });
        }
    });
});