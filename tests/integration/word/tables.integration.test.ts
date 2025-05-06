import { insertTable, insertTableFromArray, extractTableData, insertTableSchema, insertTableFromArraySchema, extractTableDataSchema } from '../../../src/tools/word/tables.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { z } from 'zod';
import * as fs from 'fs-extra';
import { Packer, Table, TableRow, TableCell, Paragraph, TextRun, Document as DocxDocument } from 'docx'; // For mocking Packer and other docx elements

// Mock dependencies
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
    createErrorResponse: jest.fn((message, code, details) => ({ success: false, error: { code, message, details } })),
}));
jest.mock('../../../src/utils/logger');
jest.mock('fs-extra', () => {
    const originalFsExtra = jest.requireActual('fs-extra');
    return {
        ...originalFsExtra, // Preserve other fs-extra exports if any are used indirectly
        pathExists: jest.fn(),
        writeFile: jest.fn(), // This is now correctly a Jest mock function
    };
});
jest.mock('docx', () => {
    const originalDocx = jest.requireActual('docx');
    return {
        ...originalDocx,
        Packer: { toBuffer: jest.fn() },
        Table: jest.fn().mockImplementation(props => ({ props, rows: props.rows || [] })),
        TableRow: jest.fn().mockImplementation(props => ({ props, cells: props.children || [] })),
        TableCell: jest.fn().mockImplementation(props => ({ props, children: props.children || [] })),
        Paragraph: jest.fn().mockImplementation(props => ({ props, children: props.children || [] })),
        TextRun: jest.fn().mockImplementation(props => ({ props, text: props.text || (typeof props === 'string' ? props : '') })),
        Document: jest.fn().mockImplementation(props => ({ props, sections: props.sections || [] })),
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
const mockCreateErrorResponse = require('../../../src/utils/errorHandler').createErrorResponse as jest.Mock;
const mockHandleToolError = require('../../../src/utils/errorHandler').handleToolError as jest.Mock;


describe('word/tables integration tests', () => {
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
            RowIndex: 1,
            ColumnIndex: 1,
            release: jest.fn(),
        };

        mockRange = {
            Start: 0, End: 10, Collapse: jest.fn(), InsertParagraphAfter: jest.fn(),
            _textValue: '', set Text(value: string) { (this as any)._textValue = value; }, get Text() { return (this as any)._textValue; },
            release: jest.fn(),
        };
        mockSelection = { Range: mockRange, Type: 1, release: jest.fn() };
        mockParagraphs = {
            Count: 5,
            Item: jest.fn(idx => ({
                Range: { ...mockRange, Start: (idx - 1) * 15, End: (idx * 15) - 1, Collapse: jest.fn(), release: jest.fn() },
                release: jest.fn()
            })),
            release: jest.fn(),
        };
        
        const mockCreatedTableInstance = {
            Index: 1, Style: '', Cell: jest.fn(() => mockCell),
            Rows: { Count: 2, Item: jest.fn().mockReturnValue({ Cells: { Count: 2, Item: jest.fn().mockReturnValue(mockCell), release: jest.fn()}, release: jest.fn()}), release: jest.fn() },
            Columns: { Count: 2, Item: jest.fn().mockReturnValue({ Cells: { Count: 2, Item: jest.fn().mockReturnValue(mockCell), release: jest.fn()}, release: jest.fn()}), release: jest.fn() },
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
            Bookmarks: jest.fn((name: string) => ({ Range: mockBookmarkRange, release: jest.fn() })),
            Save: jest.fn(), Close: jest.fn(), release: jest.fn(),
        };
        mockWordApp = {
            Documents: { Open: jest.fn().mockReturnValue(mockDoc) },
            Selection: mockSelection, Quit: jest.fn(), release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockWordApp);
        mockValidateFilePath.mockImplementation(fp => fp);
        mockFsPathExists.mockResolvedValue(true);
        mockPackerToBuffer.mockResolvedValue(Buffer.from("docx-buffer"));
        mockFsWriteFile.mockImplementation(() => Promise.resolve(undefined));
        mockCreateErrorResponse.mockImplementation((message, code) => ({ success: false, error: { code, message } }));
        mockHandleToolError.mockImplementation((err, code) => ({ success: false, error: { code, message: err.message } }));

    });

    const testModes = [
        { mode: 'COM', useComInterop: true, default: false },
        { mode: 'Library', useComInterop: false, default: true },
    ];
    
    const effectiveTestModes = testModes.map(tm => ({
        ...tm,
        effectiveUseComInterop: tm.useComInterop === undefined ? false : tm.useComInterop,
    }));

    describe.each(effectiveTestModes)('insertTable (mode: $mode, useComInterop: $effectiveUseComInterop)', ({ useComInterop, effectiveUseComInterop }) => {
        const baseParams = { filePath: 'C:/test/document.docx', rows: 3, columns: 2 };
        const getParams = (options?: Partial<z.infer<typeof insertTableSchema>>) => ({ ...baseParams, ...options, useComInterop });

        if (effectiveUseComInterop) {
            test('COM: should insert a table at the end by default and apply style', async () => {
                const params = getParams({ style: "Table Grid" });
                const result = await insertTable(params);
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath);
                expect(mockTablesCollection.Add).toHaveBeenCalled();
                const addedTable = mockTablesCollection.Add.mock.results[0].value;
                expect(addedTable.Style).toBe("Table Grid");
                expect(mockDoc.Save).toHaveBeenCalled();
                expect(result.success).toBe(true);
            });
            // Add more COM path tests for insertTable (positions, errors)
        } else {
            test('Library: should create a new file with a table', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = getParams({ filePath: 'C:/test/new_doc_lib.docx' });
                const result = await insertTable(params);
                expect(mockFsPathExists).toHaveBeenCalledWith(params.filePath);
                expect(mockPackerToBuffer).toHaveBeenCalled();
                expect(mockFsWriteFile).toHaveBeenCalledWith(params.filePath, Buffer.from("docx-buffer"));
                expect(result.success).toBe(true);
                expect(Table).toHaveBeenCalled();
                expect(TableRow).toHaveBeenCalledTimes(params.rows);
            });

            test('Library: should fail if file exists', async () => {
                mockFsPathExists.mockResolvedValue(true);
                const params = getParams();
                const result = await insertTable(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('LIB_INSERT_EXISTING_FILE_NOT_SUPPORTED');
            });
        }
    });

    describe.each(effectiveTestModes)('insertTableFromArray (mode: $mode, useComInterop: $effectiveUseComInterop)', ({ useComInterop, effectiveUseComInterop }) => {
        const data = [['A1', 'B1'], ['A2', 'B2']];
        const baseParamsArray = { filePath: 'C:/test/document_array.docx', data };
        const getParams = (options?: Partial<z.infer<typeof insertTableFromArraySchema>>) => ({ ...baseParamsArray, ...options, useComInterop });


        if (effectiveUseComInterop) {
            test('COM: should insert table from array and apply style options', async () => {
                const params = getParams({ styleName: "Cool Style", styleOptions: { headerRow: true, bandedColumns: true, firstColumn: false, bandedRows: false }});
                const result = await insertTableFromArray(params);
                expect(mockTablesCollection.Add).toHaveBeenCalled();
                const addedTable = mockTablesCollection.Add.mock.results[0].value;
                expect(addedTable.Cell).toHaveBeenCalledTimes(data.length * data[0].length);
                expect(addedTable.Style).toBe("Cool Style");
                expect(addedTable.ApplyStyleHeadingRows).toBe(true);
                expect(addedTable.ApplyStyleBandedColumns).toBe(true);
                expect(mockDoc.Save).toHaveBeenCalled();
                expect(result.success).toBe(true);
            });
             test('COM: should insert at bookmark', async () => {
                const params = getParams({ position: 'bookmark:TestBookmark' });
                await insertTableFromArray(params);
                expect(mockDoc.Bookmarks).toHaveBeenCalledWith('TestBookmark');
                expect(mockTablesCollection.Add).toHaveBeenCalledWith(mockBookmarkRange, data.length, data[0].length, 0, 0);
            });
        } else {
            test('Library: should create new file with table from array and use TextRun', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = getParams({ filePath: 'C:/test/new_doc_array_lib.docx' });
                const result = await insertTableFromArray(params);
                expect(mockPackerToBuffer).toHaveBeenCalled();
                expect(mockFsWriteFile).toHaveBeenCalled();
                expect(result.success).toBe(true);
                expect(TextRun).toHaveBeenCalledWith('A1');
                expect(TextRun).toHaveBeenCalledWith('B2');
            });

            test('Library: should fail if file exists for insertTableFromArray', async () => {
                mockFsPathExists.mockResolvedValue(true);
                const params = getParams();
                const result = await insertTableFromArray(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('LIB_INSERT_ARRAY_EXISTING_FILE_NOT_SUPPORTED');
            });
        }
    });

    describe.each(effectiveTestModes)('extractTableData (mode: $mode, useComInterop: $effectiveUseComInterop)', ({ useComInterop, effectiveUseComInterop }) => {
        const baseParamsExtract = { filePath: 'C:/test/doc_with_table.docx', tableIndex: 1 };
        const getParams = (options?: Partial<z.infer<typeof extractTableDataSchema>>) => ({ ...baseParamsExtract, ...options, useComInterop });

        if (effectiveUseComInterop) {
            test('COM: should extract table data, trimming cell text', async () => {
                const params = getParams();
                mockTablesCollection.Item(1).Cell.mockImplementation((r: number, c: number) => ({
                    Range: { Text: ` R${r}C${c} \r\n `, release: jest.fn() }, // With spaces and newlines
                    RowIndex: r, ColumnIndex: c, release: jest.fn()
                }));
                const result = await extractTableData(params);
                expect(mockDoc.Tables.Item).toHaveBeenCalledWith(1);
                expect(result.success).toBe(true);
                if(result.success) expect(result.data).toEqual([['R1C1', 'R1C2'], ['R2C1', 'R2C2']]);
            });
            
            test('COM: should handle table index out of bounds', async () => {
                mockTablesCollection.Count = 0;
                const params = getParams();
                const result = await extractTableData(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.message).toContain("Table index 1 out of bounds. Document has 0 tables.");
            });

        } else {
            test('Library: should return NOT_IMPLEMENTED_LIB_TABLE_EXTRACTION', async () => {
                const params = getParams();
                const result = await extractTableData(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_TABLE_EXTRACTION');
            });
            
            test('Library: should return FILE_NOT_FOUND_LIB if file does not exist', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = getParams();
                const result = await extractTableData(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('FILE_NOT_FOUND_LIB');
            });
        }
    });
});