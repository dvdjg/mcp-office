import { insertTable, insertTableSchema } from '../../../src/tools/word/tables.tool';
import { modifyText, modifySchema as modifyTextSchema } from '../../../src/tools/word/text.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import * as fs from 'fs-extra';
import { Packer, Table, TableRow, TableCell, Paragraph, TextRun, Document as DocxDocument } from 'docx';
import { z } from 'zod';

// Mock dependencies
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
    createErrorResponse: jest.fn((message, code) => ({ success: false, error: { code, message } })),
}));
jest.mock('../../../src/utils/logger');
jest.mock('fs-extra', () => ({
    pathExists: jest.fn(),
    writeFile: jest.fn(),
    readFile: jest.fn(),
}));
jest.mock('docx', () => {
    const originalDocx = jest.requireActual('docx');
    return {
        ...originalDocx,
        Packer: {
            toBuffer: jest.fn(),
        },
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
const mockLoggerError = jest.spyOn(logger, 'error');
const mockFsPathExists = fs.pathExists as jest.Mock;
const mockFsWriteFile = fs.writeFile as jest.MockedFunction<typeof fs.writeFile>;
const mockPackerToBuffer = Packer.toBuffer as jest.Mock;
const mockCreateErrorResponse = require('../../../src/utils/errorHandler').createErrorResponse as jest.Mock;


describe('word/tables and word/text integration tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockSelection: any;
    let mockRange: any;
    let mockTablesCollection: any;
    let mockTableInstance: any;
    let mockCellInstance: any;
    let mockCellRangeInstance: any;
    // let mockOfficeAppInstance: any; // Replaced by direct mockWordApp for COM

    beforeEach(() => {
        jest.clearAllMocks();

        mockCellRangeInstance = {
            _text: 'Initial Cell Text',
            set Text(value: string) { (this as any)._text = value; },
            get Text() { return (this as any)._text; },
            Delete: jest.fn(),
            Collapse: jest.fn(),
            release: jest.fn(),
        };

        mockCellInstance = {
            Range: mockCellRangeInstance,
            release: jest.fn(),
        };

        mockTableInstance = {
            Index: 1,
            Style: '',
            Cell: jest.fn((row: number, col: number) => {
                if (row === 1 && col === 1) return mockCellInstance;
                return { Range: { Text: `Cell ${row},${col}`, Delete: jest.fn(), Collapse: jest.fn(), release: jest.fn() }, release: jest.fn() };
            }),
            release: jest.fn(),
        };

        mockTablesCollection = {
            Add: jest.fn().mockReturnValue(mockTableInstance),
            Item: jest.fn((index: number) => (index === 1 ? mockTableInstance : undefined)),
            Count: 1,
            release: jest.fn(),
        };

        mockRange = {
            Start: 0, End: 10, Collapse: jest.fn(), InsertParagraphAfter: jest.fn(), release: jest.fn(),
        };

        mockSelection = { Range: mockRange, Type: 1, release: jest.fn() };

        mockDoc = {
            Content: { End: 50, release: jest.fn() },
            Tables: mockTablesCollection,
            Range: jest.fn().mockReturnValue(mockRange),
            Save: jest.fn(),
            Close: jest.fn(),
            release: jest.fn(),
        };

        mockWordApp = {
            Documents: {
                Open: jest.fn().mockReturnValue(mockDoc), // For COM path opening existing
                Add: jest.fn().mockReturnValue(mockDoc),    // For COM path creating new (if insertTable did that)
            },
            Selection: mockSelection,
            release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockWordApp); // Simplified: getOfficeApplication returns the app
        mockValidateFilePath.mockImplementation(fp => fp);
        mockFsPathExists.mockResolvedValue(true); // Default to file existing for COM modify
        mockPackerToBuffer.mockResolvedValue(Buffer.from("docx-buffer"));
        mockFsWriteFile.mockImplementation(() => Promise.resolve());
        mockCreateErrorResponse.mockImplementation((message, code) => ({ success: false, error: { code, message } }));
    });

    const testModes = [
        { mode: 'COM', useComInterop: true },
        { mode: 'Library', useComInterop: false },
    ];

    describe.each(testModes)('Sequential table insert and text modify (mode: $mode)', ({ useComInterop }) => {
        const filePath = 'C:/test/integration_doc.docx';
        const insertParams: z.infer<typeof insertTableSchema> = { filePath, rows: 2, columns: 2, position: 'end', useComInterop };
        const modifyParams: z.infer<typeof modifyTextSchema> = { filePath, range: 'table:1:cell:1:1', newText: 'Modified Cell Text', useComInterop };

        if (useComInterop) {
            test('COM: should insert a table and then modify text in a cell', async () => {
                // --- Step 1: Insert Table (COM) ---
                const insertResult = await insertTable(insertParams);
                expect(insertResult.success).toBe(true);
                expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(filePath, false, false); // Assuming insertTable opens for read/write
                expect(mockTablesCollection.Add).toHaveBeenCalledWith(mockRange, insertParams.rows, insertParams.columns, 0, 0);
                expect(mockDoc.Save).toHaveBeenCalledTimes(1);
                expect(mockDoc.Close).toHaveBeenCalledWith(false);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockTableInstance);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockRange);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);


                // Reset mocks for the second COM call
                jest.clearAllMocks();
                mockGetOfficeApplication.mockResolvedValue(mockWordApp); // Re-mock for next call
                mockWordApp.Documents.Open.mockReturnValue(mockDoc); // Ensure Open returns the doc again
                mockValidateFilePath.mockImplementation(fp => fp);


                // --- Step 2: Modify Text (COM) ---
                const modifyResult = await modifyText(modifyParams);
                expect(modifyResult.success).toBe(true);
                expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application'); // Called again
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(filePath, false, false); // Opened again for modify
                expect(mockTablesCollection.Item).toHaveBeenCalledWith(1);
                expect(mockTableInstance.Cell).toHaveBeenCalledWith(1, 1);
                expect(mockCellRangeInstance.Text).toBe(modifyParams.newText);
                expect(mockDoc.Save).toHaveBeenCalledTimes(1);
                expect(mockDoc.Close).toHaveBeenCalledWith(false);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockCellRangeInstance);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockCellInstance);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockTableInstance);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            });
        } else { // Library Path
            test('Library: should create a new file with a table, then attempt to modify (expecting NOT_IMPLEMENTED_LIB for modify)', async () => {
                // --- Step 1: Insert Table (Library) ---
                mockFsPathExists.mockResolvedValue(false); // Simulate file not existing for creation
                const insertResultLib = await insertTable(insertParams);

                expect(insertResultLib.success).toBe(true);
                expect(mockFsPathExists).toHaveBeenCalledWith(filePath);
                expect(mockPackerToBuffer).toHaveBeenCalled();
                expect(mockFsWriteFile).toHaveBeenCalledWith(filePath, Buffer.from("docx-buffer"));
                expect(Table).toHaveBeenCalled(); // docx Table constructor
                expect(TableRow).toHaveBeenCalledTimes(insertParams.rows);


                // --- Step 2: Modify Text (Library) ---
                mockFsPathExists.mockResolvedValue(true); // File now exists
                const modifyResultLib = await modifyText(modifyParams);

                expect(modifyResultLib.success).toBe(false);
                if (!modifyResultLib.success) {
                    expect(modifyResultLib.error.code).toBe('NOT_IMPLEMENTED_LIB');
                }
                expect(mockLoggerInfo).toHaveBeenCalledWith(expect.stringContaining(`Executing word/text/modify for file: ${filePath}, range: ${modifyParams.range}, useComInterop: false`));
                expect(mockCreateErrorResponse).toHaveBeenCalledWith(
                    'Library path for modifyText is not yet implemented. Use COM Interop (useComInterop: true) for this functionality.',
                    'NOT_IMPLEMENTED_LIB'
                );
                // Ensure no COM objects were touched for library path
                expect(mockGetOfficeApplication).not.toHaveBeenCalled();
            });
        }
    });

    // Tests for error handling in COM path (can remain largely similar, just ensure useComInterop: true)
    describe('Error handling for modifyText (COM path)', () => {
        const filePath = 'C:/test/integration_doc_errors.docx';
        const commonModifyParams = { filePath, newText: 'Modified Cell Text', useComInterop: true };

        test('COM: should handle error if specified table for modify does not exist', async () => {
            const modifyParamsTableError = { ...commonModifyParams, range: 'table:2:cell:1:1' };
            mockTablesCollection.Item.mockImplementation((index: number) => (index === 1 ? mockTableInstance : undefined));

            const modifyResult = await modifyText(modifyParamsTableError);
            expect(modifyResult.success).toBe(false);
            if (!modifyResult.success) {
                expect(modifyResult.error.message).toContain('Failed to get range from specifier'); // Updated to reflect getRangeFromSpecifier's error
            }
            expect(mockDoc.Save).not.toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
        });

        test('COM: should handle error if specified cell for modify does not exist', async () => {
            const modifyParamsCellError = { ...commonModifyParams, range: 'table:1:cell:9:9' };
            mockTableInstance.Cell.mockImplementation((row: number, col: number) => {
                if (row === 1 && col === 1) return mockCellInstance;
                throw new Error('Mock COM Error: Cell not found');
            });

            const modifyResult = await modifyText(modifyParamsCellError);
            expect(modifyResult.success).toBe(false);
            if (!modifyResult.success) {
                expect(modifyResult.error.message).toContain('Failed to get range from specifier'); // Updated to reflect getRangeFromSpecifier's error
            }
            expect(mockDoc.Save).not.toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockTableInstance); // Table was obtained
            expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
        });
    });
});