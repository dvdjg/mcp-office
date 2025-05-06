import { getText, insertText, modifyText, deleteText } from '../../../src/tools/word/text.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { z } from 'zod';
import * as fs from 'fs-extra'; // For mocking fs.pathExists
import mammoth from 'mammoth'; // For mocking mammoth
import * as path from 'path'; // Import path
import { applyMarkdownFormattingToWord as mockApplyMarkdownUtil } from '../../../src/utils/markdownToOffice';
import { resolveNaturalLanguageRange as mockResolveNaturalLanguageRangeUtil } from '../../../src/utils/wordRangeResolver';


// Mock dependencies
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
    createErrorResponse: jest.fn((message, code, details) => ({ success: false, error: { code, message, details } })),
}));
jest.mock('../../../src/utils/logger');
jest.mock('fs-extra');
jest.mock('mammoth');
jest.mock('../../../src/utils/markdownToOffice');
jest.mock('../../../src/utils/wordRangeResolver');

const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;
const mockLoggerInfo = jest.spyOn(logger, 'info');
const mockLoggerWarn = jest.spyOn(logger, 'warn');
const mockLoggerError = jest.spyOn(logger, 'error');
const mockFsPathExists = fs.pathExists as jest.Mock;
const mockMammothExtractRawText = mammoth.extractRawText as jest.Mock;
const mockCreateErrorResponse = require('../../../src/utils/errorHandler').createErrorResponse as jest.Mock;
const mockApplyMarkdownFormattingToWord = mockApplyMarkdownUtil as jest.Mock;
const mockResolveNaturalLanguageRange = mockResolveNaturalLanguageRangeUtil as jest.Mock;


describe('word/text integration tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockSelection: any;
    let mockRange: any;
    let mockParagraphs: any;
    let mockContentRange: any;

    beforeEach(() => {
        jest.clearAllMocks();

        mockRange = {
            Start: 0,
            End: 22,
            Delete: jest.fn(),
            Collapse: jest.fn(),
            _textValue: 'Sample text from range', // Initial value
            // Mock setter and getter for Text
            set Text(value: string) { (this as any)._textValue = value; },
            get Text() { return (this as any)._textValue; },
            release: jest.fn(),
        };
        mockSelection = {
            Range: mockRange,
            Type: 2, // wdSelectionIP
            release: jest.fn(),
        };
        mockParagraphs = {
            Count: 5,
            Item: jest.fn((index: number) => {
                if (index > 0 && index <= mockParagraphs.Count) {
                    return { Range: { ...mockRange, Text: `Paragraph ${index} text`, _textValue: `Paragraph ${index} text`, release: jest.fn() }, release: jest.fn() };
                }
                // Simulate COM error for out-of-bounds access
                throw new Error(`COM Error: The requested member of the collection does not exist. Index: ${index}`);
            }),
            release: jest.fn(),
        };
        mockContentRange = { ...mockRange, Text: 'Document content text', _textValue: 'Document content text', Delete: jest.fn(), release: jest.fn() };
        mockDoc = {
            Content: mockContentRange,
            Paragraphs: mockParagraphs,
            Range: jest.fn((start, end) => ({ ...mockRange, Start: start, End: end, _textValue: '', Text: '', release: jest.fn() })),
            Save: jest.fn(),
            Close: jest.fn(),
            release: jest.fn(),
            // Documents: { Add: jest.fn().mockReturnThis(), SaveAs2: jest.fn() }, // This was incorrect for mockDoc
        };
        mockWordApp = {
            Documents: {
                Open: jest.fn().mockReturnValue(mockDoc), // mockReturnValue for sync COM calls
                Add: jest.fn().mockReturnValue(mockDoc),
            },
            Selection: mockSelection,
            Quit: jest.fn(),
            release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockWordApp);
        mockValidateFilePath.mockImplementation(filePath => path.resolve(filePath)); // Resolve path
        mockFsPathExists.mockResolvedValue(true);
        mockMammothExtractRawText.mockResolvedValue({ value: 'Mammoth extracted text', messages: [] });
        mockResolveNaturalLanguageRange.mockReturnValue(null); // Default behavior
        mockApplyMarkdownFormattingToWord.mockResolvedValue(undefined); // Default behavior
    });

    const testModes = [
        { mode: 'COM', useComInterop: true, default: false },
        { mode: 'Library', useComInterop: false, default: true },
    ];

    describe.each(testModes)('getText (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParams = { filePath: 'C:/test/document.docx' };
        const getParams = (range: string) => ({ ...baseParams, range, useComInterop });

        if (useComInterop) {
            test('COM: should get text from selection', async () => {
                const params = getParams('selection');
                const result = await getText(params);
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(path.resolve(params.filePath), false, true, false, "", "", false, "", "", 0, false);
                expect(result.success).toBe(true);
                if (result.success) expect(result.data).toBe('Sample text from range');
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp.Selection.Range);
            });

            test('COM: should use resolveNaturalLanguageRange and fallback to getRangeFromSpecifier', async () => {
                mockResolveNaturalLanguageRange.mockReturnValue(null); // Simulate natural language not found
                const params = getParams('paragraph:3');
                const result = await getText(params);
                expect(mockResolveNaturalLanguageRange).toHaveBeenCalledWith(mockDoc, params.range, mockWordApp);
                expect(mockParagraphs.Item).toHaveBeenCalledWith(3);
                expect(result.success).toBe(true);
                if (result.success) expect(result.data).toBe('Paragraph 3 text');
            });
        } else { // Library path
            test('Library: should get text from document using Mammoth (default useComInterop:false)', async () => {
                const params = { ...baseParams, range: 'document' }; // No useComInterop, defaults to false
                const result = await getText(params);
                expect(mockFsPathExists).toHaveBeenCalledWith(path.resolve(params.filePath));
                expect(mockMammothExtractRawText).toHaveBeenCalledWith({ path: path.resolve(params.filePath) });
                expect(result.success).toBe(true);
                if (result.success) expect(result.data).toBe('Mammoth extracted text');
            });

            test('Library: should return error for "selection" range', async () => {
                const params = getParams('selection');
                const result = await getText(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('RANGE_REQUIRES_COM_LIB');
            });
        }
    });

    describe.each(testModes)('insertText (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParams = { filePath: 'C:/test/document.docx', text: 'New inserted text' };
        const getParams = (position: string, format?: 'plaintext' | 'markdown') => ({ ...baseParams, position, format, useComInterop });


        if (useComInterop) {
            test('COM: should insert markdown text at start by default', async () => {
                const params = getParams('start'); // format defaults to markdown
                mockFsPathExists.mockResolvedValue(true);
                const result = await insertText(params);
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(path.resolve(params.filePath), false, false);
                expect(mockDoc.Range).toHaveBeenCalledWith(0, 0);
                const insertionRange = mockDoc.Range.mock.results[0].value;
                expect(mockApplyMarkdownFormattingToWord).toHaveBeenCalledWith(insertionRange, params.text, mockWordApp, mockDoc);
                expect(mockDoc.Save).toHaveBeenCalled();
                expect(result.success).toBe(true);
            });

            test('COM: should insert plaintext at paragraph:2:end', async () => {
                const params = getParams('paragraph:2:end', 'plaintext');
                const mockParaRangeInstance = { Start: 50, End: 60, release: jest.fn() };
                mockParagraphs.Item.mockImplementation(idx => idx === 2 ? { Range: mockParaRangeInstance, release: jest.fn() } : undefined);
                
                const result = await insertText(params);
                expect(mockParagraphs.Item).toHaveBeenCalledWith(2);
                expect(mockDoc.Range).toHaveBeenCalledWith(mockParaRangeInstance.End -1, mockParaRangeInstance.End -1);
                const insertionRange = mockDoc.Range.mock.results[0].value;
                expect(insertionRange.Text).toBe(params.text);
                expect(mockApplyMarkdownFormattingToWord).not.toHaveBeenCalled();
                expect(result.success).toBe(true);
            });

            test('COM: should create file and insert text if file does not exist', async () => {
                const params = getParams('start', 'plaintext');
                params.filePath = 'C:/test/new_doc_for_insert.docx';
                mockFsPathExists.mockResolvedValue(false);

                const result = await insertText(params);
                expect(mockWordApp.Documents.Add).toHaveBeenCalled();
                expect(mockDoc.Range).toHaveBeenCalledWith(0,0);
                const insertionRange = mockDoc.Range.mock.results[0].value;
                expect(insertionRange.Text).toBe(params.text);
                expect(mockDoc.SaveAs2).toHaveBeenCalledWith(path.resolve(params.filePath), 16);
                expect(result.success).toBe(true);
            });

        } else { // Library path
            test('Library: should return NOT_IMPLEMENTED_LIB (default useComInterop:false)', async () => {
                const params = { ...baseParams, position: 'start' }; // No useComInterop
                const result = await insertText(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB');
            });
        }
    });

    describe.each(testModes)('modifyText (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParams = { filePath: 'C:/test/document.docx', newText: 'Text after modification' };
        const getParams = (range: string) => ({ ...baseParams, range, useComInterop });

        if (useComInterop) {
            test('COM: should modify text for "document" range', async () => {
                const params = getParams('document');
                const result = await modifyText(params);
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(path.resolve(params.filePath), false, false);
                expect(mockDoc.Content.Text).toBe(params.newText);
                expect(mockDoc.Save).toHaveBeenCalled();
                expect(result.success).toBe(true);
            });
        } else { // Library path
            test('Library: should return NOT_IMPLEMENTED_LIB (default useComInterop:false)', async () => {
                const params = { ...baseParams, range: 'document' }; // No useComInterop
                const result = await modifyText(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB');
            });
        }
    });

    describe.each(testModes)('deleteText (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParams = { filePath: 'C:/test/document.docx' };
        const getParams = (range: string) => ({ ...baseParams, range, useComInterop });

        if (useComInterop) {
            test('COM: should delete text for "paragraph:1" range', async () => {
                const params = getParams('paragraph:1');
                const result = await deleteText(params);
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(path.resolve(params.filePath), false, false);
                expect(mockParagraphs.Item).toHaveBeenCalledWith(1);
                const rangeToDelete = mockParagraphs.Item(1).Range;
                expect(rangeToDelete.Delete).toHaveBeenCalled();
                expect(mockDoc.Save).toHaveBeenCalled();
                expect(result.success).toBe(true);
            });
        } else { // Library path
            test('Library: should return NOT_IMPLEMENTED_LIB (default useComInterop:false)', async () => {
                const params = { ...baseParams, range: 'paragraph:1' }; // No useComInterop
                const result = await deleteText(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB');
            });
        }
    });

    // General error handling tests (can apply to any function, using getText as example)
    test('getText COM: should handle file open failure', async () => {
        mockWordApp.Documents.Open.mockReturnValue(null); // Simulate open returning null
        const params = { filePath: 'C:/test/fail_open.docx', range: 'document', useComInterop: true };
        const result = await getText(params);
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('FILE_OPEN_FAILED_COM');
        }
    });

    test('getText Library: should handle fs.pathExists failure', async () => {
        mockFsPathExists.mockRejectedValue(new Error("FS error"));
        const params = { filePath: 'C:/test/fs_error.docx', range: 'document', useComInterop: false };
        const result = await getText(params);
        expect(result.success).toBe(false);
        if (!result.success) {
            // The tool's getText for library path doesn't directly catch fs.pathExists errors before calling it.
            // It would likely manifest as a MAMMOTH_GET_TEXT_FAILED or similar if mammoth then fails.
            // For a more direct test, we'd need to adjust the tool or mock mammoth to show path issue.
            // For now, this tests that if pathExists is false, it returns FILE_NOT_FOUND.
            mockFsPathExists.mockResolvedValue(false);
            const result2 = await getText(params);
            expect(result2.success).toBe(false);
            if(!result2.success) expect(result2.error.code).toBe('FILE_NOT_FOUND');
        }
    });
});