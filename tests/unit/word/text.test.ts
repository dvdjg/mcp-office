import { getText, insertText, modifyText, deleteText, getRangeFromSpecifier } from '../../../src/tools/word/text.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { z } from 'zod';
import * as fs from 'fs-extra';
import mammoth from 'mammoth';
import * as path from 'path'; // Import path for path.resolve
import { createErrorResponse as mockCreateErrorResponseUtil } from '../../../src/utils/errorHandler';
import { applyMarkdownFormattingToWord } from '../../../src/utils/markdownToOffice'; // Added this import


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
jest.mock('../../../src/utils/wordRangeResolver', () => ({
    resolveNaturalLanguageRange: jest.fn(),
}));
jest.mock('../../../src/utils/markdownToOffice'); // Added this mock call

const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;
const mockLoggerInfo = jest.spyOn(logger, 'info');
const mockLoggerWarn = jest.spyOn(logger, 'warn');
const mockLoggerError = jest.spyOn(logger, 'error');
const mockFsPathExists = fs.pathExists as jest.Mock;
const mockMammothExtractRawText = mammoth.extractRawText as jest.Mock;
const mockCreateErrorResponse = mockCreateErrorResponseUtil as jest.Mock;
const mockResolveNaturalLanguageRange = require('../../../src/utils/wordRangeResolver').resolveNaturalLanguageRange as jest.Mock;
const mockApplyMarkdownFormattingToWord = applyMarkdownFormattingToWord as jest.Mock; // Corrected to use direct import


describe('word/text unit tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockSelection: any;
    let mockRange: any;
    let mockParagraphs: any;
    let mockContentRange: any;

    beforeEach(() => {
        jest.clearAllMocks();

        mockRange = {
            _textValue: 'Sample text from range',
            Start: 0, End: 22,
            Delete: jest.fn(), Collapse: jest.fn(),
            set Text(value: string) { (this as any)._textValue = value; },
            get Text() { return (this as any)._textValue; },
            release: jest.fn(),
        };
        mockSelection = { Range: mockRange, Type: 2, release: jest.fn() }; // wdSelectionIP
        mockParagraphs = {
            Count: 5,
            Item: jest.fn(index => {
                if (index > 0 && index <= mockParagraphs.Count) {
                    return { Range: { ...mockRange, Text: `Paragraph ${index} text`, _textValue: `Paragraph ${index} text`}, release: jest.fn() };
                }
                throw new Error(`COM Error: Paragraph index ${index} out of bounds.`); // Simulate COM error
            }),
            release: jest.fn(),
        };
        mockContentRange = { ...mockRange, Text: 'Document content text', _textValue: 'Document content text', Delete: jest.fn(), Start: 0, End: 20 };
        mockDoc = {
            Content: mockContentRange, Paragraphs: mockParagraphs,
            Range: jest.fn((start, end) => ({ ...mockRange, Start: start, End: end, _textValue: '', Text: '' })), // Ensure Text property exists for insertion
            Save: jest.fn(), Close: jest.fn(), release: jest.fn(),
            // Documents: { Add: jest.fn().mockReturnThis(), SaveAs2: jest.fn() }, // This was incorrect for mockDoc
        };
        mockWordApp = {
            Documents: {
                Open: jest.fn().mockReturnValue(mockDoc), // mockReturnValue for sync COM calls
                Add: jest.fn().mockReturnValue(mockDoc)
            },
            Selection: mockSelection, Quit: jest.fn(), release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockWordApp);
        mockValidateFilePath.mockImplementation(fp => fp);
        mockFsPathExists.mockResolvedValue(true);
        mockMammothExtractRawText.mockResolvedValue({ value: 'Mammoth extracted text', messages: [] });
        mockResolveNaturalLanguageRange.mockReturnValue(null);
        // Ensure mockApplyMarkdownFormattingToWord is reset and configured for each test if needed, or here globally
        mockApplyMarkdownFormattingToWord.mockReset().mockResolvedValue(undefined);
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
                if(result.success) expect(result.data).toBe('Sample text from range');
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp.Selection.Range);
            });

            test('COM: should get text from document', async () => {
                const params = getParams('document');
                const result = await getText(params);
                expect(result.success).toBe(true);
                if(result.success) expect(result.data).toBe('Document content text');
                expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc.Content);
            });

            test('COM: should get text from paragraph', async () => {
                const params = getParams('paragraph:2');
                const result = await getText(params);
                expect(mockParagraphs.Item).toHaveBeenCalledWith(2);
                expect(result.success).toBe(true);
                if(result.success) expect(result.data).toBe('Paragraph 2 text');
                expect(mockReleaseObject).toHaveBeenCalledWith(mockParagraphs.Item(2).Range);
            });

            test('COM: should use resolveNaturalLanguageRange first', async () => {
                const naturalRange = { ...mockRange, Text: "Natural language text", _textValue: "Natural language text", release: jest.fn() };
                mockResolveNaturalLanguageRange.mockReturnValue(naturalRange);
                const params = getParams('the first sentence');
                const result = await getText(params);
                expect(mockResolveNaturalLanguageRange).toHaveBeenCalledWith(mockDoc, params.range, mockWordApp);
                expect(result.success).toBe(true);
                if(result.success) expect(result.data).toBe("Natural language text");
                expect(mockReleaseObject).toHaveBeenCalledWith(naturalRange);
            });

            test('COM: should return error if file not found', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = getParams('document');
                const result = await getText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('FILE_NOT_FOUND');
            });

            test('COM: should return error if doc open fails', async () => {
                mockWordApp.Documents.Open.mockReturnValue(null);
                const params = getParams('document');
                const result = await getText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('FILE_OPEN_FAILED_COM');
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp); // wordApp is released
            });

            test('COM: should return error if range not found by either method and getRangeFromSpecifier throws', async () => {
                mockResolveNaturalLanguageRange.mockReturnValue(null);
                const params = getParams('paragraph:99'); // This will make getRangeFromSpecifier throw
                const result = await getText(params);
                expect(result.success).toBe(false);
                if(!result.success) {
                    expect(result.error.code).toBe('GET_TEXT_FAILED_COM');
                    expect(result.error.message).toContain("Paragraph index 99 is out of bounds");
                }
            });
             test('COM: should return error if range not found by either method and getRangeFromSpecifier returns null (though it should throw)', async () => {
                mockResolveNaturalLanguageRange.mockReturnValue(null);
                // Temporarily make getRangeFromSpecifier return null instead of throwing for a specific case
                const originalGetRange = jest.requireActual('../../../src/tools/word/text.tool').getRangeFromSpecifier;
                const mockGetRangeFromSpecifier = jest.fn(originalGetRange); // Spy on the original
                (getRangeFromSpecifier as jest.Mock) = mockGetRangeFromSpecifier; // Re-assign temporarily for this test
                mockGetRangeFromSpecifier.mockImplementationOnce(() => null); // Make it return null once

                const params = getParams('some_specific_range_that_getRange_handles_but_returns_null');
                const result = await getText(params);
                expect(result.success).toBe(false);
                if(!result.success) {
                     expect(result.error.code).toBe('RANGE_ERROR_COM'); // Specific error code when range is null
                     expect(result.error.message).toContain("Could not determine range for extraction via COM");
                }
                (getRangeFromSpecifier as jest.Mock) = jest.fn(originalGetRange); // Restore original mock behavior if needed elsewhere, or rely on beforeEach
            });


        } else { // Library path
            test('Library: should get text from document using Mammoth (default useComInterop:false)', async () => {
                const params = { ...baseParams, range: 'document' }; // No useComInterop, defaults to false
                const result = await getText(params);
                expect(mockFsPathExists).toHaveBeenCalledWith(path.resolve(params.filePath));
                expect(mockMammothExtractRawText).toHaveBeenCalledWith({ path: path.resolve(params.filePath) });
                expect(result.success).toBe(true);
                if(result.success) expect(result.data).toBe('Mammoth extracted text');
            });

            test('Library: should return error for "selection" range', async () => {
                const params = getParams('selection');
                const result = await getText(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('RANGE_REQUIRES_COM_LIB');
            });

            test('Library: should return error for "paragraph:N" range', async () => {
                const params = getParams('paragraph:1');
                const result = await getText(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('RANGE_REQUIRES_COM_LIB');
            });
            
            test('Library: should return error if mammoth fails', async () => {
                mockMammothExtractRawText.mockRejectedValue(new Error("Mammoth boom"));
                const params = getParams('document');
                const result = await getText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('MAMMOTH_GET_TEXT_FAILED');
            });
        }
    });

    describe.each(testModes)('insertText (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParams = { filePath: 'C:/test/document.docx', text: 'New inserted text' };
        const getParams = (position: string, format?: 'plaintext' | 'markdown') => ({ ...baseParams, position, format, useComInterop });

        if (useComInterop) {
            test('COM: should insert text at start (markdown by default)', async () => {
                const params = getParams('start'); // format defaults to markdown
                const result = await insertText(params);
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(path.resolve(params.filePath), false, false);
                expect(mockDoc.Range).toHaveBeenCalledWith(0, 0);
                const insertionRange = mockDoc.Range.mock.results[0].value;
                expect(mockApplyMarkdownFormattingToWord).toHaveBeenCalledWith(insertionRange, params.text, mockWordApp, mockDoc);
                expect(mockDoc.Save).toHaveBeenCalled();
                expect(result.success).toBe(true);
            });
            
            test('COM: should insert text at end (plaintext)', async () => {
                const params = getParams('end', 'plaintext');
                mockDoc.Content.End = 100;
                const result = await insertText(params);
                expect(mockDoc.Range).toHaveBeenCalledWith(100, 100);
                const insertionRange = mockDoc.Range.mock.results[0].value;
                expect(insertionRange.Text).toBe(params.text);
                expect(mockDoc.Save).toHaveBeenCalled();
                expect(result.success).toBe(true);
            });

            test('COM: should insert text at selection (markdown), Type not IP', async () => {
                const params = getParams('selection', 'markdown');
                mockWordApp.Selection.Type = 1; // wdSelectionNormal
                const result = await insertText(params);
                expect(mockWordApp.Selection.Range.Collapse).toHaveBeenCalledWith(1);
                expect(mockApplyMarkdownFormattingToWord).toHaveBeenCalledWith(mockWordApp.Selection.Range, params.text, mockWordApp, mockDoc);
                expect(result.success).toBe(true);
            });
             test('COM: should insert text at selection (markdown), Type IP', async () => {
                const params = getParams('selection', 'markdown');
                mockWordApp.Selection.Type = 2; // wdSelectionIP
                const result = await insertText(params);
                expect(mockWordApp.Selection.Range.Collapse).not.toHaveBeenCalled();
                expect(mockApplyMarkdownFormattingToWord).toHaveBeenCalledWith(mockWordApp.Selection.Range, params.text, mockWordApp, mockDoc);
                expect(result.success).toBe(true);
            });

            test('COM: should insert text at paragraph:N:start (plaintext)', async () => {
                const params = getParams('paragraph:2:start', 'plaintext');
                const mockParaRangeInstance = { Start: 50, End: 60, release: jest.fn() };
                mockParagraphs.Item.mockImplementation(idx => idx === 2 ? { Range: mockParaRangeInstance, release: jest.fn() } : null);
                const result = await insertText(params);
                expect(mockParagraphs.Item).toHaveBeenCalledWith(2);
                expect(mockDoc.Range).toHaveBeenCalledWith(mockParaRangeInstance.Start, mockParaRangeInstance.Start);
                const insertionRange = mockDoc.Range.mock.results[0].value;
                expect(insertionRange.Text).toBe(params.text);
                expect(result.success).toBe(true);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockParaRangeInstance);
            });
            
            test('COM: should insert text at paragraph:N:end (markdown)', async () => {
                const params = getParams('paragraph:3:end', 'markdown');
                const mockParaRangeInstance = { Start: 70, End: 80, release: jest.fn() };
                mockParagraphs.Item.mockImplementation(idx => idx === 3 ? { Range: mockParaRangeInstance, release: jest.fn() } : null);
                const result = await insertText(params);
                expect(mockParagraphs.Item).toHaveBeenCalledWith(3);
                expect(mockDoc.Range).toHaveBeenCalledWith(mockParaRangeInstance.End - 1, mockParaRangeInstance.End - 1);
                const insertionRange = mockDoc.Range.mock.results[0].value;
                expect(mockApplyMarkdownFormattingToWord).toHaveBeenCalledWith(insertionRange, params.text, mockWordApp, mockDoc);
                expect(result.success).toBe(true);
            });

            test('COM: should create file and insert text if file does not exist', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = getParams('start', 'plaintext');
                params.filePath = 'C:/test/new_doc_for_insert.docx';
                const result = await insertText(params);
                expect(mockWordApp.Documents.Add).toHaveBeenCalled();
                expect(mockDoc.SaveAs2).toHaveBeenCalledWith(path.resolve(params.filePath), 16);
                const insertionRange = mockDoc.Range.mock.results[0].value;
                expect(insertionRange.Text).toBe(params.text);
                expect(mockDoc.Save).not.toHaveBeenCalled();
                expect(result.success).toBe(true);
            });
            
            test('COM: should return error for invalid paragraph position format "paragraph:1:middle"', async () => {
                const params = getParams('paragraph:1:middle');
                const result = await insertText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('INVALID_POSITION_COM');
            });
             test('COM: should return error for invalid paragraph position format "paragraph:X"', async () => {
                const params = getParams('paragraph:X'); // Missing :start or :end
                const result = await insertText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('INVALID_POSITION_COM'); // Or specific error from resolveNaturalLanguageRange
            });


            test('COM: should return error if selection is null for position:selection', async () => {
                mockWordApp.Selection = null;
                const params = getParams('selection');
                const result = await insertText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('NO_SELECTION_COM');
            });
            
            test('COM: should return error if opening existing document fails', async () => {
                mockWordApp.Documents.Open.mockImplementation(() => { throw new Error("Open failed"); });
                const params = getParams('start');
                const result = await insertText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('FILE_OPERATION_FAILED_COM');
            });
            
            test('COM: should return error if creating new document fails', async () => {
                mockFsPathExists.mockResolvedValue(false);
                mockWordApp.Documents.Add.mockImplementation(() => { throw new Error("Add failed"); });
                const params = getParams('start');
                const result = await insertText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('FILE_OPERATION_FAILED_COM');
            });

        } else { // Library path
            test('Library: should return NOT_IMPLEMENTED_LIB (default useComInterop:false)', async () => {
                const params = { ...baseParams, position: 'start' };
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

            test('COM: should modify text for "paragraph:3" range', async () => {
                const params = getParams('paragraph:3');
                const result = await modifyText(params);
                expect(mockParagraphs.Item).toHaveBeenCalledWith(3);
                const targetRange = mockParagraphs.Item(3).Range;
                expect(targetRange.Text).toBe(params.newText);
                expect(mockDoc.Save).toHaveBeenCalled();
                expect(result.success).toBe(true);
            });

            test('COM: should return error if file not found', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = getParams('document');
                const result = await modifyText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('FILE_NOT_FOUND');
            });
            
            test('COM: should return error if opening document fails', async () => {
                mockWordApp.Documents.Open.mockReturnValue(null);
                const params = getParams('document');
                const result = await modifyText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('FILE_OPEN_FAILED_COM');
            });
            
            test('COM: should return error if getRangeFromSpecifier fails', async () => {
                const params = getParams('paragraph:99'); // Invalid range
                const result = await modifyText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('MODIFY_TEXT_FAILED_COM');
            });

        } else { // Library path
            test('Library: should return NOT_IMPLEMENTED_LIB (default useComInterop:false)', async () => {
                const params = { ...baseParams, range: 'document' };
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

            test('COM: should delete text for "document" range', async () => {
                const params = getParams('document');
                const result = await deleteText(params);
                expect(mockDoc.Content.Delete).toHaveBeenCalled();
                expect(mockDoc.Save).toHaveBeenCalled();
                expect(result.success).toBe(true);
            });
             test('COM: should return error if file not found', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = getParams('document');
                const result = await deleteText(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('FILE_NOT_FOUND');
            });
        } else { // Library path
            test('Library: should return NOT_IMPLEMENTED_LIB (default useComInterop:false)', async () => {
                const params = { ...baseParams, range: 'paragraph:1' };
                const result = await deleteText(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB');
            });
        }
    });

    describe('getRangeFromSpecifier (COM helper)', () => {
        beforeEach(() => {
             mockDoc.Paragraphs = mockParagraphs;
             mockWordApp.Selection = mockSelection;
        });

        test('should return selection range', () => {
            const range = getRangeFromSpecifier(mockDoc, 'selection', mockWordApp);
            expect(range).toBe(mockSelection.Range);
        });

        test('should return document content range', () => {
            const range = getRangeFromSpecifier(mockDoc, 'document', mockWordApp);
            expect(range).toBe(mockDoc.Content);
        });

        test('should return paragraph range', () => {
            const range = getRangeFromSpecifier(mockDoc, 'paragraph:1', mockWordApp);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(1);
            expect(range).toBe(mockParagraphs.Item(1).Range);
        });

        test('should throw error for invalid paragraph index format (e.g., "abc")', () => {
            expect(() => getRangeFromSpecifier(mockDoc, 'paragraph:abc', mockWordApp))
                .toThrow("Invalid paragraph index format: 'abc'");
        });
         test('should throw error for invalid paragraph index format (e.g., "0")', () => {
            expect(() => getRangeFromSpecifier(mockDoc, 'paragraph:0', mockWordApp))
                .toThrow("Invalid paragraph index format: '0'");
        });

        test('should throw error for paragraph index out of bounds', () => {
            mockParagraphs.Count = 2;
            expect(() => getRangeFromSpecifier(mockDoc, 'paragraph:3', mockWordApp))
                .toThrow("Paragraph index 3 is out of bounds. Document has 2 paragraphs.");
        });

        test('should throw error for unsupported range format', () => {
            expect(() => getRangeFromSpecifier(mockDoc, 'invalidRange', mockWordApp))
                .toThrow("Unsupported range format: 'invalidRange'");
        });

        test('should throw error if selection.Range is null and wordApp.Selection exists', () => {
            mockWordApp.Selection = { Range: null, release: jest.fn() };
            expect(() => getRangeFromSpecifier(mockDoc, 'selection', mockWordApp))
                .toThrow("Could not get range from selection");
            expect(mockLoggerWarn).toHaveBeenCalledWith(expect.stringContaining("wordApp.Selection.Range was null or undefined"));
        });
        
        test('should throw error if wordApp.Selection itself is null', () => {
            mockWordApp.Selection = null;
            expect(() => getRangeFromSpecifier(mockDoc, 'selection', mockWordApp))
                .toThrow("Could not get range from selection");
        });

        test('should throw error if doc.Content is null for "document" range', () => {
            mockDoc.Content = null;
            expect(() => getRangeFromSpecifier(mockDoc, 'document', mockWordApp))
                .toThrow("Could not get document content range.");
        });
        
        test('should throw error if doc.Paragraphs is null for "paragraph:N" range', () => {
            mockDoc.Paragraphs = null;
            expect(() => getRangeFromSpecifier(mockDoc, 'paragraph:1', mockWordApp))
                .toThrow("Could not access document paragraphs collection.");
        });
    });
});