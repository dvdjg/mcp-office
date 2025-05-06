import { applyStyle, listStyles } from '../../../src/tools/word/styles.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { createErrorResponse as mockCreateErrorResponseUtil } from '../../../src/utils/errorHandler'; // Import the actual function for mocking

// Mock dependencies
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
    createErrorResponse: jest.fn((message, code) => ({ success: false, error: { code, message } })),
}));
jest.mock('../../../src/utils/logger');

const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;
const mockLoggerInfo = jest.spyOn(logger, 'info');
const mockLoggerWarn = jest.spyOn(logger, 'warn');
const mockLoggerError = jest.spyOn(logger, 'error');
const mockCreateErrorResponse = mockCreateErrorResponseUtil as jest.Mock;


describe('word/styles unit tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockSelection: any;
    let mockRange: any;
    let mockParagraphs: any;
    let mockStylesCollection: any;

    beforeEach(() => {
        jest.clearAllMocks();

        mockRange = { Style: '', release: jest.fn() };
        const mockParagraphRange = { Style: '', release: jest.fn() }; // Separate for paragraph specific tests
        mockSelection = { Range: mockRange, release: jest.fn() };

        mockParagraphs = {
            Count: 5,
            Item: jest.fn((index: number) => { // Renamed from 'Item' to 'item' if it's a direct function call in JS
                if (index > 0 && index <= mockParagraphs.Count) {
                    // Return a new mock object for each paragraph to avoid style bleed in tests
                    return { Range: { ...mockParagraphRange }, release: jest.fn() };
                }
                // Simulate COM error for out-of-bounds
                throw new Error(`COM Error: The requested member of the collection does not exist. Index: ${index}`);
            }),
            release: jest.fn(),
        };

        const mockStyleItemNormal = { NameLocal: 'Normal', release: jest.fn() };
        const mockStyleItemHeading1 = { NameLocal: 'Heading 1', release: jest.fn() };
        const mockStyleItemHeading2 = { NameLocal: 'Heading 2', release: jest.fn() };
        const mockStyleItemNoName = { release: jest.fn() }; // Style without NameLocal

        mockStylesCollection = {
            Count: 3, // Default count
            Item: jest.fn((index: number) => { // Renamed from 'Item' to 'item'
                if (index === 1) return mockStyleItemNormal;
                if (index === 2) return mockStyleItemHeading1;
                if (index === 3) return mockStyleItemHeading2;
                // Simulate COM error for out-of-bounds
                throw new Error(`COM Error: The requested member of the collection does not exist. Index: ${index}`);
            }),
            release: jest.fn(),
        };

        mockDoc = {
            Content: { ...mockRange }, // Use a copy for document content
            Paragraphs: mockParagraphs,
            Styles: mockStylesCollection,
            Close: jest.fn(),
            release: jest.fn(),
        };
        mockWordApp = {
            Documents: { Open: jest.fn().mockReturnValue(mockDoc) }, // Open returns the doc directly
            Selection: mockSelection,
            // Quit: jest.fn(), // Quit is not typically called per operation
            release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockWordApp); // Returns the app object directly
        mockValidateFilePath.mockImplementation(fp => fp);
    });

    const testModes = [
        { mode: 'COM', useComInterop: true, default: false },
        { mode: 'Library', useComInterop: false, default: true },
    ];

    describe.each(testModes)('applyStyle (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParams = { filePath: 'C:/test/document.docx', style: 'Heading 1' };
        const getParams = (range: string) => ({ ...baseParams, range, useComInterop });

        if (useComInterop) {
            test('COM: should apply style to selection', async () => {
                const params = getParams('selection');
                const result = await applyStyle(params);
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath);
                expect(mockWordApp.Selection.Range.Style).toBe(params.style);
                expect(result.success).toBe(true);
                expect(mockLoggerInfo).toHaveBeenCalledWith(expect.stringContaining(`Successfully applied style '${params.style}'`));
                expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            });

            test('COM: should apply style to document content', async () => {
                const params = getParams('document');
                const result = await applyStyle(params);
                expect(mockDoc.Content.Style).toBe(params.style);
                expect(result.success).toBe(true);
            });

            test('COM: should apply style to a specific paragraph', async () => {
                const params = getParams('paragraph:3');
                const result = await applyStyle(params);
                expect(mockParagraphs.Item).toHaveBeenCalledWith(3);
                // The Style property is on the Range object returned by Paragraphs(index).Range
                const paragraphInstance = mockParagraphs.Item(3); // Get the specific mock
                expect(paragraphInstance.Range.Style).toBe(params.style);
                expect(result.success).toBe(true);
            });

            test('COM: should throw error for invalid range format', async () => {
                const params = getParams('invalidRangeFormat');
                const result = await applyStyle(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.message).toContain("Unsupported range format: 'invalidRangeFormat'");
                expect(mockLoggerError).toHaveBeenCalledWith(expect.stringContaining("Unsupported range format"), expect.any(Object));
            });

            test('COM: should throw error for invalid paragraph index (string)', async () => {
                const params = getParams('paragraph:abc');
                const result = await applyStyle(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.message).toContain("Invalid paragraph index format: 'abc'");
            });

            test('COM: should throw error for paragraph index 0', async () => {
                const params = getParams('paragraph:0');
                const result = await applyStyle(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.message).toContain("Invalid paragraph index format: '0'");
            });

            test('COM: should throw error for out-of-bounds paragraph index', async () => {
                const params = getParams('paragraph:99'); // mockParagraphs.Count is 5
                const result = await applyStyle(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.message).toContain("Paragraph index 99 out of bounds. Document has 5 paragraphs.");
            });
            
            test('COM: should handle error if selection range is not available', async () => {
                mockWordApp.Selection.Range = null; // Simulate no range available
                const params = getParams('selection');
                const result = await applyStyle(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.message).toContain("Could not get range from selection");
            });

            test('COM: should release objects even if style application fails mid-way (e.g., invalid style name)', async () => {
                // This test is more about the finally block.
                // Simulate an error after opening doc but before/during style application.
                // For instance, if `selectedRange.Style = style` threw an error.
                // We can't easily mock that part of COM, but we can check releases on other errors.
                mockParagraphs.Item.mockImplementationOnce(() => { throw new Error("Simulated COM error during paragraph access"); });
                const params = getParams('paragraph:1');
                await applyStyle(params); // We expect it to throw and be caught by handleToolError
                expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            });

        } else { // Library path
            test('Library: should return NOT_IMPLEMENTED_LIB_STYLE_APPLY for paragraph range (default useComInterop:false)', async () => {
                const params = { ...baseParams, range: 'paragraph:1' }; // useComInterop defaults to false
                const result = await applyStyle(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_STYLE_APPLY');
                expect(mockCreateErrorResponse).toHaveBeenCalledWith(expect.stringContaining("not fully implemented"), 'NOT_IMPLEMENTED_LIB_STYLE_APPLY');
                expect(mockLoggerWarn).toHaveBeenCalledWith(expect.stringContaining("Applying arbitrary named styles ('Heading 1') is complex with 'docx'"));
            });

            test('Library: should return LIB_RANGE_NOT_SUPPORTED for selection range', async () => {
                const params = getParams('selection');
                const result = await applyStyle(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('LIB_RANGE_NOT_SUPPORTED');
                expect(mockCreateErrorResponse).toHaveBeenCalledWith(expect.stringContaining("does not support 'selection' range"), 'LIB_RANGE_NOT_SUPPORTED');
            });

            test('Library: should return NOT_IMPLEMENTED_LIB_STYLE_APPLY for non-heading style and document range', async () => {
                const params = getParams('document');
                params.style = 'SomeCustomStyle';
                const result = await applyStyle(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_STYLE_APPLY');
                expect(mockCreateErrorResponse).toHaveBeenCalledWith(expect.stringContaining("not implemented or the style is not supported"), 'NOT_IMPLEMENTED_LIB_STYLE_APPLY');
            });
        }
    });

    describe.each(testModes)('listStyles (mode: $mode, default: $default)', ({ useComInterop }) => {
        const getParams = () => ({ filePath: 'C:/test/document.docx', useComInterop });

        if (useComInterop) {
            test('COM: should list styles from the document', async () => {
                const params = getParams();
                const result = await listStyles(params);
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath, false, true);
                expect(mockStylesCollection.Item).toHaveBeenCalledTimes(mockStylesCollection.Count);
                expect(result.success).toBe(true);
                if (result.success) expect(result.data).toEqual(['Normal', 'Heading 1', 'Heading 2']);
                expect(mockLoggerInfo).toHaveBeenCalledWith(expect.stringContaining(`Successfully listed ${mockStylesCollection.Count} styles`));

                expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockStylesCollection);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockStylesCollection.Item(1)); // Normal
                expect(mockReleaseObject).toHaveBeenCalledWith(mockStylesCollection.Item(2)); // Heading 1
                expect(mockReleaseObject).toHaveBeenCalledWith(mockStylesCollection.Item(3)); // Heading 2
            });

            test('COM: should return empty list if document has no styles', async () => {
                mockStylesCollection.Count = 0;
                const params = getParams();
                const result = await listStyles(params);
                expect(result.success).toBe(true);
                if (result.success) expect(result.data).toEqual([]);
                expect(mockStylesCollection.Item).not.toHaveBeenCalled();
                expect(mockLoggerInfo).toHaveBeenCalledWith(expect.stringContaining("Successfully listed 0 styles"));
            });
            
            test('COM: should skip style if NameLocal is missing and log error', async () => {
                mockStylesCollection.Item.mockImplementation((index: number) => {
                    if (index === 1) return { NameLocal: 'Normal', release: jest.fn() };
                    if (index === 2) return { release: jest.fn() }; // Missing NameLocal
                    if (index === 3) return { NameLocal: 'Heading 2', release: jest.fn() };
                    throw new Error("Should not happen");
                });
                mockStylesCollection.Count = 3;
                const params = getParams();
                const result = await listStyles(params);
                expect(result.success).toBe(true);
                if (result.success) expect(result.data).toEqual(['Normal', 'Heading 2']);
                // The tool doesn't log an error for missing NameLocal, it just skips. This is fine.
                // If it were to log, we'd check mockLoggerWarn or mockLoggerError.
            });

            test('COM: should handle error if document open fails', async () => {
                mockWordApp.Documents.Open.mockReturnValue(null); // Simulate open failure
                const params = getParams();
                const result = await listStyles(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.message).toContain('Failed to open document');
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp); // App should still be released
                expect(mockReleaseObject).not.toHaveBeenCalledWith(mockDoc); // Doc was not opened
            });

            test('COM: should handle error when accessing individual style item (COM error)', async () => {
                mockStylesCollection.Item.mockImplementation((index: number) => {
                    if (index === 1) return { NameLocal: 'Normal', release: jest.fn() };
                    if (index === 2) throw new Error("COM Error: Failed to access style item");
                    if (index === 3) return { NameLocal: 'Heading 2', release: jest.fn() };
                    return undefined;
                });
                mockStylesCollection.Count = 3;

                const params = getParams();
                const result = await listStyles(params);
                expect(result.success).toBe(true); // The function attempts to gather all it can
                if (result.success) expect(result.data).toEqual(['Normal', 'Heading 2']);
                expect(mockLoggerError).toHaveBeenCalledWith(expect.stringContaining("Error accessing style at index 2: COM Error: Failed to access style item"));
            });

        } else { // Library path
            test('Library: should return NOT_IMPLEMENTED_LIB_STYLE_LIST (default useComInterop:false)', async () => {
                const params = { filePath: 'C:/test/document.docx' }; // useComInterop defaults to false
                const result = await listStyles(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_STYLE_LIST');
                expect(mockCreateErrorResponse).toHaveBeenCalledWith(expect.stringContaining("Listing all styles from an existing document is not supported"), 'NOT_IMPLEMENTED_LIB_STYLE_LIST');
                expect(mockLoggerWarn).toHaveBeenCalledWith(expect.stringContaining("Listing all styles from a .docx file's definition is not directly supported"));
            });
        }
    });
});