import { listStyles, applyStyle } from '../../../src/tools/word/styles.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { createErrorResponse } from '../../../src/utils/errorHandler';


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
const mockCreateErrorResponse = createErrorResponse as jest.Mock;


describe('word/styles integration tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockStylesCollection: any;
    let mockSelectionRange: any;
    let mockParagraphRange: any;
    let mockDocumentContentRange: any;


    beforeEach(() => {
        jest.clearAllMocks();

        mockStylesCollection = {
            Count: 3,
            Item: jest.fn((index: number) => {
                if (index === 1) return { NameLocal: 'Normal', release: jest.fn() };
                if (index === 2) return { NameLocal: 'Heading 1', release: jest.fn() };
                if (index === 3) return { NameLocal: 'Heading 2', release: jest.fn() };
                return undefined;
            }),
            release: jest.fn(),
        };

        mockSelectionRange = { Style: '', release: jest.fn() };
        mockParagraphRange = { Style: '', release: jest.fn() };
        mockDocumentContentRange = { Style: '', release: jest.fn() };

        mockDoc = {
            Styles: mockStylesCollection,
            Content: mockDocumentContentRange,
            Paragraphs: jest.fn((index: number) => {
                if (index > 0 && index <= (mockDoc.Paragraphs as any).Count) {
                    return { Range: mockParagraphRange, release: jest.fn() };
                }
                throw new Error("COM Error: Invalid paragraph index."); // Simulate COM error
            }),
            Close: jest.fn(),
            release: jest.fn(),
        };
        (mockDoc.Paragraphs as any).Count = 5;


        mockWordApp = {
            Documents: {
                Open: jest.fn().mockReturnValue(mockDoc),
            },
            Selection: {
                Range: mockSelectionRange,
                release: jest.fn(), // Selection object itself might need release
            },
            release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockWordApp);
        mockValidateFilePath.mockImplementation(filePath => filePath);
    });

    const testModes = [
        { mode: 'COM', useComInterop: true, default: false },
        { mode: 'Library', useComInterop: false, default: true }, // Default behavior is library path
    ];

    // Consolidate undefined with false for testing logic, as schema defaults undefined to false
    const effectiveTestModes = testModes.map(tm => ({
        ...tm,
        effectiveUseComInterop: tm.useComInterop === undefined ? false : tm.useComInterop,
    }));


    describe.each(effectiveTestModes)('listStyles (mode: $mode, useComInterop: $effectiveUseComInterop)', ({ useComInterop, effectiveUseComInterop }) => {
        const params = { filePath: 'C:/test/document.docx', useComInterop };

        if (effectiveUseComInterop) {
            test('COM: should list styles correctly', async () => {
                const result = await listStyles(params);
                expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.filePath, false, true);
                expect(mockStylesCollection.Item).toHaveBeenCalledTimes(3);
                expect(result).toEqual({ success: true, data: ['Normal', 'Heading 1', 'Heading 2'] });
                expect(mockReleaseObject).toHaveBeenCalledWith(mockStylesCollection.Item(1));
                expect(mockReleaseObject).toHaveBeenCalledWith(mockStylesCollection.Item(2));
                expect(mockReleaseObject).toHaveBeenCalledWith(mockStylesCollection.Item(3));
                expect(mockReleaseObject).toHaveBeenCalledWith(mockStylesCollection);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            });

            test('COM: should handle error if document open fails', async () => {
                mockWordApp.Documents.Open.mockReturnValue(null);
                const result = await listStyles(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error?.message).toContain('Failed to open document');
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            });

            test('COM: should handle empty styles collection', async () => {
                mockStylesCollection.Count = 0;
                const result = await listStyles(params);
                expect(result.success).toBe(true);
                if (result.success) expect(result.data).toEqual([]);
                expect(mockStylesCollection.Item).not.toHaveBeenCalled();
            });
            
            test('COM: should skip style if NameLocal is missing', async () => {
                mockStylesCollection.Item.mockImplementation((index: number) => {
                    if (index === 1) return { NameLocal: 'Normal', release: jest.fn() };
                    if (index === 2) return { release: jest.fn() }; // No NameLocal
                    if (index === 3) return { NameLocal: 'Heading 2', release: jest.fn() };
                    return undefined;
                });
                const result = await listStyles(params);
                expect(result.success).toBe(true);
                if(result.success) expect(result.data).toEqual(['Normal', 'Heading 2']);
            });

        } else { // Library path (effectiveUseComInterop is false)
            test('Library: should return NOT_IMPLEMENTED_LIB_STYLE_LIST', async () => {
                const result = await listStyles(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_STYLE_LIST');
                expect(mockCreateErrorResponse).toHaveBeenCalledWith(expect.stringContaining("Listing all styles from an existing document is not supported"), 'NOT_IMPLEMENTED_LIB_STYLE_LIST');
                expect(mockLoggerInfo).toHaveBeenCalledWith(expect.stringContaining(`useComInterop=false`));
            });
        }
    });

    describe.each(effectiveTestModes)('applyStyle (mode: $mode, useComInterop: $effectiveUseComInterop)', ({ useComInterop, effectiveUseComInterop }) => {
        const baseParams = { filePath: 'C:/test/document.docx', style: 'Heading 1' };
        const getParams = (range: string) => ({ ...baseParams, range, useComInterop });


        if (effectiveUseComInterop) {
            test('COM: should apply style to paragraph range', async () => {
                const paramsWithRange = getParams('paragraph:1');
                const result = await applyStyle(paramsWithRange);
                expect(mockDoc.Paragraphs).toHaveBeenCalledWith(1);
                expect(mockParagraphRange.Style).toBe(paramsWithRange.style);
                expect(result.success).toBe(true);
            });

            test('COM: should apply style to selection range', async () => {
                const paramsWithRange = getParams('selection');
                const result = await applyStyle(paramsWithRange);
                expect(mockSelectionRange.Style).toBe(paramsWithRange.style);
                expect(result.success).toBe(true);
            });

            test('COM: should apply style to document range', async () => {
                const paramsWithRange = getParams('document');
                const result = await applyStyle(paramsWithRange);
                expect(mockDocumentContentRange.Style).toBe(paramsWithRange.style);
                expect(result.success).toBe(true);
            });

            test('COM: should handle invalid paragraph index (e.g., "paragraph:0")', async () => {
                const paramsWithRange = getParams('paragraph:0');
                const result = await applyStyle(paramsWithRange);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.message).toContain("Invalid paragraph index format: '0'");
            });
            
            test('COM: should handle out-of-bounds paragraph index', async () => {
                const paramsWithRange = getParams('paragraph:99');
                const result = await applyStyle(paramsWithRange);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.message).toContain("Paragraph index 99 out of bounds.");
            });
            
            test('COM: should handle error if selection range is null', async () => {
                mockWordApp.Selection.Range = null;
                const paramsWithRange = getParams('selection');
                const result = await applyStyle(paramsWithRange);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.message).toContain("Could not get range from selection");
            });


        } else { // Library path
            test('Library: should return NOT_IMPLEMENTED_LIB_STYLE_APPLY for paragraph range', async () => {
                const paramsWithRange = getParams('paragraph:1');
                const result = await applyStyle(paramsWithRange);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_STYLE_APPLY');
            });

            test('Library: should return LIB_RANGE_NOT_SUPPORTED for selection range', async () => {
                const paramsWithRange = getParams('selection');
                const result = await applyStyle(paramsWithRange);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('LIB_RANGE_NOT_SUPPORTED');
            });
        }
    });
});