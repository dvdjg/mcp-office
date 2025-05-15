import { applyAutoTitlesTool, ApplyAutoTitlesOutputData, SuggestedHeading } from '../../../src/tools/word/applyAutoTitles.tool';
import getActiveOfficeDocumentsTool from '../../../src/tools/os/getActiveOfficeDocuments.tool';
import { aiSuggestTool } from '../../../src/tools/office/aiSuggest.tool';
import { getText } from '../../../src/tools/word/text.tool';
import { applyStyle } from '../../../src/tools/word/styles.tool';
import logger from '../../../src/utils/logger';
import { createErrorResponse } from '../../../src/utils/errorHandler';
import type { ApiResponse, ToolRequestParams, FastMCPContext } from '../../../src/types/common.types';

// Mock dependencies
jest.mock('../../../src/tools/os/getActiveOfficeDocuments.tool.js');
jest.mock('../../../src/tools/office/aiSuggest.tool.js');
jest.mock('../../../src/tools/word/text.tool.js');
jest.mock('../../../src/tools/word/styles.tool.js');
jest.mock('../../../src/utils/logger.js');
jest.mock('../../../src/utils/errorHandler.js');

const mockedGetActiveOfficeDocumentsTool = getActiveOfficeDocumentsTool as jest.Mocked<typeof getActiveOfficeDocumentsTool>;
const mockedAiSuggestTool = aiSuggestTool as jest.Mocked<typeof aiSuggestTool>;
const mockedGetText = getText as jest.MockedFunction<typeof getText>;
const mockedApplyStyle = applyStyle as jest.MockedFunction<typeof applyStyle>;
const mockedLogger = logger as jest.Mocked<typeof logger>;
const mockedCreateErrorResponse = createErrorResponse as jest.MockedFunction<typeof createErrorResponse>;

describe('applyAutoTitlesTool', () => {
  let mockContext: FastMCPContext<any>;

  beforeEach(() => {
    jest.clearAllMocks();
    mockContext = {
      log: mockedLogger, // Assuming context has a log property
      session: {
        // Mock any session properties if needed by downstream tools like aiSuggestTool
        requestSampling: jest.fn().mockResolvedValue({ success: true, content: [{ type: 'text', text: '[]'}] }),
      },
    } as any; // Cast to any if FastMCPContext is complex to fully mock

    // Default mock implementations
    mockedGetActiveOfficeDocumentsTool.handler = jest.fn();
    mockedAiSuggestTool.handler = jest.fn();
    mockedGetText.mockResolvedValue({ success: true, data: 'Sample document text.' });
    mockedApplyStyle.mockResolvedValue({ success: true, data: {} });
    mockedCreateErrorResponse.mockImplementation((code, message, details) => ({
        success: false,
        error: { code, message, details },
    }));
    mockedLogger.info = jest.fn();
    mockedLogger.warn = jest.fn();
    mockedLogger.error = jest.fn();
  });

  it('should successfully apply titles when active document and suggestions are found', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: true,
      data: { documents: [{ applicationType: 'Word', filePath: 'C:\\path\\to\\active.docx' }] },
    } as ApiResponse<any>);

    const mockSuggestedHeadings: SuggestedHeading[] = [
      { text: 'Chapter 1', level: 'Heading 1', paragraphIdentifier: 'Chapter 1' },
      { text: 'Introduction', level: 'Heading 2', paragraphIdentifier: 'Introduction' },
    ];
    // aiSuggestTool's handler is expected to return ApiResponse<string> where string is JSON
    mockedAiSuggestTool.handler.mockResolvedValue({
        success: true,
        data: JSON.stringify(mockSuggestedHeadings)
    } as ApiResponse<string>);

    const params: ToolRequestParams = {};
    const result = await applyAutoTitlesTool.handler(params, mockContext);

    expect(result.success).toBe(true);
    if (result.success) {
      const resultData = result.data as ApplyAutoTitlesOutputData;
      expect(resultData.message).toContain('Successfully processed Word document.');
      expect(resultData.appliedHeadings).toHaveLength(2);
      expect(resultData.appliedHeadings[0]).toEqual({ text: 'Chapter 1', style: 'Heading 1' });
    }
    expect(mockedGetText).toHaveBeenCalledWith(
      { filePath: 'C:\\path\\to\\active.docx', range: 'document', useComInterop: false },
      mockContext
    );
    expect(mockedAiSuggestTool.handler).toHaveBeenCalled();
    expect(mockedApplyStyle).toHaveBeenCalledTimes(2);
    expect(mockedApplyStyle).toHaveBeenCalledWith(
      expect.objectContaining({ filePath: 'C:\\path\\to\\active.docx', range: 'Chapter 1', style: 'Heading 1' }),
      mockContext
    );
  });

  it('should use filePath from input if provided', async () => {
    const providedPath = 'C:\\direct\\path\\document.docx';
    const mockSuggestedHeadings: SuggestedHeading[] = [
      { text: 'Section A', level: 'Heading 1', paragraphIdentifier: 'Section A' },
    ];
     mockedAiSuggestTool.handler.mockResolvedValue({
        success: true,
        data: JSON.stringify(mockSuggestedHeadings)
    } as ApiResponse<string>);


    const params: ToolRequestParams = { filePath: providedPath };
    await applyAutoTitlesTool.handler(params, mockContext);

    expect(mockedGetActiveOfficeDocumentsTool.handler).not.toHaveBeenCalled();
    expect(mockedGetText).toHaveBeenCalledWith(
      { filePath: providedPath, range: 'document', useComInterop: false },
      mockContext
    );
    expect(mockedApplyStyle).toHaveBeenCalledWith(
      expect.objectContaining({ filePath: providedPath, range: 'Section A', style: 'Heading 1' }),
      mockContext
    );
  });

  it('should return error if no active Word document is found and no filePath provided', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: true,
      data: { documents: [] }, // No documents
    } as ApiResponse<any>);

    const params: ToolRequestParams = {};
    const result = await applyAutoTitlesTool.handler(params, mockContext);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('NO_ACTIVE_WORD_DOC');
    }
    expect(mockedCreateErrorResponse).toHaveBeenCalledWith('NO_ACTIVE_WORD_DOC', 'No active Word document found.');
  });

  it('should return error if getActiveOfficeDocumentsTool fails', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: false,
      error: { code: 'OS_ERROR', message: 'OS specific error' },
    } as ApiResponse<any>);
    
    const params: ToolRequestParams = {};
    const result = await applyAutoTitlesTool.handler(params, mockContext);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('ACTIVE_DOC_FAILED');
    }
    expect(mockedCreateErrorResponse).toHaveBeenCalledWith('ACTIVE_DOC_FAILED', 'OS specific error');
  });

  it('should return error if getText fails', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: true,
      data: { documents: [{ applicationType: 'Word', filePath: 'C:\\path\\to\\active.docx' }] },
    } as ApiResponse<any>);
    mockedGetText.mockResolvedValue({
      success: false,
      error: { code: 'READ_ERROR', message: 'Cannot read file' },
    } as ApiResponse<string>);

    const params: ToolRequestParams = {};
    const result = await applyAutoTitlesTool.handler(params, mockContext);
    
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('GET_TEXT_FAILED');
    }
    expect(mockedCreateErrorResponse).toHaveBeenCalledWith('GET_TEXT_FAILED', 'Cannot read file');
  });

  it('should return error if document text is empty', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: true,
      data: { documents: [{ applicationType: 'Word', filePath: 'C:\\path\\to\\active.docx' }] },
    } as ApiResponse<any>);
    mockedGetText.mockResolvedValue({ success: true, data: '  ' }); // Empty or whitespace only

    const params: ToolRequestParams = {};
    const result = await applyAutoTitlesTool.handler(params, mockContext);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('EMPTY_DOCUMENT');
    }
  });

  it('should return error if aiSuggestTool fails', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: true,
      data: { documents: [{ applicationType: 'Word', filePath: 'C:\\path\\to\\active.docx' }] },
    } as ApiResponse<any>);
    mockedAiSuggestTool.handler.mockResolvedValue({
      success: false,
      error: { code: 'AI_ERROR', message: 'AI suggestion failed' },
    } as ApiResponse<string>);

    const params: ToolRequestParams = {};
    const result = await applyAutoTitlesTool.handler(params, mockContext);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('AI_SUGGEST_FAILED');
    }
  });

  it('should return error if aiSuggestTool returns non-string data', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: true,
      data: { documents: [{ applicationType: 'Word', filePath: 'C:\\path\\to\\active.docx' }] },
    } as ApiResponse<any>);
    mockedAiSuggestTool.handler.mockResolvedValue({
      success: true,
      data: { someObject: true }, // Not a string
    } as any); // Cast to any to simulate incorrect data type

    const params: ToolRequestParams = {};
    const result = await applyAutoTitlesTool.handler(params, mockContext);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('AI_SUGGEST_INVALID_FORMAT');
    }
  });
  
  it('should return error if aiSuggestTool returns unparsable JSON string', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: true,
      data: { documents: [{ applicationType: 'Word', filePath: 'C:\\path\\to\\active.docx' }] },
    } as ApiResponse<any>);
    mockedAiSuggestTool.handler.mockResolvedValue({
      success: true,
      data: "this is not json",
    } as ApiResponse<string>);

    const params: ToolRequestParams = {};
    const result = await applyAutoTitlesTool.handler(params, mockContext);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('AI_SUGGEST_PARSE_FAILED');
    }
  });

  it('should return success with no applied headings if aiSuggest returns empty array', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: true,
      data: { documents: [{ applicationType: 'Word', filePath: 'C:\\path\\to\\active.docx' }] },
    } as ApiResponse<any>);
    mockedAiSuggestTool.handler.mockResolvedValue({
        success: true,
        data: JSON.stringify([])
    } as ApiResponse<string>);

    const params: ToolRequestParams = {};
    const result = await applyAutoTitlesTool.handler(params, mockContext);

    expect(result.success).toBe(true);
    if (result.success) {
      const resultData = result.data as ApplyAutoTitlesOutputData;
      expect(resultData.message).toContain('No headings were identified');
      expect(resultData.appliedHeadings).toHaveLength(0);
    }
  });

  it('should continue applying styles if one application fails', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: true,
      data: { documents: [{ applicationType: 'Word', filePath: 'C:\\path\\to\\active.docx' }] },
    } as ApiResponse<any>);
    const mockSuggestedHeadings: SuggestedHeading[] = [
      { text: 'Good Heading', level: 'Heading 1', paragraphIdentifier: 'Good Heading' },
      { text: 'Bad Heading', level: 'Heading 2', paragraphIdentifier: 'Bad Heading' }, // This one will fail
      { text: 'Another Good', level: 'Heading 1', paragraphIdentifier: 'Another Good' },
    ];
     mockedAiSuggestTool.handler.mockResolvedValue({
        success: true,
        data: JSON.stringify(mockSuggestedHeadings)
    } as ApiResponse<string>);

    mockedApplyStyle
      .mockResolvedValueOnce({ success: true, data: {} }) // Good Heading
      .mockResolvedValueOnce({ success: false, error: { code: 'STYLE_ERROR', message: 'Failed to apply' } }) // Bad Heading
      .mockResolvedValueOnce({ success: true, data: {} }); // Another Good

    const params: ToolRequestParams = {};
    const result = await applyAutoTitlesTool.handler(params, mockContext);

    expect(result.success).toBe(true);
    if (result.success) {
      const resultData = result.data as ApplyAutoTitlesOutputData;
      expect(resultData.appliedHeadings).toHaveLength(2);
      expect(resultData.appliedHeadings.find(h => h.text === 'Good Heading')).toBeDefined();
      expect(resultData.appliedHeadings.find(h => h.text === 'Another Good')).toBeDefined();
      expect(resultData.appliedHeadings.find(h => h.text === 'Bad Heading')).toBeUndefined();
    }
    expect(mockedApplyStyle).toHaveBeenCalledTimes(3);
    expect(mockedLogger.error).toHaveBeenCalledWith(
        expect.stringContaining("Error applying style 'Heading 2' to \"Bad Heading\""),
        expect.anything()
    );
  });

   it('should return appropriate message if AI suggests headings but none could be applied', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: true,
      data: { documents: [{ applicationType: 'Word', filePath: 'C:\\path\\to\\active.docx' }] },
    } as ApiResponse<any>);
    const mockSuggestedHeadings: SuggestedHeading[] = [
      { text: 'Failed Heading 1', level: 'Heading 1', paragraphIdentifier: 'Failed Heading 1' },
      { text: 'Failed Heading 2', level: 'Heading 2', paragraphIdentifier: 'Failed Heading 2' },
    ];
    mockedAiSuggestTool.handler.mockResolvedValue({
        success: true,
        data: JSON.stringify(mockSuggestedHeadings)
    } as ApiResponse<string>);

    mockedApplyStyle.mockResolvedValue({ success: false, error: { code: 'STYLE_ERROR', message: 'Failed to apply' } }); // All fail

    const params: ToolRequestParams = {};
    const result = await applyAutoTitlesTool.handler(params, mockContext);

    expect(result.success).toBe(true); // Tool itself succeeded, but operation within might not fully complete
    if (result.success) {
      const resultData = result.data as ApplyAutoTitlesOutputData;
      expect(resultData.message).toEqual('AI suggested headings, but none could be applied. Check logs for details.');
      expect(resultData.appliedHeadings).toHaveLength(0);
    }
  });

  it('should skip invalid heading suggestions (missing text, level, or identifier)', async () => {
    mockedGetActiveOfficeDocumentsTool.handler.mockResolvedValue({
      success: true,
      data: { documents: [{ applicationType: 'Word', filePath: 'C:\\path\\to\\active.docx' }] },
    } as ApiResponse<any>);
    const mockSuggestedHeadings: any[] = [
      { text: 'Valid Heading', level: 'Heading 1', paragraphIdentifier: 'Valid Heading' },
      { level: 'Heading 2', paragraphIdentifier: 'Missing Text' }, // Missing text
      { text: 'Missing Level', paragraphIdentifier: 'Missing Level' }, // Missing level
      { text: 'Missing Identifier', level: 'Heading 3'}, // Missing identifier (will fallback to text)
      { text: 'Valid Heading 2', level: 'Heading 1', paragraphIdentifier: 'Valid Heading 2' },
    ];
     mockedAiSuggestTool.handler.mockResolvedValue({
        success: true,
        data: JSON.stringify(mockSuggestedHeadings)
    } as ApiResponse<string>);
    
    const params: ToolRequestParams = {};
    await applyAutoTitlesTool.handler(params, mockContext);

    expect(mockedApplyStyle).toHaveBeenCalledTimes(3); // Valid, Missing Identifier (uses text), Valid 2
    expect(mockedLogger.warn).toHaveBeenCalledWith(expect.stringContaining("Skipping invalid heading suggestion (missing text, level, or identifier)"), expect.stringContaining("\"level\":\"Heading 2\""));
    expect(mockedLogger.warn).toHaveBeenCalledWith(expect.stringContaining("Skipping invalid heading suggestion (missing text, level, or identifier)"), expect.stringContaining("\"text\":\"Missing Level\""));
  });

});