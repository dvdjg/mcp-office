import { jest } from '@jest/globals';
import { concludeStoryInDocumentTool } from '../../../src/tools/word/concludeStoryInDocument.tool.js';
import getActiveOfficeDocumentsTool from '../../../src/tools/os/getActiveOfficeDocuments.tool.js';
import { wordTextTool } from '../../../src/tools/word/text.tool.js';
import { wordGenerateAndInsertTextTool } from '../../../src/tools/word/generateAndInsertText.tool.js';

describe('word/concludeStoryInDocument', () => {
  const context = { log: console } as any;

  beforeEach(() => {
    jest.restoreAllMocks();
  });

  test('concludes a story from an active document name', async () => {
    jest.spyOn(getActiveOfficeDocumentsTool, 'handler').mockResolvedValue({
      success: true,
      data: {
        documents: [
          { applicationType: 'Word', resolvedPath: 'C:\\Docs\\Story.docx', filePath: 'C:\\Docs\\Story.docx' },
        ],
      },
    } as any);
    const getTextTool = wordTextTool.find((tool) => tool.path === 'word/text/get');
    jest.spyOn(getTextTool!, 'handler').mockResolvedValue({
      success: true,
      data: 'Para 1\n\nPara 2\n\nPara 3',
    } as any);
    jest.spyOn(wordGenerateAndInsertTextTool, 'handler').mockResolvedValue({
      success: true,
      data: 'ok',
    } as any);

    const result = await concludeStoryInDocumentTool.handler({ documentNameOrPath: 'Story.docx' }, context);

    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data).toContain('C:\\Docs\\Story.docx');
    }
  });

  test('returns DOCUMENT_NOT_FOUND when the named document is not active', async () => {
    jest.spyOn(getActiveOfficeDocumentsTool, 'handler').mockResolvedValue({
      success: true,
      data: { documents: [] },
    } as any);

    const result = await concludeStoryInDocumentTool.handler({ documentNameOrPath: 'Missing.docx' }, context);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('DOCUMENT_NOT_FOUND');
    }
  });
});
