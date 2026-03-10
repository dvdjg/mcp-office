import { jest } from '@jest/globals';
import adaptWordToPowerpointTool from '../../../src/tools/office/adaptWordToPowerpoint.tool';
import getActiveOfficeDocumentsTool from '../../../src/tools/os/getActiveOfficeDocuments.tool';
import wordToPowerpointTool from '../../../src/tools/office/wordToPowerpoint.tool';

describe('office/adaptWordToPowerpoint', () => {
  beforeEach(() => {
    jest.restoreAllMocks();
  });

  test('uses explicit input path when provided', async () => {
    const conversionSpy = jest.spyOn(wordToPowerpointTool, 'handler').mockResolvedValue({
      success: true,
      data: { ok: true },
    } as any);

    const result = await adaptWordToPowerpointTool.handler({
      inputWordPath: 'C:/docs/report.docx',
      headingLevelForNewSlide: 2,
    });

    expect(result.success).toBe(true);
    expect(conversionSpy).toHaveBeenCalledWith(
      expect.objectContaining({
        wordFilePath: 'C:/docs/report.docx',
        powerpointFilePath: 'C:\\docs\\report.pptx',
        headingLevelForNewSlide: 2,
        operation: 'transfer',
      }),
      undefined,
    );
  });

  test('falls back to the active Word document', async () => {
    jest.spyOn(getActiveOfficeDocumentsTool, 'handler').mockResolvedValue({
      success: true,
      data: {
        documents: [{ filePath: 'C:/docs/active.docx', applicationType: 'Word' }],
      },
    } as any);
    jest.spyOn(wordToPowerpointTool, 'handler').mockResolvedValue({
      success: true,
      data: { ok: true },
    } as any);

    const result = await adaptWordToPowerpointTool.handler({});

    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data.powerpointFilePath).toBe('C:\\docs\\active.pptx');
    }
  });

  test('returns an error when no active Word document exists', async () => {
    jest.spyOn(getActiveOfficeDocumentsTool, 'handler').mockResolvedValue({
      success: true,
      data: { documents: [] },
    } as any);

    const result = await adaptWordToPowerpointTool.handler({});

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('NO_WORD_DOC_ACTIVE');
    }
  });
});
