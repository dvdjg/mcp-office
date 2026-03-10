import { jest } from '@jest/globals';
import { applyAutoTitlesTool } from '../../../src/tools/word/applyAutoTitles.tool.js';
import getActiveOfficeDocumentsTool from '../../../src/tools/os/getActiveOfficeDocuments.tool.js';

describe('word/applyAutoTitles', () => {
  const context = { log: console } as any;

  beforeEach(() => {
    jest.restoreAllMocks();
  });

  test('returns VALIDATION_ERROR when filePath is not a string', async () => {
    const result = await applyAutoTitlesTool.handler({ filePath: 123 as any }, context);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('VALIDATION_ERROR');
    }
  });

  test('returns NO_ACTIVE_WORD_DOC when there are no active Word documents', async () => {
    jest.spyOn(getActiveOfficeDocumentsTool, 'handler').mockResolvedValue({
      success: true,
      data: { documents: [] },
    } as any);

    const result = await applyAutoTitlesTool.handler({}, context);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('NO_ACTIVE_WORD_DOC');
    }
  });

  test('returns ACTIVE_DOC_FAILED when active document detection fails', async () => {
    jest.spyOn(getActiveOfficeDocumentsTool, 'handler').mockResolvedValue({
      success: false,
      error: { code: 'OS_ERROR', message: 'OS specific error' },
    } as any);

    const result = await applyAutoTitlesTool.handler({}, context);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('ACTIVE_DOC_FAILED');
    }
  });
});
