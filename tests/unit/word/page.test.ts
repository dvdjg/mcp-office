import { wordPageTool } from '../../../src/tools/word/page.tool';

describe('word/page', () => {
  test('rejects getPageCount when extra fields are provided', async () => {
    const result = await wordPageTool.handler({
      operation: 'getPageCount',
      filePath: 'C:/test.docx',
      sectionIndex: 1,
    });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('VALIDATION_ERROR');
    }
  });

  test('requires filePath', async () => {
    const result = await wordPageTool.handler({
      operation: 'getPageCount',
    } as any);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('VALIDATION_ERROR');
    }
  });
});
