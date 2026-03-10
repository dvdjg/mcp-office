import { excelWorksheetsTool } from '../../../src/tools/excel/worksheets.tool';

describe('excel/worksheets', () => {
  test('returns zero worksheets for a missing file on exceljs count path', async () => {
    const result = await excelWorksheetsTool[0].handler({
      filePath: 'C:/definitely-missing.xlsx',
      operation: 'getWorksheetCount',
      useComInterop: false,
    });

    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data).toEqual({ worksheetCount: 0 });
    }
  });
});
