import { formatTablesInWorksheetTool } from '../../../src/tools/excel/formatTablesInWorksheet.tool';

describe('excel/formatTablesInWorksheet', () => {
  test('returns an exceljs error for a missing file', async () => {
    const result = await formatTablesInWorksheetTool[0].handler({
      inputExcelPath: 'C:/missing.xlsx',
      useComInterop: false,
    });

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('FORMAT_TABLES_EXCELJS_ERROR');
    }
  });
});
