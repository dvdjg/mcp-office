import { excelDataAnalysisTool } from '../../../src/tools/excel/dataAnalysis.tool';
import { ToolRequestParams, ApiResponse } from '../../../src/types/common.types';
import * as fs from 'fs-extra';
import ExcelJS from 'exceljs';
import * as path from 'path';
import logger from '../../../src/utils/logger';
// saveResource is not directly used by dataAnalysis tool, so not mocked here unless a specific test needs it.

// Mock dependencies
jest.mock('fs-extra');
jest.mock('exceljs');
jest.mock('../../../src/utils/logger');

const mockFs = fs as jest.Mocked<typeof fs>;
const mockExcelJS = ExcelJS as jest.Mocked<typeof ExcelJS>;
const mockLogger = logger as jest.Mocked<typeof logger>;

// The tool is an array, get the handler from the first element
const dataAnalysisToolHandler = excelDataAnalysisTool[0].handler;

describe('excelDataAnalysisTool Unit Tests (exceljs path)', () => {
  let mockWorkbook: ExcelJS.Workbook;
  let mockWorksheet: ExcelJS.Worksheet;
  let mockTable: ExcelJS.Table;

  beforeEach(() => {
    jest.clearAllMocks();

    mockTable = {
        name: 'TestTable',
        ref: 'A1:C5',
        // ... other properties if needed
    } as unknown as ExcelJS.Table;

    mockWorksheet = {
      name: 'Sheet1',
      autoFilter: undefined, // To check if it gets set
      getTable: jest.fn().mockReturnValue(mockTable),
      // Add other necessary worksheet properties if used by the tool
    } as unknown as ExcelJS.Worksheet;

    mockWorkbook = {
      xlsx: {
        readFile: jest.fn().mockResolvedValue(undefined),
        writeFile: jest.fn().mockResolvedValue(undefined),
      },
      getWorksheet: jest.fn().mockReturnValue(mockWorksheet),
      addWorksheet: jest.fn().mockReturnValue(mockWorksheet), // Though not primary for this tool
      worksheets: [mockWorksheet],
    } as unknown as ExcelJS.Workbook;

    (mockExcelJS.Workbook as jest.Mock).mockImplementation(() => mockWorkbook);
    mockFs.pathExists.mockImplementation(async () => true); // Default: file exists
    mockFs.readFile.mockImplementation(async () => Buffer.from('excelcontent'));
  });

  const baseParams: Omit<ToolRequestParams, 'operation'> = {
    filePath: 'test-data-analysis.xlsx',
    useComInterop: false,
    sheetName: 'Sheet1',
    // rangeAddress or tableName will be added per test
  };

  describe('Sort Operation (exceljs)', () => {
    it('should log warning and indicate sort is not supported', async () => {
      const params = { ...baseParams, operation: 'sort', rangeAddress: 'A1:B10', sortCriteria: [{ column: 'A', order: 'Ascending' }] };
      const result = await dataAnalysisToolHandler(params);
      expect(result.success).toBe(true); // Tool "succeeds" by logging
      expect(mockLogger.warn).toHaveBeenCalledWith(expect.stringContaining('Sorting data within Excel files is not directly supported'));
      if (result.success) {
        expect(result.data).toContain('Sorting data within Excel files is not directly supported');
      }
      // writeFile should not be called if no action is taken
      expect(mockWorkbook.xlsx.writeFile).not.toHaveBeenCalled();
    });
  });

  describe('Filter Operation (exceljs)', () => {
    it('should enable AutoFilter on a given rangeAddress', async () => {
      const params = { ...baseParams, operation: 'filter', rangeAddress: 'A1:D20', filterCriteria: [{ column: 'A', criteria1: 'Test' }] };
      const result = await dataAnalysisToolHandler(params);
      expect(result.success).toBe(true);
      expect(mockWorksheet.autoFilter).toBe('A1:D20');
      if (result.success) {
        expect(result.data).toContain('AutoFilter enabled on range "A1:D20"');
        expect(result.data).toContain('Applying specific criteria programmatically is not supported');
      }
      expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalled();
    });

    it('should enable AutoFilter on a range derived from tableName', async () => {
        const params = { ...baseParams, operation: 'filter', tableName: 'TestTable', filterCriteria: [{ column: 'A', criteria1: 'Test' }] };
        const result = await dataAnalysisToolHandler(params);
        expect(result.success).toBe(true);
        expect(mockWorksheet.getTable).toHaveBeenCalledWith('TestTable');
        expect(mockWorksheet.autoFilter).toBe(mockTable.ref); // 'A1:C5'
        if (result.success) {
          expect(result.data).toContain(`AutoFilter enabled on range "${mockTable.ref}"`);
        }
        expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalled();
      });

    it('should return error if table not found for filter by tableName', async () => {
        (mockWorksheet.getTable as jest.Mock).mockReturnValue(undefined);
        const params = { ...baseParams, operation: 'filter', tableName: 'GhostTable', filterCriteria: [{ column: 'A', criteria1: 'Test' }] };
        const result = await dataAnalysisToolHandler(params);
        expect(result.success).toBe(false);
        if(!result.success) expect(result.error.message).toContain('exceljs: Table "GhostTable" not found');
    });

    it('should return error if neither rangeAddress nor tableName provided for filter', async () => {
        // Omitting rangeAddress and tableName from baseParams for this test
        const params = { filePath: 'test.xlsx', useComInterop: false, sheetName: 'Sheet1', operation: 'filter', filterCriteria: [{ column: 'A', criteria1: 'Test' }] };
        const result = await dataAnalysisToolHandler(params);
        expect(result.success).toBe(false);
        if(!result.success) expect(result.error.message).toContain('Either rangeAddress or tableName must be provided.');
    });
  });

  describe('Pivot Operation (exceljs)', () => {
    it('should log warning and indicate pivot tables are not supported', async () => {
      const params = { ...baseParams, operation: 'pivot', rangeAddress: 'A1:F100', pivotTableParameters: { pivotTableName: 'TestPivot', destination: 'Sheet2!A1' } };
      const result = await dataAnalysisToolHandler(params);
      expect(result.success).toBe(true);
      expect(mockLogger.warn).toHaveBeenCalledWith(expect.stringContaining('Creating or manipulating Pivot Tables is not supported'));
      if (result.success) {
        expect(result.data).toContain('Creating or manipulating Pivot Tables is not supported');
      }
      expect(mockWorkbook.xlsx.writeFile).not.toHaveBeenCalled(); // No file write if no action
    });
  });

  describe('Calculate Operation (exceljs)', () => {
    it('should process workbook and indicate formulas are calculated on load/save', async () => {
      // Ensure baseParams is spread here to include filePath
      const params = { ...baseParams, operation: 'calculate', formulaRange: 'A1:A10' };
      const result = await dataAnalysisToolHandler(params);
      expect(result.success).toBe(true);
      expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalledWith(path.resolve(params.filePath));
      if (result.success) {
        expect(result.data).toContain('Formulas are generally calculated upon file load/save');
      }
    });

    it('should handle calculate when file did not exist (no write should occur if no file to read/calc)', async () => {
        mockFs.pathExists.mockImplementation(async () => false);
        const params = { ...baseParams, operation: 'calculate', formulaRange: 'A1:A10' };
        const result = await dataAnalysisToolHandler(params);
        expect(result.success).toBe(true);
        expect(mockWorkbook.xlsx.writeFile).not.toHaveBeenCalled();
        if (result.success) {
          expect(result.data).toContain(`Workbook at "${params.filePath}" processed.`);
        }
      });
  });

  describe('Error Handling and Edge Cases (exceljs)', () => {
    it('should return error if file not found for operations other than pivot (which might create)', async () => {
      mockFs.pathExists.mockImplementation(async () => false);
      const params = { ...baseParams, operation: 'sort', rangeAddress: 'A1:B2', sortCriteria: [{column:'A', order:'Ascending'}] };
      const result = await dataAnalysisToolHandler(params);
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.message).toContain('exceljs: File not found');
      }
    });

    it('should return error if worksheet not found (and not pivot op)', async () => {
      (mockWorkbook.getWorksheet as jest.Mock).mockReturnValue(undefined);
      mockWorkbook.worksheets = []; // Ensure no default sheet can be picked
      const params = { ...baseParams, operation: 'filter', rangeAddress: 'A1:B2', sheetName: 'NonExistentSheet', filterCriteria: [{column:'A', criteria1:'Val'}] };
      const result = await dataAnalysisToolHandler(params);
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.message).toContain('exceljs: Worksheet "NonExistentSheet" not found.');
      }
    });
  });
});