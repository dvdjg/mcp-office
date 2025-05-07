import excelChartsTool from '../../../src/tools/excel/charts.tool';
import { ToolRequestParams, ApiResponse } from '../../../src/types/common.types';
import * as fs from 'fs-extra';
import ExcelJS from 'exceljs';
import * as path from 'path';
import logger from '../../../src/utils/logger';
import { saveResource } from '../../../src/tools/dynamic/resources.tool';

// Mock dependencies
jest.mock('fs-extra');
jest.mock('exceljs');
jest.mock('../../../src/utils/logger');
jest.mock('../../../src/tools/dynamic/resources.tool');

const mockFs = fs as jest.Mocked<typeof fs>;
const mockExcelJS = ExcelJS as jest.Mocked<typeof ExcelJS>;
const mockLogger = logger as jest.Mocked<typeof logger>;
const mockSaveResource = saveResource as jest.MockedFunction<typeof saveResource>;

describe('excelChartsTool Unit Tests (exceljs path)', () => {
  let mockWorkbook: ExcelJS.Workbook;
  let mockWorksheet: any; // Use 'any' for worksheet to mock 'addChart'

  beforeEach(() => {
    jest.clearAllMocks();

    mockWorksheet = {
      name: 'Sheet1',
      addChart: jest.fn(), // Mock addChart specifically
      // Add other necessary worksheet properties if used by the tool
    };

    mockWorkbook = {
      xlsx: {
        readFile: jest.fn().mockResolvedValue(undefined),
        writeFile: jest.fn().mockResolvedValue(undefined),
      },
      getWorksheet: jest.fn().mockReturnValue(mockWorksheet) as jest.MockedFunction<typeof ExcelJS.Workbook.prototype.getWorksheet>,
      addWorksheet: jest.fn().mockReturnValue(mockWorksheet),
      worksheets: [mockWorksheet as ExcelJS.Worksheet], // Cast for the array
    } as unknown as ExcelJS.Workbook;

    (mockExcelJS.Workbook as jest.Mock).mockImplementation(() => mockWorkbook);
    mockFs.pathExists.mockImplementation(async () => true); // Default: file exists
    mockFs.readFile.mockImplementation(async () => Buffer.from('excelcontent'));
    mockSaveResource.mockResolvedValue({} as any);
  });

  const baseParams: Omit<ToolRequestParams, 'operation'> = {
    filePath: 'test-charts.xlsx',
    useComInterop: false, // Explicitly test exceljs path
    sheetName: 'Sheet1',
  };

  describe('Insert Operation (exceljs)', () => {
    it('should attempt to insert a chart with required parameters', async () => {
      const params = {
        ...baseParams,
        filePath: 'test-charts-insert.xlsx', // Ensure filePath is part of params
        operation: 'insert',
        rangeAddress: 'A1:B5',
        chartType: 'xlColumnClustered',
        chartTitle: 'My Test Chart',
      };
      const result = await excelChartsTool.handler(params);

      expect(result.success).toBe(true);
      expect(mockWorksheet.addChart).toHaveBeenCalled();
      expect(mockLogger.warn).toHaveBeenCalledWith(expect.stringContaining('Chart data series configuration from rangeAddress "A1:B5" is using a placeholder'));
      if (result.success) {
        expect(result.data).toContain('exceljs: Chart of type "bar" inserted.'); // bar is mapping for xlColumnClustered
      }
      expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalledWith(path.resolve(params.filePath));
      expect(mockSaveResource).toHaveBeenCalled();
    });

    it('should return error if rangeAddress is missing for insert', async () => {
      const params = { ...baseParams, filePath: 'test-charts-no-range.xlsx', operation: 'insert', chartType: 'xlLine' };
      const result = await excelChartsTool.handler(params);
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.message).toContain('exceljs: rangeAddress and chartType are required for insert.');
      }
    });

    it('should return error if chartType is missing for insert', async () => {
      const params = { ...baseParams, filePath: 'test-charts-no-type.xlsx', operation: 'insert', rangeAddress: 'A1:B5' };
      const result = await excelChartsTool.handler(params);
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.message).toContain('exceljs: rangeAddress and chartType are required for insert.');
      }
    });
    
    it('should return error for unsupported chartType in exceljs', async () => {
        const params = { ...baseParams, filePath: 'test-charts-unsupported.xlsx', operation: 'insert', rangeAddress: 'A1:B5', chartType: 'xlUnsupportedChartTypeForExcelJS' };
        const result = await excelChartsTool.handler(params);
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.message).toContain('exceljs: Unsupported chartType "xlUnsupportedChartTypeForExcelJS"');
        }
      });

    it('should create worksheet if it does not exist for insert on a new file', async () => {
        mockFs.pathExists.mockImplementation(async () => false); // File does not exist
        (mockWorkbook.getWorksheet as jest.Mock).mockReturnValue(undefined); // Sheet does not exist
        mockWorkbook.worksheets = []; // No sheets initially

        const params = {
            ...baseParams,
            filePath: 'new-chart-file.xlsx',
            sheetName: 'AutoCreatedSheet',
            operation: 'insert',
            rangeAddress: 'C1:D10',
            chartType: 'xlPie',
        };
        await excelChartsTool.handler(params);
        expect(mockWorkbook.addWorksheet).toHaveBeenCalledWith('AutoCreatedSheet');
        expect(mockWorksheet.addChart).toHaveBeenCalled();
    });
  });

  describe('Limited Support Operations (exceljs)', () => {
    const limitedSupportOps: ('modify' | 'delete' | 'reposition' | 'list')[] = ['modify', 'delete', 'reposition', 'list'];

    limitedSupportOps.forEach(op => {
      it(`should log warning and indicate limited support for "${op}" operation`, async () => {
        const params = { ...baseParams, filePath: `test-charts-${op}.xlsx`, operation: op, chartIndex: 1 };
        const result = await excelChartsTool.handler(params);
        
        expect(result.success).toBe(true);
        expect(mockLogger.warn).toHaveBeenCalledWith(expect.stringContaining(`Operation '${op}' for charts has significant limitations or is not supported with exceljs`));
        if (result.success) {
            if (op === 'list') {
                expect(result.data).toEqual([]);
            } else {
                expect(result.data).toContain(`Operation '${op}' for charts has significant limitations`);
            }
        }
        expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalledWith(path.resolve(params.filePath));
        if (op !== 'list') {
            expect(mockSaveResource).toHaveBeenCalled();
        }
      });
    });
  });

  describe('File and Sheet Error Handling (exceljs)', () => {
    it('should return error if file not found for non-insert operation', async () => {
      mockFs.pathExists.mockImplementation(async () => false);
      const params = { ...baseParams, filePath: 'test-charts-list-no-file.xlsx', operation: 'list' };
      const result = await excelChartsTool.handler(params);
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.message).toContain(`exceljs: File not found at ${path.resolve(params.filePath)} for operation 'list'.`);
      }
    });

    it('should return error if worksheet not found', async () => {
      (mockWorkbook.getWorksheet as jest.Mock).mockReturnValue(undefined); // Sheet does not exist
      mockWorkbook.worksheets = []; // No sheets to default to
      const params = { ...baseParams, filePath: 'test-charts-no-sheet.xlsx', operation: 'insert', rangeAddress: 'A1:B2', chartType: 'xlLine', sheetName: 'NonExistentSheet' };
      const result = await excelChartsTool.handler(params);
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.message).toContain('exceljs: Worksheet "NonExistentSheet" not found or could not be created.');
      }
    });
  });

  describe('Unsupported Operation (exceljs)', () => {
    it('should return error for truly unsupported operation string', async () => {
        const params = { ...baseParams, filePath: 'test-charts-unsupported-op.xlsx', operation: 'unsupported_op_string' as any };
        const result = await excelChartsTool.handler(params);
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('VALIDATION_ERROR');
            expect(result.error.message).toContain('Input validation failed');
        }
    });
  });
});