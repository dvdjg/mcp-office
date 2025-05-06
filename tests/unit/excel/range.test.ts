import excelRangeTool from '../../../src/tools/excel/range.tool';
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

describe('excelRangeTool Unit Tests (exceljs path)', () => {
  let mockWorkbook: ExcelJS.Workbook;
  let mockWorksheet: ExcelJS.Worksheet;
  let mockCell: ExcelJS.Cell;

  beforeEach(() => {
    jest.clearAllMocks();

    mockCell = {
      value: null,
      style: {},
      text: '',
      address: 'A1',
      // Add other properties and methods as needed by the tool
    } as unknown as ExcelJS.Cell; // Type assertion for simplicity

    mockWorksheet = {
      getCell: jest.fn().mockReturnValue(mockCell),
      name: 'Sheet1',
      // Add other properties and methods as needed
    } as unknown as ExcelJS.Worksheet;

    mockWorkbook = {
      xlsx: {
        readFile: jest.fn().mockResolvedValue(undefined),
        writeFile: jest.fn().mockResolvedValue(undefined),
      },
      getWorksheet: jest.fn().mockReturnValue(mockWorksheet),
      addWorksheet: jest.fn().mockReturnValue(mockWorksheet),
      worksheets: [mockWorksheet],
      // Add other properties and methods as needed
    } as unknown as ExcelJS.Workbook;

    (mockExcelJS.Workbook as jest.Mock).mockImplementation(() => mockWorkbook);
    // Correcting based on the assumption that the linter sees the problematic state for lines 54 & 55
    mockFs.pathExists.mockImplementation(async () => false);
    mockFs.readFile.mockImplementation(async () => Buffer.from('excelcontent'));
    (saveResource as jest.Mock).mockResolvedValue({} as any);
  });

  const baseParams: ToolRequestParams = {
    filePath: 'test.xlsx',
    rangeAddress: 'A1',
    operation: 'read', // Default operation
    useComInterop: false, // Explicitly test exceljs path
  };

  describe('File Handling', () => {
    it('should attempt to read an existing file for "read" operation', async () => {
      mockFs.pathExists.mockImplementation(async () => true);
      await excelRangeTool.handler({ ...baseParams, operation: 'read' });
      expect(mockFs.pathExists).toHaveBeenCalledWith(path.resolve('test.xlsx'));
      expect(mockWorkbook.xlsx.readFile).toHaveBeenCalledWith(path.resolve('test.xlsx'));
    });

    it('should return FILE_NOT_FOUND_EXCELJS if file does not exist for "read" operation', async () => {
      mockFs.pathExists.mockImplementation(async () => false);
      const result = await excelRangeTool.handler({ ...baseParams, operation: 'read' });
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('FILE_NOT_FOUND_EXCELJS');
      }
    });

    it('should create a new workbook if file does not exist for "write" operation', async () => {
      mockFs.pathExists.mockImplementation(async () => false);
      await excelRangeTool.handler({ ...baseParams, operation: 'write', values: [['test']] });
      expect(mockFs.pathExists).toHaveBeenCalledWith(path.resolve('test.xlsx'));
      expect(mockWorkbook.xlsx.readFile).not.toHaveBeenCalled(); // Should not try to read
      expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalledWith(path.resolve('test.xlsx'));
    });

    it('should use existing workbook if file exists for "write" operation', async () => {
      mockFs.pathExists.mockImplementation(async () => true);
      await excelRangeTool.handler({ ...baseParams, operation: 'write', values: [['test']] });
      expect(mockWorkbook.xlsx.readFile).toHaveBeenCalledWith(path.resolve('test.xlsx'));
      expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalledWith(path.resolve('test.xlsx'));
    });
  });

  describe('Sheet Handling (exceljs)', () => {
    it('should get worksheet by name if provided and exists', async () => {
      mockFs.pathExists.mockImplementation(async () => true);
      (mockWorkbook.getWorksheet as jest.Mock).mockReturnValueOnce(mockWorksheet);
      await excelRangeTool.handler({ ...baseParams, sheetName: 'Sheet1', operation: 'read' });
      expect(mockWorkbook.getWorksheet).toHaveBeenCalledWith('Sheet1');
    });

    it('should add worksheet by name if provided for "write" op and does not exist', async () => {
      mockFs.pathExists.mockImplementation(async () => true); // File exists
      (mockWorkbook.getWorksheet as jest.Mock).mockReturnValueOnce(undefined); // Sheet does not exist
      (mockWorkbook.addWorksheet as jest.Mock).mockReturnValueOnce(mockWorksheet);
      await excelRangeTool.handler({ ...baseParams, sheetName: 'NewSheet', operation: 'write', values: [['data']] });
      expect(mockWorkbook.addWorksheet).toHaveBeenCalledWith('NewSheet');
    });
    
    it('should return SHEET_NOT_FOUND_EXCELJS if sheetName provided for "read" op and does not exist', async () => {
        mockFs.pathExists.mockImplementation(async () => true);
        (mockWorkbook.getWorksheet as jest.Mock).mockReturnValueOnce(undefined); // Sheet does not exist
        const result = await excelRangeTool.handler({ ...baseParams, sheetName: 'NonExistentSheet', operation: 'read' });
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('SHEET_NOT_FOUND_EXCELJS');
        }
    });

    it('should use first worksheet if no sheetName/sheetIndex and worksheets exist', async () => {
        mockFs.pathExists.mockImplementation(async () => true);
        mockWorkbook.worksheets = [mockWorksheet, {} as ExcelJS.Worksheet]; // Ensure worksheets array is populated
        (mockWorkbook.getWorksheet as jest.Mock).mockImplementation(nameOrId => {
            if (typeof nameOrId === 'string') return mockWorkbook.worksheets.find(ws => ws.name === nameOrId);
            if (typeof nameOrId === 'number') return mockWorkbook.worksheets[nameOrId-1];
            return undefined;
        });

        await excelRangeTool.handler({ ...baseParams, operation: 'read' });
        expect(mockWorkbook.addWorksheet).not.toHaveBeenCalled();
        expect(mockWorksheet.getCell).toHaveBeenCalledWith('A1');
    });

    it('should add "Sheet1" if no sheetName/sheetIndex, no worksheets exist, and operation is "write"', async () => {
        mockFs.pathExists.mockImplementation(async () => false); // New file
        mockWorkbook.worksheets = []; // No sheets exist initially
        (mockWorkbook.addWorksheet as jest.Mock).mockReturnValueOnce(mockWorksheet);

        await excelRangeTool.handler({ ...baseParams, operation: 'write', values: [['data']] });
        expect(mockWorkbook.addWorksheet).toHaveBeenCalledWith('Sheet1');
    });
    
    it('should return SHEET_NOT_FOUND_EXCELJS for "read" if no sheetName/sheetIndex and no worksheets exist', async () => {
        mockFs.pathExists.mockImplementation(async () => true); // File exists
        mockWorkbook.worksheets = []; // No sheets in the file
        const result = await excelRangeTool.handler({ ...baseParams, operation: 'read' });
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('SHEET_NOT_FOUND_EXCELJS');
            expect(result.error.message).toContain('No worksheets found in the file');
        }
    });

    // TODO: Add tests for sheetIndex handling
  });

  describe('Read Operation (exceljs)', () => {
    it('should read value from the specified cell (top-left of rangeAddress)', async () => {
      mockFs.pathExists.mockImplementation(async () => true);
      mockCell.value = 'Test Value';
      const result = await excelRangeTool.handler({ ...baseParams, operation: 'read', rangeAddress: 'B5' });
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toBe('Test Value');
      }
      expect(mockWorksheet.getCell).toHaveBeenCalledWith('B5');
    });
  });

  describe('Write Operation (exceljs)', () => {
    it('should write value to the specified cell (top-left of rangeAddress)', async () => {
      mockFs.pathExists.mockImplementation(async () => false); // Create new file
      const values = [['Hello'], ['World']];
      const result = await excelRangeTool.handler({ ...baseParams, operation: 'write', values, rangeAddress: 'C3' });
      expect(result.success).toBe(true);
      expect(mockWorksheet.getCell).toHaveBeenCalledWith('C3');
      expect(mockCell.value).toBe('Hello');
      expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalledWith(path.resolve('test.xlsx'));
      expect(saveResource).toHaveBeenCalled();
    });

    it('should throw error if values are not provided for "write" operation', async () => {
        const result = await excelRangeTool.handler({ ...baseParams, operation: 'write', values: undefined });
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.message).toContain('Values are required for the "write" operation.');
        }
      });
  });
  
  describe('Format Operation (exceljs)', () => {
    it('should attempt to format the specified cell (top-left of rangeAddress)', async () => {
        mockFs.pathExists.mockImplementation(async () => true);
        const formatProps = { font: { bold: true } };
        const result = await excelRangeTool.handler({ ...baseParams, operation: 'format', formatProperties: formatProps, rangeAddress: 'D1' });
        
        expect(result.success).toBe(true);
        expect(mockWorksheet.getCell).toHaveBeenCalledWith('D1');
        expect(logger.warn).toHaveBeenCalledWith(expect.stringContaining("Format application is currently very basic"));
        expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalledWith(path.resolve('test.xlsx'));
        expect(saveResource).toHaveBeenCalled();
    });

    it('should throw error if formatProperties are not provided for "format" operation', async () => {
        const result = await excelRangeTool.handler({ ...baseParams, operation: 'format', formatProperties: undefined });
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.message).toContain('Format properties are required for the "format" operation.');
        }
    });
  });

  describe('Apply Operation (exceljs)', () => {
    it('should attempt to apply properties to the specified cell (top-left of rangeAddress)', async () => {
        mockFs.pathExists.mockImplementation(async () => true);
        const applyProps = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } } };
        const result = await excelRangeTool.handler({ ...baseParams, operation: 'apply', formatProperties: applyProps, rangeAddress: 'E2' });

        expect(result.success).toBe(true);
        expect(mockWorksheet.getCell).toHaveBeenCalledWith('E2');
        expect(logger.warn).toHaveBeenCalledWith(expect.stringContaining("Apply operation is currently very basic."));
        expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalledWith(path.resolve('test.xlsx'));
        expect(saveResource).toHaveBeenCalled();
    });
    
    it('should throw error if formatProperties are not provided for "apply" operation', async () => {
        const result = await excelRangeTool.handler({ ...baseParams, operation: 'apply', formatProperties: undefined });
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.message).toContain('Properties are required for the "apply" operation.');
        }
    });
  });

  describe('Unsupported Operation (exceljs)', () => {
    it('should return error for unsupported operation', async () => {
        const result = await excelRangeTool.handler({ ...baseParams, operation: 'unsupported_op' as any });
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.message).toContain('exceljs: Unsupported operation: "unsupported_op".');
        }
    });
  });

  describe('Dynamic Resource Saving (exceljs)', () => {
    it('should call saveResource after successful write operation', async () => {
        mockFs.pathExists.mockImplementation(async () => false);
        await excelRangeTool.handler({ ...baseParams, operation: 'write', values: [['data']] });
        expect(saveResource).toHaveBeenCalledWith(
            'excel/range',
            'test.xlsx',
            expect.any(Buffer)
        );
    });

    it('should include warning in message if saveResource fails', async () => {
        mockFs.pathExists.mockImplementation(async () => false);
        (saveResource as jest.Mock).mockRejectedValueOnce(new Error('Failed to save'));
        const result = await excelRangeTool.handler({ ...baseParams, operation: 'write', values: [['data']] });
        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.message).toContain('Warning: Failed to save as dynamic resource');
        }
    });
  });

});