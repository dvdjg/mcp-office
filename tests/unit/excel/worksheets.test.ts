import { excelWorksheetsTool } from '../../../src/tools/excel/worksheets.tool';
import { ToolRequestParams, ApiResponse } from '../../../src/types/common.types';
import * as fs from 'fs-extra';
import ExcelJS from 'exceljs';
import * as path from 'path';
import { saveResource } from '../../../src/tools/dynamic/resources.tool';

// Mock dependencies
jest.mock('fs-extra');
jest.mock('exceljs');
jest.mock('../../../src/tools/dynamic/resources.tool');
jest.mock('../../../src/utils/logger', () => ({
  info: jest.fn(),
  warn: jest.fn(),
  error: jest.fn(),
  debug: jest.fn(),
}));


const mockFs = fs as jest.Mocked<typeof fs>;
const mockExcelJS = ExcelJS as jest.Mocked<typeof ExcelJS>;
const mockSaveResource = saveResource as jest.MockedFunction<typeof saveResource>;

// The tool is an array, get the handler from the first element
const worksheetToolHandler = excelWorksheetsTool[0].handler;

describe('excelWorksheetsTool Unit Tests (exceljs path)', () => {
  let mockWorkbook: ExcelJS.Workbook;
  let mockSheet1: ExcelJS.Worksheet;
  let mockSheet2: ExcelJS.Worksheet;

  beforeEach(() => {
    jest.clearAllMocks();

    mockSheet1 = { name: 'Sheet1', id: 1, views: [] } as unknown as ExcelJS.Worksheet;
    mockSheet2 = { name: 'Sheet2', id: 2, views: [] } as unknown as ExcelJS.Worksheet;

    mockWorkbook = {
      xlsx: {
        readFile: jest.fn().mockResolvedValue(undefined),
        writeFile: jest.fn().mockResolvedValue(undefined),
      },
      getWorksheet: jest.fn(nameOrId => {
        if (nameOrId === 'Sheet1' || nameOrId === 1) return mockSheet1;
        if (nameOrId === 'Sheet2' || nameOrId === 2) return mockSheet2;
        return undefined;
      }),
      addWorksheet: jest.fn(name => ({ name, id: (mockWorkbook.worksheets.length + 1), views:[] } as unknown as ExcelJS.Worksheet)),
      removeWorksheet: jest.fn(),
      worksheets: [mockSheet1, mockSheet2],
      views: [],
    } as unknown as ExcelJS.Workbook;

    (mockExcelJS.Workbook as jest.Mock).mockImplementation(() => mockWorkbook);
    mockFs.pathExists.mockImplementation(async () => false); // Default: file does not exist
    mockFs.readFile.mockImplementation(async () => Buffer.from('excelcontent'));
    mockSaveResource.mockResolvedValue({} as any);
  });

  const baseParams: Omit<ToolRequestParams, 'operation'> = {
    filePath: 'test.xlsx',
    useComInterop: false, // Explicitly test exceljs path
  };

  describe('Add Operation (exceljs)', () => {
    it('should add a new sheet with a specified name', async () => {
      mockFs.pathExists.mockImplementation(async () => true); // File exists
      const newSheetName = 'NewSheet';
      const result = await worksheetToolHandler({ ...baseParams, operation: 'add', sheetName: newSheetName });
      expect(result.success).toBe(true);
      expect(mockWorkbook.addWorksheet).toHaveBeenCalledWith(newSheetName);
      if(result.success) expect(result.data.message).toContain(`Sheet '${newSheetName}' added successfully`);
      expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalledWith(path.resolve('test.xlsx'));
      expect(mockSaveResource).toHaveBeenCalled();
    });

    it('should add a new sheet with a default name if none specified', async () => {
      mockFs.pathExists.mockImplementation(async () => false); // New file
      mockWorkbook.worksheets = []; // Start with no sheets for default naming
      (mockWorkbook.addWorksheet as jest.Mock).mockImplementationOnce(name => ({ name, id:1, views:[] }));
      const result = await worksheetToolHandler({ ...baseParams, operation: 'add' });
      expect(result.success).toBe(true);
      expect(mockWorkbook.addWorksheet).toHaveBeenCalledWith('Sheet1'); // Default name
      if(result.success) expect(result.data.message).toContain(`Sheet 'Sheet1' added successfully`);
    });

    it('should handle adding a sheet to a new file', async () => {
        mockFs.pathExists.mockImplementation(async () => false); // File does not exist
        const newSheetName = 'FirstSheet';
        const result = await worksheetToolHandler({ ...baseParams, operation: 'add', sheetName: newSheetName });
        expect(result.success).toBe(true);
        expect(mockFs.pathExists).toHaveBeenCalledWith(path.resolve('test.xlsx'));
        expect(mockWorkbook.xlsx.readFile).not.toHaveBeenCalled();
        expect(mockWorkbook.addWorksheet).toHaveBeenCalledWith(newSheetName);
        expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalledWith(path.resolve('test.xlsx'));
      });
  });

  describe('Delete Operation (exceljs)', () => {
    it('should delete an existing sheet by name', async () => {
      mockFs.pathExists.mockImplementation(async () => true);
      const result = await worksheetToolHandler({ ...baseParams, operation: 'delete', sheetName: 'Sheet1' });
      expect(result.success).toBe(true);
      expect(mockWorkbook.removeWorksheet).toHaveBeenCalledWith(mockSheet1.id);
      if(result.success) expect(result.data.message).toContain("Sheet 'Sheet1' (identified by name 'Sheet1') deleted successfully");
      expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalled();
      expect(mockSaveResource).toHaveBeenCalled();
    });

    it('should delete an existing sheet by 1-based index', async () => {
        mockFs.pathExists.mockImplementation(async () => true);
        const result = await worksheetToolHandler({ ...baseParams, operation: 'delete', sheetIndex: 2 });
        expect(result.success).toBe(true);
        expect(mockWorkbook.removeWorksheet).toHaveBeenCalledWith(mockSheet2.id);
        if(result.success) expect(result.data.message).toContain("Sheet 'Sheet2' (identified by index 2) deleted successfully");
      });

    it('should return error if sheet to delete is not found by name', async () => {
      mockFs.pathExists.mockImplementation(async () => true);
      const result = await worksheetToolHandler({ ...baseParams, operation: 'delete', sheetName: 'NonExistentSheet' });
      expect(result.success).toBe(false);
      if(!result.success) expect(result.error.message).toContain("Sheet not found with name 'NonExistentSheet'");
    });

    it('should return error if sheet to delete is not found by index (out of bounds)', async () => {
        mockFs.pathExists.mockImplementation(async () => true);
        const result = await worksheetToolHandler({ ...baseParams, operation: 'delete', sheetIndex: 99 });
        expect(result.success).toBe(false);
        if(!result.success) expect(result.error.message).toContain("Sheet not found with name 'undefined' or index 99");
      });

    it('should require sheetName or sheetIndex for delete', async () => {
        const result = await worksheetToolHandler({ ...baseParams, operation: 'delete' });
        expect(result.success).toBe(false);
        if(!result.success) expect(result.error.message).toContain('sheetName or sheetIndex is required');
    });
  });

  describe('Rename Operation (exceljs)', () => {
    it('should rename an existing sheet by name', async () => {
      mockFs.pathExists.mockImplementation(async () => true);
      const oldName = 'Sheet1';
      const newName = 'RenamedSheet';
      const result = await worksheetToolHandler({ ...baseParams, operation: 'rename', sheetName: oldName, newSheetName: newName });
      expect(result.success).toBe(true);
      expect(mockSheet1.name).toBe(newName);
      if(result.success) expect(result.data.message).toContain(`Sheet '${oldName}' renamed to '${newName}'`);
      expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalled();
      expect(mockSaveResource).toHaveBeenCalled();
    });

    it('should rename an existing sheet by 1-based index', async () => {
        mockFs.pathExists.mockImplementation(async () => true);
        const oldName = mockSheet2.name; // 'Sheet2'
        const newName = 'IndexRenamedSheet';
        const result = await worksheetToolHandler({ ...baseParams, operation: 'rename', sheetIndex: 2, newSheetName: newName });
        expect(result.success).toBe(true);
        expect(mockSheet2.name).toBe(newName);
        if(result.success) expect(result.data.message).toContain(`Sheet '${oldName}' renamed to '${newName}'`);
      });

    it('should return error if sheet to rename is not found', async () => {
      mockFs.pathExists.mockImplementation(async () => true);
      const result = await worksheetToolHandler({ ...baseParams, operation: 'rename', sheetName: 'NonExistent', newSheetName: 'WontMatter' });
      expect(result.success).toBe(false);
      if(!result.success) expect(result.error.message).toContain("Sheet not found");
    });

    it('should require newSheetName for rename', async () => {
        const result = await worksheetToolHandler({ ...baseParams, operation: 'rename', sheetName: 'Sheet1' });
        expect(result.success).toBe(false);
        if(!result.success) expect(result.error.message).toContain('newSheetName is required');
    });
  });

  describe('Set Active Operation (exceljs)', () => {
    it('should set active sheet by name', async () => {
      mockFs.pathExists.mockImplementation(async () => true);
      const result = await worksheetToolHandler({ ...baseParams, operation: 'set', sheetName: 'Sheet2' });
      expect(result.success).toBe(true);
      expect(mockWorkbook.views).toEqual([{
        x: 0, y: 0, width: 10000, height: 20000,
        firstSheet: 0, activeTab: 1, visibility: 'visible' // Sheet2 is index 1
      }]);
      if(result.success) expect(result.data.message).toContain("Sheet 'Sheet2' (identified by name 'Sheet2') set as active");
      expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalled();
      // 'set' operation does not save as dynamic resource in the tool's current logic
      expect(mockSaveResource).not.toHaveBeenCalled();
    });

    it('should set active sheet by 1-based index', async () => {
        mockFs.pathExists.mockImplementation(async () => true);
        const result = await worksheetToolHandler({ ...baseParams, operation: 'set', sheetIndex: 1 }); // Sheet1
        expect(result.success).toBe(true);
        expect(mockWorkbook.views).toEqual([{
          x: 0, y: 0, width: 10000, height: 20000,
          firstSheet: 0, activeTab: 0, visibility: 'visible' // Sheet1 is index 0
        }]);
        if(result.success) expect(result.data.message).toContain("Sheet 'Sheet1' (identified by index 1) set as active");
      });

    it('should return error if sheet to set active is not found', async () => {
      mockFs.pathExists.mockImplementation(async () => true);
      const result = await worksheetToolHandler({ ...baseParams, operation: 'set', sheetName: 'NonExistent' });
      expect(result.success).toBe(false);
      if(!result.success) expect(result.error.message).toContain("Sheet not found");
    });
  });

  describe('File Existence and Error Handling', () => {
    it('should return error if file does not exist for "delete" operation', async () => {
      mockFs.pathExists.mockImplementation(async () => false);
      const result = await worksheetToolHandler({ ...baseParams, operation: 'delete', sheetName: 'AnySheet' });
      expect(result.success).toBe(false);
      if(!result.success) expect(result.error.message).toContain('file was not found');
    });

    it('should return error for unsupported operation', async () => {
        const result = await worksheetToolHandler({ ...baseParams, operation: 'unsupported_op' as any });
        expect(result.success).toBe(false);
        if(!result.success) expect(result.error.message).toContain('Unsupported operation: unsupported_op');
    });
  });
});