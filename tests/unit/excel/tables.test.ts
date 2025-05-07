import excelTablesTool from '../../../src/tools/excel/tables.tool';
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

describe('excelTablesTool Unit Tests (exceljs path)', () => {
  let mockWorkbook: ExcelJS.Workbook;
  let mockWorksheet: ExcelJS.Worksheet;
  let mockTable: ExcelJS.Table;

  beforeEach(() => {
    jest.clearAllMocks();

    mockTable = {
      name: 'TestTable',
      ref: 'A1:C5',
      headerRow: true,
      totalsRow: false,
      style: { theme: 'TableStyleMedium9', showRowStripes: true },
      columns: [{ name: 'Col1', filterButton: true }, { name: 'Col2', filterButton: true }],
      rows: [['a', 1], ['b', 2]],
      addRow: jest.fn(),
      removeRows: jest.fn(),
      removeColumns: jest.fn(),
      // Add other table properties/methods if used by the tool
    } as unknown as ExcelJS.Table;

    mockWorksheet = {
      name: 'Sheet1',
      addTable: jest.fn().mockReturnValue(mockTable),
      getTable: jest.fn().mockReturnValue(mockTable),
      removeTable: jest.fn(),
      // Add other necessary worksheet properties if used by the tool
    } as unknown as ExcelJS.Worksheet;

    mockWorkbook = {
      xlsx: {
        readFile: jest.fn().mockResolvedValue(undefined),
        writeFile: jest.fn().mockResolvedValue(undefined),
      },
      getWorksheet: jest.fn().mockReturnValue(mockWorksheet),
      addWorksheet: jest.fn().mockReturnValue(mockWorksheet),
      worksheets: [mockWorksheet],
    } as unknown as ExcelJS.Workbook;

    (mockExcelJS.Workbook as jest.Mock).mockImplementation(() => mockWorkbook);
    mockFs.pathExists.mockImplementation(async () => true); // Default: file exists
    mockFs.readFile.mockImplementation(async () => Buffer.from('excelcontent'));
    mockSaveResource.mockResolvedValue({} as any);
  });

  const baseParams: Omit<ToolRequestParams, 'operation'> = {
    filePath: 'test-tables.xlsx',
    useComInterop: false,
    sheetName: 'Sheet1',
  };

  describe('Insert Operation (exceljs)', () => {
    it('should insert a table with given rangeAddress and tableName', async () => {
      const params = {
        ...baseParams,
        operation: 'insert',
        rangeAddress: 'A1:D10',
        tableName: 'NewSalesTable',
        data: [['Header1', 'Header2'], ['Val1', 'Val2']] // data for columns/rows
      };
      const result = await excelTablesTool.handler(params);
      expect(result.success).toBe(true);
      expect(mockWorksheet.addTable).toHaveBeenCalledWith(expect.objectContaining({
        name: 'NewSalesTable',
        ref: 'A1:D10',
        columns: [{ name: 'Header1', filterButton: true }, { name: 'Header2', filterButton: true }],
        rows: [['Val1', 'Val2']],
      }));
      if (result.success) {
        expect(result.data).toContain('exceljs: Table inserted in range A1:D10. Name: NewSalesTable');
      }
      expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalled();
      expect(mockSaveResource).toHaveBeenCalled();
    });

    it('should insert a table with auto-generated name if tableName is not provided', async () => {
        const params = {
          ...baseParams,
          operation: 'insert',
          rangeAddress: 'B2:E5',
          data: [['ID'], [1]]
        };
        Date.now = jest.fn(() => 1234567890123); // Mock Date.now for predictable name
        const result = await excelTablesTool.handler(params);
        expect(result.success).toBe(true);
        expect(mockWorksheet.addTable).toHaveBeenCalledWith(expect.objectContaining({
          name: 'Table1234567890123',
          ref: 'B2:E5',
        }));
        if (result.success) {
            expect(result.data).toContain('exceljs: Table inserted in range B2:E5. Name: Table1234567890123');
          }
      });

    it('should require rangeAddress for insert operation', async () => {
      const params = { ...baseParams, operation: 'insert', tableName: 'NoRangeTable' };
      const result = await excelTablesTool.handler(params);
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.message).toContain('exceljs: rangeAddress is required for "insert".');
      }
    });
  });

  describe('Add Rows/Columns Operation (exceljs)', () => {
    it('should add rows to an existing table', async () => {
      const params = {
        ...baseParams,
        operation: 'add',
        tableName: 'TestTable',
        location: 'rows',
        data: [['c', 3], ['d', 4]],
      };
      const result = await excelTablesTool.handler(params);
      expect(result.success).toBe(true);
      expect(mockTable.addRow).toHaveBeenCalledTimes(2);
      expect(mockTable.addRow).toHaveBeenCalledWith(['c', 3]);
      expect(mockTable.addRow).toHaveBeenCalledWith(['d', 4]);
      if (result.success) {
        expect(result.data).toContain('exceljs: 2 row(s) added to table "TestTable".');
      }
    });

    it('should log warning for adding columns (limitation)', async () => {
      const params = {
        ...baseParams,
        operation: 'add',
        tableName: 'TestTable',
        location: 'columns',
        count: 2,
      };
      const result = await excelTablesTool.handler(params);
      expect(result.success).toBe(true); // Still "succeeds" by logging warning
      expect(mockLogger.warn).toHaveBeenCalledWith(expect.stringContaining('Adding columns to an existing table is complex and not directly supported'));
      if (result.success) {
        expect(result.data).toContain('Operation for table "TestTable" noted as a limitation.');
      }
    });

    it('should require data for adding rows', async () => {
        const params = { ...baseParams, operation: 'add', tableName: 'TestTable', location: 'rows', data: [] };
        const result = await excelTablesTool.handler(params);
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.message).toContain('exceljs: data is required for adding rows.');
        }
      });
  });

  describe('Delete Rows/Columns/Table Operation (exceljs)', () => {
    it('should delete rows from a table', async () => {
      mockTable.rows = [['a',1], ['b',2], ['c',3], ['d',4]]; // Ensure enough rows for deletion
      const params = {
        ...baseParams,
        operation: 'delete',
        tableName: 'TestTable',
        location: 'rows',
        count: 2,
        position: 1, // Delete first 2 rows
      };
      const result = await excelTablesTool.handler(params);
      expect(result.success).toBe(true);
      expect(mockTable.removeRows).toHaveBeenCalledWith(0, 2); // 0-indexed
      if (result.success) {
        expect(result.data).toContain('exceljs: 2 row(s) deleted from table "TestTable".');
      }
    });

    it('should delete columns from a table', async () => {
        mockTable.columns = [{name:'A'}, {name:'B'}, {name:'C'}];
        const params = {
          ...baseParams,
          operation: 'delete',
          tableName: 'TestTable',
          location: 'columns',
          count: 1,
          position: 2, // Delete 2nd column
        };
        const result = await excelTablesTool.handler(params);
        expect(result.success).toBe(true);
        expect(mockTable.removeColumns).toHaveBeenCalledWith(1, 1); // 0-indexed
        if (result.success) {
          expect(result.data).toContain('exceljs: 1 column(s) deleted from table "TestTable".');
        }
      });

    it('should delete the entire table if location is not rows/columns', async () => {
      const params = { ...baseParams, operation: 'delete', tableName: 'TestTable' };
      const result = await excelTablesTool.handler(params);
      expect(result.success).toBe(true);
      expect(mockWorksheet.removeTable).toHaveBeenCalledWith('TestTable');
      if (result.success) {
        expect(result.data).toContain('exceljs: Table "TestTable" deleted.');
      }
    });

    it('should return error if table not found for delete rows/columns', async () => {
        (mockWorksheet.getTable as jest.Mock).mockReturnValue(undefined);
        const params = { ...baseParams, operation: 'delete', tableName: 'GhostTable', location: 'rows', count:1 };
        const result = await excelTablesTool.handler(params);
        expect(result.success).toBe(false);
        if(!result.success) expect(result.error.message).toContain('exceljs: Table "GhostTable" not found.');
    });
  });
  
  describe('Modify Operation (exceljs)', () => {
    it('should indicate modification capabilities for existing table', async () => {
        const params = { ...baseParams, operation: 'modify', tableName: 'TestTable' };
        const result = await excelTablesTool.handler(params);
        expect(result.success).toBe(true);
        expect(mockWorksheet.getTable).toHaveBeenCalledWith('TestTable');
        if(result.success) {
            expect(result.data).toContain('exceljs: Table "TestTable" found. Modification capabilities are specific');
        }
        expect(mockWorkbook.xlsx.writeFile).toHaveBeenCalled();
        expect(mockSaveResource).toHaveBeenCalled();
    });

    it('should return error if table to modify is not found', async () => {
        (mockWorksheet.getTable as jest.Mock).mockReturnValue(undefined);
        const params = { ...baseParams, operation: 'modify', tableName: 'NonExistentTable' };
        const result = await excelTablesTool.handler(params);
        expect(result.success).toBe(false);
        if(!result.success) {
            expect(result.error.message).toContain('exceljs: Table "NonExistentTable" not found.');
        }
    });
  });

  describe('Error Handling and Edge Cases (exceljs)', () => {
    it('should return error if file not found and operation is not insert', async () => {
        mockFs.pathExists.mockImplementation(async () => false);
        const params = { ...baseParams, operation: 'modify', tableName: 'AnyTable' };
        const result = await excelTablesTool.handler(params);
        expect(result.success).toBe(false);
        if(!result.success) expect(result.error.message).toContain('Worksheet "Sheet1" could not be found or created');
    });

    it('should create sheet if not found during insert', async () => {
        mockFs.pathExists.mockImplementation(async () => true); // File exists
        (mockWorkbook.getWorksheet as jest.Mock).mockReturnValue(undefined); // Sheet does not
        const params = { ...baseParams, sheetName: 'NewSheetForTable', operation: 'insert', rangeAddress: 'A1:B2' };
        await excelTablesTool.handler(params);
        expect(mockWorkbook.addWorksheet).toHaveBeenCalledWith('NewSheetForTable');
        expect(mockWorksheet.addTable).toHaveBeenCalled(); // addTable is on the returned sheet
    });
  });
});