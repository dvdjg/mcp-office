import { excelWorksheetsTool } from '../../../src/tools/excel/worksheets.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import * as fs from 'fs-extra';
import ExcelJS from 'exceljs';
import path from 'path';
import logger from '../../../src/utils/logger';

// Mock dependencies
jest.mock('../../../src/utils/officeInterop', () => ({
  getOfficeApplication: jest.fn(),
  releaseObject: jest.fn(),
}));

jest.mock('fs-extra', () => ({
  pathExists: jest.fn(),
  readFile: jest.fn(),
}));

const mockExcelJSWorkbook = {
  xlsx: {
    readFile: jest.fn(),
    writeFile: jest.fn(),
  },
  addWorksheet: jest.fn(),
  getWorksheet: jest.fn(),
  removeWorksheet: jest.fn(),
  worksheets: [],
  views: [],
};
jest.mock('exceljs', () => {
  return {
    Workbook: jest.fn(() => mockExcelJSWorkbook),
  };
});

jest.mock('../../../src/tools/dynamic/resources.tool', () => ({
  saveResource: jest.fn().mockResolvedValue({ success: true, data: { message: "Resource saved" } }),
}));

jest.mock('../../../src/utils/logger', () => ({
  info: jest.fn(),
  warn: jest.fn(),
  error: jest.fn(),
  debug: jest.fn(),
}));

// Define top-level mock functions for COM methods
const mockExcelOpen = jest.fn();
const mockExcelAdd = jest.fn();

// Define the structure of the application mock (defined once)
const mockExcelApplicationObject = {
  Workbooks: {
    Open: mockExcelOpen,
    Add: mockExcelAdd,
  },
  DisplayAlerts: true,
  Visible: false,
};

// Define the structure of the office instance mock (defined once)
const mockOfficeAppInstanceObject = {
  app: mockExcelApplicationObject,
  release: jest.fn(),
};

// Define TypeScript types for the mock objects created in beforeEach
type MockExcelSheetsState = {
  Item: jest.Mock<any, any>;
  Add: jest.Mock<any, any>;
  Count: number;
  Delete: jest.Mock<any, any>;
};

type MockExcelWorkbookState = {
  Sheets: MockExcelSheetsState;
  Save: jest.Mock<any, any>;
  SaveAs: jest.Mock<any, any>;
  Close: jest.Mock<any, any>;
};

describe('excelWorksheetsTool - getWorksheetCount', () => {
  let currentExcelSheetsState: MockExcelSheetsState;
  let currentExcelWorkbookState: MockExcelWorkbookState;

  beforeEach(() => {
    jest.clearAllMocks();

    currentExcelSheetsState = {
      Item: jest.fn(),
      Add: jest.fn(),
      Count: 0,
      Delete: jest.fn(),
    };
    Object.defineProperty(currentExcelSheetsState, 'Count', {
        value: 0,
        writable: true,
        configurable: true,
    });

    currentExcelWorkbookState = {
      Sheets: currentExcelSheetsState,
      Save: jest.fn(),
      SaveAs: jest.fn(),
      Close: jest.fn(),
    };

    mockExcelOpen.mockResolvedValue(currentExcelWorkbookState);
    mockExcelAdd.mockResolvedValue(currentExcelWorkbookState);
    mockOfficeAppInstanceObject.release.mockReset();
    
    (getOfficeApplication as jest.Mock).mockResolvedValue(mockOfficeAppInstanceObject);
    (fs.pathExists as jest.Mock).mockResolvedValue(true);
    (mockExcelJSWorkbook.xlsx.readFile as jest.Mock).mockResolvedValue(undefined);
  });

  // --- COM Interop Tests ---
  describe('COM Interop Path (useComInterop: true)', () => {
    it('should successfully get worksheet count using COM', async () => {
      const mockFilePath = 'C:/test/workbook.xlsx';
      const expectedSheetCount = 3;
      currentExcelSheetsState.Count = expectedSheetCount;

      const params = {
        filePath: mockFilePath,
        operation: 'getWorksheetCount',
        useComInterop: true,
      };

      const result = await excelWorksheetsTool[0].handler(params as any);

      expect(getOfficeApplication).toHaveBeenCalledWith('Excel.Application');
      expect(mockExcelOpen).toHaveBeenCalledWith(path.resolve(mockFilePath));
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toEqual({ worksheetCount: expectedSheetCount });
      }
      expect(currentExcelWorkbookState.Close).toHaveBeenCalledWith(false);
      expect(mockOfficeAppInstanceObject.release).toHaveBeenCalled();
    });

    it('should return an error if COM Sheets.Count access fails after Open', async () => {
      const mockFilePath = 'C:/test/workbook.xlsx';
      const comError = new Error('COM Error during Sheets.Count');
      mockExcelOpen.mockResolvedValue(currentExcelWorkbookState); // Open succeeds
      Object.defineProperty(currentExcelSheetsState, 'Count', {
        get: jest.fn(() => { throw comError; }),
        configurable: true
      });

      const params = {
        filePath: mockFilePath,
        operation: 'getWorksheetCount',
        useComInterop: true,
      };

      const result = await excelWorksheetsTool[0].handler(params as any);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error?.code).toBe('EXCEL_WORKSHEETS_COM_ERROR');
        expect(result.error?.message).toBe(comError.message);
      }
      expect(mockExcelOpen).toHaveBeenCalledWith(path.resolve(mockFilePath));
      expect(mockOfficeAppInstanceObject.release).toHaveBeenCalled();
    });

    it('should return "File not found" error if COM Open fails for getWorksheetCount', async () => {
      const mockFilePath = 'C:/test/non_existent_workbook.xlsx';
      const openError = new Error('Simulated COM Open failure');
      mockExcelOpen.mockRejectedValue(openError);

      const params = {
        filePath: mockFilePath,
        operation: 'getWorksheetCount',
        useComInterop: true,
      };
      const result = await excelWorksheetsTool[0].handler(params as any);
      expect(result.success).toBe(false);
      if(!result.success){
        expect(result.error?.code).toBe('EXCEL_WORKSHEETS_COM_ERROR');
        expect(result.error?.message).toBe(`File not found: ${path.resolve(mockFilePath)}`);
      }
      expect(mockOfficeAppInstanceObject.release).toHaveBeenCalled();
    });
  });

  // --- ExcelJS Tests ---
  describe('ExcelJS Path (useComInterop: false)', () => {
    it('should successfully get worksheet count using exceljs', async () => {
      const mockFilePath = 'C:/test/workbook.xlsx';
      const expectedSheetCount = 2;
      mockExcelJSWorkbook.worksheets = new Array(expectedSheetCount) as any;

      const params = {
        filePath: mockFilePath,
        operation: 'getWorksheetCount',
        useComInterop: false,
      };

      const result = await excelWorksheetsTool[0].handler(params as any);

      expect(ExcelJS.Workbook).toHaveBeenCalled();
      expect(mockExcelJSWorkbook.xlsx.readFile).toHaveBeenCalledWith(path.resolve(mockFilePath));
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toEqual({ worksheetCount: expectedSheetCount });
      }
    });

    it('should return an error if exceljs fails to read file for count', async () => {
      const mockFilePath = 'C:/test/workbook.xlsx';
      const exceljsError = new Error('ExcelJS error reading file');
      (mockExcelJSWorkbook.xlsx.readFile as jest.Mock).mockRejectedValue(exceljsError);

      const params = {
        filePath: mockFilePath,
        operation: 'getWorksheetCount',
        useComInterop: false,
      };

      const result = await excelWorksheetsTool[0].handler(params as any);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error?.code).toBe('EXCEL_WORKSHEETS_EXCELJS_ERROR');
        expect(result.error?.message).toBe(exceljsError.message);
      }
    });

    it('should return worksheet count as 0 if file does not exist (exceljs path)', async () => {
        const mockFilePath = 'C:/test/non_existent_workbook.xlsx';
        (fs.pathExists as jest.Mock).mockResolvedValue(false);
        mockExcelJSWorkbook.worksheets = [] as any;

        const params = {
            filePath: mockFilePath,
            operation: 'getWorksheetCount',
            useComInterop: false,
        };

        const result = await excelWorksheetsTool[0].handler(params as any);

        expect(ExcelJS.Workbook).toHaveBeenCalled();
        expect(fs.pathExists).toHaveBeenCalledWith(path.resolve(mockFilePath));
        expect(mockExcelJSWorkbook.xlsx.readFile).not.toHaveBeenCalled();
        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data).toEqual({ worksheetCount: 0 });
        }
    });
  });

  it('should return a validation error if filePath is missing for getWorksheetCount', async () => {
    const params = {
      operation: 'getWorksheetCount',
    };
    const result = await excelWorksheetsTool[0].handler(params as any);
    expect(result.success).toBe(false);
    if (!result.success) {
        expect(result.error?.code).toBe('VALIDATION_ERROR');
        expect(result.error?.message).toContain("Input validation failed");
        expect(JSON.stringify(result.error?.details)).toContain("filePath");
    }
  });
});