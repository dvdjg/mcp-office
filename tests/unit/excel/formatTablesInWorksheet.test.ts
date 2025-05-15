import { formatTablesInWorksheetTool } from '../../../src/tools/excel/formatTablesInWorksheet.tool';
import getActiveOfficeDocumentsTool from '../../../src/tools/os/getActiveOfficeDocuments.tool';
import excelRangeTool from '../../../src/tools/excel/range.tool';
import { aiSuggestTool } from '../../../src/tools/office/aiSuggest.tool';
import { FastMCPContext, ApiResponse } from '../../../src/types/common.types'; // Added ApiResponse
import ExcelJS from 'exceljs';
import * as fs from 'fs-extra';
import * as path from 'path'; // Import path

// Mocks
jest.mock('../../../src/tools/os/getActiveOfficeDocuments.tool');
jest.mock('../../../src/tools/excel/range.tool');
jest.mock('../../../src/tools/office/aiSuggest.tool');
jest.mock('fs-extra');
jest.mock('exceljs');

const mockGetActiveOfficeDocumentsTool = getActiveOfficeDocumentsTool as jest.Mocked<typeof getActiveOfficeDocumentsTool>;
const mockExcelRangeTool = excelRangeTool as jest.Mocked<typeof excelRangeTool>;
const mockAiSuggestTool = aiSuggestTool as jest.Mocked<typeof aiSuggestTool>;
const mockFs = fs as jest.Mocked<typeof fs>;

const mockExcelJSWorkbook = {
  xlsx: {
    readFile: jest.fn().mockResolvedValue(undefined),
    writeFile: jest.fn().mockResolvedValue(undefined),
  },
  getWorksheet: jest.fn(),
  addWorksheet: jest.fn(),
  worksheets: [],
} as unknown as ExcelJS.Workbook;

// Mock ExcelJS constructor
(ExcelJS.Workbook as jest.Mock).mockImplementation(() => mockExcelJSWorkbook);


describe('formatTablesInWorksheetTool', () => {
  const tool = formatTablesInWorksheetTool[0]; // Assuming it's an array with one tool
  let mockContext: FastMCPContext<any>;

  beforeEach(() => {
    jest.clearAllMocks();
    mockContext = {
      session: {
        requestSampling: jest.fn().mockResolvedValue({ success: true, content: [{ type: 'text', text: 'AI Suggestion' }] }),
      },
    } as unknown as FastMCPContext<any>;

    // Resetting workbook mock for each test
    mockExcelJSWorkbook.worksheets = [];
    (mockExcelJSWorkbook.getWorksheet as jest.Mock).mockImplementation((idOrName: string | number) => {
        if (typeof idOrName === 'string') {
            return (mockExcelJSWorkbook.worksheets as any[]).find(ws => ws.name === idOrName);
        }
        return (mockExcelJSWorkbook.worksheets as any[])[idOrName -1];
    });
    (mockExcelJSWorkbook.addWorksheet as jest.Mock).mockImplementation((name: string) => {
        const newWs = { name, getTables: jest.fn().mockReturnValue([]), getRow: jest.fn().mockReturnValue({ getCell: jest.fn().mockReturnValue({value: ''}) }) };
        (mockExcelJSWorkbook.worksheets as any[]).push(newWs);
        return newWs;
    });
  });

  describe('Input Validation', () => {
    // inputExcelPath is optional, so direct validation failure for non-existence won't occur at schema parse.
    // Behavior for non-existent files is tested in other sections.

    it('should use default worksheetIdentifier (2) if not provided', async () => {
      mockGetActiveOfficeDocumentsTool.handler = jest.fn().mockResolvedValue({
        success: true,
        data: { documents: [{ filePath: 'C:\\path\\to\\active.xlsx', applicationType: 'Excel' }] },
      });
      (mockFs.pathExists as jest.Mock).mockImplementation(async () => true); // For ExcelJS path
      
      // Mock worksheet and table identification for ExcelJS
      const mockSheet2 = { name: 'Sheet2', getTables: jest.fn().mockReturnValue([]), getRow: jest.fn().mockReturnValue({ getCell: jest.fn().mockReturnValue({value: ''}) }) };
      (mockExcelJSWorkbook.worksheets as any[]) = [
        { name: 'Sheet1', getTables: jest.fn().mockReturnValue([]), getRow: jest.fn().mockReturnValue({ getCell: jest.fn().mockReturnValue({value: ''}) }) },
        mockSheet2,
      ];
      (mockExcelJSWorkbook.getWorksheet as jest.Mock).mockImplementation((id: number | string) => {
        if (id === 2) return mockSheet2;
        if (id === 'Sheet2') return mockSheet2;
        return (mockExcelJSWorkbook.worksheets as any[]).find(ws => ws.name === id || ws.id === id);
      });


      await tool.handler({ useComInterop: false }, mockContext); // No inputExcelPath, no worksheetIdentifier
      expect(mockExcelJSWorkbook.getWorksheet).toHaveBeenCalledWith(2); // Default is 2
    });
  });

  describe('Active Document Detection (useComInterop: false)', () => {
    beforeEach(() => {
        (mockFs.pathExists as jest.Mock).mockImplementation(async () => true); // Assume file exists for these tests
         // Setup a default worksheet for ExcelJS path to avoid errors in later stages
        const defaultWs = { name: 'DefaultSheet', getTables: jest.fn().mockReturnValue([]), getRow: jest.fn().mockReturnValue({ getCell: jest.fn().mockReturnValue({value: ''}) }) };
        (mockExcelJSWorkbook.worksheets as any[]) = [defaultWs];
        (mockExcelJSWorkbook.getWorksheet as jest.Mock).mockReturnValue(defaultWs);
    });

    it('should use inputExcelPath if provided', async () => {
      const params = { inputExcelPath: 'provided/path.xlsx', useComInterop: false };
      await tool.handler(params, mockContext);
      expect(mockGetActiveOfficeDocumentsTool.handler).not.toHaveBeenCalled();
      expect(mockExcelJSWorkbook.xlsx.readFile).toHaveBeenCalledWith(expect.stringContaining('provided' + path.sep + 'path.xlsx'));
    });

    it('should call getActiveOfficeDocumentsTool if inputExcelPath is not provided', async () => {
      mockGetActiveOfficeDocumentsTool.handler = jest.fn().mockResolvedValue({
        success: true,
        data: { documents: [{ filePath: 'C:\\active\\doc.xlsx', applicationType: 'Excel' }] },
      });
      await tool.handler({ useComInterop: false }, mockContext);
      expect(mockGetActiveOfficeDocumentsTool.handler).toHaveBeenCalledTimes(1);
      expect(mockExcelJSWorkbook.xlsx.readFile).toHaveBeenCalledWith(expect.stringContaining('C:' + path.sep + 'active' + path.sep + 'doc.xlsx'));
    });

    it('should throw error if no active Excel document is found', async () => {
      mockGetActiveOfficeDocumentsTool.handler = jest.fn().mockResolvedValue({
        success: true,
        data: { documents: [] }, // No documents
      });
      const result = await tool.handler({ useComInterop: false }, mockContext) as ApiResponse<any>;
      expect(result.success).toBe(false);
      if (!result.success) { // Type guard
        expect(result.error?.message).toContain('No active Excel document found');
      }
    });

    it('should throw error if multiple active Excel documents are found', async () => {
      mockGetActiveOfficeDocumentsTool.handler = jest.fn().mockResolvedValue({
        success: true,
        data: {
          documents: [
            { filePath: 'C:\\doc1.xlsx', applicationType: 'Excel' },
            { filePath: 'C:\\doc2.xlsx', applicationType: 'Excel' },
          ],
        },
      });
      const result = await tool.handler({ useComInterop: false }, mockContext) as ApiResponse<any>;
      expect(result.success).toBe(false);
      if (!result.success) { // Type guard
        expect(result.error?.message).toContain('Multiple active Excel documents found');
      }
    });

     it('should throw error if getActiveOfficeDocumentsTool fails', async () => {
      mockGetActiveOfficeDocumentsTool.handler = jest.fn().mockResolvedValue({
        success: false,
        error: { code: 'OS_ERROR', message: 'Failed to query OS' },
      });
      const result = await tool.handler({ useComInterop: false }, mockContext) as ApiResponse<any>;
      expect(result.success).toBe(false);
      if (!result.success) { // Type guard
        expect(result.error?.message).toContain('Failed to get active office documents: Failed to query OS');
      }
    });
  });

  // TODO: Add tests for worksheet identification (COM and ExcelJS)
  // TODO: Add tests for table identification (COM and ExcelJS)
  // TODO: Add tests for AI suggestion and application flow (COM and ExcelJS)
  //    - Mocking excelRangeTool.handler and aiSuggestTool.handler
  //    - Verifying calls and handling of their responses
  //    - Mocking COM table formatting calls
  //    - Mocking ExcelJS table style application

  describe('ExcelJS Path - Table Processing', () => {
    const filePath = 'test.xlsx';
    beforeEach(() => {
        (mockFs.pathExists as jest.Mock).mockImplementation(async () => true);
        mockGetActiveOfficeDocumentsTool.handler = jest.fn().mockResolvedValue({
            success: true,
            data: { documents: [{ filePath, applicationType: 'Excel' }] },
        });
    });

    it('should report no tables found if worksheet has no tables (ExcelJS)', async () => {
        const mockSheet = { name: 'Sheet1', getTables: jest.fn().mockReturnValue([]), getRow: jest.fn().mockReturnValue({ getCell: jest.fn().mockReturnValue({value: ''}) }) };
        (mockExcelJSWorkbook.worksheets as any[]) = [mockSheet];
        (mockExcelJSWorkbook.getWorksheet as jest.Mock).mockReturnValue(mockSheet);

        const result = await tool.handler({ inputExcelPath: filePath, worksheetIdentifier: 'Sheet1', useComInterop: false }, mockContext) as ApiResponse<any>;
        expect(result.success).toBe(true);
        if (result.success) { // Type guard
            expect(result.data?.message).toContain("No tables found in worksheet 'Sheet1'");
        }
        expect(mockSheet.getTables).toHaveBeenCalled();
    });

    it('should process tables, get AI suggestions, and attempt to apply them (ExcelJS)', async () => {
        const mockTable1 = { 
            name: 'SalesTable', 
            ref: 'A1:C5', 
            headerRow: true, 
            columns: [{name: 'Product'}, {name: 'Region'}, {name: 'Sales'}],
            style: {}
        };
        const mockSheetData = {
            name: 'DataSheet',
            getTables: jest.fn().mockReturnValue([[mockTable1, undefined]]), // Simulating the tuple structure
            getRow: jest.fn((rowIndex: number) => {
                if (rowIndex === 1) return { getCell: (colIndex: number) => ({ value: mockTable1.columns[colIndex-1].name }) }; // Header
                return { getCell: (colIndex: number) => ({ value: `Data${rowIndex}${colIndex}` }) }; // Data
            }),
        };
        (mockExcelJSWorkbook.worksheets as any[]) = [mockSheetData];
        (mockExcelJSWorkbook.getWorksheet as jest.Mock).mockReturnValue(mockSheetData);
        
        mockAiSuggestTool.handler = jest.fn().mockResolvedValue({
            success: true,
            data: "Apply table style 'TableStyleMedium9'\nEnable filter buttons"
        });

        const result = await tool.handler({ inputExcelPath: filePath, worksheetIdentifier: 'DataSheet', useComInterop: false }, mockContext) as ApiResponse<any>;
        
        expect(result.success).toBe(true);
        if (result.success) { // Type guard
            expect(result.data?.message).toContain("ExcelJS: Table formatting process completed.");
            expect(result.data?.results[0].tableName).toBe('SalesTable');
            expect(result.data?.results[0].appliedFormats).toContain("Applied table style: TableStyleMedium9");
            expect(result.data?.results[0].appliedFormats).toContain("Enabled filter buttons on header row");
        }
        expect(mockAiSuggestTool.handler).toHaveBeenCalled();
        expect(mockExcelJSWorkbook.xlsx.writeFile).toHaveBeenCalledWith(expect.stringContaining(filePath)); // Check save
    });
  });
  
  // Note: COM path tests are more complex due to the need to mock COM objects extensively.
  // They would follow a similar structure, mocking the return values of COM calls.
});