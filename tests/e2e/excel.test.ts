import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_excel_dir');
const TEST_XLSX_PATH = path.join(TEMP_DIR, 'test_data.xlsx');
const FIXTURE_XLSX_PATH = path.join(__dirname, '../fixtures/sample_excel_data.xlsx'); // Assuming a fixture file exists

// Define a basic type for the expected successful response
interface SuccessResponse {
  success: true;
  data: any; // Use a more specific type if the data structure is known
  message?: string;
}

// Define a basic type for the expected error response
interface ErrorResponse {
  success: false;
  error: {
    code: string;
    message: string;
    details?: any;
  };
}

type ToolResponse = SuccessResponse | ErrorResponse;


describe('excel e2e tests', () => {

  beforeAll(async () => {
    // Ensure temp directory does not exist before tests
    try {
      await fs.rm(TEMP_DIR, { recursive: true });
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        console.error(`Error cleaning up temp directory ${TEMP_DIR}:`, error);
      }
    }
    // Create temp directory
    await fs.mkdir(TEMP_DIR, { recursive: true });
    // Note: Creating a valid .xlsx file programmatically for tests is complex.
    // For now, we will rely on the 'write' test to create a file,
    // or assume a fixture file exists for 'read' tests.
    // A more robust approach would involve creating a test fixture .xlsx file here.
    // For tests that require a pre-existing file, we will use FIXTURE_XLSX_PATH.
  });

  afterAll(async () => {
    // Clean up temp directory after tests
    try {
      await fs.rm(TEMP_DIR, { recursive: true });
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        console.error(`Error cleaning up temp directory ${TEMP_DIR}:`, error);
      }
    }
  });

  describe('excel/range', () => {
    test('should write data to a new excel file', async () => {
      const dataToWrite = [['Header1', 'Header2'], [1, 2], [3, 4]];
      const outputPath = path.join(TEMP_DIR, 'write_test.xlsx');

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'excel/range/write',
          arguments: {
            document: 'tests/temp_excel_dir/write_test.xlsx', // Use path relative to workspace
            range: 'A1', // Start writing from A1
            values: dataToWrite,
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the file was created (basic check)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();

      // TODO: Add verification of the content written to the Excel file
      // This would require reading the Excel file content, which might need another tool or library.
    });

    test('should read data from an excel file (requires fixture)', async () => {
      // Assuming 'tests/fixtures/sample_excel_data.xlsx' exists with data in Sheet1!A1:B2
      const sampleExcelPath = 'tests/fixtures/sample_excel_data.xlsx';

      // Check if fixture exists
      try {
        await fs.stat(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'));
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.warn(`Fixture file ${sampleExcelPath} not found. Skipping read test.`);
           return; // Skip the test if fixture is missing
        }
        throw error;
      }

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'excel/range/read',
          arguments: {
            document: sampleExcelPath, // Use path relative to workspace
            range: 'Sheet1!A1:B2', // Specify the range to read
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('values');
        expect(Array.isArray(result.data.values)).toBe(true);
        // Basic check for expected data structure - adjust based on fixture content
        // expect(result.data.values.length).toBeGreaterThan(0);
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    test('should format a range in an excel file (requires fixture)', async () => {
      // Assuming 'tests/fixtures/sample_excel_data.xlsx' exists
      const sampleExcelPath = 'tests/fixtures/sample_excel_data.xlsx';
      const outputPath = path.join(TEMP_DIR, 'format_test.xlsx');

      // Check if fixture exists
      try {
        await fs.stat(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'));
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.warn(`Fixture file ${sampleExcelPath} not found. Skipping format test.`);
           return; // Skip the test if fixture is missing
        }
        throw error;
      }

      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'), outputPath);


      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'excel/range/format',
          arguments: {
            document: 'tests/temp_excel_dir/format_test.xlsx', // Use path relative to workspace
            range: 'Sheet1!A1:A1', // Specify the range to format
            format: { bold: true, color: '#FF0000' }, // Example format
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // TODO: Add verification that the formatting was applied
      // This would likely require reading the formatting information from the Excel file,
      // which might need a more advanced tool or library.
    });

    test('should apply a style to a range in an excel file (requires fixture)', async () => {
      // Assuming 'tests/fixtures/sample_excel_data.xlsx' exists
      const sampleExcelPath = 'tests/fixtures/sample_excel_data.xlsx';
      const outputPath = path.join(TEMP_DIR, 'apply_style_test.xlsx');

      // Check if fixture exists
      try {
        await fs.stat(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'));
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.warn(`Fixture file ${sampleExcelPath} not found. Skipping apply style test.`);
           return; // Skip the test if fixture is missing
        }
        throw error;
      }

      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'), outputPath);

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'excel/range/apply',
          arguments: {
            document: 'tests/temp_excel_dir/apply_style_test.xlsx', // Use path relative to workspace
            range: 'Sheet1!A1:B2', // Specify the range
            style: 'Good', // Example built-in style
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // TODO: Add verification that the style was applied
      // Similar to formatting, this requires reading style information.
    });
  });

  describe('excel/data-analysis', () => {
    test('should filter data based on criteria (requires fixture)', async () => {
      // Assuming 'tests/fixtures/data_analysis_sample.xlsx' exists with data
      const dataAnalysisExcelPath = 'tests/fixtures/data_analysis_sample.xlsx'; // Replace with actual fixture path
      const outputPath = path.join(TEMP_DIR, 'filtered_data.xlsx');


       // Check if fixture exists
       try {
        await fs.stat(path.join(__dirname, '../fixtures/data_analysis_sample.xlsx'));
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.warn(`Fixture file ${dataAnalysisExcelPath} not found. Skipping data-analysis filter test.`);
           return; // Skip the test if fixture is missing
        }
        throw error;
      }

      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(path.join(__dirname, '../fixtures/data_analysis_sample.xlsx'), outputPath);


      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'excel/data-analysis',
          arguments: {
            document: 'tests/temp_excel_dir/filtered_data.xlsx', // Use path relative to workspace
            sheetName: 'Sheet1', // Specify the sheet
            rangeAddress: 'A1:C10', // Specify the range
            operation: 'filter',
            filterCriteria: [{ column: 'Value', criteria1: 100, operator: 'xlGreater' }], // Example criteria
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('filteredValues');
        expect(Array.isArray(result.data.filteredValues)).toBe(true);
        // TODO: Add more specific checks based on expected filtered data
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    // TODO: Add tests for other data-analysis operations and variations
  });

  describe('excel/worksheets', () => {
    test('should list worksheets (requires fixture)', async () => {
      const sampleExcelPath = 'tests/fixtures/sample_excel_data.xlsx';

      // Check if fixture exists
      try {
        await fs.stat(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'));
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.warn(`Fixture file ${sampleExcelPath} not found. Skipping list worksheets test.`);
           return; // Skip the test if fixture is missing
        }
        throw error;
      }

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'excel/worksheets/list',
          arguments: {
            document: sampleExcelPath, // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('sheetNames');
        expect(Array.isArray(result.data.sheetNames)).toBe(true);
        expect(result.data.sheetNames).toContain('Sheet1'); // Assuming Sheet1 exists in the fixture
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    test('should add a new worksheet (requires fixture)', async () => {
      const sampleExcelPath = 'tests/fixtures/sample_excel_data.xlsx';
      const outputPath = path.join(TEMP_DIR, 'add_sheet_test.xlsx');

      // Check if fixture exists
      try {
        await fs.stat(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'));
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.warn(`Fixture file ${sampleExcelPath} not found. Skipping add worksheet test.`);
           return; // Skip the test if fixture is missing
        }
        throw error;
      }

      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'), outputPath);

      const newSheetName = 'NewTestSheet';

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'excel/worksheets/add',
          arguments: {
            document: 'tests/temp_excel_dir/add_sheet_test.xlsx', // Use path relative to workspace
            sheetName: newSheetName,
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // TODO: Verify the new sheet was added. This might require listing sheets again
      // or reading the file content in a way that shows sheet names.
    });

    test('should delete a worksheet (requires fixture)', async () => {
      const sampleExcelPath = 'tests/fixtures/sample_excel_data.xlsx';
      const outputPath = path.join(TEMP_DIR, 'delete_sheet_test.xlsx');

      // Check if fixture exists
      try {
        await fs.stat(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'));
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.warn(`Fixture file ${sampleExcelPath} not found. Skipping delete worksheet test.`);
           return; // Skip the test if fixture is missing
        }
        throw error;
      }

      // Copy the fixture to the temp directory and add a sheet to delete
      await fs.copyFile(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'), outputPath);
      // Manually add a sheet to the copied file for deletion test (requires a tool call or manual step)
      // For now, let's assume a sheet named 'SheetToDelete' exists in the fixture or is added by a prior step.
      // A more robust test would add the sheet using the 'add' tool before deleting.
      const sheetToDeleteName = 'SheetToDelete'; // Assuming this sheet exists or is added

      // Note: To make this test reliable, we should add 'SheetToDelete' first using the 'add' tool.
      // This makes the tests dependent, which is not ideal for unit/integration but acceptable for e2e.
      // Alternatively, ensure the fixture has a sheet specifically for deletion tests.

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'excel/worksheets/delete',
          arguments: {
            document: 'tests/temp_excel_dir/delete_sheet_test.xlsx', // Use path relative to workspace
            sheetName: sheetToDeleteName,
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // TODO: Verify the sheet was deleted. This might require listing sheets again.
    });
  });

  describe('excel/charts', () => {
    test('should insert a chart (requires fixture)', async () => {
      const sampleExcelPath = 'tests/fixtures/sample_excel_data.xlsx';
      const outputPath = path.join(TEMP_DIR, 'insert_chart_test.xlsx');

      // Check if fixture exists
      try {
        await fs.stat(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'));
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.warn(`Fixture file ${sampleExcelPath} not found. Skipping insert chart test.`);
           return; // Skip the test if fixture is missing
        }
        throw error;
      }

      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'), outputPath);

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'excel/charts/insert',
          arguments: {
            document: 'tests/temp_excel_dir/insert_chart_test.xlsx', // Use path relative to workspace
            sheetName: 'Sheet1', // Specify the sheet
            rangeAddress: 'A1:B5', // Specify the data range for the chart
            chartType: 'ColumnClustered', // Example chart type
            // Optional: position, title, etc.
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // TODO: Verify the chart was inserted. This might require inspecting the Excel file structure.
    });
  });

  describe('excel/tables', () => {
    test('should insert a table (requires fixture)', async () => {
      const sampleExcelPath = 'tests/fixtures/sample_excel_data.xlsx';
      const outputPath = path.join(TEMP_DIR, 'insert_table_test.xlsx');

      // Check if fixture exists
      try {
        await fs.stat(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'));
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.warn(`Fixture file ${sampleExcelPath} not found. Skipping insert table test.`);
           return; // Skip the test if fixture is missing
        }
        throw error;
      }

      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(path.join(__dirname, '../fixtures/sample_excel_data.xlsx'), outputPath);

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'excel/tables/insert',
          arguments: {
            document: 'tests/temp_excel_dir/insert_table_test.xlsx', // Use path relative to workspace
            sheetName: 'Sheet1', // Specify the sheet
            rangeAddress: 'A1:C5', // Specify the data range for the table
            hasHeaders: true, // Assuming the range includes headers
            // Optional: tableName, tableStyle, etc.
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // TODO: Verify the table was inserted. This might require inspecting the Excel file structure.
    });
  });

  // TODO: Add tests for complex variations and error handling for all Excel tools
});