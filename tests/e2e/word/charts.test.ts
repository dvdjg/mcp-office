import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_charts_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_WORD_DOC_EMPTY = path.join(FIXTURES_DIR, 'empty_doc.docx'); // Assuming an empty Word doc fixture exists
const FIXTURE_EXCEL_DATA = path.join(FIXTURES_DIR, 'sample_chart_data.xlsx'); // Assuming an Excel file with chart data exists


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


describe('word/charts e2e tests', () => {

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

    // Check if fixture files exist
    const fixtures = [FIXTURE_WORD_DOC_EMPTY, FIXTURE_EXCEL_DATA];
    for (const fixturePath of fixtures) {
      try {
        await fs.stat(fixturePath);
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.error(`Fixture file ${fixturePath} not found. Please ensure it exists for word/charts tests.`);
           // Depending on test setup, might throw or skip tests
        } else {
          console.error(`Error checking fixture file ${fixturePath}:`, error);
        }
      }
    }
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

  test('should insert a chart into a Word document from Excel data (requires fixtures)', async () => {
    const outputPath = path.join(TEMP_DIR, 'doc_with_chart.docx');
    // Copy the empty fixture to the temp directory to avoid modifying the original
    await fs.copyFile(FIXTURE_WORD_DOC_EMPTY, outputPath);

    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/charts/insert', // Assuming the operation name is insert
        arguments: {
          document: 'tests/temp_word_charts_dir/doc_with_chart.docx', // Use path relative to workspace
          chartType: 'ColumnClustered', // Example chart type - need to confirm supported types
          data: {
            sourceType: 'excel',
            excelPath: 'tests/fixtures/sample_chart_data.xlsx', // Use path relative to workspace
            range: 'Sheet1!A1:B5', // Example data range in Excel
          },
          // Optional: position, title, size, etc.
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    // Verify the document was modified (basic check - could check modification time)
    await expect(fs.stat(outputPath)).resolves.toBeTruthy();
    // TODO: Add verification that the chart was actually inserted into the document
  });

  // TODO: Add tests for inserting charts with inline data
  // TODO: Add tests for other chart types and parameters
  // TODO: Add tests for error handling (e.g., non-existent document, invalid data source)
});