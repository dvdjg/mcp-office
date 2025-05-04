import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_tables_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_WORD_DOC_EMPTY = path.join(FIXTURES_DIR, 'empty_doc.docx'); // Assuming an empty Word doc fixture exists
const FIXTURE_DOC_WITH_TABLE = path.join(FIXTURES_DIR, 'doc_with_table.docx'); // Assuming a Word doc fixture with a table exists


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


describe('word/tables e2e tests', () => {

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
    const fixtures = [FIXTURE_WORD_DOC_EMPTY, FIXTURE_DOC_WITH_TABLE];
    for (const fixturePath of fixtures) {
      try {
        await fs.stat(fixturePath);
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.error(`Fixture file ${fixturePath} not found. Please ensure it exists for word/tables tests.`);
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

  describe('word/tables/insert', () => {
    test('should insert a table into a Word document (requires fixture)', async () => {
      const outputPath = path.join(TEMP_DIR, 'doc_with_table.docx');
      // Copy the empty fixture to the temp directory to avoid modifying the original
      await fs.copyFile(FIXTURE_WORD_DOC_EMPTY, outputPath);

      const tableData = [['Header 1', 'Header 2'], ['Data 1', 'Data 2']];

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/tables/insert',
          arguments: {
            document: 'tests/temp_word_tables_dir/doc_with_table.docx', // Use path relative to workspace
            data: tableData,
            // Optional: position, style, etc. - need to confirm from API spec
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the document was modified (basic check - could check modification time)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification that the table was actually inserted (requires inspecting doc content/structure)
    });

    // TODO: Add tests for inserting tables with different data, positions, and styles
    // TODO: Add tests for error handling (e.g., non-existent document, invalid data format)
  });

  describe('word/tables/read', () => {
    test('should read data from a table in a Word document (requires fixture)', async () => {
      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/tables/read',
          arguments: {
            document: 'tests/fixtures/doc_with_table.docx', // Use path relative to workspace
            tableIndex: 1, // Assuming the first table
            // Optional: range within the table
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('values');
        expect(Array.isArray(result.data.values)).toBe(true);
        // TODO: Add more specific checks based on the expected table content in the fixture
        // expect(result.data.values.length).toBeGreaterThan(0);
        // expect(result.data.values[0][0]).toBe('Expected Header');
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    // TODO: Add tests for reading specific tables by index, reading ranges within tables
    // TODO: Add tests for error handling (e.g., non-existent document, invalid table index)
  });

  // TODO: Add tests for update and delete table operations
});