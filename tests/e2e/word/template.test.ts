import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_template_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_DOC_WITH_PLACEHOLDERS = path.join(FIXTURES_DIR, 'doc_with_placeholders.docx'); // Assuming this fixture exists with placeholders


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


describe('word/template e2e tests', () => {

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

    // Check if fixture file exists
    try {
      await fs.stat(FIXTURE_DOC_WITH_PLACEHOLDERS);
    } catch (error: any) {
      if (error.code === 'ENOENT') {
         console.error(`Fixture file ${FIXTURE_DOC_WITH_PLACEHOLDERS} not found. Please ensure it exists for word/template tests.`);
         // Depending on test setup, might throw or skip tests
      } else {
        console.error(`Error checking fixture file ${FIXTURE_DOC_WITH_PLACEHOLDERS}:`, error);
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

  test('should fill placeholders in a Word document template (requires fixture)', async () => {
    const outputPath = path.join(TEMP_DIR, 'filled_template.docx');
    // Copy the fixture to the temp directory to avoid modifying the original
    await fs.copyFile(FIXTURE_DOC_WITH_PLACEHOLDERS, outputPath);

    const placeholders = {
      '{{name}}': 'John Doe',
      '{{date}}': '2023-10-27',
    }; // Example placeholders and values - need to match fixture

    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/template/fill', // Assuming the operation name is fill
        arguments: {
          document: 'tests/temp_word_template_dir/filled_template.docx', // Use path relative to workspace
          placeholders: placeholders,
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    // Verify the document was modified (basic check - could check modification time)
    await expect(fs.stat(outputPath)).resolves.toBeTruthy();
    // TODO: Add verification that the placeholders were actually filled in the document
  });

  // TODO: Add tests for different types of placeholders, missing placeholders, extra placeholders
  // TODO: Add tests for error handling (e.g., non-existent document)
});