import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_code_format_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_DOC_WITH_CODE = path.join(FIXTURES_DIR, 'doc_with_code.docx'); // Assuming this fixture exists with code snippets


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


describe('word/code-format e2e tests', () => {

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
      await fs.stat(FIXTURE_DOC_WITH_CODE);
    } catch (error: any) {
      if (error.code === 'ENOENT') {
         console.error(`Fixture file ${FIXTURE_DOC_WITH_CODE} not found. Please ensure it exists for word/code-format tests.`);
         // Depending on test setup, might throw or skip tests
      } else {
        console.error(`Error checking fixture file ${FIXTURE_DOC_WITH_CODE}:`, error);
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

  test('should format code snippets in a Word document (requires fixture)', async () => {
    const outputPath = path.join(TEMP_DIR, 'formatted_code_doc.docx');
    // Copy the fixture to the temp directory to avoid modifying the original
    await fs.copyFile(FIXTURE_DOC_WITH_CODE, outputPath);

    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/code-format',
        arguments: {
          filePath: 'tests/temp_word_code_format_dir/formatted_code_doc.docx', // Use path relative to workspace
          // Optional: style, font, language (if not auto-detected)
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    // Verify the document was modified (basic check - could check modification time)
    await expect(fs.stat(outputPath)).resolves.toBeTruthy();
    // TODO: Add verification that the code was actually formatted (requires inspecting doc content/structure)
  });

  // TODO: Add tests for specifying style, font, and language
  // TODO: Add tests for error handling (e.g., non-existent document, document with no code)
});