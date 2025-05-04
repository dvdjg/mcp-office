import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_reformat_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_POORLY_FORMATTED_DOC = path.join(FIXTURES_DIR, 'poorly_formatted.docx'); // Assuming this fixture exists


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


describe('word/reformat e2e tests', () => {

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
      await fs.stat(FIXTURE_POORLY_FORMATTED_DOC);
    } catch (error: any) {
      if (error.code === 'ENOENT') {
         console.error(`Fixture file ${FIXTURE_POORLY_FORMATTED_DOC} not found. Please ensure it exists for word/reformat tests.`);
         // Depending on test setup, might throw or skip tests
      } else {
        console.error(`Error checking fixture file ${FIXTURE_POORLY_FORMATTED_DOC}:`, error);
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

  test('should reformat a poorly formatted Word document (requires fixture)', async () => {
    const outputPath = path.join(TEMP_DIR, 'reformatted_doc.docx');
    // Copy the fixture to the temp directory to avoid modifying the original
    await fs.copyFile(FIXTURE_POORLY_FORMATTED_DOC, outputPath);

    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/reformat',
        arguments: {
          document: 'tests/temp_word_reformat_dir/reformatted_doc.docx', // Use path relative to workspace
          styleSet: 'Professional', // Example style set - need to confirm supported sets from API spec
          // Optional: specific styles to apply, etc.
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    // Verify the document was modified (basic check - could check modification time)
    await expect(fs.stat(outputPath)).resolves.toBeTruthy();
    // TODO: Add verification that the document was actually reformatted (requires inspecting doc content/structure)
  });

  // TODO: Add tests for specifying specific styles or other reformatting options
  // TODO: Add tests for error handling (e.g., non-existent document, invalid style set)
});