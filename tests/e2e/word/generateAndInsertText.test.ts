import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_generate_text_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_WORD_DOC_EMPTY = path.join(FIXTURES_DIR, 'empty_doc.docx'); // Assuming an empty Word doc fixture exists


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


describe('word/generate-and-insert-text e2e tests', () => {

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
      await fs.stat(FIXTURE_WORD_DOC_EMPTY);
    } catch (error: any) {
      if (error.code === 'ENOENT') {
         console.error(`Fixture file ${FIXTURE_WORD_DOC_EMPTY} not found. Please ensure it exists for word/generate-and-insert-text tests.`);
         // Depending on test setup, might throw or skip tests
      } else {
        console.error(`Error checking fixture file ${FIXTURE_WORD_DOC_EMPTY}:`, error);
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

  test('should generate and insert text into a Word document (requires fixture)', async () => {
    const outputPath = path.join(TEMP_DIR, 'doc_with_generated_text.docx');
    // Copy the empty fixture to the temp directory to avoid modifying the original
    await fs.copyFile(FIXTURE_WORD_DOC_EMPTY, outputPath);

    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/generate-and-insert-text',
        arguments: {
          document: 'tests/temp_word_generate_text_dir/doc_with_generated_text.docx', // Use path relative to workspace
          prompt: 'Write a short paragraph about cloud computing.', // Example prompt
          position: 'end', // Insert at the end of the document
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    // Verify the document was modified (basic check - could check modification time)
    await expect(fs.stat(outputPath)).resolves.toBeTruthy();
    // TODO: Add verification that text was actually inserted (e.g., check file size increased significantly)
  });

  // TODO: Add tests for inserting at different positions (start, specific range)
  // TODO: Add tests for error handling (e.g., non-existent document, invalid position)
});