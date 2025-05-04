import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_headers_footers_dir');
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


describe('word/headersFooters e2e tests', () => {

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
         console.error(`Fixture file ${FIXTURE_WORD_DOC_EMPTY} not found. Please ensure it exists for word/headersFooters tests.`);
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

  test('should insert a header into a Word document (requires fixture)', async () => {
    const outputPath = path.join(TEMP_DIR, 'doc_with_header.docx');
    // Copy the empty fixture to the temp directory to avoid modifying the original
    await fs.copyFile(FIXTURE_WORD_DOC_EMPTY, outputPath);

    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/headersFooters/insert', // Assuming the operation name is insert
        arguments: {
          document: 'tests/temp_word_headers_footers_dir/doc_with_header.docx', // Use path relative to workspace
          type: 'header',
          content: 'This is a test header.',
          // Optional: sectionIndex, alignment, etc. - need to confirm from API spec
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    // Verify the document was modified (basic check - could check modification time)
    await expect(fs.stat(outputPath)).resolves.toBeTruthy();
    // TODO: Add verification that the header was actually inserted (requires inspecting doc content/structure)
  });

  test('should insert a footer into a Word document (requires fixture)', async () => {
    const outputPath = path.join(TEMP_DIR, 'doc_with_footer.docx');
    // Copy the empty fixture to the temp directory to avoid modifying the original
    await fs.copyFile(FIXTURE_WORD_DOC_EMPTY, outputPath);

    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/headersFooters/insert', // Assuming the operation name is insert
        arguments: {
          document: 'tests/temp_word_headers_footers_dir/doc_with_footer.docx', // Use path relative to workspace
          type: 'footer',
          content: 'This is a test footer.',
          // Optional: sectionIndex, alignment, etc. - need to confirm from API spec
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    // Verify the document was modified (basic check - could check modification time)
    await expect(fs.stat(outputPath)).resolves.toBeTruthy();
    // TODO: Add verification that the footer was actually inserted (requires inspecting doc content/structure)
  });

  // TODO: Add tests for inserting headers/footers with different options (sectionIndex, alignment, etc.)
  // TODO: Add tests for deleting headers/footers
  // TODO: Add tests for error handling (e.g., non-existent document, invalid type)
});