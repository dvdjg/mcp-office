import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_text_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_WORD_DOC = path.join(FIXTURES_DIR, 'CV.docx'); // Assuming a Word doc fixture exists


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


describe('word/text e2e tests', () => {

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
      await fs.stat(FIXTURE_WORD_DOC);
    } catch (error: any) {
      if (error.code === 'ENOENT') {
         console.error(`Fixture file ${FIXTURE_WORD_DOC} not found. Please ensure it exists for word/text tests.`);
         // Depending on test setup, might throw or skip tests
      } else {
        console.error(`Error checking fixture file ${FIXTURE_WORD_DOC}:`, error);
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

  describe('word/text/insert', () => {
    test('should insert text into a Word document (requires fixture)', async () => {
      const outputPath = path.join(TEMP_DIR, 'doc_with_inserted_text.docx');
      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(FIXTURE_WORD_DOC, outputPath);

      const textToInsert = 'This is some text to insert.';

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/text/insert',
          arguments: {
            document: 'tests/temp_word_text_dir/doc_with_inserted_text.docx', // Use path relative to workspace
            text: textToInsert,
            position: 'end', // Insert at the end
            // Optional: range, style
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the document was modified (basic check - could check modification time)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification that the text was actually inserted (requires inspecting doc content)
    });

    // TODO: Add tests for inserting at different positions (start, specific range)
    // TODO: Add tests for inserting with a specific style
    // TODO: Add tests for error handling (e.g., non-existent document, invalid position)
  });

  describe('word/text/read', () => {
    test('should read text from a Word document (requires fixture)', async () => {
      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/text/read',
          arguments: {
            document: 'tests/fixtures/CV.docx', // Use path relative to workspace
            // Optional: range
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('text');
        expect(typeof result.data.text).toBe('string');
        expect(result.data.text.length).toBeGreaterThan(0);
        // TODO: Add more specific checks based on the expected content of the fixture document
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    // TODO: Add tests for reading text from a specific range
    // TODO: Add tests for error handling (e.g., non-existent document, invalid range)
  });

  describe('word/text/replace', () => {
    test('should replace text in a Word document (requires fixture)', async () => {
      const outputPath = path.join(TEMP_DIR, 'doc_with_replaced_text.docx');
      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(FIXTURE_WORD_DOC, outputPath);

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/text/replace',
          arguments: {
            document: 'tests/temp_word_text_dir/doc_with_replaced_text.docx', // Use path relative to workspace
            searchText: 'Curriculum Vitae', // Example text to replace - ensure it exists in the fixture
            replaceText: 'Resume',
            // Optional: matchCase, matchWholeWord, range
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the document was modified (basic check - could check modification time)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification that the text was actually replaced (requires inspecting doc content)
    });

    // TODO: Add tests for replacing with different options (matchCase, matchWholeWord, range)
    // TODO: Add tests for replacing text that does not exist
    // TODO: Add tests for error handling (e.g., non-existent document, invalid range)
  });
});