import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_markdown_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_MD_FILE = path.join(FIXTURES_DIR, 'sample.md'); // Assuming a sample markdown fixture exists
const FIXTURE_DOTX_FILE = path.join(FIXTURES_DIR, 'sample_template.dotx'); // Assuming a sample template fixture exists


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


describe('word/markdown e2e tests', () => {

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
    const fixtures = [FIXTURE_MD_FILE, FIXTURE_DOTX_FILE];
    for (const fixturePath of fixtures) {
      try {
        await fs.stat(fixturePath);
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.error(`Fixture file ${fixturePath} not found. Please ensure it exists for word/markdown tests.`);
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

  describe('word/markdown/import', () => {
    test('should import a markdown file into a new Word document', async () => {
      const outputPath = path.join(TEMP_DIR, 'imported_markdown.docx');

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/markdown/import',
          arguments: {
            path: 'tests/fixtures/sample.md', // Use path relative to workspace
            output: 'tests/temp_word_markdown_dir/imported_markdown.docx', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the Word file was created (basic check)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification of the imported content in the Word document
    });

    test('should import a markdown file into a Word document using a template (requires fixture)', async () => {
      const outputPath = path.join(TEMP_DIR, 'imported_markdown_with_template.docx');

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/markdown/import',
          arguments: {
            path: 'tests/fixtures/sample.md', // Use path relative to workspace
            template: 'tests/fixtures/sample_template.dotx', // Use path relative to workspace
            output: 'tests/temp_word_markdown_dir/imported_markdown_with_template.docx', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the Word file was created (basic check)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification that the template was applied and content imported correctly
    });

    // TODO: Add tests for error handling (e.g., non-existent markdown file, non-existent template)
  });

  // TODO: Add tests for word/markdown/export variations if not fully covered by existing tests
});