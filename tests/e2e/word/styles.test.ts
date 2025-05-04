import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_styles_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_WORD_DOC = path.join(FIXTURES_DIR, 'CV.docx'); // Assuming this fixture exists

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


describe('word/styles e2e tests', () => {

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
         console.error(`Fixture file ${FIXTURE_WORD_DOC} not found. Please ensure it exists for word/styles tests.`);
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

  test('should list styles in a Word document (requires fixture)', async () => {
    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/styles/list',
        arguments: {
          document: 'tests/fixtures/CV.docx', // Use path relative to workspace
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    if (result.success) {
      expect(result.data).toHaveProperty('styles');
      expect(Array.isArray(result.data.styles)).toBe(true);
      // Basic check for some common styles
      const styleNames = result.data.styles.map((s: any) => s.name);
      expect(styleNames).toContain('Normal');
      // Add more style names if known to be in the fixture document
    } else {
      fail('Expected test to succeed but it failed.');
    }
  });

  test('should apply a style to a range in a Word document (requires fixture)', async () => {
    const outputPath = path.join(TEMP_DIR, 'apply_style_test.docx');
    // Copy the fixture to the temp directory to avoid modifying the original
    await fs.copyFile(FIXTURE_WORD_DOC, outputPath);

    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/styles/apply',
        arguments: {
          document: 'tests/temp_word_styles_dir/apply_style_test.docx', // Use path relative to workspace
          style: 'Heading1', // Example style - ensure this style exists in the fixture or is a default
          range: 'paragraph:1', // Apply to the first paragraph
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    // Verify the document was modified (basic check - could check modification time)
    await expect(fs.stat(outputPath)).resolves.toBeTruthy();
    // TODO: Add verification that the style was actually applied to the specified range
  });

  // TODO: Add tests for create, modify, and delete style operations
  // TODO: Add tests for error handling (e.g., applying non-existent style, invalid range)
});