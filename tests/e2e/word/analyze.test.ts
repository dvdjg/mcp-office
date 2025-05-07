import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_analyze_dir');
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


describe('word/analyze e2e tests', () => {

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
         console.error(`Fixture file ${FIXTURE_WORD_DOC} not found. Please ensure it exists for word/analyze tests.`);
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

  test('should analyze a Word document and add comments (requires fixture)', async () => {
    const outputPath = path.join(TEMP_DIR, 'analyzed_doc.docx');
    // Copy the fixture to the temp directory to avoid modifying the original
    await fs.copyFile(FIXTURE_WORD_DOC, outputPath);

    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/analyze',
        arguments: {
          document: 'tests/temp_word_analyze_dir/analyzed_doc.docx', // Use path relative to workspace
          criteria: 'grammar,spelling', // Example criteria - need to confirm supported criteria from API spec
          operation: 'addComments', // Assuming the operation name is addComments
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    // Verify the document was modified (basic check - could check modification time)
    await expect(fs.stat(outputPath)).resolves.toBeTruthy();
    // TODO: Add verification that comments were actually added (requires inspecting doc content/structure)
  });

  test('should analyze a Word document and return analysis results (requires fixture)', async () => {
    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/analyze',
        arguments: {
          document: 'tests/fixtures/CV.docx', // Use path relative to workspace
          criteria: 'readability', // Example criteria
          operation: 'analyze', // Assuming the operation name is analyze
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    if (result.success) {
      expect(result.data).toHaveProperty('analysisResults');
      // TODO: Add more specific checks based on the expected analysis results structure
      // expect(result.data.analysisResults).toHaveProperty('readabilityScore');
    } else {
      fail('Expected test to succeed but it failed.');
    }
  });

  // TODO: Add tests for other criteria and error handling
});