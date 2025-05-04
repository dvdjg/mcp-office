import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_mermaid_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_WORD_DOC_EMPTY = path.join(FIXTURES_DIR, 'empty_doc.docx'); // Assuming an empty Word doc fixture exists
const FIXTURE_DOC_WITH_MERMAID = path.join(FIXTURES_DIR, 'doc_with_mermaid.docx'); // Assuming a Word doc fixture with a Mermaid diagram exists


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


describe('word/mermaid e2e tests', () => {

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
    const fixtures = [FIXTURE_WORD_DOC_EMPTY, FIXTURE_DOC_WITH_MERMAID];
    for (const fixturePath of fixtures) {
      try {
        await fs.stat(fixturePath);
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.error(`Fixture file ${fixturePath} not found. Please ensure it exists for word/mermaid tests.`);
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

  describe('word/mermaid/import', () => {
    test('should import a Mermaid diagram into a Word document (requires fixture)', async () => {
      const outputPath = path.join(TEMP_DIR, 'doc_with_mermaid.docx');
      // Copy the empty fixture to the temp directory to avoid modifying the original
      await fs.copyFile(FIXTURE_WORD_DOC_EMPTY, outputPath);

      const mermaidSyntax = `
graph TD
    A[Start] --> B{Is it?};
    B -->|Yes| C[OK];
    C --> D[End];
    B -->|No| E[Error];
    E --> D;
`;

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/mermaid/import',
          arguments: {
            document: 'tests/temp_word_mermaid_dir/doc_with_mermaid.docx', // Use path relative to workspace
            syntax: mermaidSyntax,
            format: 'svg', // Example format - need to confirm supported formats
            // Optional: position, scale
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the document was modified (basic check - could check modification time)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification that the Mermaid diagram was actually inserted (requires inspecting doc content/structure)
    });

    // TODO: Add tests for importing with different formats (e.g., png) and positions
    // TODO: Add tests for error handling (e.g., non-existent document, invalid syntax)
  });

  describe('word/mermaid/export', () => {
    test('should export a Mermaid diagram from a Word document (requires fixture)', async () => {
      const outputPath = path.join(TEMP_DIR, 'exported_mermaid.svg');

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/mermaid/export',
          arguments: {
            document: 'tests/fixtures/doc_with_mermaid.docx', // Use path relative to workspace
            diagramIndex: 1, // Assuming the first diagram
            output_path: 'tests/temp_word_mermaid_dir/exported_mermaid.svg', // Use path relative to workspace
            format: 'svg', // Example format
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the output file was created (basic check)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification of the exported file content (e.g., check for SVG/PNG markers)
    });

    // TODO: Add tests for exporting different diagrams by index, with different formats
    // TODO: Add tests for error handling (e.g., non-existent document, invalid diagram index)
  });
});