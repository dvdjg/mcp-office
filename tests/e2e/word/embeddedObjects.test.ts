import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_embedded_objects_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_DOC_WITH_EMBEDDED = path.join(FIXTURES_DIR, 'doc_with_embedded_objects.docx'); // Assuming this fixture exists with embedded objects
const FIXTURE_WORD_DOC_EMPTY = path.join(FIXTURES_DIR, 'empty_doc.docx'); // Assuming an empty Word doc fixture exists
const FIXTURE_EXCEL_FILE = path.join(FIXTURES_DIR, 'sample_excel_data.xlsx'); // Assuming a sample Excel file fixture exists


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


describe('word/embedded-objects e2e tests', () => {

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
    const fixtures = [FIXTURE_DOC_WITH_EMBEDDED, FIXTURE_WORD_DOC_EMPTY, FIXTURE_EXCEL_FILE];
    for (const fixturePath of fixtures) {
      try {
        await fs.stat(fixturePath);
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.error(`Fixture file ${fixturePath} not found. Please ensure it exists for word/embedded-objects tests.`);
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

  describe('word/embedded-objects/extractAll', () => {
    test('should extract all embedded objects from a Word document (requires fixture)', async () => {
      const outputDir = path.join(TEMP_DIR, 'extracted_objects');

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/embedded-objects/extractAll',
          arguments: {
            filePath: 'tests/fixtures/doc_with_embedded_objects.docx', // Use path relative to workspace
            outputDirectory: 'tests/temp_word_embedded_objects_dir/extracted_objects', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('extractedFiles');
        expect(Array.isArray(result.data.extractedFiles)).toBe(true);
        // TODO: Add more specific checks based on the expected number and names of embedded files in the fixture
        // expect(result.data.extractedFiles.length).toBeGreaterThan(0);
        // expect(result.data.extractedFiles).toContain(path.normalize('tests/temp_word_embedded_objects_dir/extracted_objects/embedded_file_name.ext'));
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    // TODO: Add tests for extracting specific images by index or other criteria if supported
    // TODO: Add tests for error handling (e.g., non-existent document)
  });

  describe('word/embedded-objects/insert', () => {
    test('should insert an embedded object into a Word document (requires fixtures)', async () => {
      const outputPath = path.join(TEMP_DIR, 'doc_with_inserted_object.docx');
      // Copy the empty fixture to the temp directory to avoid modifying the original
      await fs.copyFile(FIXTURE_WORD_DOC_EMPTY, outputPath);

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/embedded-objects/insert',
          arguments: {
            filePath: 'tests/temp_word_embedded_objects_dir/doc_with_inserted_object.docx', // Use path relative to workspace
            objectPath: 'tests/fixtures/sample_excel_data.xlsx', // Use path relative to workspace
            // Optional: position, displayAsIcon, iconPath, iconLabel
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the document was modified (basic check - could check modification time)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification that the object was actually inserted into the document
    });

    // TODO: Add tests for inserting with different options (position, displayAsIcon, etc.)
    // TODO: Add tests for error handling (e.g., non-existent document, non-existent object file)
  });

  describe('word/embedded-objects/delete', () => {
    test('should delete an embedded object from a Word document (requires fixture)', async () => {
      const outputPath = path.join(TEMP_DIR, 'doc_after_delete_object.docx');
      // Copy the fixture with embedded objects to the temp directory
      await fs.copyFile(FIXTURE_DOC_WITH_EMBEDDED, outputPath);
      const objectIndexToDelete = 1; // Assuming there's at least one embedded object at index 1

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/embedded-objects/delete',
          arguments: {
            filePath: 'tests/temp_word_embedded_objects_dir/doc_after_delete_object.docx', // Use path relative to workspace
            objectIndex: objectIndexToDelete,
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the document was modified (basic check - could check modification time)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification that the object was actually deleted (requires inspecting doc content/structure)
    });

    // TODO: Add tests for deleting non-existent objects, error handling
  });

  describe('word/embedded-objects/modify', () => {
    test('should modify an embedded object in a Word document (requires fixtures)', async () => {
      const outputPath = path.join(TEMP_DIR, 'doc_after_modify_object.docx');
      // Copy the fixture with embedded objects to the temp directory
      await fs.copyFile(FIXTURE_DOC_WITH_EMBEDDED, outputPath);
      const objectIndexToModify = 1; // Assuming there's at least one embedded object at index 1
      const newObjectPath = path.join(FIXTURES_DIR, 'sample_excel_data.xlsx'); // Use a different Excel file or modified one


      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/embedded-objects/modify',
          arguments: {
            filePath: 'tests/temp_word_embedded_objects_dir/doc_after_modify_object.docx', // Use path relative to workspace
            objectIndex: objectIndexToModify,
            newObjectPath: 'tests/fixtures/sample_excel_data.xlsx', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the document was modified (basic check - could check modification time)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification that the object was actually modified (requires inspecting doc content/structure)
    });

    // TODO: Add tests for modifying with different object types, error handling
  });
});