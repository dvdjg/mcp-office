import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_image_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_DOC_WITH_IMAGE = path.join(FIXTURES_DIR, 'doc_with_image.docx'); // Assuming this fixture exists with an image
const FIXTURE_IMAGE_FILE = path.join(FIXTURES_DIR, 'sample_image.png'); // Assuming a sample image fixture exists
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


describe('word/image e2e tests', () => {

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
    const fixtures = [FIXTURE_DOC_WITH_IMAGE, FIXTURE_IMAGE_FILE, FIXTURE_WORD_DOC_EMPTY];
    for (const fixturePath of fixtures) {
      try {
        await fs.stat(fixturePath);
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.error(`Fixture file ${fixturePath} not found. Please ensure it exists for word/image tests.`);
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

  describe('word/image/extract', () => {
    test('should extract images from a Word document (requires fixture)', async () => {
      const outputDir = path.join(TEMP_DIR, 'extracted_images');

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/image/extract',
          arguments: {
            document: 'tests/fixtures/doc_with_image.docx', // Use path relative to workspace
            outputDirectory: 'tests/temp_word_image_dir/extracted_images', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('extractedFiles');
        expect(Array.isArray(result.data.extractedFiles)).toBe(true);
        // TODO: Add more specific checks based on the expected number and names of images in the fixture
        // expect(result.data.extractedFiles.length).toBeGreaterThan(0);
        // expect(result.data.extractedFiles).toContain(path.normalize('tests/temp_word_image_dir/extracted_images/image1.png'));
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    // TODO: Add tests for extracting specific images by index or other criteria if supported
    // TODO: Add tests for error handling (e.g., non-existent document, document with no images)
  });

  describe('word/image/insert', () => {
    test('should insert an image into a Word document (requires fixtures)', async () => {
      const outputPath = path.join(TEMP_DIR, 'doc_with_inserted_image.docx');
      // Copy the empty fixture to the temp directory to avoid modifying the original
      await fs.copyFile(FIXTURE_WORD_DOC_EMPTY, outputPath);

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'word/image/insert',
          arguments: {
            document: 'tests/temp_word_image_dir/doc_with_inserted_image.docx', // Use path relative to workspace
            imagePath: 'tests/fixtures/sample_image.png', // Use path relative to workspace
            // Optional: position, width, height, altText
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the document was modified (basic check - could check modification time)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification that the image was actually inserted into the document
    });

    // TODO: Add tests for inserting image data (base64), inserting at specific positions, with size/alt text
    // TODO: Add tests for error handling (e.g., non-existent document, non-existent image file)
  });
});