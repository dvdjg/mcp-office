import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000';
const TEMP_DIR = path.join(__dirname, '../temp_word_image_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');

const FIXTURE_DOC_WITH_IMAGE_NAME = 'doc_with_image.docx'; // Contains at least one image
const FIXTURE_DOC_WITH_IMAGE_PATH = path.join(FIXTURES_DIR, FIXTURE_DOC_WITH_IMAGE_NAME);
const RELATIVE_FIXTURE_DOC_WITH_IMAGE_PATH = `tests/fixtures/${FIXTURE_DOC_WITH_IMAGE_NAME}`;

const FIXTURE_IMAGE_NAME = 'sample_image.png';
const FIXTURE_IMAGE_PATH = path.join(FIXTURES_DIR, FIXTURE_IMAGE_NAME);
// const RELATIVE_FIXTURE_IMAGE_PATH = `tests/fixtures/${FIXTURE_IMAGE_NAME}`; // Not used directly by tool, image data is base64

const FIXTURE_EMPTY_DOC_NAME = 'empty_doc.docx';
const FIXTURE_EMPTY_DOC_PATH = path.join(FIXTURES_DIR, FIXTURE_EMPTY_DOC_NAME);
const RELATIVE_FIXTURE_EMPTY_DOC_PATH = `tests/fixtures/${FIXTURE_EMPTY_DOC_NAME}`;


interface SuccessResponse {
  success: true;
  data: any;
  message?: string;
}

interface ErrorResponse {
  success: false;
  error: {
    code: string;
    message: string;
    details?: any;
  };
}

type ToolResponse = SuccessResponse | ErrorResponse;

const callTool = async (toolName: string, args: Record<string, any>): Promise<ToolResponse> => {
  const response = await fetch(`${MCP_SERVER_URL}/tool`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ tool_name: toolName, arguments: args }),
  });
  return response.json() as Promise<ToolResponse>;
};

describe('word/image e2e tests', () => {
  let sampleImageBase64: string;

  beforeAll(async () => {
    await fs.rm(TEMP_DIR, { recursive: true, force: true });
    await fs.mkdir(TEMP_DIR, { recursive: true });

    const fixturesToVerify = [FIXTURE_DOC_WITH_IMAGE_PATH, FIXTURE_IMAGE_PATH, FIXTURE_EMPTY_DOC_PATH];
    for (const fixturePath of fixturesToVerify) {
      try {
        await fs.stat(fixturePath);
      } catch (error: any) {
        if (error.code === 'ENOENT') {
          console.error(`Required fixture file ${fixturePath} not found.`);
          throw new Error(`Fixture file ${fixturePath} not found.`);
        }
        throw error;
      }
    }
    // Load sample image data for insertion tests
    const imageBuffer = await fs.readFile(FIXTURE_IMAGE_PATH);
    sampleImageBase64 = imageBuffer.toString('base64');
  });

  afterAll(async () => {
    await fs.rm(TEMP_DIR, { recursive: true, force: true });
  });

  const testModes = [
    { mode: 'COM', useComInterop: true },
    { mode: 'Library', useComInterop: false },
    { mode: 'Library (Default)', useComInterop: undefined },
  ];

  describe.each(testModes)('word/image/extract (mode: $mode)', ({ useComInterop }) => {
    test('should extract an image from a Word document by index', async () => {
      const args = {
        filePath: RELATIVE_FIXTURE_DOC_WITH_IMAGE_PATH,
        identifier: 1, // Extract the first image
        outputFormat: 'png', // outputFormat might be ignored by library path
        useComInterop,
      };
      const result = await callTool('word/image/extract', args);

      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toBeInstanceOf(Buffer); // Expect Buffer in response data
        expect(result.data.length).toBeGreaterThan(0);
        // TODO: Could save buffer to a temp file and verify it's a valid image
      }
    });

    test('should extract an image from a Word document by alt text (if available)', async () => {
        // This test assumes 'doc_with_image.docx' has an image with identifiable alt text.
        // Let's assume an image has "Test Alt Text" or similar.
        const altTextIdentifier = "A sample image used for testing purposes."; // Adjust if your fixture has different alt text
        const args = {
            filePath: RELATIVE_FIXTURE_DOC_WITH_IMAGE_PATH,
            identifier: altTextIdentifier,
            outputFormat: 'jpeg',
            useComInterop,
        };
        const result = await callTool('word/image/extract', args);

        // Library path might struggle with alt text if Mammoth doesn't expose it well for this tool's logic
        if (useComInterop) {
            expect(result.success).toBe(true);
            if (result.success) {
                expect(result.data).toBeInstanceOf(Buffer);
                expect(result.data.length).toBeGreaterThan(0);
            }
        } else {
            // Library path's alt text identification is best-effort.
            // It might succeed if alt text is simple and Mammoth picks it up, or fail.
            if (result.success) {
                expect(result.data).toBeInstanceOf(Buffer);
                expect(result.data.length).toBeGreaterThan(0);
                console.warn("Library path succeeded extracting by alt text, which is a good sign.");
            } else {
                expect(result.success).toBe(false);
                // Expect specific error if alt text not found or not supported well by library path
                expect(result.error.code).toMatch(/IMAGE_EXTRACTION_FAILED_LIB|USER_ERROR/);
                 console.warn(`Library path failed to extract by alt text as potentially expected: ${result.error.message}`);
            }
        }
    });
  });

  describe.each(testModes)('word/image/insert (mode: $mode)', ({ useComInterop }) => {
    test('should insert an image into a Word document', async () => {
      const outputFileName = `doc_inserted_image_${useComInterop}.docx`;
      const outputPath = path.join(TEMP_DIR, outputFileName);
      const relativeOutputPath = `tests/temp_word_image_dir/${outputFileName}`;

      if (useComInterop) { // COM modifies existing
        await fs.copyFile(FIXTURE_EMPTY_DOC_PATH, outputPath);
      }

      const args = {
        filePath: relativeOutputPath,
        imageDataBase64: sampleImageBase64,
        position: 'end',
        width: 150,
        height: 100,
        altText: 'A test image inserted via E2E test',
        useComInterop,
      };
      const result = await callTool('word/image/insert', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        await expect(fs.stat(outputPath)).resolves.toBeTruthy();
        // TODO: Verify image presence in COM path (complex)
      } else { // Library path
        // Library path creates new, or fails if file exists and not 'end' position
        let fileExisted = false;
        try { await fs.stat(outputPath); fileExisted = true; } catch (e) { /* fine */ }

        if (fileExisted) { // Should fail as it's not a new file
            expect(result.success).toBe(false);
            if(!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_MODIFY');
        } else { // New file creation
            expect(result.success).toBe(true);
            await expect(fs.stat(outputPath)).resolves.toBeTruthy();
            // TODO: Verify image presence in new file for Library path (complex)
        }
      }
    });

    test('Library path should fail to insert into existing file at specific position', async () => {
        if (useComInterop === false || useComInterop === undefined) {
            const outputFileName = `existing_doc_lib_img_insert_fail.docx`;
            const outputPath = path.join(TEMP_DIR, outputFileName);
            const relativeOutputPath = `tests/temp_word_image_dir/${outputFileName}`;
            await fs.copyFile(FIXTURE_EMPTY_DOC_PATH, outputPath); // Create an existing file

            const args = {
                filePath: relativeOutputPath,
                imageDataBase64: sampleImageBase64,
                position: 'paragraph:1', // Specific position
                useComInterop,
            };
            const result = await callTool('word/image/insert', args);
            expect(result.success).toBe(false);
            if (!result.success) {
                expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_MODIFY');
            }
        } else {
            expect(true).toBe(true); // Test only for library path
        }
    });
  });

  test('word/image/extract should fail for non-existent file', async () => {
    const args = {
      filePath: 'non_existent_doc_for_image_extract.docx',
      identifier: 1,
      useComInterop: true, // COM path
    };
    const result = await callTool('word/image/extract', args);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('IMAGE_EXTRACTION_FAILED_COM'); // COM error for file open
      expect(result.error.message).toMatch(/Failed to open document|File not found/i);
    }
  });
});