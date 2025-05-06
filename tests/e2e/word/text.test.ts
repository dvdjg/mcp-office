import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_word_text_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');
const FIXTURE_WORD_DOC_NAME = 'CV.docx';
const FIXTURE_WORD_DOC_PATH = path.join(FIXTURES_DIR, FIXTURE_WORD_DOC_NAME);
const RELATIVE_FIXTURE_PATH = `tests/fixtures/${FIXTURE_WORD_DOC_NAME}`; // Relative to workspace root

// Define a basic type for the expected successful response
interface SuccessResponse {
  success: true;
  data: any;
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

const callTool = async (toolName: string, args: Record<string, any>): Promise<ToolResponse> => {
  const response = await fetch(`${MCP_SERVER_URL}/tool`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ tool_name: toolName, arguments: args }),
  });
  // if (!response.ok) { // It's better to check response.ok inside the test for more specific error messages
  //   const errorBody = await response.text();
  //   throw new Error(`Tool call failed with status ${response.status}: ${errorBody}`);
  // }
  return response.json() as Promise<ToolResponse>;
};


describe('word/text e2e tests', () => {
  beforeAll(async () => {
    await fs.rm(TEMP_DIR, { recursive: true, force: true });
    await fs.mkdir(TEMP_DIR, { recursive: true });
    try {
      await fs.stat(FIXTURE_WORD_DOC_PATH);
    } catch (error: any) {
      if (error.code === 'ENOENT') {
        console.error(`Fixture file ${FIXTURE_WORD_DOC_PATH} not found. Please ensure it exists.`);
        throw new Error(`Fixture file ${FIXTURE_WORD_DOC_PATH} not found.`);
      }
      throw error;
    }
  });

  afterAll(async () => {
    await fs.rm(TEMP_DIR, { recursive: true, force: true });
  });

  const testModes = [
    { mode: 'COM', useComInterop: true },
    { mode: 'Library', useComInterop: false },
    { mode: 'Library (Default)', useComInterop: undefined },
  ];

  describe.each(testModes)('word/text/insert (mode: $mode)', ({ useComInterop }) => {
    test('should insert text into a Word document or report not implemented', async () => {
      const testFileName = `doc_with_inserted_text_${useComInterop}.docx`;
      const outputPath = path.join(TEMP_DIR, testFileName);
      const relativeOutputPath = `tests/temp_word_text_dir/${testFileName}`;
      await fs.copyFile(FIXTURE_WORD_DOC_PATH, outputPath);

      const textToInsert = 'This is some text to insert via test.';
      const args = {
        filePath: relativeOutputPath,
        text: textToInsert,
        position: 'end',
        useComInterop,
      };

      const result = await callTool('word/text/insert', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        // Basic check, deeper content verification is complex for E2E
        const stats = await fs.stat(outputPath);
        expect(stats.size).toBeGreaterThan(0); // Or check modification time
      } else {
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB');
        }
      }
    });

    test('should create a new file if it does not exist and insert text (COM only) or report not implemented', async () => {
        const testFileName = `new_doc_inserted_text_${useComInterop}.docx`;
        const outputPath = path.join(TEMP_DIR, testFileName);
        const relativeOutputPath = `tests/temp_word_text_dir/${testFileName}`;

        const textToInsert = 'Text in a new document.';
        const args = {
            filePath: relativeOutputPath,
            text: textToInsert,
            position: 'start',
            useComInterop,
        };

        const result = await callTool('word/text/insert', args);

        if (useComInterop) {
            expect(result.success).toBe(true);
            await expect(fs.stat(outputPath)).resolves.toBeTruthy();
            // TODO: Verify content if possible
        } else {
            expect(result.success).toBe(false);
            if (!result.success) {
                expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB');
            }
        }
    });
  });

  describe.each(testModes)('word/text/get (mode: $mode)', ({ useComInterop }) => {
    test('should get text from a Word document (full document)', async () => {
      const args = {
        filePath: RELATIVE_FIXTURE_PATH,
        range: 'document',
        useComInterop,
      };
      const result = await callTool('word/text/get', args);

      expect(result.success).toBe(true);
      if (result.success) {
        expect(typeof result.data).toBe('string');
        expect(result.data.length).toBeGreaterThan(0);
        // TODO: More specific content checks for CV.docx
      }
    });

    test('should handle specific ranges (COM) or report error (Library)', async () => {
      const args = {
        filePath: RELATIVE_FIXTURE_PATH,
        range: 'paragraph:1', // Specific range
        useComInterop,
      };
      const result = await callTool('word/text/get', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        if (result.success) {
          expect(typeof result.data).toBe('string');
          // Content might be empty for a specific paragraph if it's just a newline
        }
      } else {
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('RANGE_REQUIRES_COM_LIB');
        }
      }
    });

     test('should handle "selection" range (COM) or report error (Library)', async () => {
      const args = {
        filePath: RELATIVE_FIXTURE_PATH,
        range: 'selection',
        useComInterop,
      };
      const result = await callTool('word/text/get', args);

      if (useComInterop) {
        // This test is tricky for E2E as selection depends on Word's state.
        // It might fail if Word is not active or has no selection.
        // For robust E2E, this might need a mock or a way to ensure selection.
        // We expect it to either succeed (if selection exists) or fail gracefully.
        if (result.success) {
            expect(typeof result.data).toBe('string');
        } else {
            expect(result.error.code).toBe('GET_TEXT_FAILED_COM'); // Or a more specific selection error
            expect(result.error.message).toContain("Could not get range from selection");
        }
      } else {
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('RANGE_REQUIRES_COM_LIB');
        }
      }
    });
  });

  describe.each(testModes)('word/text/modify (mode: $mode)', ({ useComInterop }) => {
    test('should modify text in a Word document or report not implemented', async () => {
      const testFileName = `doc_with_modified_text_${useComInterop}.docx`;
      const outputPath = path.join(TEMP_DIR, testFileName);
      const relativeOutputPath = `tests/temp_word_text_dir/${testFileName}`;
      await fs.copyFile(FIXTURE_WORD_DOC_PATH, outputPath);

      const newText = 'This document content has been entirely replaced.';
      const args = {
        filePath: relativeOutputPath,
        range: 'document', // Modify the entire document's content
        newText,
        useComInterop,
      };

      const result = await callTool('word/text/modify', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        // TODO: Verify content change
      } else {
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB');
        }
      }
    });
  });

  describe.each(testModes)('word/text/delete (mode: $mode)', ({ useComInterop }) => {
    test('should delete text in a Word document or report not implemented', async () => {
      const testFileName = `doc_with_deleted_text_${useComInterop}.docx`;
      const outputPath = path.join(TEMP_DIR, testFileName);
      const relativeOutputPath = `tests/temp_word_text_dir/${testFileName}`;
      await fs.copyFile(FIXTURE_WORD_DOC_PATH, outputPath);

      const args = {
        filePath: relativeOutputPath,
        range: 'paragraph:1', // Delete the first paragraph
        useComInterop,
      };

      const result = await callTool('word/text/delete', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        // TODO: Verify content change (e.g., document is smaller or specific text is gone)
      } else {
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB');
        }
      }
    });
  });

  test('word/text/get should fail for non-existent file', async () => {
    const args = {
      filePath: 'non_existent_document.docx',
      range: 'document',
      useComInterop: false, // Library path chosen as it's simpler for this check
    };
    const result = await callTool('word/text/get', args);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('FILE_NOT_FOUND');
    }
  });

});