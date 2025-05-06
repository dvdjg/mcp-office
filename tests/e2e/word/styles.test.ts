import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000';
const TEMP_DIR = path.join(__dirname, '../temp_word_styles_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');

const FIXTURE_WORD_DOC_NAME = 'CV.docx'; // A document with various styles
const FIXTURE_WORD_DOC_PATH = path.join(FIXTURES_DIR, FIXTURE_WORD_DOC_NAME);
const RELATIVE_FIXTURE_WORD_DOC_PATH = `tests/fixtures/${FIXTURE_WORD_DOC_NAME}`;

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

describe('word/styles e2e tests', () => {
  beforeAll(async () => {
    await fs.rm(TEMP_DIR, { recursive: true, force: true });
    await fs.mkdir(TEMP_DIR, { recursive: true });

    try {
      await fs.stat(FIXTURE_WORD_DOC_PATH);
    } catch (error: any) {
      if (error.code === 'ENOENT') {
        console.error(`Required fixture file ${FIXTURE_WORD_DOC_PATH} not found.`);
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

  describe.each(testModes)('word/styles/list (mode: $mode)', ({ useComInterop }) => {
    test('should list styles in a Word document or report not implemented', async () => {
      const args = {
        filePath: RELATIVE_FIXTURE_WORD_DOC_PATH,
        useComInterop,
      };
      const result = await callTool('word/styles/list', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        if (result.success) {
          expect(Array.isArray(result.data)).toBe(true);
          // CV.docx should have 'Normal' and some heading styles
          expect(result.data).toContain('Normal');
          expect(result.data.some((style: string) => style.startsWith('Heading'))).toBe(true);
        }
      } else { // Library path
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_STYLE_LIST');
        }
      }
    });
  });

  describe.each(testModes)('word/styles/apply (mode: $mode)', ({ useComInterop }) => {
    test('should apply a style to a range in a Word document or report not implemented', async () => {
      const outputFileName = `doc_style_applied_${useComInterop}.docx`;
      const outputPath = path.join(TEMP_DIR, outputFileName);
      const relativeOutputPath = `tests/temp_word_styles_dir/${outputFileName}`;
      await fs.copyFile(FIXTURE_WORD_DOC_PATH, outputPath); // Operate on a copy

      const styleToApply = 'Heading 1'; // A common style
      const rangeToApply = 'paragraph:1'; // Apply to the first paragraph

      const args = {
        filePath: relativeOutputPath,
        style: styleToApply,
        range: rangeToApply,
        useComInterop,
      };
      const result = await callTool('word/styles/apply', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        // TODO: Verify style application in COM path (complex, requires reading doc structure)
        // For now, just check that the file was likely modified (e.g., timestamp or size if possible)
        const stats = await fs.stat(outputPath);
        expect(stats.size).toBeGreaterThan(0); // Basic check
      } else { // Library path
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_STYLE_APPLY');
        }
      }
    });

     test('Library path should fail for "selection" range', async () => {
        if (useComInterop === false || useComInterop === undefined) {
            const outputFileName = `doc_style_selection_fail_lib.docx`;
            const relativeOutputPath = `tests/temp_word_styles_dir/${outputFileName}`;
             // No need to copy file as it should fail before file ops for this specific check

            const args = {
                filePath: relativeOutputPath, // File might not even be read
                style: "Normal",
                range: "selection",
                useComInterop,
            };
            const result = await callTool('word/styles/apply', args);
            expect(result.success).toBe(false);
            if (!result.success) {
                expect(result.error.code).toBe('LIB_RANGE_NOT_SUPPORTED');
            }
        } else {
            expect(true).toBe(true); // Test only for library path
        }
    });
  });

  test('word/styles/list should fail for non-existent file (COM)', async () => {
    const args = {
      filePath: 'non_existent_doc_for_styles.docx',
      useComInterop: true, // COM path to check file existence handling
    };
    const result = await callTool('word/styles/list', args);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('WORD_STYLE_ERROR_COM'); // Generic COM error
      expect(result.error.message).toMatch(/Failed to open document|File not found/i);
    }
  });

  test('word/styles/apply should fail for non-existent file (COM)', async () => {
    const args = {
      filePath: 'non_existent_doc_for_style_apply.docx',
      style: 'Normal',
      range: 'paragraph:1',
      useComInterop: true, // COM path
    };
    const result = await callTool('word/styles/apply', args);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('WORD_STYLE_ERROR_COM');
      expect(result.error.message).toMatch(/Failed to open document|File not found/i);
    }
  });
});