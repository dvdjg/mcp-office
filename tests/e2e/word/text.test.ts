import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
import * as fs from 'fs/promises';
import * as path from 'path';
// import fetch from 'node-fetch'; // Replaced with MCP SDK
import { Client } from '@modelcontextprotocol/sdk/client'; // Ensuring this is the path used
import { SSEClientTransport } from '@modelcontextprotocol/sdk/client/sse.js';
// Note: We might need to import specific error types from the SDK later if generic error handling is insufficient.

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

const callTool = async (toolName: string, toolArgs: Record<string, any>): Promise<ToolResponse> => {
  const client = new Client(
    {
      name: "e2e-test-client",
      version: "1.0.0",
    },
    { capabilities: {} },
    // Authentication method to be determined and added here
  );

  // const apiKey = process.env.MCP_AUTH_TOKEN || 'default-secret-token'; // API key will be handled by SDK auth mechanism
  const transport = new SSEClientTransport(
    new URL(`${MCP_SERVER_URL}/tool`)
    // Options for SSEClientTransport, if any, to be determined for auth.
    // The 'headers' option was found to be invalid.
  );

  try {
    await client.connect(transport);
    // Assuming client.tools.invoke is the method to call tools.
    // The 'toolArgs' are passed directly as the parameters for the tool.
    const sdkResult = await client.tools.invoke(toolName, toolArgs);

    // Adapt the SDK's result (likely ContentResult) to the existing ToolResponse structure.
    // ContentResult is typically { content: [{ type: 'text', text: '...' }, ...] }
    let responseData: any = null;
    if (sdkResult && sdkResult.content) {
      if (sdkResult.content.length > 0) {
        // For simplicity, if the first part is text, use its text. Otherwise, use the first part.
        // Tests might need adjustment if they expect more complex data structures.
        if (sdkResult.content[0].type === 'text') {
          responseData = sdkResult.content[0].text;
        } else {
          responseData = sdkResult.content[0];
        }
      } else {
        responseData = ""; // Or an empty object, if tools return empty content array for no data
      }
    } else {
      // If the result is not a standard ContentResult, pass it as is.
      // This might happen for tools that return simple strings directly (though FastMCP wraps them).
      responseData = sdkResult;
    }

    return {
      success: true,
      data: responseData,
    };
  } catch (error: any) {
    // Map errors from the SDK to the existing ErrorResponse structure.
    let errorCode = "TOOL_EXECUTION_FAILED_SDK";
    let errorMessage = "Unknown error during SDK tool call";
    let errorDetails: any = null;

    if (error && typeof error.message === 'string') {
      errorMessage = error.message;
    }

    // Attempt to extract code and details, this is speculative based on common error patterns.
    // FastMCP server throws UserError with message and details. SDK might preserve these.
    if (error && error.details) {
        errorDetails = error.details;
        if (typeof error.details.code === 'string') {
            errorCode = error.details.code;
        } else if (error.name === 'UserError') { // From fastmcp UserError on server
             errorCode = "USER_ERROR_FROM_SERVER";
        }
    } else if (error && typeof error.code === 'string') {
        errorCode = error.code;
    } else if (error && error.name === 'UserError') { // Check error.name if details.code is not present
        errorCode = "USER_ERROR_SDK";
    }

    // If error itself might be the details (e.g., a simple string error from some part of the stack)
    if (!errorDetails && error && !(error instanceof Error) && typeof error !== 'function') {
        errorDetails = error;
    }


    return {
      success: false,
      error: {
        code: errorCode,
        message: errorMessage,
        details: errorDetails,
      },
    };
  } finally {
    if (client.connected) {
      await client.disconnect();
    }
  }
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