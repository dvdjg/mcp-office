import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000';
const TEMP_DIR = path.join(__dirname, '../temp_word_metadata_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');

const FIXTURE_CV_DOC_NAME = 'CV.docx'; // Document with existing metadata/comments
const FIXTURE_CV_DOC_PATH = path.join(FIXTURES_DIR, FIXTURE_CV_DOC_NAME);
const RELATIVE_FIXTURE_CV_DOC_PATH = `tests/fixtures/${FIXTURE_CV_DOC_NAME}`;

const FIXTURE_EMPTY_DOC_NAME = 'empty_doc.docx';
const FIXTURE_EMPTY_DOC_PATH = path.join(FIXTURES_DIR, FIXTURE_EMPTY_DOC_NAME);
// const RELATIVE_FIXTURE_EMPTY_DOC_PATH = `tests/fixtures/${FIXTURE_EMPTY_DOC_NAME}`; // Not used directly in requests, but for setup

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

const callTool = async (args: Record<string, any>): Promise<ToolResponse> => {
  const response = await fetch(`${MCP_SERVER_URL}/tool`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ tool_name: 'word/metadata', arguments: args }),
  });
  return response.json() as Promise<ToolResponse>;
};

describe('word/metadata e2e tests', () => {
  beforeAll(async () => {
    await fs.rm(TEMP_DIR, { recursive: true, force: true });
    await fs.mkdir(TEMP_DIR, { recursive: true });

    const fixturesToVerify = [FIXTURE_CV_DOC_PATH, FIXTURE_EMPTY_DOC_PATH];
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
  });

  afterAll(async () => {
    await fs.rm(TEMP_DIR, { recursive: true, force: true });
  });

  const testModes = [
    { mode: 'COM', useComInterop: true },
    { mode: 'Library', useComInterop: false },
    { mode: 'Library (Default)', useComInterop: undefined },
  ];

  describe.each(testModes)('operation: set (mode: $mode)', ({ useComInterop }) => {
    test('should set a document property', async () => {
      const outputFileName = `doc_set_property_${useComInterop}.docx`;
      const outputPath = path.join(TEMP_DIR, outputFileName);
      const relativeOutputPath = `tests/temp_word_metadata_dir/${outputFileName}`;
      
      const propertyName = 'Title';
      const propertyValue = `Test Document Title - ${useComInterop}`;

      if (useComInterop) {
        await fs.copyFile(FIXTURE_CV_DOC_PATH, outputPath); // COM modifies existing
      }
      // Library path for 'set' creates a new file

      const args = {
        filePath: relativeOutputPath,
        operation: 'set',
        propertyName,
        propertyValue,
        useComInterop,
      };
      const result = await callTool(args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        await expect(fs.stat(outputPath)).resolves.toBeTruthy();
        // TODO: Verify property was actually set in COM path (requires reading it back)
      } else { // Library path
        let fileExisted = false;
        try { await fs.stat(outputPath); fileExisted = true; } catch(e) {/*ok*/}

        if (fileExisted && (useComInterop === false || useComInterop === undefined)) { // Library path 'set' fails if file exists
            expect(result.success).toBe(false);
            if(!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_MODIFY');
        } else { // Library path creates new file
            expect(result.success).toBe(true);
            await expect(fs.stat(outputPath)).resolves.toBeTruthy();
            // TODO: Verify property in new file for Library path (requires reading it back, which is also limited)
        }
      }
    });

    test('Library path should fail to set property on existing file', async () => {
        if (useComInterop === false || useComInterop === undefined) {
            const outputFileName = `existing_doc_lib_set_prop_fail.docx`;
            const outputPath = path.join(TEMP_DIR, outputFileName);
            const relativeOutputPath = `tests/temp_word_metadata_dir/${outputFileName}`;
            await fs.copyFile(FIXTURE_EMPTY_DOC_PATH, outputPath); // Create an existing file

            const args = {
                filePath: relativeOutputPath,
                operation: 'set',
                propertyName: 'Author',
                propertyValue: 'Test Author Lib Fail',
                useComInterop,
            };
            const result = await callTool(args);
            expect(result.success).toBe(false);
            if (!result.success) {
                expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_MODIFY');
            }
        } else {
            expect(true).toBe(true); // Test only for library path
        }
    });
  });

  describe.each(testModes)('operation: get (property) (mode: $mode)', ({ useComInterop }) => {
    test('should get a specific document property', async () => {
      const propertyName = 'Author'; // Assuming CV.docx has an Author
      const args = {
        filePath: RELATIVE_FIXTURE_CV_DOC_PATH,
        operation: 'get',
        propertyName,
        useComInterop,
      };
      const result = await callTool(args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        if (result.success) {
          expect(result.data.name).toBe(propertyName);
          expect(result.data.value).toBeDefined(); // Value can be anything
        }
      } else { // Library path
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_PROP_GET');
        }
      }
    });
  });

  describe.each(testModes)('operation: manage (list properties/comments) (mode: $mode)', ({ useComInterop }) => {
    test('should list all document properties and comments', async () => {
      const args = {
        filePath: RELATIVE_FIXTURE_CV_DOC_PATH,
        operation: 'manage',
        useComInterop,
      };
      const result = await callTool(args);
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toHaveProperty('properties');
        expect(result.data).toHaveProperty('comments');
        expect(Array.isArray(result.data.comments)).toBe(true);

        if (useComInterop) {
          expect(Array.isArray(result.data.properties)).toBe(true);
          expect(result.data.properties.length).toBeGreaterThan(0);
        } else {
          // Library path for properties is limited
          expect(result.data.properties).toBe("Property listing via library path is limited/not implemented.");
        }
        // Further checks for comments if CV.docx is known to have them
      }
    });
  });
  
  // Comment-related tests (add, get, remove) are primarily for COM path due to library limitations
  describe('COM Only: Comment Operations', () => {
    const useComInterop = true;
    const tempDocName = 'comment_test_doc.docx';
    const tempDocPath = path.join(TEMP_DIR, tempDocName);
    const relativeTempDocPath = `tests/temp_word_metadata_dir/${tempDocName}`;

    beforeEach(async () => {
      await fs.copyFile(FIXTURE_EMPTY_DOC_PATH, tempDocPath);
      // Add some initial text to comment on
      const addTextResponse = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          tool_name: 'word/text/insert',
          arguments: {
            filePath: relativeTempDocPath,
            text: "This is the first paragraph for comments.\nThis is the second paragraph.",
            position: "start",
            useComInterop: true // Use COM to prepare the doc
          }
        }),
      });
      const addTextResult = await addTextResponse.json() as ToolResponse;
      if (!addTextResult.success) {
        console.error("Failed to prepare document for comment tests:", addTextResult.error);
        throw new Error("Comment test setup failed: could not insert initial text.");
      }
    });

    test('should add, get, and remove a comment (COM)', async () => {
      // Add comment
      const addArgs = {
        filePath: relativeTempDocPath,
        operation: 'add',
        commentText: 'This is a test comment.',
        range: { start: 0, end: 10 }, // Comment on "This is th"
        useComInterop,
      };
      let result = await callTool(addArgs);
      expect(result.success).toBe(true);
      if(!result.success) console.error("Add comment failed:", result.error);


      // Get the added comment (assuming it's the first one)
      const getArgs = {
        filePath: relativeTempDocPath,
        operation: 'get',
        commentIndex: 1,
        useComInterop,
      };
      result = await callTool(getArgs);
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data.index).toBe(1);
        expect(result.data.text).toContain('This is a test comment.'); // COM might add context
      } else {
        console.error("Get comment failed:", result.error);
      }

      // Remove the comment
      const removeArgs = {
        filePath: relativeTempDocPath,
        operation: 'remove',
        commentIndex: 1,
        useComInterop,
      };
      result = await callTool(removeArgs);
      expect(result.success).toBe(true);
       if(!result.success) console.error("Remove comment failed:", result.error);

      // Verify comment is gone by trying to get it again
      result = await callTool(getArgs);
      expect(result.success).toBe(false); // Should fail as comment is removed
      if (!result.success) {
        expect(result.error.code).toBe('OUT_OF_BOUNDS');
      }
    });
  });

   describe.each(testModes)('operation: get (comment) (mode: $mode)', ({ useComInterop }) => {
    // This test relies on CV.docx having comments or Mammoth being able to parse some.
    // It's a bit fragile for library path.
    test('should get a specific comment or report error/not implemented', async () => {
      const args = {
        filePath: RELATIVE_FIXTURE_CV_DOC_PATH, // Assuming CV.docx might have comments
        operation: 'get',
        commentIndex: 1,
        useComInterop,
      };
      const result = await callTool(args);

      if (useComInterop) {
        // COM path might succeed if comment exists, or fail if not.
        // This depends heavily on CV.docx content.
        if (result.success) {
            expect(result.data.index).toBe(1);
            expect(result.data.text).toBeDefined();
        } else {
            expect(result.error.code).toBe('OUT_OF_BOUNDS'); // If no comments or index too high
        }
      } else { // Library path
        // Library path comment extraction is best-effort via HTML parsing.
        // It might find something or report an error if parsing fails or no comments are found.
        if (result.success) {
            expect(result.data.index).toBe(1);
            expect(result.data.text).toBeDefined();
            console.warn("Library path succeeded getting comment, HTML parsing worked.");
        } else {
            expect(result.error.code).toMatch(/OUT_OF_BOUNDS_LIB|NOT_FOUND_LIB/); // Or other parsing related error
            console.warn(`Library path failed to get comment as potentially expected: ${result.error.message}`);
        }
      }
    });
  });


  test('word/metadata (set) should fail for non-existent file (COM)', async () => {
    const args = {
      filePath: 'non_existent_doc_for_metadata_set.docx',
      operation: 'set',
      propertyName: 'Title',
      propertyValue: 'Wont be set',
      useComInterop: true,
    };
    const result = await callTool(args);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('METADATA_COM_OPERATION_ERROR'); // Generic COM error for file open
      expect(result.error.message).toMatch(/Failed to open document|File not found/i);
    }
  });
});