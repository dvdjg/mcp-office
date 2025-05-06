import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000';
const TEMP_DIR = path.join(__dirname, '../temp_word_tables_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');

const FIXTURE_EMPTY_DOC_NAME = 'empty_doc.docx';
const FIXTURE_EMPTY_DOC_PATH = path.join(FIXTURES_DIR, FIXTURE_EMPTY_DOC_NAME);
const RELATIVE_FIXTURE_EMPTY_DOC_PATH = `tests/fixtures/${FIXTURE_EMPTY_DOC_NAME}`;

const FIXTURE_DOC_WITH_TABLE_NAME = 'doc_with_table.docx'; // Contains at least one table
const FIXTURE_DOC_WITH_TABLE_PATH = path.join(FIXTURES_DIR, FIXTURE_DOC_WITH_TABLE_NAME);
const RELATIVE_FIXTURE_DOC_WITH_TABLE_PATH = `tests/fixtures/${FIXTURE_DOC_WITH_TABLE_NAME}`;

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

describe('word/tables e2e tests', () => {
  beforeAll(async () => {
    await fs.rm(TEMP_DIR, { recursive: true, force: true });
    await fs.mkdir(TEMP_DIR, { recursive: true });

    const fixturesToVerify = [FIXTURE_EMPTY_DOC_PATH, FIXTURE_DOC_WITH_TABLE_PATH];
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

  describe.each(testModes)('word/tables/insert (mode: $mode)', ({ useComInterop }) => {
    test('should insert a table into a Word document', async () => {
      const outputFileName = `doc_inserted_table_${useComInterop}.docx`;
      const outputPath = path.join(TEMP_DIR, outputFileName);
      const relativeOutputPath = `tests/temp_word_tables_dir/${outputFileName}`;

      // COM path modifies existing, Library path creates new
      if (useComInterop) {
        await fs.copyFile(FIXTURE_EMPTY_DOC_PATH, outputPath);
      }

      const args = {
        filePath: relativeOutputPath,
        rows: 3,
        columns: 4,
        position: 'end',
        style: 'Table Grid', // Style might only be effective in COM
        useComInterop,
      };
      const result = await callTool('word/tables/insert', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        await expect(fs.stat(outputPath)).resolves.toBeTruthy();
        // TODO: Verify table structure in COM path
      } else { // Library path
        let fileExistedBeforeToolCall = false;
        try {
            await fs.stat(outputPath);
            fileExistedBeforeToolCall = true;
        } catch (e) {
            // File didn't exist, which is expected for library path to succeed
        }

        if (fileExistedBeforeToolCall) {
             expect(result.success).toBe(false);
             if(!result.success) expect(result.error.code).toBe('LIB_INSERT_EXISTING_FILE_NOT_SUPPORTED');
        } else {
            expect(result.success).toBe(true);
            await expect(fs.stat(outputPath)).resolves.toBeTruthy();
             // TODO: Verify table structure in new file for Library path
        }
      }
    });

    test('Library path should fail if trying to insert into existing file', async () => {
        if (useComInterop === false || useComInterop === undefined) {
            const outputFileName = `existing_doc_lib_insert_fail.docx`;
            const outputPath = path.join(TEMP_DIR, outputFileName);
            const relativeOutputPath = `tests/temp_word_tables_dir/${outputFileName}`;
            await fs.copyFile(FIXTURE_EMPTY_DOC_PATH, outputPath); // Create an existing file

            const args = {
                filePath: relativeOutputPath,
                rows: 2,
                columns: 2,
                useComInterop,
            };
            const result = await callTool('word/tables/insert', args);
            expect(result.success).toBe(false);
            if (!result.success) {
                expect(result.error.code).toBe('LIB_INSERT_EXISTING_FILE_NOT_SUPPORTED');
            }
        } else {
            // This test is only for library path
            expect(true).toBe(true);
        }
    });
  });

  describe.each(testModes)('word/tables/insertFromArray (mode: $mode)', ({ useComInterop }) => {
    const tableData = [['Name', 'Age'], ['Alice', 30], ['Bob', 24]];
    test('should insert a table from array into a Word document', async () => {
      const outputFileName = `doc_inserted_array_table_${useComInterop}.docx`;
      const outputPath = path.join(TEMP_DIR, outputFileName);
      const relativeOutputPath = `tests/temp_word_tables_dir/${outputFileName}`;

      if (useComInterop) {
        await fs.copyFile(FIXTURE_EMPTY_DOC_PATH, outputPath);
      }

      const args = {
        filePath: relativeOutputPath,
        data: tableData,
        position: 'end',
        styleName: 'Light Shading - Accent 1', // Style might only be effective in COM
        useComInterop,
      };
      const result = await callTool('word/tables/insertFromArray', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        await expect(fs.stat(outputPath)).resolves.toBeTruthy();
        // TODO: Verify table content and structure in COM path
      } else { // Library path
        let fileExistedBeforeToolCall = false;
        try {
            await fs.stat(outputPath);
            fileExistedBeforeToolCall = true;
        } catch (e) {
            // File didn't exist, which is expected for library path to succeed
        }
        if (fileExistedBeforeToolCall) {
             expect(result.success).toBe(false);
             if(!result.success) expect(result.error.code).toBe('LIB_INSERT_ARRAY_EXISTING_FILE_NOT_SUPPORTED');
        } else {
            expect(result.success).toBe(true);
            await expect(fs.stat(outputPath)).resolves.toBeTruthy();
            // TODO: Verify table content and structure in new file for Library path
        }
      }
    });
     test('Library path should fail if trying to insertFromArray into existing file', async () => {
        if (useComInterop === false || useComInterop === undefined) {
            const outputFileName = `existing_doc_lib_insertarray_fail.docx`;
            const outputPath = path.join(TEMP_DIR, outputFileName);
            const relativeOutputPath = `tests/temp_word_tables_dir/${outputFileName}`;
            await fs.copyFile(FIXTURE_EMPTY_DOC_PATH, outputPath); // Create an existing file

            const args = {
                filePath: relativeOutputPath,
                data: [['test']],
                useComInterop,
            };
            const result = await callTool('word/tables/insertFromArray', args);
            expect(result.success).toBe(false);
            if (!result.success) {
                expect(result.error.code).toBe('LIB_INSERT_ARRAY_EXISTING_FILE_NOT_SUPPORTED');
            }
        } else {
            expect(true).toBe(true);
        }
    });
  });

  describe.each(testModes)('word/tables/extractData (mode: $mode)', ({ useComInterop }) => {
    test('should extract data from a table in a Word document', async () => {
      const args = {
        filePath: RELATIVE_FIXTURE_DOC_WITH_TABLE_PATH,
        tableIndex: 1, // Assuming the first table in FIXTURE_DOC_WITH_TABLE_NAME
        useComInterop,
      };
      const result = await callTool('word/tables/extractData', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        if (result.success) {
          expect(Array.isArray(result.data)).toBe(true);
          // Example check, assuming FIXTURE_DOC_WITH_TABLE_NAME has a known table
          // For 'doc_with_table.docx', let's assume a 2x2 table:
          // Header1 | Header2
          // Cell1   | Cell2
          // This check needs to be adjusted based on the actual content of 'doc_with_table.docx'
          // For now, a generic check:
          expect(result.data.length).toBeGreaterThanOrEqual(1); // At least one row
          if (result.data.length > 0) {
            expect(result.data[0].length).toBeGreaterThanOrEqual(1); // At least one cell in the first row
          }
        }
      } else { // Library path
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_TABLE_EXTRACTION');
        }
      }
    });

    test('should return error for out-of-bounds table index (COM)', async () => {
        if (useComInterop) {
            const args = {
                filePath: RELATIVE_FIXTURE_DOC_WITH_TABLE_PATH,
                tableIndex: 999, // Assuming this index is out of bounds
                useComInterop: true,
            };
            const result = await callTool('word/tables/extractData', args);
            expect(result.success).toBe(false);
            if (!result.success) {
                expect(result.error.code).toBe('WORD_TABLE_EXTRACT_DATA_FAILED_COM'); // Or a more specific out of bounds error
                expect(result.error.message).toContain('out of bounds');
            }
        } else {
            expect(true).toBe(true); // Test only for COM path
        }
    });
  });

  test('word/tables/extractData should fail for non-existent file', async () => {
    const args = {
      filePath: 'non_existent_document_for_table_extract.docx',
      tableIndex: 1,
      useComInterop: true, // COM path to check file existence handling
    };
    const result = await callTool('word/tables/extractData', args);
    expect(result.success).toBe(false);
    if (!result.success) {
      // COM error for file open might be generic
      expect(result.error.code).toBe('WORD_TABLE_EXTRACT_DATA_FAILED_COM');
      expect(result.error.message).toMatch(/Failed to open document|File not found/i);
    }
  });
});