import path from 'path';
import fs from 'fs-extra';
import { insertTable, insertTableFromArray, extractTableData } from '../../../src/tools/word/tables.tool.js';
import { getText } from '../../../src/tools/word/text.tool.js';
import { createTempWordWorkspace } from '../../helpers/wordFixtures.js';

describe('word/tables integration', () => {
  let tempDir: string;

  beforeAll(async () => {
    tempDir = await createTempWordWorkspace('mcp-office-tables-');
  });

  afterAll(async () => {
    if (tempDir) {
      await fs.remove(tempDir);
    }
  });

  test('insertTable creates a new document on the library path', async () => {
    const outputPath = path.join(tempDir, 'blank-table.docx');

    const result = await insertTable(
      {
        filePath: outputPath,
        rows: 2,
        columns: 3,
        useComInterop: false,
      },
      { log: console } as any,
    );

    expect(result.success).toBe(true);
    await expect(fs.pathExists(outputPath)).resolves.toBe(true);
  });

  test('insertTableFromArray creates a readable table document on the library path', async () => {
    const outputPath = path.join(tempDir, 'array-table.docx');

    const result = await insertTableFromArray(
      {
        filePath: outputPath,
        data: [
          ['Region', 'Revenue', 'Status'],
          ['North', '1.2M', 'Recovered'],
          ['South', '0.8M', 'At Risk'],
        ],
        useComInterop: false,
      },
      { log: console } as any,
    );

    expect(result.success).toBe(true);
    await expect(fs.pathExists(outputPath)).resolves.toBe(true);

    const textResult = await getText(
      {
        filePath: outputPath,
        range: 'document',
        useComInterop: false,
      },
      { log: console } as any,
    );

    expect(textResult.success).toBe(true);
    if (textResult.success) {
      expect(textResult.data).toContain('Region');
      expect(textResult.data).toContain('North');
      expect(textResult.data).toContain('At Risk');
    }
  });

  test('library path rejects insertion into existing files and keeps extractData marked as pending', async () => {
    const existingPath = path.join(tempDir, 'existing-table.docx');
    await fs.writeFile(existingPath, 'placeholder');

    const insertResult = await insertTable(
      {
        filePath: existingPath,
        rows: 1,
        columns: 1,
        useComInterop: false,
      },
      { log: console } as any,
    );
    expect(insertResult.success).toBe(false);
    if (!insertResult.success) {
      expect(insertResult.error.code).toBe('LIB_INSERT_EXISTING_FILE_NOT_SUPPORTED');
    }

    const extractResult = await extractTableData(
      {
        filePath: existingPath,
        tableIndex: 1,
        useComInterop: false,
      },
      { log: console } as any,
    );
    expect(extractResult.success).toBe(false);
    if (!extractResult.success) {
      expect(extractResult.error.code).toBe('NOT_IMPLEMENTED_LIB_TABLE_EXTRACTION');
    }
  });
});
