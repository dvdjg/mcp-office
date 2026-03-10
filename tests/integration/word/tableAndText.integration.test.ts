import path from 'path';
import fs from 'fs-extra';
import { exportToMarkdown } from '../../../src/tools/word/markdown.tool.js';
import { insertTableFromArray } from '../../../src/tools/word/tables.tool.js';
import { createTempWordWorkspace } from '../../helpers/wordFixtures.js';

describe('word table + text integration', () => {
  let tempDir: string;

  beforeAll(async () => {
    tempDir = await createTempWordWorkspace('mcp-office-table-text-');
  });

  afterAll(async () => {
    if (tempDir) {
      await fs.remove(tempDir);
    }
  });

  test('a table created from array can be exported as markdown/html for downstream conversions', async () => {
    const docxPath = path.join(tempDir, 'table-pipeline.docx');
    const mdPath = path.join(tempDir, 'table-pipeline.md');

    const tableResult = await insertTableFromArray(
      {
        filePath: docxPath,
        data: [
          ['Milestone', 'Owner'],
          ['Prototype', 'Ana'],
          ['Validation', 'Luis'],
        ],
        useComInterop: false,
      },
      { log: console } as any,
    );

    expect(tableResult.success).toBe(true);

    const exportResult = await exportToMarkdown(
      {
        filePath: docxPath,
        output: mdPath,
        tableFormat: 'html_tables',
        imageDir: 'images',
        comments: 'ignore',
        useComInterop: false,
      },
      { log: console } as any,
    );

    expect(exportResult.success).toBe(true);
    const markdown = await fs.readFile(mdPath, 'utf8');
    expect(markdown).toContain('<table>');
    expect(markdown).toContain('Milestone');
    expect(markdown).toContain('Prototype');
    expect(markdown).toContain('Luis');
  });
});
