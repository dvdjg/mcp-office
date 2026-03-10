import path from 'path';
import fs from 'fs-extra';
import AdmZip from 'adm-zip';
import { exportToMarkdown } from '../../../src/tools/word/markdown.tool.js';
import { getText } from '../../../src/tools/word/text.tool.js';
import { createComplexWordFixture, createTempWordWorkspace } from '../../helpers/wordFixtures.js';

describe('word functional conversions with generated documents', () => {
  let tempDir: string;
  let docxPath: string;

  beforeAll(async () => {
    tempDir = await createTempWordWorkspace();
    ({ docxPath } = await createComplexWordFixture(tempDir));
  });

  afterAll(async () => {
    if (tempDir) {
      await fs.remove(tempDir);
    }
  });

  test('getText extracts the full textual structure from a generated complex document', async () => {
    const result = await getText(
      {
        filePath: docxPath,
        range: 'document',
        useComInterop: false,
      },
      { log: console } as any,
    );

    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data).toContain('Quarterly Review');
      expect(result.data).toContain('Regional Insights');
      expect(result.data).toContain('North region momentum recovered');
      expect(result.data).toContain('Prepare discount sensitivity analysis.');
      expect(result.data).toContain('Action Tracker');
      expect(result.data).toContain('Recovered');
    }
  });

  test('exportToMarkdown keeps headings, list items and table content using html_tables mode', async () => {
    const outputPath = path.join(tempDir, 'complex-structure-html.md');
    const result = await exportToMarkdown(
      {
        filePath: docxPath,
        output: outputPath,
        tableFormat: 'html_tables',
        imageDir: 'images',
        comments: 'ignore',
        useComInterop: false,
      },
      { log: console } as any,
    );

    expect(result.success).toBe(true);
    await expect(fs.pathExists(outputPath)).resolves.toBe(true);

    const markdown = await fs.readFile(outputPath, 'utf8');
    expect(markdown).toContain('# Quarterly Review');
    expect(markdown).toMatch(/\*\*critical\*\*/);
    expect(markdown).toMatch(/[_*]follow-up actions[_*]/);
    expect(markdown).toMatch(/[-*]\s+South region requires pricing review\./);
    expect(markdown).toContain('<table>');
    expect(markdown).toContain('<td><p>North</p></td>');
    expect(markdown).toContain('<td><p>At Risk</p></td>');
  });

  test('exportToMarkdown can package a complex conversion into a zip artifact', async () => {
    const outputPath = path.join(tempDir, 'complex-structure-zip.md');
    const zipPath = path.join(tempDir, 'complex-structure.zip');
    const result = await exportToMarkdown(
      {
        filePath: docxPath,
        output: outputPath,
        tableFormat: 'html_tables',
        imageDir: 'images',
        comments: 'ignore',
        zipOutput: true,
        zipFileName: path.basename(zipPath),
        useComInterop: false,
      },
      { log: console } as any,
    );

    expect(result.success).toBe(true);
    await expect(fs.pathExists(zipPath)).resolves.toBe(true);

    const zip = new AdmZip(zipPath);
    const entryNames = zip.getEntries().map((entry) => entry.entryName);
    expect(entryNames).toContain('complex-structure-zip.md');

    const markdown = zip.readAsText('complex-structure-zip.md');
    expect(markdown).toContain('Quarterly Review');
    expect(markdown).toContain('<table>');
    expect(markdown).toContain('<td><p>Region</p></td>');
    expect(markdown).toContain('North');
    expect(markdown).toContain('South');
  });
});
