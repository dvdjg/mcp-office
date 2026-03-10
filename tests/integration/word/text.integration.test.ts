import path from 'path';
import fs from 'fs-extra';
import { getText, insertText, modifyText, deleteText } from '../../../src/tools/word/text.tool.js';
import { createComplexWordFixture, createTempWordWorkspace } from '../../helpers/wordFixtures.js';

describe('word/text integration', () => {
  let tempDir: string;
  let docxPath: string;

  beforeAll(async () => {
    tempDir = await createTempWordWorkspace('mcp-office-text-');
    ({ docxPath } = await createComplexWordFixture(tempDir, 'text-source.docx'));
  });

  afterAll(async () => {
    if (tempDir) {
      await fs.remove(tempDir);
    }
  });

  test('getText returns the full document content through the library path', async () => {
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
      expect(result.data).toContain('Action Tracker');
    }
  });

  test('getText reports that paragraph ranges require COM on the library path', async () => {
    const result = await getText(
      {
        filePath: docxPath,
        range: 'paragraph:1',
        useComInterop: false,
      },
      { log: console } as any,
    );

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('RANGE_REQUIRES_COM_LIB');
    }
  });

  test('insertText, modifyText and deleteText leave an explicit not-implemented contract on the library path', async () => {
    const missingInsertTarget = path.join(tempDir, 'library-insert.docx');

    const insertResult = await insertText(
      {
        filePath: missingInsertTarget,
        text: 'Append this text',
        position: 'end',
        useComInterop: false,
      },
      { log: console } as any,
    );
    expect(insertResult.success).toBe(false);
    if (!insertResult.success) {
      expect(insertResult.error.code).toBe('NOT_IMPLEMENTED_LIB');
    }

    const modifyResult = await modifyText(
      {
        filePath: docxPath,
        range: 'document',
        newText: 'Replacement text',
        useComInterop: false,
      },
      { log: console } as any,
    );
    expect(modifyResult.success).toBe(false);
    if (!modifyResult.success) {
      expect(modifyResult.error.code).toBe('NOT_IMPLEMENTED_LIB');
    }

    const deleteResult = await deleteText(
      {
        filePath: docxPath,
        range: 'paragraph:1',
        useComInterop: false,
      },
      { log: console } as any,
    );
    expect(deleteResult.success).toBe(false);
    if (!deleteResult.success) {
      expect(deleteResult.error.code).toBe('NOT_IMPLEMENTED_LIB');
    }
  });
});
