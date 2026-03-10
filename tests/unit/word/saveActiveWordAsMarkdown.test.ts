import { jest } from '@jest/globals';
import path from 'path';
import fs from 'fs-extra';
import saveActiveWordAsMarkdownTool from '../../../src/tools/word/saveActiveWordAsMarkdown.tool';
import getActiveOfficeDocumentsTool from '../../../src/tools/os/getActiveOfficeDocuments.tool';
import { wordMarkdownExportTool } from '../../../src/tools/word/markdown.tool';

describe('word/saveActiveWordAsMarkdown', () => {
  const context = {
    log: {
      info: jest.fn(),
      warn: jest.fn(),
      error: jest.fn(),
      debug: jest.fn(),
    },
    reportProgress: jest.fn().mockResolvedValue(undefined),
    session: new Map(),
  };

  beforeEach(() => {
    jest.restoreAllMocks();
    jest.spyOn(fs, 'pathExists').mockResolvedValue(true as any);
    jest.spyOn(fs, 'ensureDir').mockResolvedValue(undefined as any);
  });

  test('exports the only active Word document', async () => {
    jest.spyOn(getActiveOfficeDocumentsTool, 'handler').mockResolvedValue({
      success: true,
      data: { documents: [{ filePath: 'C:\\Docs\\report.docx', applicationType: 'Word' }] },
    } as any);
    const exportSpy = jest.spyOn(wordMarkdownExportTool, 'handler').mockResolvedValue({
      success: true,
      data: { outputPath: 'C:\\Docs\\report.md' },
    } as any);

    const result = await saveActiveWordAsMarkdownTool.handler({}, context as any);

    expect(result.success).toBe(true);
    expect(exportSpy).toHaveBeenCalledWith(
      expect.objectContaining({
        filePath: 'C:\\Docs\\report.docx',
        output: 'C:\\Docs\\report.md',
      }),
      context,
    );
  });

  test('requires inputFilePath when multiple Word documents are active', async () => {
    jest.spyOn(getActiveOfficeDocumentsTool, 'handler').mockResolvedValue({
      success: true,
      data: {
        documents: [
          { filePath: 'C:\\Docs\\one.docx', applicationType: 'Word' },
          { filePath: 'C:\\Docs\\two.docx', applicationType: 'Word' },
        ],
      },
    } as any);

    const result = await saveActiveWordAsMarkdownTool.handler({}, context as any);

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('MULTIPLE_ACTIVE_WORD_DOCS');
    }
  });

  test('creates the output directory when needed', async () => {
    jest.spyOn(getActiveOfficeDocumentsTool, 'handler').mockResolvedValue({
      success: true,
      data: { documents: [{ filePath: 'C:\\Docs\\report.docx', applicationType: 'Word' }] },
    } as any);
    jest.spyOn(wordMarkdownExportTool, 'handler').mockResolvedValue({
      success: true,
      data: { outputPath: 'C:\\Out\\report.md' },
    } as any);
    const pathExistsSpy = jest.spyOn(fs, 'pathExists').mockResolvedValue(false as any);
    const ensureDirSpy = jest.spyOn(fs, 'ensureDir').mockResolvedValue(undefined as any);

    const result = await saveActiveWordAsMarkdownTool.handler({ outputFilePath: 'C:\\Out\\report.md' }, context as any);

    expect(result.success).toBe(true);
    expect(pathExistsSpy).toHaveBeenCalledWith(path.dirname('C:\\Out\\report.md'));
    expect(ensureDirSpy).toHaveBeenCalledWith(path.dirname('C:\\Out\\report.md'));
  });
});
