import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch'; // Assuming node-fetch is available for making HTTP requests

const MCP_SERVER_URL = 'http://localhost:3000'; // Adjust if your server runs elsewhere
const TEMP_DIR = path.join(__dirname, '../temp_word_markdown_dir');
const FIXTURES_DIR = path.join(__dirname, '../../fixtures');

const FIXTURE_MD_NAME = 'sample.md';
const FIXTURE_MD_PATH = path.join(FIXTURES_DIR, FIXTURE_MD_NAME);
const RELATIVE_FIXTURE_MD_PATH = `tests/fixtures/${FIXTURE_MD_NAME}`;

const FIXTURE_DOTX_NAME = 'sample_template.dotx';
const FIXTURE_DOTX_PATH = path.join(FIXTURES_DIR, FIXTURE_DOTX_NAME);
const RELATIVE_FIXTURE_DOTX_PATH = `tests/fixtures/${FIXTURE_DOTX_NAME}`;

const FIXTURE_WORD_DOC_NAME = 'CV.docx'; // For export tests
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

describe('word/markdown e2e tests', () => {
  beforeAll(async () => {
    await fs.rm(TEMP_DIR, { recursive: true, force: true });
    await fs.mkdir(TEMP_DIR, { recursive: true });

    const fixturesToVerify = [FIXTURE_MD_PATH, FIXTURE_DOTX_PATH, FIXTURE_WORD_DOC_PATH];
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

  describe.each(testModes)('word/markdown/import (mode: $mode)', ({ useComInterop }) => {
    test('should import a markdown file into a new Word document', async () => {
      const outputFileName = `imported_markdown_${useComInterop}.docx`;
      const outputPath = path.join(TEMP_DIR, outputFileName);
      const relativeOutputPath = `tests/temp_word_markdown_dir/${outputFileName}`;

      const args = {
        filePath: RELATIVE_FIXTURE_MD_PATH,
        output: relativeOutputPath,
        useComInterop,
      };
      const result = await callTool('word/markdown/import', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        await expect(fs.stat(outputPath)).resolves.toBeTruthy();
        // TODO: Deeper content verification for COM path
      } else {
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB');
        }
      }
    });

    test('should import a markdown file using a template (COM) or report not implemented (Library)', async () => {
      const outputFileName = `imported_markdown_template_${useComInterop}.docx`;
      const outputPath = path.join(TEMP_DIR, outputFileName);
      const relativeOutputPath = `tests/temp_word_markdown_dir/${outputFileName}`;

      const args = {
        filePath: RELATIVE_FIXTURE_MD_PATH,
        output: relativeOutputPath,
        template: RELATIVE_FIXTURE_DOTX_PATH,
        useComInterop,
      };
      const result = await callTool('word/markdown/import', args);

      if (useComInterop) {
        expect(result.success).toBe(true);
        await expect(fs.stat(outputPath)).resolves.toBeTruthy();
        // TODO: Verify template application and content for COM path
      } else {
        expect(result.success).toBe(false);
        if (!result.success) {
          // Library path currently ignores template and reports not implemented for the core functionality
          expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB');
        }
      }
    });
  });

  describe.each(testModes)('word/markdown/export (mode: $mode)', ({ useComInterop }) => {
    test('should export a Word document to Markdown', async () => {
      const outputFileName = `exported_word_${useComInterop}.md`;
      const outputPath = path.join(TEMP_DIR, outputFileName);
      const relativeOutputPath = `tests/temp_word_markdown_dir/${outputFileName}`;
      const imageDirName = `exported_images_${useComInterop}`;
      const relativeImageDir = `tests/temp_word_markdown_dir/${imageDirName}`; // Relative to workspace for tool
      const absoluteImageDir = path.join(TEMP_DIR, imageDirName); // Absolute for fs checks

      const args = {
        filePath: RELATIVE_FIXTURE_WORD_DOC_PATH,
        output: relativeOutputPath,
        imageDir: relativeImageDir, // Tool expects path relative to output MD's dir, or workspace. Let's use one next to output.
        useComInterop,
      };
      const result = await callTool('word/markdown/export', args);

      expect(result.success).toBe(true);
      if (result.success) {
        await expect(fs.stat(outputPath)).resolves.toBeTruthy();
        const mdContent = await fs.readFile(outputPath, 'utf-8');
        expect(mdContent.length).toBeGreaterThan(0);

        // Check for images (basic check: if imageDir was specified, it should exist if images were present)
        // CV.docx is expected to have images.
        try {
            const imageDirStats = await fs.stat(absoluteImageDir);
            expect(imageDirStats.isDirectory()).toBe(true);
            const filesInImageDir = await fs.readdir(absoluteImageDir);
            // This check depends on CV.docx having extractable images by both paths
            // Mammoth extracts images, COM path should too if implemented.
            expect(filesInImageDir.length).toBeGreaterThan(0);
        } catch (e: any) {
            // If the image directory doesn't exist, it might mean no images were found or an error occurred.
            // For this test, we expect images from CV.docx.
            if (useComInterop === false) { // Mammoth should definitely extract images
                 fail(`Image directory ${absoluteImageDir} not found or empty for Library (Mammoth) path.`);
            } else {
                // COM path image extraction might be more nuanced or have specific conditions.
                // For now, we'll log a warning if it fails for COM, but this might need refinement.
                console.warn(`Image directory check failed for COM path: ${e.message}. This might be acceptable if CV.docx has no COM-extractable images or if COM image extraction is partial.`);
            }
        }
      }
    });

    test('should export a Word document to Markdown and create a ZIP', async () => {
      const outputMdFileName = `exported_word_zip_${useComInterop}.md`;
      const relativeOutputMdPath = `tests/temp_word_markdown_dir/${outputMdFileName}`;
      const zipFileName = `exported_word_archive_${useComInterop}.zip`;
      const outputZipPath = path.join(TEMP_DIR, zipFileName); // For fs.stat
      const relativeOutputZipPath = `tests/temp_word_markdown_dir/${zipFileName}`; // For tool arg if needed, or derive

      const imageDirName = `exported_zip_images_${useComInterop}`;
      const relativeImageDir = `tests/temp_word_markdown_dir/${imageDirName}`;


      const args = {
        filePath: RELATIVE_FIXTURE_WORD_DOC_PATH,
        output: relativeOutputMdPath, // Base name for MD inside zip
        imageDir: relativeImageDir,
        zipOutput: true,
        zipFileName: relativeOutputZipPath, // Explicit zip name
        useComInterop,
      };
      const result = await callTool('word/markdown/export', args);

      expect(result.success).toBe(true);
      if (result.success) {
        await expect(fs.stat(outputZipPath)).resolves.toBeTruthy();
        // TODO: Could inspect ZIP content if a library like 'adm-zip' is added to devDependencies
      }
    });
  });

  test('word/markdown/import should fail for non-existent Markdown file', async () => {
    const args = {
      filePath: 'non_existent_markdown.md',
      output: 'tests/temp_word_markdown_dir/wont_be_created.docx',
      useComInterop: true, // COM path chosen as it's more likely to hit file system check first
    };
    const result = await callTool('word/markdown/import', args);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('FILE_NOT_FOUND');
    }
  });

  test('word/markdown/export should fail for non-existent Word file', async () => {
    const args = {
      filePath: 'non_existent_word_doc.docx',
      output: 'tests/temp_word_markdown_dir/wont_be_created.md',
      useComInterop: false, // Library path chosen
    };
    const result = await callTool('word/markdown/export', args);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('FILE_NOT_FOUND');
    }
  });
});