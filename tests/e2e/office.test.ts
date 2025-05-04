import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_office_dir');
const FIXTURES_DIR = path.join(__dirname, '../fixtures');
const FIXTURE_WORD_DOC = path.join(FIXTURES_DIR, 'CV.docx'); // Assuming this fixture exists
const FIXTURE_EXCEL_DOC = path.join(FIXTURES_DIR, 'sample_excel_data.xlsx'); // Assuming this fixture exists
const FIXTURE_PPT_DOC = path.join(FIXTURES_DIR, 'sample_presentation.pptx'); // Assuming this fixture exists


// Define a basic type for the expected successful response
interface SuccessResponse {
  success: true;
  data: any; // Use a more specific type if the data structure is known
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


describe('office (cross-application) e2e tests', () => {

  beforeAll(async () => {
    // Ensure temp directory does not exist before tests
    try {
      await fs.rm(TEMP_DIR, { recursive: true });
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        console.error(`Error cleaning up temp directory ${TEMP_DIR}:`, error);
      }
    }
    // Create temp directory
    await fs.mkdir(TEMP_DIR, { recursive: true });

    // Check if fixture files exist
    const fixtures = [FIXTURE_WORD_DOC, FIXTURE_EXCEL_DOC, FIXTURE_PPT_DOC];
    for (const fixturePath of fixtures) {
      try {
        await fs.stat(fixturePath);
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.error(`Fixture file ${fixturePath} not found. Please ensure it exists for cross-application tests.`);
           // Depending on test setup, might throw or skip tests
        } else {
          console.error(`Error checking fixture file ${fixturePath}:`, error);
        }
      }
    }
  });

  afterAll(async () => {
    // Clean up temp directory after tests
    try {
      await fs.rm(TEMP_DIR, { recursive: true });
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        console.error(`Error cleaning up temp directory ${TEMP_DIR}:`, error);
      }
    }
  });

  describe('office/pdf/export', () => {
    test('should export a Word document to PDF (requires fixture)', async () => {
      const outputPath = path.join(TEMP_DIR, 'exported_word.pdf');

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'office/pdf/export',
          arguments: {
            document: 'tests/fixtures/CV.docx', // Use path relative to workspace
            output_path: 'tests/temp_office_dir/exported_word.pdf', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the PDF file was created (basic check)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification of PDF content if possible
    });

    test('should export an Excel document to PDF (requires fixture)', async () => {
      const outputPath = path.join(TEMP_DIR, 'exported_excel.pdf');

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'office/pdf/export',
          arguments: {
            document: 'tests/fixtures/sample_excel_data.xlsx', // Use path relative to workspace
            output_path: 'tests/temp_office_dir/exported_excel.pdf', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the PDF file was created (basic check)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification of PDF content if possible
    });

    test('should export a PowerPoint document to PDF (requires fixture)', async () => {
      const outputPath = path.join(TEMP_DIR, 'exported_ppt.pdf');

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'office/pdf/export',
          arguments: {
            document: 'tests/fixtures/sample_presentation.pptx', // Use path relative to workspace
            output_path: 'tests/temp_office_dir/exported_ppt.pdf', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the PDF file was created (basic check)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification of PDF content if possible
    });
  });

  describe('office/combine', () => {
    test('should combine two Word documents (requires fixtures)', async () => {
      const outputPath = path.join(TEMP_DIR, 'combined_word.docx');

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'office/combine',
          arguments: {
            documents: [
              'tests/fixtures/CV.docx', // Use path relative to workspace
              'tests/fixtures/cuentoAladdin_draft1.docx', // Use path relative to workspace
            ],
            output_path: 'tests/temp_office_dir/combined_word.docx', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the combined file was created (basic check)
      await expect(fs.stat(outputPath)).resolves.toBeTruthy();
      // TODO: Add verification of combined document content
    });

    // TODO: Add tests for combining different file types if supported by the tool
  });

  describe('office/transfer', () => {
    test('should transfer data from Excel to Word (requires fixtures)', async () => {
      const targetWordPath = path.join(TEMP_DIR, 'word_with_excel_data.docx');
      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(FIXTURE_WORD_DOC, targetWordPath);


      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'office/transfer',
          arguments: {
            source: 'excel:tests/fixtures/sample_excel_data.xlsx:Sheet1!A1:B2', // Use path relative to workspace
            target: 'word:tests/temp_office_dir/word_with_excel_data.docx:end', // Use path relative to workspace
            operation: 'insert', // Assuming 'insert' is the operation for transferring data
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the target Word file was modified (basic check - could check modification time)
      await expect(fs.stat(targetWordPath)).resolves.toBeTruthy();
      // TODO: Add verification that the Excel data was inserted into the Word document
    });

    test('should transfer data from Excel to Word with linking (requires fixtures)', async () => {
      const targetWordPath = path.join(TEMP_DIR, 'word_with_linked_excel_data.docx');
      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(FIXTURE_WORD_DOC, targetWordPath);

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'office/transfer',
          arguments: {
            source: 'excel:tests/fixtures/sample_excel_data.xlsx:Sheet1!A1:B2', // Use path relative to workspace
            target: 'word:tests/temp_office_dir/word_with_linked_excel_data.docx:end', // Use path relative to workspace
            operation: 'insert',
            link: true, // Enable linking
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the target Word file was modified
      await expect(fs.stat(targetWordPath)).resolves.toBeTruthy();
      // TODO: Add verification that the Excel data was inserted as a linked object
    });

    test('should embed a Word document into a PowerPoint slide (requires fixtures)', async () => {
      const targetPptPath = path.join(TEMP_DIR, 'ppt_with_embedded_word.pptx');
      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(FIXTURE_PPT_DOC, targetPptPath);
      const slideIndex = 1; // Assuming the fixture has at least one slide at index 1

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'office/transfer',
          arguments: {
            source: 'word:tests/fixtures/CV.docx', // Use path relative to workspace
            target: `powerpoint:tests/temp_office_dir/ppt_with_embedded_word.pptx:slide:${slideIndex}`, // Use path relative to workspace
            operation: 'embed', // Embed the document
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the target PowerPoint file was modified
      await expect(fs.stat(targetPptPath)).resolves.toBeTruthy();
      // TODO: Add verification that the Word document was embedded into the PowerPoint slide
    });

    // TODO: Add tests for other transfer scenarios and error handling
  });

  // TODO: Add tests for Static and Dynamic Resources tools
});