import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const FIXTURES_PATH = path.join(__dirname, '../fixtures');
const TEMP_DIR = path.join(__dirname, '../temp_fs_dir');
const NESTED_TEMP_DIR = path.join(TEMP_DIR, 'nested/directory');
const TEST_FILE_PATH = path.join(TEMP_DIR, 'test_file.txt');
const TEST_FILE_CONTENT = 'This is a test file content.';
const TEST_BLOB_PATH = path.join(TEMP_DIR, 'test_blob.bin');
const TEST_BLOB_CONTENT = Buffer.from('This is some binary data for the blob test.');
const TEST_ZIP_PATH = path.join(TEMP_DIR, 'test_archive.zip');
const FILE_TO_ZIP_PATH = path.join(TEMP_DIR, 'file_to_zip.txt');
const FILE_TO_ZIP_CONTENT = 'Content of the file to be zipped.';


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


describe('fs e2e tests', () => {

  beforeAll(async () => {
    // Ensure fixtures directory exists (should be part of the project)
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
    // Create a nested directory and a file for recursive listing test
    await fs.mkdir(NESTED_TEMP_DIR, { recursive: true });
    await fs.writeFile(path.join(NESTED_TEMP_DIR, 'test_file.txt'), 'hello');
    // Create a test file for read/delete tests
    await fs.writeFile(TEST_FILE_PATH, TEST_FILE_CONTENT);
    // Create a file to be zipped
    await fs.writeFile(FILE_TO_ZIP_PATH, FILE_TO_ZIP_CONTENT);
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

  describe('fs/directory', () => {
    test('should list files in a directory', async () => {
      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/directory/list',
          arguments: {
            path: 'tests/fixtures/', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('files');
        expect(Array.isArray(result.data.files)).toBe(true);

        // Basic check for some expected files/directories in fixtures
        const fileNames = result.data.files.map((f: any) => f.name);
        expect(fileNames).toContain('CV.docx');
        expect(fileNames).toContain('cuentoAladdin_draft1.docx');
        expect(fileNames).toContain('cuentoAladdin_draft2.docx');
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    test('should list files in a directory with a filter', async () => {
      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/directory/list',
          arguments: {
            path: 'tests/fixtures/', // Use path relative to workspace
            filter: '*.docx',
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('files');
        expect(Array.isArray(result.data.files)).toBe(true);

        // Check that only .docx files are listed
        const fileNames = result.data.files.map((f: any) => f.name);
        expect(fileNames.every((name: string) => name.endsWith('.docx'))).toBe(true);
        expect(fileNames).toContain('CV.docx');
        expect(fileNames).toContain('cuentoAladdin_draft1.docx');
        expect(fileNames).toContain('cuentoAladdin_draft2.docx');
        // Assuming there are non-docx files in fixtures, check they are not present
        // This check might need adjustment based on actual fixture content
        // expect(fileNames).not.toContain('some_other_file.txt');
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    test('should list files recursively in a directory', async () => {
      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/directory/list',
          arguments: {
            path: 'tests/temp_fs_dir', // Use path relative to workspace
            recursive: true,
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('files');
        expect(Array.isArray(result.data.files)).toBe(true);

        // Check for the nested directory and the file within it
        const filePaths = result.data.files.map((f: any) => f.path);
        // Use path.normalize to handle potential differences in path separators
        expect(filePaths).toContain(path.normalize('tests/temp_fs_dir/nested'));
        expect(filePaths).toContain(path.normalize('tests/temp_fs_dir/nested/directory'));
        expect(filePaths).toContain(path.normalize('tests/temp_fs_dir/nested/directory/test_file.txt'));
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });


    test('should create a new directory', async () => {
      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/directory/create',
          arguments: {
            path: 'tests/temp_fs_dir/new_dir', // Use a new path for this test
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the directory was created on the file system
      await expect(fs.stat(path.join(TEMP_DIR, 'new_dir'))).resolves.toBeTruthy();
      await expect(fs.stat(path.join(TEMP_DIR, 'new_dir'))).resolves.toHaveProperty('isDirectory', true);
    });

    test('should delete an empty directory', async () => {
      // Ensure the directory exists before attempting to delete it
      const dirToDelete = path.join(TEMP_DIR, 'dir_to_delete');
      await fs.mkdir(dirToDelete, { recursive: true });

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/directory/delete',
          arguments: {
            path: 'tests/temp_fs_dir/dir_to_delete', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the directory was deleted from the file system
      await expect(fs.stat(dirToDelete)).rejects.toThrow();
    });
  });

  describe('fs/file', () => {
    test('should read a file', async () => {
      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/file/read',
          arguments: {
            path: 'tests/temp_fs_dir/test_file.txt', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('content', TEST_FILE_CONTENT);
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    test('should write a new file', async () => {
      const newFilePath = path.join(TEMP_DIR, 'new_test_file.txt');
      const newFileContent = 'Content for the new file.';

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/file/write',
          arguments: {
            path: 'tests/temp_fs_dir/new_test_file.txt', // Use path relative to workspace
            content: newFileContent,
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the file was created and has the correct content
      const fileContent = await fs.readFile(newFilePath, 'utf-8');
      expect(fileContent).toBe(newFileContent);
    });

    test('should delete a file', async () => {
      // Ensure the file exists before attempting to delete it
      const fileToDeletePath = path.join(TEMP_DIR, 'file_to_delete.txt');
      await fs.writeFile(fileToDeletePath, 'delete me');

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/file/delete',
          arguments: {
            path: 'tests/temp_fs_dir/file_to_delete.txt', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the file was deleted
      await expect(fs.stat(fileToDeletePath)).rejects.toThrow();
    });
  });

  describe('fs/blob', () => {
    test('should save a blob', async () => {
      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/blob/save',
          arguments: {
            path: 'tests/temp_fs_dir/test_blob.bin', // Use path relative to workspace
            content: TEST_BLOB_CONTENT.toString('base64'), // Send as base64
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the blob was saved
      await expect(fs.stat(TEST_BLOB_PATH)).resolves.toBeTruthy();
      const savedContent = await fs.readFile(TEST_BLOB_PATH);
      expect(savedContent).toEqual(TEST_BLOB_CONTENT);
    });

    test('should read a blob', async () => {
      // Ensure the blob exists before attempting to read it
      await fs.writeFile(TEST_BLOB_PATH, TEST_BLOB_CONTENT);

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/blob/read',
          arguments: {
            path: 'tests/temp_fs_dir/test_blob.bin', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('content');
        const decodedContent = Buffer.from(result.data.content, 'base64');
        expect(decodedContent).toEqual(TEST_BLOB_CONTENT);
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });
  });

  describe('fs/archive', () => {
    test('should create a zip archive', async () => {
      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/archive/create',
          arguments: {
            archive_path: 'tests/temp_fs_dir/test_archive.zip', // Use path relative to workspace
            source_paths: ['tests/temp_fs_dir/file_to_zip.txt'], // Use path relative to workspace
            format: 'zip',
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // Verify the archive was created (basic check)
      await expect(fs.stat(TEST_ZIP_PATH)).resolves.toBeTruthy();
    });

    test('should list contents of a zip archive', async () => {
      // Ensure the archive exists before listing
      // This might require creating a dummy zip file here or relying on the create test
      // For now, let's assume the create test ran or a fixture exists.
      // A more robust approach would be to create a zip fixture in beforeAll.

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'fs/archive/list',
          arguments: {
            archive_path: 'tests/temp_fs_dir/test_archive.zip', // Use path relative to workspace
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('files');
        expect(Array.isArray(result.data.files)).toBe(true);
        const fileNames = result.data.files.map((f: any) => f.name);
        expect(fileNames).toContain('file_to_zip.txt');
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    test('should extract a file from a zip archive', async () => {
       // Ensure the archive exists before extracting
       // Similar to list test, might need a fixture or rely on create test.

       const extractedFilePath = path.join(TEMP_DIR, 'extracted_file_to_zip.txt');

       const response = await fetch(`${MCP_SERVER_URL}/tool`, {
         method: 'POST',
         headers: {
           'Content-Type': 'application/json',
         },
         body: JSON.stringify({
           tool_name: 'fs/archive/extract',
           arguments: {
             archive_path: 'tests/temp_fs_dir/test_archive.zip', // Use path relative to workspace
             file_path_in_archive: 'file_to_zip.txt',
             output_path: 'tests/temp_fs_dir/extracted_file_to_zip.txt', // Use path relative to workspace
           },
         }),
       });

       expect(response.ok).toBe(true);
       const result = await response.json() as ToolResponse;

       expect(result.success).toBe(true);

       // Verify the file was extracted and has the correct content
       await expect(fs.stat(extractedFilePath)).resolves.toBeTruthy();
       const extractedContent = await fs.readFile(extractedFilePath, 'utf-8');
       expect(extractedContent).toBe(FILE_TO_ZIP_CONTENT);
    });
  });
});