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


describe('fs/directory e2e tests', () => {

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
    // Create a nested directory and a file for recursive listing test
    await fs.mkdir(NESTED_TEMP_DIR, { recursive: true });
    await fs.writeFile(path.join(NESTED_TEMP_DIR, 'test_file.txt'), 'hello');
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
          path: 'tests/temp_fs_dir', // Use path relative to workspace
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    // Verify the directory was created on the file system
    await expect(fs.stat(TEMP_DIR)).resolves.toBeTruthy();
    await expect(fs.stat(TEMP_DIR)).resolves.toHaveProperty('isDirectory', true);
  });

  test('should delete an empty directory', async () => {
    // Ensure the directory exists before attempting to delete it
    await fs.mkdir(TEMP_DIR, { recursive: true });

    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'fs/directory/delete',
        arguments: {
          path: 'tests/temp_fs_dir', // Use path relative to workspace
        },
      }),
    });

    expect(response.ok).toBe(true);
    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(true);

    // Verify the directory was deleted from the file system
    await expect(fs.stat(TEMP_DIR)).rejects.toThrow();
  });

  // Add tests for complex variations in subsequent steps
});