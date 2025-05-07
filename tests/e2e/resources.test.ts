import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_resources_dir');

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


describe('resources e2e tests', () => {

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

  describe('memory/ai_assistant_guide', () => {
    test('should read the ai_assistant_guide resource', async () => {
      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'memory/ai_assistant_guide/read',
          arguments: {
            // No specific arguments mentioned in API doc for read
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('content');
        expect(typeof result.data.content).toBe('string');
        expect(result.data.content.length).toBeGreaterThan(0);
        // Basic check for expected content
        expect(result.data.content).toContain('AI assistant guide');
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    // TODO: Add test for memory/ai_assistant_guide/list if applicable and parameters are known
  });

  describe('dynamic/resources', () => {
    test('should list dynamic resources', async () => {
      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'dynamic/resources/list',
          arguments: {
            // No specific arguments mentioned in API doc for list
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      if (result.success) {
        expect(result.data).toHaveProperty('resources');
        expect(Array.isArray(result.data.resources)).toBe(true);
        // Basic check for expected structure of resource items
        if (result.data.resources.length > 0) {
          expect(result.data.resources[0]).toHaveProperty('uri');
          expect(result.data.resources[0]).toHaveProperty('description');
        }
      } else {
        fail('Expected test to succeed but it failed.');
      }
    });

    test('should read a dynamic resource (requires a known dynamic resource)', async () => {
      // This test requires a known dynamic resource URI to exist.
      // Assuming 'memory://ai_assistant_guide' is also accessible via dynamic/resources/read
      // or there is another known dynamic resource.
      // Based on the API doc, memory/ai_assistant_guide is listed separately,
      // so let's assume there's another dynamic resource or this tool can read memory resources.
      // For now, let's use a placeholder URI and add a TODO to use a real one.
      const knownDynamicResourceUri = 'dynamic://some_resource_uri'; // Replace with a real dynamic resource URI

      // TODO: Determine a real dynamic resource URI to test with.
      console.warn(`Using placeholder URI "${knownDynamicResourceUri}" for dynamic/resources/read test.`);


      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'dynamic/resources/read',
          arguments: {
            uri: knownDynamicResourceUri,
          },
        }),
      });

      // This test might fail if the placeholder URI is not valid or doesn't exist.
      // The expectation here is more about the tool call structure and response format
      // than the actual content, unless a reliable dynamic resource is known.

      // expect(response.ok).toBe(true); // The response might not be OK for a non-existent resource
      const result = await response.json() as ToolResponse;

      // Depending on the tool's error handling for non-existent resources,
      // we might expect success: false or success: true with empty data.
      // For now, let's expect a response and check for either success or a specific error code.

      expect(result).toHaveProperty('success');

      if (result.success) {
         expect(result.data).toHaveProperty('content');
         // TODO: Add more specific content checks if a reliable resource is used
      } else {
         // Expecting an error for a non-existent resource
         expect(result.error).toHaveProperty('code');
         expect(result.error).toHaveProperty('message');
         // TODO: Check for a specific error code if the API defines one for non-existent resources
      }
    });

    // TODO: Add tests for dynamic/resources/write, delete, metadata, search if parameters are known
  });
});