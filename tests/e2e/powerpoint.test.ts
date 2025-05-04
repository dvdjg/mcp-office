import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Assuming the server runs on localhost:3000
const TEMP_DIR = path.join(__dirname, '../temp_powerpoint_dir');
const FIXTURE_PPTX_PATH = path.join(__dirname, '../fixtures/sample_presentation.pptx'); // Assuming a fixture file exists

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


describe('powerpoint e2e tests', () => {

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
    // Note: Creating a valid .pptx file programmatically for tests is complex.
    // For now, we will rely on a fixture file for most tests.
    // Ensure the fixture file exists.
    try {
      await fs.stat(FIXTURE_PPTX_PATH);
    } catch (error: any) {
      if (error.code === 'ENOENT') {
         console.error(`Fixture file ${FIXTURE_PPTX_PATH} not found. Please ensure it exists for PowerPoint tests.`);
         // Depending on test setup, might throw or skip tests
      } else {
        console.error(`Error checking fixture file ${FIXTURE_PPTX_PATH}:`, error);
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

  describe('powerpoint/slides', () => {
    test('should add a new slide', async () => {
      const outputPath = path.join(TEMP_DIR, 'add_slide_test.pptx');
      // Copy the fixture to the temp directory to avoid modifying the original
      await fs.copyFile(FIXTURE_PPTX_PATH, outputPath);

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'powerpoint/slides/add',
          arguments: {
            document: 'tests/temp_powerpoint_dir/add_slide_test.pptx', // Use path relative to workspace
            layout: 'TitleAndContent', // Example layout
            // Optional: position
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // TODO: Verify the slide was added. This might require inspecting the presentation structure.
    });

    test('should delete a slide', async () => {
      const outputPath = path.join(TEMP_DIR, 'delete_slide_test.pptx');
      // Copy the fixture to the temp directory and ensure it has at least one slide to delete
      await fs.copyFile(FIXTURE_PPTX_PATH, outputPath);
      // Note: To make this test reliable, we should ensure the fixture has multiple slides
      // or add a slide using the 'add' tool before deleting.
      const slideIndexToDelete = 1; // Assuming the fixture has at least one slide at index 1 (2nd slide)

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'powerpoint/slides/delete',
          arguments: {
            document: 'tests/temp_powerpoint_dir/delete_slide_test.pptx', // Use path relative to workspace
            slideIndex: slideIndexToDelete,
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // TODO: Verify the slide was deleted. This might require inspecting the presentation structure.
    });

    test('should set properties of a slide', async () => {
      const outputPath = path.join(TEMP_DIR, 'set_slide_props_test.pptx');
      // Copy the fixture to the temp directory
      await fs.copyFile(FIXTURE_PPTX_PATH, outputPath);
      const slideIndexToModify = 1; // Assuming the fixture has at least one slide at index 1

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'powerpoint/slides/set',
          arguments: {
            document: 'tests/temp_powerpoint_dir/set_slide_props_test.pptx', // Use path relative to workspace
            slideIndex: slideIndexToModify,
            properties: {
              // Example properties - need to confirm actual supported properties from API spec
              // background: { type: 'color', color: '#FFFF00' },
            },
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // TODO: Verify the properties were set. This requires inspecting the presentation structure.
    });
  });

  describe('powerpoint/shapes', () => {
    test('should add a shape to a slide', async () => {
      const outputPath = path.join(TEMP_DIR, 'add_shape_test.pptx');
      // Copy the fixture to the temp directory
      await fs.copyFile(FIXTURE_PPTX_PATH, outputPath);
      const slideIndex = 1; // Assuming the fixture has at least one slide at index 1

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'powerpoint/shapes/add',
          arguments: {
            document: 'tests/temp_powerpoint_dir/add_shape_test.pptx', // Use path relative to workspace
            slideIndex: slideIndex,
            shapeType: 'Rectangle', // Example shape type - need to confirm supported types
            // Optional: position, size, text, imagePath, etc.
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // TODO: Verify the shape was added. This requires inspecting the presentation structure.
    });
  });

  describe('powerpoint/charts', () => {
    test('should insert a chart into a slide', async () => {
      const outputPath = path.join(TEMP_DIR, 'insert_chart_test.pptx');
      // Copy the fixture to the temp directory
      await fs.copyFile(FIXTURE_PPTX_PATH, outputPath);
      const slideIndex = 1; // Assuming the fixture has at least one slide at index 1
      // Note: Inserting a chart usually requires data. This test assumes the tool
      // can insert a default chart or takes data as an argument (need to check API spec).

      const response = await fetch(`${MCP_SERVER_URL}/tool`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          tool_name: 'powerpoint/charts/insert',
          arguments: {
            document: 'tests/temp_powerpoint_dir/insert_chart_test.pptx', // Use path relative to workspace
            slideIndex: slideIndex,
            chartType: 'ColumnClustered', // Example chart type - need to confirm supported types
            // Optional: data, position, size, title, etc.
          },
        }),
      });

      expect(response.ok).toBe(true);
      const result = await response.json() as ToolResponse;

      expect(result.success).toBe(true);

      // TODO: Verify the chart was inserted. This requires inspecting the presentation structure.
    });
  });

  // TODO: Add tests for complex variations and error handling for all PowerPoint tools
});