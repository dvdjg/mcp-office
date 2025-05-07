import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch'; // Assuming node-fetch is available for making HTTP requests

const MCP_SERVER_URL = 'http://localhost:3000'; // Adjust if your test server runs elsewhere
const TEMP_INTEGRATION_DIR = path.join(__dirname, '../temp_ppt_integration_dir');
const FIXTURE_PPTX_PATH = path.join(__dirname, '../../fixtures/sample_presentation.pptx'); // Adjust path to your fixture

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

describe('PowerPoint Tools - Integration Tests', () => {
  beforeAll(async () => {
    await fs.rm(TEMP_INTEGRATION_DIR, { recursive: true, force: true });
    await fs.mkdir(TEMP_INTEGRATION_DIR, { recursive: true });
    try {
      await fs.stat(FIXTURE_PPTX_PATH);
    } catch (error: any) {
      if (error.code === 'ENOENT') {
        throw new Error(`Fixture file ${FIXTURE_PPTX_PATH} not found. Critical for integration tests.`);
      }
      throw error;
    }
  });

  afterAll(async () => {
    await fs.rm(TEMP_INTEGRATION_DIR, { recursive: true, force: true });
  });

  const runIntegrationTestForTool = (
    description: string,
    toolName: string,
    baseArguments: Record<string, any>,
    comPathAssertions: (response: ToolResponse, outputPath?: string) => Promise<void> | void,
    libPathAssertions: (response: ToolResponse, outputPath?: string) => Promise<void> | void,
    // If true, the library path test expects to create a new file rather than modify FIXTURE_PPTX_PATH
    libPathCreatesNewFile: boolean = false
  ) => {
    describe(description, () => {
      // Test COM Path
      test(`COM Path: ${description}`, async () => {
        const uniqueId = Date.now();
        const outputFileName = `${toolName.replace(/\//g, '_')}_com_int_${uniqueId}.pptx`;
        const outputPath = path.join(TEMP_INTEGRATION_DIR, outputFileName);
        // Relative path for the tool argument, assuming workspace is project root
        const relativeOutputPath = path.join(path.basename(path.dirname(TEMP_INTEGRATION_DIR)), path.basename(TEMP_INTEGRATION_DIR), outputFileName);


        // For COM, usually operates on a copy of a fixture or an existing file.
        // If the operation is meant to create, this might not be needed.
        if (baseArguments.filePath || !libPathCreatesNewFile) { // Heuristic: if filePath is given, or lib path doesn't create new, COM probably needs a fixture.
             await fs.copyFile(FIXTURE_PPTX_PATH, outputPath);
        }


        const requestBody = {
          tool_name: toolName,
          arguments: {
            ...baseArguments,
            filePath: baseArguments.filePath || relativeOutputPath, // Use provided filePath or the dynamic one
            useComInterop: true,
          },
        };

        const response = await fetch(`${MCP_SERVER_URL}/tool`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(requestBody),
        });
        expect(response.ok).toBe(true); // Basic check that the request was accepted
        const result = await response.json() as ToolResponse;
        await comPathAssertions(result, outputPath);
      });

      // Test Library Path
      test(`Library Path: ${description}`, async () => {
        const uniqueId = Date.now();
        const outputFileName = `${toolName.replace(/\//g, '_')}_lib_int_${uniqueId}.pptx`;
        const outputPath = path.join(TEMP_INTEGRATION_DIR, outputFileName);
        const relativeOutputPath = path.join(path.basename(path.dirname(TEMP_INTEGRATION_DIR)), path.basename(TEMP_INTEGRATION_DIR), outputFileName);

        let docPathForLib = relativeOutputPath;

        if (!libPathCreatesNewFile && baseArguments.filePath) {
            // If lib path modifies, and a filePath is given, copy fixture to that path for the lib to use.
            // This assumes baseArguments.filePath is relative to workspace.
            const targetPathForFixtureCopy = path.resolve(baseArguments.filePath); // Make absolute if relative
            await fs.mkdir(path.dirname(targetPathForFixtureCopy), {recursive: true});
            await fs.copyFile(FIXTURE_PPTX_PATH, targetPathForFixtureCopy);
            docPathForLib = baseArguments.filePath; // Use the user-provided path
        } else if (!libPathCreatesNewFile) {
            // Lib path modifies, but no specific filePath given, so use a dynamic one based on fixture.
            await fs.copyFile(FIXTURE_PPTX_PATH, outputPath);
            docPathForLib = relativeOutputPath;
        }
        // If libPathCreatesNewFile is true, docPathForLib remains relativeOutputPath, and no fixture is copied here.

        const requestBody = {
          tool_name: toolName,
          arguments: {
            ...baseArguments,
            filePath: docPathForLib,
            useComInterop: false,
          },
        };
        
        const response = await fetch(`${MCP_SERVER_URL}/tool`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(requestBody),
        });
        const result = await response.json() as ToolResponse;
        await libPathAssertions(result, outputPath); // outputPath is where a new file might be saved by the lib
      });
    });
  };

  // --- slides.tool.ts Integration Tests ---
  describe('Integration Tests for powerpoint/slides', () => {
    runIntegrationTestForTool(
      'Add a new slide',
      'powerpoint/slides', // Assuming the tool path is 'powerpoint/slides' and operation is in args
      { operation: 'add', slideLayout: 'TITLE_SLIDE' /* PptxGenJS equivalent: TITLE_SLIDE */ },
      async (result) => {
        expect(result.success).toBe(true);
        // COM: Further verification might involve trying to get properties of the new slide or counting slides.
      },
      async (result, outputPath) => {
        expect(result.success).toBe(true);
        // Lib: Verify file was created/modified. PptxGenJS should add a slide.
        // Check if outputPath (if new file) or the original path (if modified) exists and is a valid pptx.
        // For now, success:true is the primary check.
        if (result.success) {
            expect(result.data).toContain('Slide added using pptxgenjs');
        }
        if (outputPath) { // Ensure outputPath is defined
            const stats = await fs.stat(outputPath); // outputPath is where pptxgenjs saves
            expect(stats.isFile()).toBe(true);
            expect(stats.size).toBeGreaterThan(0);
        } else {
            throw new Error('outputPath was undefined in libPathAssertions for Add a new slide');
        }
      },
      true // pptxgenjs path for 'add slide' typically creates/overwrites a file.
    );

    runIntegrationTestForTool(
      'Delete a slide',
      'powerpoint/slides',
      { operation: 'delete', slideIndex: 1, filePath: 'tests/temp_ppt_integration_dir/deleteSlideTest.pptx' }, // Provide explicit path
      async (result) => {
        expect(result.success).toBe(true);
        // COM: Verify slide was deleted (e.g., count slides before/after or try to access deleted slide).
      },
      async (result) => {
        expect(result.success).toBe(false);
        expect((result as ErrorResponse).error.code).toBe('POWERPOINT_LIB_UNSUPPORTED');
        expect((result as ErrorResponse).error.message).toMatch(/Deleting slides from existing files is not supported/i);
      },
      false // Library path attempts to modify an existing file (which it can't for delete)
    );
  });

  // --- shapes.tool.ts Integration Tests ---
  describe('Integration Tests for powerpoint/shapes', () => {
    runIntegrationTestForTool(
      'Insert a shape',
      'powerpoint/shapes',
      { 
        operation: 'insert', 
        slideIndex: 1, 
        shapeType: 'msoShapeRectangle', // COM type
        position: { left: 100, top: 100 }, 
        size: { width: 100, height: 50 },
        text: 'Integration Shape'
      },
      async (result) => {
        expect(result.success).toBe(true);
        // COM: Verify shape was added.
      },
      async (result, outputPath) => {
        expect(result.success).toBe(true);
        // Lib: PptxGenJS adds shape to a new/overwritten file.
        if (result.success) {
            expect(result.data).toContain('PptxGenJS: Shape inserted');
        }
        if (outputPath) { // Ensure outputPath is defined
            const stats = await fs.stat(outputPath);
            expect(stats.isFile()).toBe(true);
            expect(stats.size).toBeGreaterThan(0);
        } else {
            throw new Error('outputPath was undefined in libPathAssertions for Insert a shape');
        }

      },
      true // pptxgenjs path for 'insert shape' creates/overwrites.
    );

    runIntegrationTestForTool(
      'Delete a shape from an existing file',
      'powerpoint/shapes',
      { 
        operation: 'delete', 
        slideIndex: 1, 
        shapeIndex: 1, // Assuming a shape exists at index 1 in the fixture
        filePath: 'tests/temp_ppt_integration_dir/deleteShapeTest.pptx' 
      },
      async (result) => {
        expect(result.success).toBe(true); // COM should delete
      },
      async (result) => {
        expect(result.success).toBe(false);
        expect((result as ErrorResponse).error.code).toBe('POWERPOINT_LIB_UNSUPPORTED');
        expect((result as ErrorResponse).error.message).toMatch(/Operation 'delete' for existing shapes is not supported/i);
      },
      false
    );
  });

  // --- animations.tool.ts Integration Tests ---
  describe('Integration Tests for powerpoint/animations', () => {
    runIntegrationTestForTool(
      'Add animation to a new shape (text object)',
      'powerpoint/animations',
      {
        operation: 'add',
        animationType: 'fadeIn', // PptxGenJS type
        newObjectText: 'Animated via Integration Test',
        newObjectOptions: { x: 0.5, y: 0.5, w: 4, h: 1 },
        duration: 1,
      },
      async (result) => {
        // COM path for 'add' animation requires existing shape. This test setup is more for lib path.
        // For COM, a separate test adding animation to an existing shape in fixture would be better.
        // For now, let's assume COM path might fail gracefully or succeed if it creates a shape then animates.
        // This needs alignment with how the COM part of 'add animation' is implemented.
        // If COM 'add' animation requires existing shape, this test for COM should expect failure or be different.
        // Let's assume for this specific setup, COM path is not the primary target.
        // A more robust COM test would ensure a shape exists first.
         expect(result.success).toBe(true); // Placeholder, COM test needs refinement
      },
      async (result, outputPath) => {
        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data.message).toContain('PptxGenJS: Added new text object with animation');
        }
        if (outputPath) { // Ensure outputPath is defined
            const stats = await fs.stat(outputPath);
            expect(stats.isFile()).toBe(true);
        } else {
            throw new Error('outputPath was undefined in libPathAssertions for Add animation');
        }
      },
      true // PptxGenJS creates new presentation with animated object.
    );
  });

  // --- properties.tool.ts Integration Tests ---
  describe('Integration Tests for powerpoint/properties', () => {
    runIntegrationTestForTool(
      'Set presentation title property',
      'powerpoint/properties',
      {
        operation: 'set',
        propertyName: 'title',
        propertyValue: 'Integration Test Title',
      },
      async (result) => {
        expect(result.success).toBe(true);
        // COM: Verify property was set (e.g., by trying to 'get' it or inspecting file if possible).
      },
      async (result, outputPath) => {
        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data).toContain("PptxGenJS: Operation 'set' completed.");
        }
        // PptxGenJS sets property on new/overwritten file.
        if (outputPath) { // Ensure outputPath is defined
            const stats = await fs.stat(outputPath); // File where properties are set
            expect(stats.isFile()).toBe(true);
        } else {
            throw new Error('outputPath was undefined in libPathAssertions for Set presentation title property');
        }
      },
      true // PptxGenJS creates/overwrites file with new properties.
    );
  });

  // TODO: Add more integration tests for other operations and tools,
  // focusing on the dual-path logic and specific limitations.
});