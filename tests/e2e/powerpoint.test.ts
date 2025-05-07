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


describe('PowerPoint E2E Tests', () => {
  beforeAll(async () => {
    try {
      await fs.rm(TEMP_DIR, { recursive: true, force: true });
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        console.error(`Error cleaning up temp directory ${TEMP_DIR} before tests:`, error);
      }
    }
    await fs.mkdir(TEMP_DIR, { recursive: true });
    try {
      await fs.stat(FIXTURE_PPTX_PATH);
    } catch (error: any) {
      if (error.code === 'ENOENT') {
        console.error(`Fixture file ${FIXTURE_PPTX_PATH} not found. Please ensure it exists for PowerPoint tests.`);
        // This is a critical failure for E2E tests, consider throwing to stop tests
        throw new Error(`Fixture file ${FIXTURE_PPTX_PATH} not found.`);
      } else {
        console.error(`Error checking fixture file ${FIXTURE_PPTX_PATH}:`, error);
        throw error; // Rethrow other errors
      }
    }
  });

  afterAll(async () => {
    try {
      await fs.rm(TEMP_DIR, { recursive: true, force: true });
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        console.error(`Error cleaning up temp directory ${TEMP_DIR} after tests:`, error);
      }
    }
  });

  const runTestForBothPaths = (
    testName: string,
    toolName: string,
    baseArguments: Record<string, any>,
    comPathAssertions: (result: ToolResponse, outputPath?: string) => void,
    libPathAssertions: (result: ToolResponse, outputPath?: string) => void,
    setupNewFileForLibPath: boolean = false // If true, lib path test will create a new file
  ) => {
    describe(testName, () => {
      // COM Path Test
      test(`COM Path: ${testName}`, async () => {
        const uniqueId = Date.now();
        const outputFileName = `${toolName.replace(/\//g, '_')}_com_${uniqueId}.pptx`;
        const outputPath = path.join(TEMP_DIR, outputFileName);
        const relativeOutputPath = path.join('tests/temp_powerpoint_dir', outputFileName); // Relative to workspace

        if (!setupNewFileForLibPath) { // For COM, always copy fixture unless it's a "create new" test
             await fs.copyFile(FIXTURE_PPTX_PATH, outputPath);
        }


        const response = await fetch(`${MCP_SERVER_URL}/tool`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            tool_name: toolName,
            arguments: {
              ...baseArguments,
              document: setupNewFileForLibPath && !baseArguments.document ? relativeOutputPath : baseArguments.document || relativeOutputPath, // if setupNewFileForLibPath and no doc, use output. Else use provided or output.
              useComInterop: true,
            },
          }),
        });
        expect(response.ok).toBe(true);
        const result = await response.json() as ToolResponse;
        comPathAssertions(result, outputPath);
      });

      // Library Path Test
      test(`Library Path: ${testName}`, async () => {
        const uniqueId = Date.now();
        const outputFileName = `${toolName.replace(/\//g, '_')}_lib_${uniqueId}.pptx`;
        const outputPath = path.join(TEMP_DIR, outputFileName);
        const relativeOutputPath = path.join('tests/temp_powerpoint_dir', outputFileName); // Relative to workspace

        if (!setupNewFileForLibPath && baseArguments.document) { // Only copy if not creating new and document is specified
            // For library path, some operations might create a new file or operate on an existing one.
            // If the operation is meant to modify an existing file, copy the fixture.
            // If it's meant to create a new one, this copy might not be needed or could be detrimental.
            // This logic will need refinement based on specific test cases.
            await fs.copyFile(FIXTURE_PPTX_PATH, outputPath);
        }


        const response = await fetch(`${MCP_SERVER_URL}/tool`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            tool_name: toolName,
            arguments: {
              ...baseArguments,
              document: setupNewFileForLibPath && !baseArguments.document ? relativeOutputPath : baseArguments.document || relativeOutputPath,
              useComInterop: false,
            },
          }),
        });
        // For library path, response.ok might be true even for "not supported" if the tool handles it gracefully.
        // The key is in the result.success and error messages.
        const result = await response.json() as ToolResponse;
        libPathAssertions(result, outputPath);
      });
    });
  };

  describe('powerpoint/slides', () => {
    runTestForBothPaths(
      'should add a new slide',
      'powerpoint/slides/add',
      {
        // document will be set by runTestForBothPaths logic
        layout: 'TitleAndContent',
      },
      // COM Path Assertions
      (result) => {
        expect(result.success).toBe(true);
        // TODO: COM: Verify slide was actually added (e.g., by trying to get its properties or count slides)
      },
      // Library Path Assertions
      (result, outputPath) => {
        expect(result.success).toBe(true); // pptxgenjs should succeed in adding a slide to a new/existing presentation
        // TODO: Lib: Verify slide was added using officeparser or by checking pptxgenjs mock
        // For now, we assume if success is true, it worked as expected by pptxgenjs.
        // A more robust check would involve parsing outputPath if it's a new file.
      },
      true // For 'add slide', lib path often means creating a new presentation or adding to one it just made.
           // If baseArguments.document is provided, it will use that.
    );

    runTestForBothPaths(
      'should delete a slide',
      'powerpoint/slides/delete',
      {
        // document will be set by runTestForBothPaths
        slideIndex: 1, // Assuming 0-based index for deletion, or 1-based as per tool spec
      },
      // COM Path Assertions
      (result) => {
        expect(result.success).toBe(true);
        // TODO: COM: Verify slide was actually deleted
      },
      // Library Path Assertions
      (result) => {
        // pptxgenjs does NOT support deleting slides from an existing presentation.
        // It can only build presentations from scratch.
        expect(result.success).toBe(false);
        expect((result as ErrorResponse).error.message).toMatch(/not supported/i);
      }
    );

    runTestForBothPaths(
      'should set properties of a slide',
      'powerpoint/slides/set',
      {
        // document will be set by runTestForBothPaths
        slideIndex: 1,
        properties: { background: { color: 'FFFF00' } }, // Example property
      },
      // COM Path Assertions
      (result) => {
        expect(result.success).toBe(true);
        // TODO: COM: Verify slide properties were set
      },
      // Library Path Assertions
      (result) => {
        // pptxgenjs has some support for slide properties, but modifying an existing loaded slide's props is tricky.
        // This test assumes it might be limited or not fully supported for arbitrary existing files.
        // If it's about setting properties on a newly added slide via pptxgenjs, it might succeed.
        // For now, let's assume modifying an arbitrary existing slide's properties is not supported.
        // This needs to align with actual capabilities.
        // If pptxgenjs *can* set properties on a slide it controls (e.g. just added), this assertion needs change.
        // Based on "modifying existing shapes in a loaded file" being a limitation, this is likely also limited.
        expect(result.success).toBe(false); // Or true, if pptxgenjs supports this for slides it creates.
        if (!result.success) {
            expect((result as ErrorResponse).error.message).toMatch(/not supported|limited support/i);
        }
        // If it *is* supported for slides pptxgenjs creates:
        // expect(result.success).toBe(true);
        // TODO: Lib: Verify properties if successful
      }
    );
  });

  describe('powerpoint/shapes', () => {
    runTestForBothPaths(
      'should add a shape to a slide',
      'powerpoint/shapes/add',
      {
        // document will be set by runTestForBothPaths
        slideIndex: 1,
        shapeType: 'Rectangle',
        options: { x: 1, y: 1, w: 2, h: 1, text: 'Hello Shape' } // Example options
      },
      // COM Path Assertions
      (result) => {
        expect(result.success).toBe(true);
        // TODO: COM: Verify shape was added
      },
      // Library Path Assertions
      (result) => {
        // pptxgenjs can add shapes to slides it creates.
        expect(result.success).toBe(true);
        // TODO: Lib: Verify shape was added (e.g. text content via officeparser if it's a new file)
      },
      true // Assume adding shape with lib path implies it's to a presentation pptxgenjs controls/creates
    );

    // TODO: Add tests for other shape operations like get, set, delete, focusing on limitations for lib path
    // e.g., deleting a specific shape from an *existing* loaded presentation (not supported by pptxgenjs)
    runTestForBothPaths(
        'should fail to delete a specific shape from an existing presentation using library path',
        'powerpoint/shapes/delete', // Assuming a tool like this exists
        {
            // document will be set by runTestForBothPaths
            slideIndex: 1,
            shapeId: 'someShapeId', // or name
        },
        // COM Path Assertions
        (result) => {
            expect(result.success).toBe(true); // Assuming COM can delete shapes
            // TODO: COM: Verify shape was deleted
        },
        // Library Path Assertions
        (result) => {
            expect(result.success).toBe(false);
            expect((result as ErrorResponse).error.message).toMatch(/not supported/i);
        }
    );
  });

  describe('powerpoint/animations', () => {
    // TODO: Add tests for animation tools, focusing on limitations for lib path
    // e.g., adding complex animations (likely not supported by pptxgenjs)
    // e.g., reading existing animations (likely not supported by officeparser/pptxgenjs)
    runTestForBothPaths(
        'should fail to add complex animation using library path',
        'powerpoint/animations/add', // Assuming a tool like this
        {
            slideIndex: 1,
            shapeId: 'someShapeId',
            animationType: 'FlyIn',
            // ...other complex animation properties
        },
        (result) => { // COM
            expect(result.success).toBe(true); // Assuming COM supports it
        },
        (result) => { // Lib
            expect(result.success).toBe(false);
            expect((result as ErrorResponse).error.message).toMatch(/not supported|limited support/i);
        }
    );
  });

  describe('powerpoint/properties', () => {
    // TODO: Add tests for presentation properties tools
    // e.g., setting author, title (pptxgenjs likely supports this for new files)
    // e.g., reading properties from an existing file (officeparser might get some, pptxgenjs won't read existing)
     runTestForBothPaths(
        'should set document properties (e.g., title, author) for a new presentation',
        'powerpoint/properties/set', // Assuming a tool like this
        {
            properties: { title: 'Test Title', author: 'Test Author' },
            // document will be a new file for lib path
        },
        (result) => { // COM
            expect(result.success).toBe(true);
            // TODO: COM: Verify properties set
        },
        (result) => { // Lib
            expect(result.success).toBe(true); // pptxgenjs should support this for new files
            // TODO: Lib: Verify properties (might need to parse the file or check mock)
        },
        true // This operation makes most sense for a new file with pptxgenjs
    );

    runTestForBothPaths(
        'should get document properties',
        'powerpoint/properties/get', // Assuming a tool like this
        {
            // document will be an existing file
        },
        (result) => { // COM
            expect(result.success).toBe(true);
            if (result.success) { // Type guard
                expect((result as SuccessResponse).data).toHaveProperty('title');
            }
            // TODO: COM: Verify more properties
        },
        (result) => { // Lib
            // officeparser might extract some basic properties. pptxgenjs doesn't read existing files.
            // This depends on what the 'get' tool uses for the library path.
            // Assuming it uses officeparser for basic info:
            if (result.success) {
                expect(result.data).toBeDefined();
                // Check for properties officeparser can extract, e.g. text content based, not metadata like author directly
                // This assertion needs to be very specific to what officeparser provides.
                // For now, let's assume it might succeed with limited data or fail if it tries to get structured props.
                // If it's expected to fail for structured props:
                // expect(result.success).toBe(false);
                // expect((result as ErrorResponse).error.message).toMatch(/limited support/i);
            } else {
                 expect((result as ErrorResponse).error.message).toMatch(/not supported|limited support/i);
            }
            // A more realistic test for lib path 'get properties' might be:
            // 1. Create a presentation with pptxgenjs and set properties.
            // 2. THEN try to 'get' them using the lib path. If it internally uses officeparser, it might only get text.
            //    If it tries to use pptxgenjs to "read", it will fail.
            // This highlights the importance of knowing the tool's internal logic for the lib path.
        }
    );
  });


  // describe('powerpoint/charts', () => { // Commenting out as per thought process, may not be in scope
  //   test('should insert a chart into a slide', async () => {
  //     const outputPath = path.join(TEMP_DIR, 'insert_chart_test.pptx');
  //     // Copy the fixture to the temp directory
  //     await fs.copyFile(FIXTURE_PPTX_PATH, outputPath);
  //     const slideIndex = 1; // Assuming the fixture has at least one slide at index 1
  //     // Note: Inserting a chart usually requires data. This test assumes the tool
  //     // can insert a default chart or takes data as an argument (need to check API spec).

  //     const response = await fetch(`${MCP_SERVER_URL}/tool`, {
  //       method: 'POST',
  //       headers: {
  //         'Content-Type': 'application/json',
  //       },
  //       body: JSON.stringify({
  //         tool_name: 'powerpoint/charts/insert',
  //         arguments: {
  //           document: 'tests/temp_powerpoint_dir/insert_chart_test.pptx', // Use path relative to workspace
  //           slideIndex: slideIndex,
  //           chartType: 'ColumnClustered', // Example chart type - need to confirm supported types
  //           // Optional: data, position, size, title, etc.
  //           useComInterop: true, // Assuming default or explicit for existing tests
  //         },
  //       }),
  //     });

  //     expect(response.ok).toBe(true);
  //     const result = await response.json() as ToolResponse;

  //     expect(result.success).toBe(true);

  //     // TODO: Verify the chart was inserted. This requires inspecting the presentation structure.
  //   });
  // });

  // TODO: Add tests for complex variations and error handling for all PowerPoint tools
});