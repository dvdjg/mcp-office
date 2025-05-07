import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
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
    comPathConfig: {
      requiresFixture: boolean;
      assertions: (result: ToolResponse, effectiveDocPath: string) => Promise<void> | void;
    },
    libPathConfig: {
      // if true, FIXTURE_PPTX_PATH is copied to a unique path for the test.
      // if false, tool is expected to create a new file at effectiveDocPath or operate without a pre-existing file.
      requiresFixture: boolean;
      // if true, the library path is expected to create a new file.
      // This helps in deciding if the absence of a file before the test is an error or expected.
      createsNewFile?: boolean;
      assertions: (result: ToolResponse, effectiveDocPath: string) => Promise<void> | void;
    }
  ) => {
    describe(testName, () => {
      // COM Path Test
      test(`COM Path: ${testName}`, async () => {
        const uniqueId = `com_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
        const docFileName = `${toolName.replace(/\//g, '_')}_${uniqueId}.pptx`;
        const effectiveDocPath = path.join(TEMP_DIR, docFileName);
        const relativeDocPath = path.join('tests/temp_powerpoint_dir', docFileName);

        if (comPathConfig.requiresFixture) {
          await fs.copyFile(FIXTURE_PPTX_PATH, effectiveDocPath);
        } else {
          // Ensure the path is clear if we are not starting with a fixture (e.g. tool creates new file)
           try { await fs.unlink(effectiveDocPath); } catch (e:any) { if (e.code !== 'ENOENT') throw e; }
        }

        const response = await fetch(`${MCP_SERVER_URL}/tool`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            tool_name: toolName,
            arguments: {
              ...baseArguments,
              document: relativeDocPath, // Tool always operates on this path
              useComInterop: true,
            },
          }),
        });
        // COM path should generally respond with ok if the tool is found and args are valid,
        // success/failure is in the JSON body.
        const resultText = await response.text();
        let result: ToolResponse;
        try {
          result = JSON.parse(resultText) as ToolResponse;
        } catch (e) {
          console.error("Failed to parse COM path response JSON:", resultText);
          throw e;
        }
        
        if (!response.ok && result.success !== false) {
            // If response is not ok, but our parsed result doesn't explicitly state success is false,
            // then something unexpected happened at the HTTP level or with the tool's basic error handling.
            console.error(`COM Path HTTP error for ${testName}: ${response.status} ${response.statusText}`, resultText);
            expect(response.ok).toBe(true); // This will fail and show the error
        }
        await comPathConfig.assertions(result, effectiveDocPath);
      });

      // Library Path Test
      test(`Library Path: ${testName}`, async () => {
        const uniqueId = `lib_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
        const docFileName = `${toolName.replace(/\//g, '_')}_${uniqueId}.pptx`;
        const effectiveDocPath = path.join(TEMP_DIR, docFileName);
        const relativeDocPath = path.join('tests/temp_powerpoint_dir', docFileName);

        if (libPathConfig.requiresFixture) {
          await fs.copyFile(FIXTURE_PPTX_PATH, effectiveDocPath);
        } else {
          // If not requiring a fixture, ensure the path is clear, especially if createsNewFile is true.
          // If createsNewFile is false and requiresFixture is false, it implies an in-memory operation
          // or an operation that doesn't need a source file but might save to `document`.
           try { await fs.unlink(effectiveDocPath); } catch (e:any) { if (e.code !== 'ENOENT') throw e; }
        }

        const response = await fetch(`${MCP_SERVER_URL}/tool`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            tool_name: toolName,
            arguments: {
              ...baseArguments,
              document: relativeDocPath, // Tool operates on/creates this path
              useComInterop: false,
            },
          }),
        });
        
        const resultText = await response.text();
        let result: ToolResponse;
        try {
          result = JSON.parse(resultText) as ToolResponse;
        } catch (e) {
          console.error("Failed to parse Library path response JSON:", resultText);
          throw e;
        }

        // For library path, response.ok might be true even for "not supported" if the tool handles it gracefully.
        // The key is in the result.success and error messages.
        // However, if it's a genuine server error (500), response.ok will be false.
        if (!response.ok && result.success !== false) {
             console.error(`Library Path HTTP error for ${testName}: ${response.status} ${response.statusText}`, resultText);
             expect(response.ok).toBe(true); // This will fail and show the error
        }
        await libPathConfig.assertions(result, effectiveDocPath);
      });
    });
  };

  describe('powerpoint/slides', () => {
    runTestForBothPaths(
      'should add a new slide',
      'powerpoint/slides/add',
      {
        layout: 'TitleAndContent', // Example layout
      },
      // COM Path Config
      {
        requiresFixture: false, // COM add slide can create a new file or add to an existing one.
                               // For this test, let's assume it creates/uses the specified 'document' path.
                               // If it needs an existing one, set to true and it will use a copy of FIXTURE_PPTX_PATH.
                               // Let's test adding to a new document.
        assertions: async (result, effectiveDocPath) => {
          expect(result.success).toBe(true);
          // TODO: COM: Verify slide was actually added.
          // This might involve trying to get slide count or properties from effectiveDocPath.
          // For now, check if file was created.
          const stats = await fs.stat(effectiveDocPath);
          expect(stats.isFile()).toBe(true);
          expect(stats.size).toBeGreaterThan(0); // Basic check
        },
      },
      // Library Path Config
      {
        requiresFixture: false, // pptxgenjs will create a new presentation if document doesn't exist or add to one it manages.
        createsNewFile: true,   // Explicitly state it creates a new file.
        assertions: async (result, effectiveDocPath) => {
          expect(result.success).toBe(true); // pptxgenjs should succeed in adding a slide.
          // TODO: Lib: Verify slide was added.
          // This could involve checking pptxgenjs mocks in unit tests.
          // For E2E, check if file was created and has content.
          const stats = await fs.stat(effectiveDocPath);
          expect(stats.isFile()).toBe(true);
          expect(stats.size).toBeGreaterThan(0); // Basic check for pptxgenjs output
        },
      }
    );

    runTestForBothPaths(
      'should delete a slide from an existing presentation',
      'powerpoint/slides/delete',
      {
        slideIndex: 1, // Assuming 1-based index for deletion as per typical user expectation.
                       // Tool should clarify if 0-based or 1-based.
      },
      // COM Path Config
      {
        requiresFixture: true, // Deleting a slide requires an existing presentation with slides.
        assertions: (result, effectiveDocPath) => {
          expect(result.success).toBe(true);
          // TODO: COM: Verify slide was actually deleted (e.g., by checking slide count before/after).
        },
      },
      // Library Path Config
      {
        requiresFixture: true, // Test deleting from an "existing" file for lib path.
        createsNewFile: false,
        assertions: (result, effectiveDocPath) => {
          // pptxgenjs does NOT support deleting slides from an arbitrary existing presentation.
          expect(result.success).toBe(false);
          expect((result as ErrorResponse).error.message).toMatch(/not supported/i);
          // Ensure the original file is not inadvertently modified or deleted if the operation fails.
          // This might be tricky if the tool attempts to save over it.
          // For now, focus on the error message.
        },
      }
    );

    runTestForBothPaths(
      'should set properties of a slide',
      'powerpoint/slides/set',
      {
        slideIndex: 1,
        properties: { background: { color: 'FFFF00' } }, // Example property
      },
      // COM Path Config
      {
        requiresFixture: true, // Setting properties requires an existing slide.
        assertions: (result, effectiveDocPath) => {
          expect(result.success).toBe(true);
          // TODO: COM: Verify slide properties were set (e.g., by getting slide properties).
        },
      },
      // Library Path Config
      {
        // Test 1: Setting properties on a slide in a NEWLY CREATED presentation by pptxgenjs
        // This would typically be part of an 'add slide with properties' flow.
        // For a dedicated 'set' tool on lib path, if it implies loading an existing file, it should fail.
        // Let's assume 'set' for lib path means operating on a file pptxgenjs *could* have just made.
        // However, the most common interpretation of a 'set' tool is modifying an *existing* entity.
        requiresFixture: true, // To test modification of an "existing" slide for lib path.
        createsNewFile: false,
        assertions: (result, effectiveDocPath) => {
          // Modifying an arbitrary existing slide's properties using pptxgenjs (by loading the file) is not supported.
          // pptxgenjs can set properties when it *creates* the slide.
          expect(result.success).toBe(false);
          expect((result as ErrorResponse).error.message).toMatch(/not supported|limited support/i);
        },
      }
    );

    // New test case for library path: setting properties on a slide during creation (if applicable)
    // This might be better as a unit test for the tool's internal logic when using pptxgenjs.
    // Or, if 'powerpoint/slides/add' can take properties, test it there.
    // For now, the 'set' tool on an existing file for lib path should report not supported.
  });

  describe('powerpoint/shapes', () => {
    runTestForBothPaths(
      'should add a shape to a slide',
      'powerpoint/shapes/add',
      {
        slideIndex: 1,
        shapeType: 'Rectangle', // Example, ensure this is a valid type for the tool
        options: { x: 1, y: 1, w: 2, h: 1, text: 'Hello Shape' }
      },
      // COM Path Config
      {
        requiresFixture: true, // Adding a shape requires an existing slide in an existing presentation.
        assertions: async (result, effectiveDocPath) => {
          expect(result.success).toBe(true);
          // TODO: COM: Verify shape was added (e.g., by trying to get shape properties or count shapes).
          // For now, check the file was modified (timestamp or size might change)
          // More robust: use another tool to get shape count or specific shape details.
        }
      },
      // Library Path Config
      {
        requiresFixture: false, // pptxgenjs adds shapes to the presentation it's building.
        createsNewFile: true,   // Assumes it will create a new file if one isn't "active" in its context.
        assertions: async (result, effectiveDocPath) => {
          expect(result.success).toBe(true); // pptxgenjs can add shapes.
          const stats = await fs.stat(effectiveDocPath);
          expect(stats.isFile()).toBe(true);
          expect(stats.size).toBeGreaterThan(0);
          // TODO: Lib: Verify shape was added. For E2E, if text was added, officeparser might be used on effectiveDocPath to check.
        }
      }
    );

    runTestForBothPaths(
      'should fail to delete a specific shape from an existing presentation using library path',
      'powerpoint/shapes/delete', // Assuming a tool like this exists
      {
        slideIndex: 1,
        shapeId: 'someShapeId', // or name, depends on tool's API
      },
      // COM Path Config
      {
        requiresFixture: true, // Deleting a shape requires an existing presentation and shape.
        assertions: (result, effectiveDocPath) => {
          expect(result.success).toBe(true); // Assuming COM can delete shapes.
          // TODO: COM: Verify shape was actually deleted.
        }
      },
      // Library Path Config
      {
        requiresFixture: true, // Test deleting from an "existing" file.
        createsNewFile: false,
        assertions: (result, effectiveDocPath) => {
          // pptxgenjs does not support modifying/deleting shapes in an arbitrary existing presentation.
          expect(result.success).toBe(false);
          expect((result as ErrorResponse).error.message).toMatch(/not supported/i);
        }
      }
    );
    // TODO: Add tests for other shape operations like get, set, focusing on limitations for lib path
  });

  describe('powerpoint/animations', () => {
    runTestForBothPaths(
      'should fail to add complex animation using library path',
      'powerpoint/animations/add', // Assuming a tool for adding animations
      {
        slideIndex: 1,
        shapeId: 'someShapeId', // ID of the shape to animate
        animationType: 'FlyIn', // Example animation
        // ...other complex animation properties that pptxgenjs might not support
      },
      // COM Path Config
      {
        requiresFixture: true, // Animations apply to shapes on slides in existing presentations.
        assertions: (result, effectiveDocPath) => {
          expect(result.success).toBe(true); // Assuming COM supports adding this animation.
          // TODO: COM: Verify animation was added.
        }
      },
      // Library Path Config
      {
        requiresFixture: true, // Test applying to an "existing" file context.
        createsNewFile: false,
        assertions: (result, effectiveDocPath) => {
          // pptxgenjs has very limited or no support for complex animations or reading existing ones.
          expect(result.success).toBe(false);
          expect((result as ErrorResponse).error.message).toMatch(/not supported|limited support/i);
        }
      }
    );
    // TODO: Add tests for reading animations (likely not supported by lib path).
  });

  describe('powerpoint/properties', () => {
    runTestForBothPaths(
      'should set document properties (e.g., title, author) for a presentation',
      'powerpoint/properties/set', // Assuming a tool for setting presentation-level properties
      {
        properties: { title: 'Test Presentation Title', author: 'E2E Test Author' },
      },
      // COM Path Config
      {
        requiresFixture: false, // COM can set properties on a new or existing file. Test on a new one.
        assertions: async (result, effectiveDocPath) => {
          expect(result.success).toBe(true);
          const stats = await fs.stat(effectiveDocPath);
          expect(stats.isFile()).toBe(true);
          expect(stats.size).toBeGreaterThan(0); // Basic check it created something
          // TODO: COM: Verify properties were set (e.g., by using a 'get' properties tool or inspecting file).
        }
      },
      // Library Path Config
      {
        requiresFixture: false, // pptxgenjs sets these on the new presentation it creates.
        createsNewFile: true,
        assertions: async (result, effectiveDocPath) => {
          expect(result.success).toBe(true); // pptxgenjs supports setting author, title, etc.
          const stats = await fs.stat(effectiveDocPath);
          expect(stats.isFile()).toBe(true);
          expect(stats.size).toBeGreaterThan(0);
          // TODO: Lib: Verify properties. officeparser might not easily get these metadata fields from the generated file.
          // Unit tests would mock pptxgenjs.
        }
      }
    );

    runTestForBothPaths(
      'should get document properties from an existing presentation',
      'powerpoint/properties/get', // Assuming a tool for getting presentation-level properties
      {}, // No specific arguments other than the document path.
      // COM Path Config
      {
        requiresFixture: true, // Getting properties requires an existing presentation.
        assertions: (result, effectiveDocPath) => {
          expect(result.success).toBe(true);
          if (result.success) { // Type guard
            expect((result as SuccessResponse).data).toBeDefined();
            // Example: expect((result as SuccessResponse).data.title).toBeDefined();
            // This depends on the FIXTURE_PPTX_PATH content.
          }
          // TODO: COM: Verify specific properties based on the FIXTURE_PPTX_PATH content.
        }
      },
      // Library Path Config
      {
        requiresFixture: true, // Test getting from an "existing" file.
        createsNewFile: false,
        assertions: (result, effectiveDocPath) => {
          // officeparser (if used by the tool for lib path 'get') might extract some text-based content,
          // but not necessarily structured metadata like 'author' or 'title' in a clean way.
          // pptxgenjs does not read/parse existing presentations.
          // So, this operation is likely "not supported" or "limited" for the library path.
          expect(result.success).toBe(false);
          if (result.success === false) {
            expect((result as ErrorResponse).error.message).toMatch(/not supported|limited support/i);
          } else {
            // This case should ideally not be hit if the expectation is 'not supported'.
            // If it can succeed with partial data via officeparser, the tool's contract should be clear.
            console.warn("Library path for 'get properties' unexpectedly succeeded. Review assertions and tool behavior.");
            expect((result as SuccessResponse).data).toBeDefined();
          }
        }
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