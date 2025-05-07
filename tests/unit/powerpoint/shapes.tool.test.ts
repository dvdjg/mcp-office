import powerpointShapesTool from '../../../src/tools/powerpoint/shapes.tool'; // Default import
import PptxGenJS from 'pptxgenjs';
// officeparser is not directly used for shape manipulation in the library path of shapes.tool
import fs from 'fs-extra';
import { ToolRequestParams } from '../../../src/types/common.types';
import * as path from 'path'; // Import path for resolving paths if needed by tool

// Mock PptxGenJS
jest.mock('pptxgenjs');

// Mock fs-extra
jest.mock('fs-extra', () => ({
  ...jest.requireActual('fs-extra'),
  readFile: jest.fn(),
  writeFile: jest.fn(), // Though pptxgenjs.writeFile is usually used
  pathExists: jest.fn(),
  ensureDir: jest.fn().mockResolvedValue(undefined),
}));

// Mock the dynamic resource saver
jest.mock('../../../src/tools/dynamic/resources.tool', () => ({
    saveResource: jest.fn().mockResolvedValue({ success: true, message: 'Resource saved' }),
}));


describe('PowerPoint Shapes Tool Unit Tests (Library Path)', () => {
  let mockPptxInstance: any;
  let mockSlideInstance: any;

  beforeEach(() => {
    jest.clearAllMocks();

    mockSlideInstance = {
      addText: jest.fn(),
      addShape: jest.fn(),
      // ... other slide methods
    };

    mockPptxInstance = {
      addSlide: jest.fn().mockReturnValue(mockSlideInstance),
      writeFile: jest.fn().mockResolvedValue(undefined),
      // Mock PptxGenJS.ShapeType if it's accessed directly, e.g. PptxGenJS.ShapeType.rect
      // This is needed because the actual PptxGenJS.ShapeType is an enum or object.
      ShapeType: { // Replicate the structure PptxGenJS uses if needed by the tool
        rect: 'rect', // Placeholder, actual values might be different
        ellipse: 'ellipse',
        // ... other shapes
      },
      // ... other presentation methods
    };
    (PptxGenJS as unknown as jest.Mock).mockImplementation(() => mockPptxInstance);
    // Also, ensure static properties like ShapeType are available on the mock constructor
     Object.defineProperty(PptxGenJS, 'ShapeType', {
        value: {
            rect: 'rect', // these are examples, use actual values from PptxGenJS if critical
            ellipse: 'ellipse',
            // add other shapes your tool might use
        },
        writable: true, // if you need to change it per test
        configurable: true,
    });


    (fs.pathExists as jest.Mock).mockResolvedValue(false); // Default: file does not exist
    (fs.readFile as unknown as jest.Mock).mockResolvedValue(Buffer.from('fake pptx content for resource saving'));
  });

  // --- Insert Shape Tests (Library Path) ---
  describe('powerpointShapesTool: insert operation (Library Path)', () => {
    test('should insert a shape (textbox) into a new presentation', async () => {
      const params: ToolRequestParams = {
        filePath: 'new_shape_test.pptx',
        operation: 'insert',
        slideIndex: 1, // PptxGenJS path in tool adds to first/new slide
        shapeType: 'msoShapeTextbox', // or just 'textbox'
        text: 'Hello World',
        position: { left: 72, top: 72 }, // 1 inch = 72 points
        size: { width: 144, height: 72 },
        useComInterop: false,
      };
      const result = await powerpointShapesTool.handler(params);

      expect(result.success).toBe(true);
      expect(PptxGenJS).toHaveBeenCalledTimes(1);
      expect(mockPptxInstance.addSlide).toHaveBeenCalledTimes(1);
      expect(mockSlideInstance.addText).toHaveBeenCalledWith(
        'Hello World',
        expect.objectContaining({
          x: 1, // 72pts / 72 = 1 inch
          y: 1,
          w: 2, // 144pts / 72 = 2 inches
          h: 1,
          fontSize: 18, // Default from tool's PptxGenJS path
        })
      );
      expect(mockPptxInstance.writeFile).toHaveBeenCalledWith({ fileName: path.resolve(params.filePath) });
    });

    test('should insert a rectangle shape into a new presentation', async () => {
      const params: ToolRequestParams = {
        filePath: 'new_rect_test.pptx',
        operation: 'insert',
        slideIndex: 1,
        shapeType: 'msoShapeRectangle', // or 'rectangle'
        position: { left: 100, top: 100 },
        size: { width: 200, height: 100 },
        useComInterop: false,
      };
      const result = await powerpointShapesTool.handler(params);

      expect(result.success).toBe(true);
      expect(PptxGenJS).toHaveBeenCalledTimes(1);
      expect(mockPptxInstance.addSlide).toHaveBeenCalledTimes(1);
      expect(mockSlideInstance.addShape).toHaveBeenCalledWith(
        PptxGenJS.ShapeType.rect, // Expecting the tool to map to PptxGenJS.ShapeType
        expect.objectContaining({
          x: 100 / 72,
          y: 100 / 72,
          w: 200 / 72,
          h: 100 / 72,
        })
      );
      expect(mockPptxInstance.writeFile).toHaveBeenCalledWith({ fileName: path.resolve(params.filePath) });
    });

    test('should insert an ellipse shape into a new presentation', async () => {
        const params: ToolRequestParams = {
          filePath: 'new_ellipse_test.pptx',
          operation: 'insert',
          slideIndex: 1,
          shapeType: 'msoShapeOval', // or 'ellipse'
          position: { left: 50, top: 50 },
          size: { width: 150, height: 75 },
          useComInterop: false,
        };
        const result = await powerpointShapesTool.handler(params);
  
        expect(result.success).toBe(true);
        expect(mockSlideInstance.addShape).toHaveBeenCalledWith(
          PptxGenJS.ShapeType.ellipse,
          expect.objectContaining({
            x: 50 / 72,
            y: 50 / 72,
            w: 150 / 72,
            h: 75 / 72,
          })
        );
        expect(mockPptxInstance.writeFile).toHaveBeenCalledWith({ fileName: path.resolve(params.filePath) });
      });

    test('should return error for unsupported shapeType on library path', async () => {
      const params: ToolRequestParams = {
        filePath: 'unsupported_shape.pptx',
        operation: 'insert',
        slideIndex: 1,
        shapeType: 'msoShapeStar', // Assuming this is not directly supported by the tool's lib path
        useComInterop: false,
      };
      const result = await powerpointShapesTool.handler(params);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('INVALID_PARAM'); // Or specific code from tool
        expect(result.error.message).toMatch(/Unsupported shapeType/i);
      }
    });

    test('should require shapeType for insert operation on library path', async () => {
        const params: ToolRequestParams = {
          filePath: 'no_shape_type.pptx',
          operation: 'insert',
          slideIndex: 1,
          // shapeType is missing
          useComInterop: false,
        };
        const result = await powerpointShapesTool.handler(params);
  
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('MISSING_PARAM');
          expect(result.error.message).toMatch(/shapeType is required/i);
        }
      });
  });

  // --- Modify, Format, Delete, List Shape Tests (Library Path) ---
  // These operations are generally not supported or very limited for PptxGenJS on existing files.
  // Tests should confirm they return the appropriate "not supported" or "limited support" messages.

  ['modify', 'format', 'delete'].forEach(operation => {
    describe(`powerpointShapesTool: ${operation} operation (Library Path)`, () => {
      test(`should report "not supported" for '${operation}' operation`, async () => {
        const params: ToolRequestParams = {
          filePath: 'existing_file.pptx',
          operation: operation as 'modify' | 'format' | 'delete',
          slideIndex: 1,
          shapeIndex: 1, // or shapeName
          // formatProperties: operation === 'format' ? { fillColor: "FF0000" } : undefined,
          useComInterop: false,
        };
        (fs.pathExists as jest.Mock).mockResolvedValue(true); // File "exists"
        const result = await powerpointShapesTool.handler(params);

        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('POWERPOINT_LIB_UNSUPPORTED');
          expect(result.error.message).toMatch(/not supported/i);
        }
      });
    });
  });

  describe('powerpointShapesTool: list operation (Library Path)', () => {
    test('should report "not supported" for list operation', async () => {
      const params: ToolRequestParams = {
        filePath: 'existing_file.pptx',
        operation: 'list',
        slideIndex: 1, // Optional for list
        useComInterop: false,
      };
      (fs.pathExists as jest.Mock).mockResolvedValue(true); // File "exists"
      const result = await powerpointShapesTool.handler(params);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('POWERPOINT_LIB_UNSUPPORTED');
        expect(result.error.message).toMatch(/not supported/i);
      }
    });
  });
});