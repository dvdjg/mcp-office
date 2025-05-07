import PptxGenJS from 'pptxgenjs';
import fs from 'fs-extra';
import powerpointShapesTool from '../../../src/tools/powerpoint/shapes.tool'; // Default import
import { validateFilePath } from '../../../src/utils/security';

// Mock PptxGenJS
jest.mock('pptxgenjs');
const mockPptxGenJSWriteFile = jest.fn().mockResolvedValue(undefined);
const mockAddSlide = jest.fn();
const mockAddShape = jest.fn();
const mockAddText = jest.fn();

// Assign mocks to the PptxGenJS prototype and class/static properties if needed
(PptxGenJS.prototype as any).writeFile = mockPptxGenJSWriteFile;
(PptxGenJS.prototype as any).addSlide = mockAddSlide;

// Mocking methods on the slide object returned by addSlide
mockAddSlide.mockImplementation(() => ({
  addShape: mockAddShape,
  addText: mockAddText,
  // ... other slide methods if used by the tool for shapes
}));

// Mock PptxGenJS.ShapeType if it's directly used by the tool
(PptxGenJS as any).ShapeType = {
    rect: 'rect',
    ellipse: 'ellipse',
    // Add other shapes if the tool uses them directly from PptxGenJS.ShapeType
};


// Mock fs-extra
const mockPathExists = jest.fn();
const mockReadFile = jest.fn();
// No need to mock ensureDir or writeFile for fs-extra if PptxGenJS's writeFile is handling file creation.
jest.mock('fs-extra', () => ({
  ...jest.requireActual('fs-extra'),
  pathExists: mockPathExists,
  readFile: mockReadFile,
}));

// Mock security utility
jest.mock('../../../src/utils/security', () => ({
    validateFilePath: jest.fn().mockReturnValue(true), // Assume valid path by default
}));


describe('PowerPoint Shapes Tool - Unit Tests (Library Path)', () => {
  const mockDocumentPath = 'secure/mock/document.pptx'; // Path that would pass basic validation
  const mockNewDocumentPath = 'secure/mock/new_document.pptx';

  beforeEach(() => {
    jest.clearAllMocks();
    // Default mock implementations for each test
    mockPathExists.mockResolvedValue(true); // Assume file exists for modification tests
    mockReadFile.mockResolvedValue(Buffer.from('mock pptx content')); // For saveResource
    (validateFilePath as jest.Mock).mockReturnValue(true); // Reset to true for each test
  });

  describe('powerpointShapesTool.handler with operation "insert" (useComInterop: false)', () => {
    test('should insert a shape into a new presentation if document does not exist', async () => {
      mockPathExists.mockResolvedValue(false); // Simulate document does not exist

      const result = await powerpointShapesTool.handler({
        filePath: mockNewDocumentPath,
        operation: 'insert',
        slideIndex: 1, // PptxGenJS path in tool might simplify to always first/new slide
        shapeType: 'msoShapeRectangle', // Tool should map this
        position: { left: 72, top: 72 }, // 1 inch, 1 inch
        size: { width: 144, height: 72 }, // 2 inch, 1 inch
        useComInterop: false,
      });

      expect(validateFilePath).toHaveBeenCalledWith(mockNewDocumentPath);
      expect(mockPathExists).toHaveBeenCalledWith(expect.stringContaining(mockNewDocumentPath));
      expect(PptxGenJS).toHaveBeenCalledTimes(1);
      expect(mockAddSlide).toHaveBeenCalledTimes(1);
      // Tool maps 'msoShapeRectangle' to PptxGenJS.ShapeType.rect
      expect(mockAddShape).toHaveBeenCalledWith(PptxGenJS.ShapeType.rect, expect.objectContaining({ x: 1, y: 1, w: 2, h: 1 }));
      expect(mockPptxGenJSWriteFile).toHaveBeenCalledTimes(1);
      expect(mockPptxGenJSWriteFile).toHaveBeenCalledWith({ fileName: expect.stringContaining(mockNewDocumentPath) });
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toContain('PptxGenJS: Shape inserted');
      }
    });

    test('should insert a textbox shape with text', async () => {
        mockPathExists.mockResolvedValue(false);
        const textToInsert = "Hello PptxGenJS";
        const result = await powerpointShapesTool.handler({ // Declare result here
            filePath: mockNewDocumentPath,
            operation: 'insert',
            slideIndex: 1,
            shapeType: 'msoShapeTextbox', // or just rely on 'text' field
            text: textToInsert,
            position: { left: 72, top: 144 },
            size: { width: 288, height: 36 },
            useComInterop: false,
        });
        expect(mockAddText).toHaveBeenCalledWith(textToInsert, expect.objectContaining({ x: 1, y: 2, w: 4, h: 0.5 }));
        expect(result.success).toBe(true);
    });

    test('should return error if shapeType is missing for insert', async () => {
        const result = await powerpointShapesTool.handler({
            filePath: mockNewDocumentPath,
            operation: 'insert',
            slideIndex: 1,
            // shapeType is missing
            useComInterop: false,
        });
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('MISSING_PARAM'); // Or as defined in the tool
            expect(result.error.message).toContain('shapeType is required');
        }
    });

    test('should return error for unsupported shapeType for insert', async () => {
        const result = await powerpointShapesTool.handler({
            filePath: mockNewDocumentPath,
            operation: 'insert',
            slideIndex: 1,
            shapeType: 'unsupportedShapeType123',
            useComInterop: false,
        });
        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('INVALID_PARAM');
            expect(result.error.message).toContain('Unsupported shapeType');
        }
    });
  });

  describe('powerpointShapesTool.handler with "modify", "format", "delete" (useComInterop: false)', () => {
    const operationsToTest: ('modify' | 'format' | 'delete')[] = ['modify', 'format', 'delete'];
    operationsToTest.forEach(operation => {
      test(`should return "not supported" error for operation "${operation}"`, async () => {
        mockPathExists.mockResolvedValue(true); // Document exists

        const result = await powerpointShapesTool.handler({
          filePath: mockDocumentPath,
          operation: operation,
          slideIndex: 1,
          shapeIndex: 1, // or shapeName
          // formatProperties needed for 'format' but error should occur before validation
          useComInterop: false,
        });

        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.message).toMatch(/not supported/i);
          expect(result.error.code).toBe('POWERPOINT_LIB_UNSUPPORTED');
        }
        expect(mockPptxGenJSWriteFile).not.toHaveBeenCalled();
      });
    });
  });

  describe('powerpointShapesTool.handler with operation "list" (useComInterop: false)', () => {
    test('should return "not supported" error for library path', async () => {
      mockPathExists.mockResolvedValue(true); // Document exists

      const result = await powerpointShapesTool.handler({
        filePath: mockDocumentPath,
        operation: 'list',
        // slideIndex is optional for list
        useComInterop: false,
      });

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.message).toMatch(/not supported/i);
        expect(result.error.code).toBe('POWERPOINT_LIB_UNSUPPORTED');
      }
    });
  });

  test('should return validation error for invalid filePath', async () => {
    (validateFilePath as jest.Mock).mockReturnValue(false); // Simulate invalid path
    const result = await powerpointShapesTool.handler({
        filePath: 'invalid/../path.pptx',
        operation: 'insert',
        slideIndex: 1,
        shapeType: 'msoShapeRectangle',
        useComInterop: false,
    });
    expect(validateFilePath).toHaveBeenCalledWith('invalid/../path.pptx');
    expect(result.success).toBe(false);
    if(!result.success){
        // This depends on how Zod errors are translated by the main handler or if caught by createErrorResponse
        // Assuming it's caught by the tool's Zod parsing and returned as VALIDATION_ERROR
        expect(result.error.code).toBe('VALIDATION_ERROR'); // Or specific code from tool
        expect(result.error.message).toContain('Input validation failed'); // Or "Invalid or potentially unsafe file path"
    }
  });

  // TODO: Add tests for PptxGenJS specific formatting options if the tool implements them for 'insert'.
  // TODO: Test edge cases for position and size conversions (e.g., undefined values).
});