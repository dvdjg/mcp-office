import PptxGenJS from 'pptxgenjs';
import officeParser from 'officeparser';
import fs from 'fs-extra';
import slidesTool from '../../../src/tools/powerpoint/slides.tool'; // Import default
// import { OfficeError } from '../../../src/utils/errorHandler'; // OfficeError is not exported

// Mock PptxGenJS
jest.mock('pptxgenjs');
const mockPptxGenJSWriteFile = jest.fn().mockResolvedValue(undefined); // writeFile is used by the tool
const mockAddSlide = jest.fn();
const mockDefineSlideMaster = jest.fn();
// const mockGetSlide = jest.fn(); // getSlide is not directly used by the library path in slides.tool.ts

// Assign mocks to the PptxGenJS prototype
(PptxGenJS.prototype as any).writeFile = mockPptxGenJSWriteFile;
(PptxGenJS.prototype as any).addSlide = mockAddSlide;
(PptxGenJS.prototype as any).defineSlideMaster = mockDefineSlideMaster;


// Mock officeParser
jest.mock('officeparser');
const mockParseOfficeAsync = officeParser.parseOfficeAsync as jest.Mock;

// Mock fs-extra
const mockEnsureDir = jest.fn();
const mockPathExists = jest.fn();
const mockReadFile = jest.fn();
const mockWriteFile = jest.fn();

jest.mock('fs-extra', () => ({
  ...jest.requireActual('fs-extra'), // Import and retain default behavior for other fs-extra functions if needed
  ensureDir: mockEnsureDir,
  pathExists: mockPathExists,
  readFile: mockReadFile,
  writeFile: mockWriteFile,
}));

describe('PowerPoint Slides Tool - Unit Tests (Library Path)', () => {
  const mockDocumentPath = 'mock/document.pptx';
  const mockNewDocumentPath = 'mock/new_document.pptx';

  beforeEach(() => {
    // Reset mocks before each test
    jest.clearAllMocks();

    // Default mock implementations
    mockEnsureDir.mockResolvedValue(undefined);
    mockPathExists.mockResolvedValue(true); // Assume file exists by default for modification tests
    mockAddSlide.mockImplementation(() => ({ // Mock what addSlide returns if chained
        addText: jest.fn(),
        addShape: jest.fn(),
        // ... other chainable methods
    }));
    // mockGetSlide.mockReturnValue(null); // Not used directly by lib path
  });

  describe('slidesTool.handler with operation "add" (useComInterop: false)', () => {
    test('should create a new presentation and add a slide if document does not exist', async () => {
      mockPathExists.mockResolvedValue(false); // Simulate document does not exist
      mockReadFile.mockResolvedValue(Buffer.from('mock pptx content')); // For saveResource

      const result = await slidesTool.handler({
        filePath: mockNewDocumentPath,
        operation: 'add',
        slideLayout: 'TITLE_SLIDE', // This will be mapped internally by the tool
        useComInterop: false,
      });

      expect(mockPathExists).toHaveBeenCalledWith(expect.stringContaining(mockNewDocumentPath));
      expect(PptxGenJS).toHaveBeenCalledTimes(1); // Constructor called for new presentation
      expect(mockAddSlide).toHaveBeenCalledTimes(1);
      // The tool maps 'TITLE_SLIDE' to 'TITLE_SLIDE' (or similar) for PptxGenJS
      expect(mockAddSlide).toHaveBeenCalledWith({ masterName: 'TITLE_SLIDE' });
      expect(mockPptxGenJSWriteFile).toHaveBeenCalledTimes(1);
      expect(mockPptxGenJSWriteFile).toHaveBeenCalledWith({ fileName: expect.stringContaining(mockNewDocumentPath) });
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toContain('Slide added using pptxgenjs');
      }
    });

    test('should "add" a slide to an "existing" presentation by creating a new one and overwriting', async () => {
      mockPathExists.mockResolvedValue(true);
      mockParseOfficeAsync.mockResolvedValue('Some text content from existing file');
      mockReadFile.mockResolvedValue(Buffer.from('mock pptx content')); // For saveResource

      const result = await slidesTool.handler({
        filePath: mockDocumentPath,
        operation: 'add',
        slideLayout: 'BLANK', // Tool maps to 'BLANK'
        useComInterop: false,
      });

      expect(PptxGenJS).toHaveBeenCalledTimes(1);
      expect(mockAddSlide).toHaveBeenCalledTimes(1);
      expect(mockAddSlide).toHaveBeenCalledWith({ masterName: 'BLANK' });
      expect(mockPptxGenJSWriteFile).toHaveBeenCalledTimes(1);
      expect(mockPptxGenJSWriteFile).toHaveBeenCalledWith({ fileName: expect.stringContaining(mockDocumentPath) });
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toContain('Slide added using pptxgenjs');
      }
    });

     test('should use a mapped default layout if none provided or mapping results in default', async () => {
      mockPathExists.mockResolvedValue(false);
      mockReadFile.mockResolvedValue(Buffer.from('mock pptx content')); // For saveResource
      await slidesTool.handler({
        filePath: mockNewDocumentPath,
        operation: 'add',
        // No slideLayout provided, tool should use a default mapping
        useComInterop: false,
      });
      // The tool's internal logic for default layout mapping will be tested here.
      // e.g. if undefined layout maps to 'BLANK' or 'TITLE_SLIDE' in the tool
      expect(mockAddSlide).toHaveBeenCalledWith({ masterName: expect.any(String) });
    });
  });

  describe('slidesTool.handler with operation "delete" (useComInterop: false)', () => {
    test('should return "not supported" error for library path', async () => {
      mockPathExists.mockResolvedValue(true); // Document exists

      const result = await slidesTool.handler({
        filePath: mockDocumentPath,
        operation: 'delete',
        slideIndex: 1,
        useComInterop: false,
      });

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.message).toMatch(/not supported/i);
        expect(result.error.code).toBe('POWERPOINT_LIB_UNSUPPORTED');
      }
      expect(mockPptxGenJSWriteFile).not.toHaveBeenCalled(); // No save attempt
    });
  });

  describe('slidesTool.handler with operation "set" (useComInterop: false)', () => {
    // The 'set' operation for library path is documented to return 'limited support' or 'not supported'
    // as pptxgenjs is for generation, not modification of arbitrary existing files.
    // Unit tests should confirm this documented behavior.

    test('should return "limited support" error when trying to "set" properties on an existing file', async () => {
        mockPathExists.mockResolvedValue(true); // Document exists

        const result = await slidesTool.handler({
            filePath: mockDocumentPath,
            operation: 'set',
            slideIndex: 1,
            // properties: { background: { color: '00FF00' } }, // Properties are not part of schema for set
            useComInterop: false,
        });

        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.message).toMatch(/limited support/i);
            expect(result.error.code).toBe('POWERPOINT_LIB_LIMITED_SUPPORT');
        }
        expect(mockPptxGenJSWriteFile).not.toHaveBeenCalled();
    });


    // If 'set' was intended to work on a NEWLY created presentation context by pptxgenjs (e.g. add a slide then set its props)
    // the tests would be different. But current slides.tool.ts logic for 'set' with lib path directly returns error.
    // So, the above test reflects the current implementation.
  });

  describe('slidesTool.handler with operation "getText" (useComInterop: false)', () => {
    test('should extract text using officeparser if file exists', async () => {
        mockPathExists.mockResolvedValue(true);
        const mockFileBuffer = Buffer.from('mock file content');
        mockReadFile.mockResolvedValue(mockFileBuffer);
        const expectedText = "This is text from the PowerPoint.";
        mockParseOfficeAsync.mockResolvedValue(expectedText);

        const result = await slidesTool.handler({
            filePath: mockDocumentPath,
            operation: 'getText',
            useComInterop: false,
        });

        expect(mockPathExists).toHaveBeenCalledWith(mockDocumentPath);
        expect(mockReadFile).toHaveBeenCalledWith(mockDocumentPath);
        expect(mockParseOfficeAsync).toHaveBeenCalledWith(mockFileBuffer);
        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data).toBe(expectedText);
        }
    });

    test('should return error if file does not exist for getText', async () => {
        mockPathExists.mockResolvedValue(false);

        const result = await slidesTool.handler({
            filePath: mockDocumentPath,
            operation: 'getText',
            useComInterop: false,
        });

        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('FILE_NOT_FOUND');
        }
    });
test('should return error if officeparser fails', async () => {
    mockPathExists.mockResolvedValue(true);
    const mockFileBuffer = Buffer.from('mock file content');
    mockReadFile.mockResolvedValue(mockFileBuffer);
    mockParseOfficeAsync.mockRejectedValue(new Error("Parsing failed"));


        const result = await slidesTool.handler({
            filePath: mockDocumentPath,
            operation: 'getText',
            useComInterop: false,
        });

        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('OFFICEPARSER_ERROR');
            expect(result.error.message).toContain("Parsing failed");
        }
    });
  });
  // TODO: Add more tests for specific properties and edge cases
  // e.g., what happens if slideIndex is out of bounds for a "new" presentation context?
  // e.g., testing various property types (colors, booleans, numbers)
}); // This closes the main 'PowerPoint Slides Tool - Unit Tests (Library Path)' describe block