import slidesTool from '../../../src/tools/powerpoint/slides.tool'; // Default import
import PptxGenJS from 'pptxgenjs';
import officeParser from 'officeparser';
import fs from 'fs-extra';
import { ToolRequestParams } from '../../../src/types/common.types'; // For typing params

// Mock PptxGenJS
// PptxGenJS is a class, so we mock its constructor and instance methods.
jest.mock('pptxgenjs');

// Mock officeParser
jest.mock('officeparser', () => ({
  parseOfficeAsync: jest.fn(),
}));

// Mock fs-extra
jest.mock('fs-extra', () => ({
  ...jest.requireActual('fs-extra'),
  readFile: jest.fn(),
  writeFile: jest.fn(), // Though pptxgenjs.writeFile is usually used for saving
  pathExists: jest.fn(),
  ensureDir: jest.fn().mockResolvedValue(undefined),
}));

describe('PowerPoint Slides Tool Unit Tests (Library Path)', () => {
  let mockPptxInstance: any;

  beforeEach(() => {
    jest.clearAllMocks();

    // Setup default mock for PptxGenJS instance
    mockPptxInstance = {
      addSlide: jest.fn((options) => {
        // console.log('Mock PptxGenJS instance.addSlide called with:', options);
        return { // Return a mock slide object
          addText: jest.fn(),
          addShape: jest.fn(),
          // ... other slide methods that might be chained or used
        };
      }),
      // save: jest.fn().mockResolvedValue('mocked_pptx_data'), // .writeFile is used in slides.tool.ts
      writeFile: jest.fn().mockResolvedValue(undefined), // Mock writeFile
      defineLayout: jest.fn(),
      defineSlideMaster: jest.fn(),
      layouts: {
        LAYOUT_WIDE: { name: 'LAYOUT_WIDE', width: 1280, height: 720 },
        TITLE_SLIDE: { name: 'TITLE_SLIDE', width: 1280, height: 720 },
        BLANK: { name: 'BLANK', width: 1280, height: 720 },
        // Add other layouts your code might expect or map to
      },
      masters: {}, // Mock masters if used
    };
    (PptxGenJS as unknown as jest.Mock).mockImplementation(() => mockPptxInstance);

    (officeParser.parseOfficeAsync as jest.Mock).mockResolvedValue('Parsed text content from PowerPoint');
    (fs.pathExists as jest.Mock).mockResolvedValue(false); // Default: file does not exist (for creation tests)
    (fs.readFile as unknown as jest.Mock).mockResolvedValue(Buffer.from('fake pptx content')); // For officeParser
  });

  // --- Add Slide Tests ---
  describe('slidesTool: add operation (Library Path)', () => {
    test('should create a new presentation and add a slide if filePath does not exist', async () => {
      (fs.pathExists as jest.Mock).mockResolvedValue(false);
      const params: ToolRequestParams = {
        filePath: 'new_presentation.pptx',
        operation: 'add',
        slideLayout: 'LAYOUT_WIDE', // This will be mapped by the tool
        useComInterop: false,
      };
      const result = await slidesTool.handler(params);

      expect(result.success).toBe(true);
      expect(PptxGenJS).toHaveBeenCalledTimes(1); // Constructor called
      expect(mockPptxInstance.addSlide).toHaveBeenCalled(); // Check if addSlide was called
      // Check layout mapping if possible, e.g. if 'LAYOUT_WIDE' maps to 'TITLE_SLIDE' or similar in the tool
      // For now, just check it was called. Specifics depend on tool's mapping logic.
      // expect(mockPptxInstance.addSlide).toHaveBeenCalledWith(expect.objectContaining({ masterName: 'TITLE_SLIDE' })); // if mapped
      expect(mockPptxInstance.writeFile).toHaveBeenCalledWith({ fileName: params.filePath });
      expect(fs.writeFile).not.toHaveBeenCalled(); // pptxgenjs writeFile handles file system interaction
    });

    test('should "add" a slide (effectively creating new) even if filePath exists for pptxgenjs', async () => {
      (fs.pathExists as jest.Mock).mockResolvedValue(true); // File "exists"
      const params: ToolRequestParams = {
        filePath: 'existing_presentation.pptx',
        operation: 'add',
        slideLayout: 'BLANK',
        useComInterop: false,
      };
      const result = await slidesTool.handler(params);

      expect(result.success).toBe(true);
      expect(PptxGenJS).toHaveBeenCalledTimes(1);
      expect(mockPptxInstance.addSlide).toHaveBeenCalled();
      expect(mockPptxInstance.writeFile).toHaveBeenCalledWith({ fileName: params.filePath });
    });

    test('should use specified slideLayout (mapped) when adding a slide', async () => {
        const params: ToolRequestParams = {
            filePath: 'layout_test.pptx',
            operation: 'add',
            slideLayout: 'Custom Title Layout', // Tool should map this
            useComInterop: false,
        };
        // Assume 'Custom Title Layout' maps to 'TITLE_SLIDE' in the tool's logic
        await slidesTool.handler(params);
        expect(mockPptxInstance.addSlide).toHaveBeenCalledWith(expect.objectContaining({ masterName: 'TITLE_SLIDE' }));
    });
  });

  // --- Delete Slide Tests ---
  describe('slidesTool: delete operation (Library Path)', () => {
    test('should report "not supported" for library path', async () => {
      const params: ToolRequestParams = {
        filePath: 'any_presentation.pptx',
        operation: 'delete',
        slideIndex: 1,
        useComInterop: false,
      };
      (fs.pathExists as jest.Mock).mockResolvedValue(true);
      const result = await slidesTool.handler(params);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.message).toMatch(/not supported/i);
      }
    });
  });

  // --- Get Slide/Text Tests ---
  describe('slidesTool: getText operation (Library Path)', () => {
    test('should extract text using officeParser if file exists', async () => {
      (fs.pathExists as jest.Mock).mockResolvedValue(true);
      const mockFileContent = Buffer.from('mock pptx file content for getText');
      (fs.readFile as unknown as jest.Mock).mockResolvedValue(mockFileContent);
      const mockParsedText = 'Extracted text via officeParser for getText';
      (officeParser.parseOfficeAsync as jest.Mock).mockResolvedValue(mockParsedText);

      const params: ToolRequestParams = {
        filePath: 'text_extract.pptx',
        operation: 'getText',
        useComInterop: false,
      };
      const result = await slidesTool.handler(params);

      expect(result.success).toBe(true);
      if (result.success) { // Type guard
        expect(fs.readFile).toHaveBeenCalledWith(params.filePath);
        expect(officeParser.parseOfficeAsync).toHaveBeenCalledWith(mockFileContent);
        expect(result.data).toBe(mockParsedText);
      } else {
        throw new Error('Test failed: Expected success true for getText operation'); // Should not happen if success is true
      }
    });

    test('should return file not found if filePath does not exist for getText', async () => {
        (fs.pathExists as jest.Mock).mockResolvedValue(false);
        const params: ToolRequestParams = {
            filePath: 'non_existent.pptx',
            operation: 'getText',
            useComInterop: false,
        };
        const result = await slidesTool.handler(params);
        expect(result.success).toBe(false);
        if(!result.success) {
            expect(result.error.code).toBe('FILE_NOT_FOUND');
        }
    });
  });


  // --- Set Slide Properties Tests ---
  describe('slidesTool: set operation (Library Path)', () => {
    test('should report "limited support" or "not supported" when trying to set properties for library path', async () => {
      const params: ToolRequestParams = {
        filePath: 'existing_presentation.pptx',
        operation: 'set',
        slideIndex: 1,
        // properties: { background: { color: 'FF0000' } }, // 'properties' is not in SlidesToolInputSchema
        useComInterop: false,
      };
      (fs.pathExists as jest.Mock).mockResolvedValue(true);
      const result = await slidesTool.handler(params);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.message).toMatch(/limited support|not supported/i);
      }
    });
  });
  // TODO: Add more tests for edge cases, error handling (e.g. officeParser failure), and different configurations.
});