import PptxGenJS from 'pptxgenjs';
import fs from 'fs-extra';
import animationsTool from '../../../src/tools/powerpoint/animations.tool'; // Default import
import { validateFilePath } from '../../../src/utils/security'; // Assuming this might be used or added

// Mock PptxGenJS
jest.mock('pptxgenjs');
const mockPptxGenJSWriteFile = jest.fn().mockResolvedValue(undefined);
const mockAddSlide = jest.fn();
const mockAddTextWithAnimation = jest.fn(); // Specific for testing animation on text

// Assign mocks to the PptxGenJS prototype
(PptxGenJS.prototype as any).writeFile = mockPptxGenJSWriteFile;
(PptxGenJS.prototype as any).addSlide = mockAddSlide;

// Mocking methods on the slide object returned by addSlide
mockAddSlide.mockImplementation(() => ({
  addText: mockAddTextWithAnimation,
  // other slide methods if used by the tool for animations
}));

// Mock fs-extra
const mockPathExists = jest.fn();
const mockReadFile = jest.fn();
jest.mock('fs-extra', () => ({
  ...jest.requireActual('fs-extra'),
  pathExists: mockPathExists,
  readFile: mockReadFile,
}));

// Mock security utility (if animations tool uses it, good practice to have it ready)
jest.mock('../../../src/utils/security', () => ({
    validateFilePath: jest.fn().mockReturnValue(true),
}));

describe('PowerPoint Animations Tool - Unit Tests (Library Path)', () => {
  const mockDocumentPath = 'secure/mock/animated_document.pptx';
  const mockNewDocumentPath = 'secure/mock/new_animated_document.pptx';

  beforeEach(() => {
    jest.clearAllMocks();
    mockPathExists.mockResolvedValue(false); // Default to document not existing for add/configure tests
    mockReadFile.mockResolvedValue(Buffer.from('mock pptx content'));
    (validateFilePath as jest.Mock).mockReturnValue(true);
  });

  describe('animationsTool.handler with operation "add" (useComInterop: false)', () => {
    test('should add a new text object with animation to a new slide', async () => {
      const result = await animationsTool.handler({
        filePath: mockNewDocumentPath,
        operation: 'add',
        animationType: 'fadeIn',
        newObjectText: 'Animated Text!',
        newObjectOptions: { x: 1, y: 1, w: 5, h: 0.5 },
        duration: 2,
        effectParameters: { delay: 0.5, direction: 'fromBottom' },
        useComInterop: false,
      });

      expect(PptxGenJS).toHaveBeenCalledTimes(1);
      expect(mockAddSlide).toHaveBeenCalledTimes(1);
      expect(mockAddTextWithAnimation).toHaveBeenCalledWith(
        'Animated Text!',
        expect.objectContaining({
          x: 1, y: 1, w: 5, h: 0.5,
          animation: {
            type: 'fadeIn',
            duration: 2,
            delay: 0.5,
            direction: 'fromBottom',
          },
        })
      );
      expect(mockPptxGenJSWriteFile).toHaveBeenCalledWith({ fileName: expect.stringContaining(mockNewDocumentPath) });
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data.message).toContain('PptxGenJS: Added new text object with animation');
      }
    });

    test('should return error if animationType or newObjectText is missing for "add"', async () => {
      let result = await animationsTool.handler({
        filePath: mockNewDocumentPath,
        operation: 'add',
        // animationType missing
        newObjectText: 'Some Text',
        useComInterop: false,
      });
      expect(result.success).toBe(false);
      if (!result.success) expect(result.error.code).toBe('MISSING_PARAM_LIB');

      result = await animationsTool.handler({
        filePath: mockNewDocumentPath,
        operation: 'add',
        animationType: 'flyIn',
        // newObjectText missing
        useComInterop: false,
      });
      expect(result.success).toBe(false);
      if (!result.success) expect(result.error.code).toBe('MISSING_PARAM_LIB');
    });
  });

  describe('animationsTool.handler with operation "configure" (useComInterop: false)', () => {
    test('should add a new slide with specified transition', async () => {
      const result = await animationsTool.handler({
        filePath: mockNewDocumentPath,
        operation: 'configure',
        transitionType: 'fade',
        duration: 1.5,
        effectParameters: { direction: 'thruBlk' }, // Example PptxGenJS transition option
        useComInterop: false,
      });

      expect(PptxGenJS).toHaveBeenCalledTimes(1);
      expect(mockAddSlide).toHaveBeenCalledWith(
        expect.objectContaining({
          transition: {
            type: 'fade',
            duration: 1.5,
            direction: 'thruBlk',
          },
        })
      );
      expect(mockPptxGenJSWriteFile).toHaveBeenCalledWith({ fileName: expect.stringContaining(mockNewDocumentPath) });
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data.message).toContain('PptxGenJS: Added a new slide with transition');
      }
    });

    test('should return error if transitionType is missing for "configure"', async () => {
      const result = await animationsTool.handler({
        filePath: mockNewDocumentPath,
        operation: 'configure',
        // transitionType missing
        useComInterop: false,
      });
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('MISSING_PARAM_LIB');
      }
    });
  });

  describe('animationsTool.handler with "remove" or "list" (useComInterop: false)', () => {
    const operationsToTest: ('remove' | 'list')[] = ['remove', 'list'];
    operationsToTest.forEach(operation => {
      test(`should return "not supported" error for operation "${operation}"`, async () => {
        mockPathExists.mockResolvedValue(true); // Simulate existing file

        const result = await animationsTool.handler({
          filePath: mockDocumentPath,
          operation: operation,
          slideIndex: 1, // Relevant for COM, but error should occur due to lib path
          shapeIndex: 1, // Relevant for COM
          useComInterop: false,
        });

        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.message).toMatch(/not supported/i);
          expect(result.error.code).toBe('POWERPOINT_LIB_UNSUPPORTED');
        }
      });
    });
  });

  test('should return validation error for invalid filePath', async () => {
    (validateFilePath as jest.Mock).mockReturnValue(false);
    const result = await animationsTool.handler({
        filePath: 'invalid/../path.pptx',
        operation: 'add',
        animationType: 'fadeIn',
        newObjectText: 'test',
        useComInterop: false,
    });
    expect(validateFilePath).toHaveBeenCalledWith('invalid/../path.pptx');
    expect(result.success).toBe(false);
    if(!result.success){
        expect(result.error.code).toBe('VALIDATION_ERROR');
    }
  });
});