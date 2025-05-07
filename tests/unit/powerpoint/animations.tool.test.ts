import animationsTool from '../../../src/tools/powerpoint/animations.tool'; // Default import
import PptxGenJS from 'pptxgenjs';
import fs from 'fs-extra';
import { ToolRequestParams } from '../../../src/types/common.types';
import * as path from 'path';

// Mock PptxGenJS
jest.mock('pptxgenjs');

// Mock fs-extra
jest.mock('fs-extra', () => ({
  ...jest.requireActual('fs-extra'),
  readFile: jest.fn(),
  writeFile: jest.fn(),
  pathExists: jest.fn(),
  ensureDir: jest.fn().mockResolvedValue(undefined),
}));

// Mock the dynamic resource saver
jest.mock('../../../src/tools/dynamic/resources.tool', () => ({
    saveResource: jest.fn().mockResolvedValue({ success: true, message: 'Resource saved' }),
}));

describe('PowerPoint Animations Tool Unit Tests (Library Path)', () => {
  let mockPptxInstance: any;
  let mockSlideInstance: any;

  beforeEach(() => {
    jest.clearAllMocks();

    mockSlideInstance = {
      addText: jest.fn(),
      // ... other slide methods
    };

    mockPptxInstance = {
      addSlide: jest.fn().mockReturnValue(mockSlideInstance),
      writeFile: jest.fn().mockResolvedValue(undefined),
      // ... other presentation methods
    };
    (PptxGenJS as unknown as jest.Mock).mockImplementation(() => mockPptxInstance);

    (fs.pathExists as jest.Mock).mockResolvedValue(false); // Default: file does not exist
    (fs.readFile as unknown as jest.Mock).mockResolvedValue(Buffer.from('fake pptx content for resource saving'));
  });

  // --- Add Animation Tests (Library Path) ---
  describe('animationsTool: add operation (Library Path)', () => {
    test('should add animation to a new text object on a new slide', async () => {
      const params: ToolRequestParams = {
        filePath: 'new_animated_text.pptx',
        operation: 'add',
        animationType: 'fadeIn',
        newObjectText: 'Animated Text!',
        newObjectOptions: { x: 1, y: 1, w: 5, h: 0.5 },
        duration: 2,
        effectParameters: { delay: 0.5, direction: 'fromBottom' },
        useComInterop: false,
      };
      const result = await animationsTool.handler(params);

      expect(result.success).toBe(true);
      expect(PptxGenJS).toHaveBeenCalledTimes(1);
      expect(mockPptxInstance.addSlide).toHaveBeenCalledTimes(1);
      expect(mockSlideInstance.addText).toHaveBeenCalledWith(
        params.newObjectText,
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
      expect(mockPptxInstance.writeFile).toHaveBeenCalledWith({ fileName: path.resolve(params.filePath) });
    });

    test('should require animationType and newObjectText for add operation', async () => {
      const paramsMissingType: ToolRequestParams = {
        filePath: 'anim_missing_type.pptx',
        operation: 'add',
        newObjectText: 'Some text',
        useComInterop: false,
      };
      let result = await animationsTool.handler(paramsMissingType);
      expect(result.success).toBe(false);
      if(!result.success) expect(result.error.code).toBe('MISSING_PARAM_LIB');

      const paramsMissingText: ToolRequestParams = {
        filePath: 'anim_missing_text.pptx',
        operation: 'add',
        animationType: 'flyIn',
        useComInterop: false,
      };
      result = await animationsTool.handler(paramsMissingText);
      expect(result.success).toBe(false);
      if(!result.success) expect(result.error.code).toBe('MISSING_PARAM_LIB');
    });
  });

  // --- Configure Transition Tests (Library Path) ---
  describe('animationsTool: configure operation (Library Path)', () => {
    test('should configure transition for a new slide', async () => {
      const params: ToolRequestParams = {
        filePath: 'new_slide_transition.pptx',
        operation: 'configure',
        transitionType: 'fade',
        duration: 1.5,
        effectParameters: { direction: 'smoothly' }, // Example PptxGenJS might take specific ones
        useComInterop: false,
      };
      const result = await animationsTool.handler(params);

      expect(result.success).toBe(true);
      expect(PptxGenJS).toHaveBeenCalledTimes(1);
      expect(mockPptxInstance.addSlide).toHaveBeenCalledWith(expect.objectContaining({
        transition: {
          type: 'fade',
          duration: 1.5,
          direction: 'smoothly',
        },
      }));
      expect(mockPptxInstance.writeFile).toHaveBeenCalledWith({ fileName: path.resolve(params.filePath) });
    });

    test('should require transitionType for configure operation', async () => {
        const params: ToolRequestParams = {
          filePath: 'trans_missing_type.pptx',
          operation: 'configure',
          useComInterop: false,
        };
        const result = await animationsTool.handler(params);
        expect(result.success).toBe(false);
        if(!result.success) expect(result.error.code).toBe('MISSING_PARAM_LIB');
      });
  });

  // --- Remove and List Operations (Library Path) ---
  ['remove', 'list'].forEach(operation => {
    describe(`animationsTool: ${operation} operation (Library Path)`, () => {
      test(`should report "not supported" for '${operation}' operation`, async () => {
        const params: ToolRequestParams = {
          filePath: 'existing_file_anim.pptx',
          operation: operation as 'remove' | 'list',
          slideIndex: 1, // These might be provided but are irrelevant for lib path failure
          shapeIndex: 1,
          useComInterop: false,
        };
        (fs.pathExists as jest.Mock).mockResolvedValue(true); // File "exists"
        const result = await animationsTool.handler(params);

        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('POWERPOINT_LIB_UNSUPPORTED');
          expect(result.error.message).toMatch(/not supported/i);
        }
      });
    });
  });
});