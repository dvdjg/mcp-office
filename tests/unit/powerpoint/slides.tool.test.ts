import slidesTool from '../../../src/tools/powerpoint/slides.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import PptxGenJS from 'pptxgenjs';
import officeParser from 'officeparser';
import fs from 'fs-extra';
import path from 'path';
import logger from '../../../src/utils/logger';

// Mock dependencies
jest.mock('../../../src/utils/officeInterop', () => ({
  getOfficeApplication: jest.fn(),
  releaseObject: jest.fn(),
}));

jest.mock('pptxgenjs');
jest.mock('officeparser', () => ({
  parseOfficeAsync: jest.fn(),
}));

jest.mock('fs-extra', () => ({
  readFile: jest.fn(),
  pathExists: jest.fn(),
}));

jest.mock('../../../src/tools/dynamic/resources.tool', () => ({
  saveResource: jest.fn().mockResolvedValue({ success: true, data: { message: "Resource saved" } }),
}));

jest.mock('../../../src/utils/logger', () => ({
  info: jest.fn(),
  warn: jest.fn(),
  error: jest.fn(),
  debug: jest.fn(),
}));

// Define top-level mock functions for COM methods
const mockPptOpen = jest.fn();
const mockPptAdd = jest.fn();

// Define the structure of the application mock (defined once)
const mockPptApplicationObject = {
  Presentations: {
    Open: mockPptOpen,
    Add: mockPptAdd,
  },
  Visible: false,
};

// Define the structure of the office instance mock (defined once)
const mockOfficeAppInstanceObject = {
  app: mockPptApplicationObject,
  release: jest.fn(),
};

// Define TypeScript types for the mock objects created in beforeEach
type MockSlidesState = {
  Count: number;
  Item: jest.Mock<any, any>;
  Add: jest.Mock<any, any>;
};

type MockPresentationState = {
  Slides: MockSlidesState;
  SlideMaster: any;
  SlideLayouts: any;
  Save: jest.Mock<any, any>;
  SaveAs: jest.Mock<any, any>;
  Close: jest.Mock<any, any>;
};

describe('slidesTool - getSlideCount', () => {
  let currentSlidesState: MockSlidesState;
  let currentPresentationState: MockPresentationState;

  beforeEach(() => {
    jest.clearAllMocks();

    currentSlidesState = {
      Count: 0,
      Item: jest.fn(() => ({ Delete: jest.fn(), Shapes: { Count: 0, Item: jest.fn() } })),
      Add: jest.fn(),
    };
    // Ensure Count is writable for tests that set it directly or redefine it
    Object.defineProperty(currentSlidesState, 'Count', {
        value: 0,
        writable: true,
        configurable: true,
    });

    currentPresentationState = {
      Slides: currentSlidesState,
      SlideMaster: { CustomLayouts: { Item: jest.fn(() => ({ Layout: {} })) } },
      SlideLayouts: { Item: jest.fn(() => ({})) },
      Save: jest.fn(),
      SaveAs: jest.fn(),
      Close: jest.fn(),
    };

    mockPptOpen.mockResolvedValue(currentPresentationState);
    mockPptAdd.mockResolvedValue(currentPresentationState);
    mockOfficeAppInstanceObject.release.mockReset(); // Reset release mock
    
    (getOfficeApplication as jest.Mock).mockResolvedValue(mockOfficeAppInstanceObject);
    (fs.pathExists as jest.Mock).mockResolvedValue(true);
  });

  // --- COM Interop Tests ---
  describe('COM Interop Path (useComInterop: true)', () => {
    it('should successfully get slide count using COM', async () => {
      const mockFilePath = 'C:/test/presentation.pptx';
      const expectedSlideCount = 7;
      currentSlidesState.Count = expectedSlideCount;

      const params = {
        filePath: mockFilePath,
        operation: 'getSlideCount',
        useComInterop: true,
      };

      const result = await slidesTool.handler(params as any);

      expect(getOfficeApplication).toHaveBeenCalledWith('PowerPoint.Application');
      expect(mockPptOpen).toHaveBeenCalledWith(mockFilePath);
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data).toEqual({ slideCount: expectedSlideCount });
      }
      expect(currentPresentationState.Close).toHaveBeenCalled();
      expect(mockOfficeAppInstanceObject.release).toHaveBeenCalled();
    });

    it('should return an error if COM Presentations.Open fails', async () => {
      const mockFilePath = 'C:/test/presentation.pptx';
      const comError = new Error('COM Error during Presentations.Open');
      mockPptOpen.mockRejectedValue(comError);

      const params = {
        filePath: mockFilePath,
        operation: 'getSlideCount',
        useComInterop: true,
      };

      const result = await slidesTool.handler(params as any);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error?.code).toBe('POWERPOINT_COM_ERROR');
        expect(result.error?.message).toBe(`COM Interop Error: ${comError.message}`);
      }
      expect(mockOfficeAppInstanceObject.release).toHaveBeenCalled();
    });

    it('should return an error if COM Slides.Count access fails after Open', async () => {
      const mockFilePath = 'C:/test/presentation.pptx';
      const countError = new Error('COM Error accessing Slides.Count');
      mockPptOpen.mockResolvedValue(currentPresentationState); // Open succeeds
      Object.defineProperty(currentSlidesState, 'Count', {
        get: jest.fn(() => { throw countError; }),
        configurable: true
      });

      const params = {
        filePath: mockFilePath,
        operation: 'getSlideCount',
        useComInterop: true,
      };
      const result = await slidesTool.handler(params as any);
      expect(result.success).toBe(false);
      if(!result.success){
        expect(result.error?.code).toBe('POWERPOINT_COM_ERROR');
        expect(result.error?.message).toBe(`COM Interop Error: ${countError.message}`);
      }
      expect(mockPptOpen).toHaveBeenCalledWith(mockFilePath);
      expect(mockOfficeAppInstanceObject.release).toHaveBeenCalled();
    });
  });

  // --- Library Path Tests ---
  describe('Library Path (useComInterop: false)', () => {
    it('should return an error indicating getSlideCount is not supported via library', async () => {
      const mockFilePath = 'C:/test/presentation.pptx';
      const params = {
        filePath: mockFilePath,
        operation: 'getSlideCount',
        useComInterop: false,
      };

      const result = await slidesTool.handler(params as any);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error?.code).toBe('POWERPOINT_LIB_UNSUPPORTED');
        expect(result.error?.message).toBe('getSlideCount is not accurately supported via officeparser/pptxgenjs. Please use COM Interop.');
      }
      expect(logger.warn).toHaveBeenCalledWith('Operation "getSlideCount" using the library path is not directly supported for accurate counts. Use COM Interop (useComInterop: true) for reliable slide count.');
    });
  });

  // --- Validation Tests ---
  it('should return a validation error if filePath is missing for getSlideCount', async () => {
    const params = {
      operation: 'getSlideCount',
      useComInterop: true,
    };

    let errorThrown = false;
    try {
      await slidesTool.handler(params as any);
    } catch (e: any) {
      errorThrown = true;
      expect(e.issues[0].message).toBe('Required');
      expect(e.issues[0].path).toEqual(['filePath']);
    }
    expect(errorThrown).toBe(true);
  });
});
