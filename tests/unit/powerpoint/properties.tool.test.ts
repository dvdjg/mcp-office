import PptxGenJS from 'pptxgenjs';
import fs from 'fs-extra';
import { powerpointPropertiesTool } from '../../../src/tools/powerpoint/properties.tool'; // Named import
import { validateFilePath } from '../../../src/utils/security';

// The tool is an array, get the first element which is the McpResource object
const propertiesToolHandler = powerpointPropertiesTool[0].handler;

// Mock PptxGenJS
jest.mock('pptxgenjs');
const mockPptxGenJSWriteFile = jest.fn().mockResolvedValue(undefined);
const mockAddSlide = jest.fn(); // For when a default slide is added

// Assign mocks to the PptxGenJS prototype
// We need to mock the instance that PptxGenJS constructor returns
let mockPptxGenJSInstance: any;

// Cast PptxGenJS to jest.Mock for constructor mocking
const PptxGenJSConstructorMock = PptxGenJS as jest.Mock;

PptxGenJSConstructorMock.mockImplementation(() => {
    mockPptxGenJSInstance = {
        writeFile: mockPptxGenJSWriteFile,
        addSlide: mockAddSlide,
        title: undefined,
        author: undefined,
        subject: undefined,
        company: undefined,
        revision: undefined,
        layout: undefined,
    };
    return mockPptxGenJSInstance;
});


// Mock fs-extra
const mockPathExists = jest.fn();
const mockReadFile = jest.fn();
jest.mock('fs-extra', () => ({
  ...jest.requireActual('fs-extra'),
  pathExists: mockPathExists,
  readFile: mockReadFile,
}));

// Mock security utility
jest.mock('../../../src/utils/security', () => ({
    validateFilePath: jest.fn().mockReturnValue(true),
}));

describe('PowerPoint Properties Tool - Unit Tests (Library Path)', () => {
  const mockDocumentPath = 'secure/mock/properties_document.pptx';

  beforeEach(() => {
    jest.clearAllMocks();
    mockPathExists.mockResolvedValue(false); // Default to document not existing
    mockReadFile.mockResolvedValue(Buffer.from('mock pptx content'));
    (validateFilePath as jest.Mock).mockReturnValue(true);
    // Re-initialize PptxGenJS mock for each test to reset instance properties
    PptxGenJSConstructorMock.mockClear(); // Use the constructor mock for clear
    mockPptxGenJSWriteFile.mockClear();
    mockAddSlide.mockClear();
  });

  describe('propertiesToolHandler with operation "set" (useComInterop: false)', () => {
    test('should set supported built-in properties on PptxGenJS instance', async () => {
      await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'set', propertyName: 'title', propertyValue: 'Test Title', useComInterop: false,
      });
      expect(mockPptxGenJSInstance.title).toBe('Test Title'); // Access through the instance captured

      await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'set', propertyName: 'author', propertyValue: 'Test Author', useComInterop: false,
      });
      expect(mockPptxGenJSInstance.author).toBe('Test Author');
      
      await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'set', propertyName: 'subject', propertyValue: 'Test Subject', useComInterop: false,
      });
      expect(mockPptxGenJSInstance.subject).toBe('Test Subject');

      await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'set', propertyName: 'company', propertyValue: 'Test Company', useComInterop: false,
      });
      expect(mockPptxGenJSInstance.company).toBe('Test Company');
      
      await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'set', propertyName: 'revision', propertyValue: '5', useComInterop: false,
      });
      expect(mockPptxGenJSInstance.revision).toBe('5');

      // All these 'set' operations should trigger a file write with a default slide
      expect(mockAddSlide).toHaveBeenCalled();
      expect(mockPptxGenJSWriteFile).toHaveBeenCalledWith({ fileName: expect.stringContaining(mockDocumentPath) });
    });

    test('should return error for unsupported propertyName for "set"', async () => {
      const result = await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'set', propertyName: 'unsupportedProp', propertyValue: 'some value', useComInterop: false,
      });
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('INVALID_PARAM_LIB');
      }
    });

    test('should return error if propertyName or propertyValue is missing for "set"', async () => {
      let result = await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'set', /* propertyName missing */ propertyValue: 'Test Title', useComInterop: false,
      });
      expect(result.success).toBe(false);
      if (!result.success) expect(result.error.code).toBe('VALIDATION_ERROR_LIB'); // or Zod validation error

      result = await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'set', propertyName: 'title', /* propertyValue missing */ useComInterop: false,
      });
      expect(result.success).toBe(false);
      if (!result.success) expect(result.error.code).toBe('VALIDATION_ERROR_LIB');
    });
  });

  describe('propertiesToolHandler with operation "configure" (useComInterop: false)', () => {
    test('should set slide layout on PptxGenJS instance', async () => {
      await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'configure', size: '16:9', useComInterop: false,
      });
      expect(mockPptxGenJSInstance.layout).toBe('LAYOUT_16x9');
      expect(mockAddSlide).toHaveBeenCalledTimes(1); // Configure adds a slide
      expect(mockPptxGenJSWriteFile).toHaveBeenCalledWith({ fileName: expect.stringContaining(mockDocumentPath) });
    });

    test('should return error for unsupported layout for "configure"', async () => {
      const result = await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'configure', size: 'unsupportedLayout', useComInterop: false,
      });
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('INVALID_PARAM_LIB');
      }
    });
  });

  describe('propertiesToolHandler with operation "add" (useComInterop: false)', () => {
    // 'add' is treated like 'set' for standard properties, and unsupported for custom.
    test('should set standard properties like "set" for "add"', async () => {
      await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'add', propertyName: 'title', propertyValue: 'Added Title', useComInterop: false,
      });
      expect(mockPptxGenJSInstance.title).toBe('Added Title');
      expect(mockAddSlide).toHaveBeenCalled();
      expect(mockPptxGenJSWriteFile).toHaveBeenCalled();
    });

    test('should return "unsupported operation" error for custom property "add"', async () => {
      const result = await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'add', propertyName: 'customProp', propertyValue: 'customValue', useComInterop: false,
      });
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('UNSUPPORTED_OPERATION_LIB');
      }
    });
  });

  describe('propertiesToolHandler with operation "get" (useComInterop: false)', () => {
    test('should return "unsupported operation" error for "get"', async () => {
      const result = await propertiesToolHandler({
        filePath: mockDocumentPath, operation: 'get', propertyName: 'title', useComInterop: false,
      });
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('UNSUPPORTED_OPERATION_LIB');
      }
    });
  });

  test('should return validation error for invalid filePath', async () => {
    (validateFilePath as jest.Mock).mockReturnValue(false);
    const result = await propertiesToolHandler({
        filePath: 'invalid/../path.pptx',
        operation: 'set',
        propertyName: 'title',
        propertyValue: 'test',
        useComInterop: false,
    });
    expect(validateFilePath).toHaveBeenCalledWith('invalid/../path.pptx');
    expect(result.success).toBe(false);
    if(!result.success){
        expect(result.error.code).toBe('VALIDATION_ERROR'); // Zod error from schema parsing
    }
  });
});