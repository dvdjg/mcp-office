import { powerpointPropertiesTool } from '../../../src/tools/powerpoint/properties.tool'; // Named import from array
import PptxGenJS from 'pptxgenjs';
import fs from 'fs-extra';
import { ToolRequestParams } from '../../../src/types/common.types';
import * as path from 'path';

// The tool is an array, get the first element which is the McpResource
const propertiesToolHandler = powerpointPropertiesTool[0].handler;

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

describe('PowerPoint Properties Tool Unit Tests (Library Path)', () => {
  let mockPptxInstance: any;

  beforeEach(() => {
    jest.clearAllMocks();

    mockPptxInstance = {
      addSlide: jest.fn(), // Called when properties/layout are modified
      writeFile: jest.fn().mockResolvedValue(undefined),
      // Properties that can be set on the PptxGenJS instance
      title: '',
      author: '',
      subject: '',
      company: '',
      revision: '',
      layout: '', // For slide size/layout configuration
    };
    (PptxGenJS as unknown as jest.Mock).mockImplementation(() => mockPptxInstance);

    (fs.pathExists as jest.Mock).mockResolvedValue(false); // Default: file does not exist
    (fs.readFile as unknown as jest.Mock).mockResolvedValue(Buffer.from('fake pptx content for resource saving'));
  });

  // --- Set Property Tests (Library Path) ---
  describe('propertiesTool: set operation (Library Path)', () => {
    test('should set "title" property on a new presentation', async () => {
      const params: ToolRequestParams = {
        filePath: 'new_props_title.pptx',
        operation: 'set',
        propertyName: 'title',
        propertyValue: 'My Test Presentation',
        useComInterop: false,
      };
      const result = await propertiesToolHandler(params);

      expect(result.success).toBe(true);
      expect(PptxGenJS).toHaveBeenCalledTimes(1);
      expect(mockPptxInstance.title).toBe('My Test Presentation');
      expect(mockPptxInstance.addSlide).toHaveBeenCalledTimes(1); // Default slide added
      expect(mockPptxInstance.writeFile).toHaveBeenCalledWith({ fileName: path.resolve(params.filePath) });
    });

    test('should set "author" property on a new presentation', async () => {
        const params: ToolRequestParams = {
          filePath: 'new_props_author.pptx',
          operation: 'set',
          propertyName: 'author',
          propertyValue: 'Roo The AI',
          useComInterop: false,
        };
        const result = await propertiesToolHandler(params);
  
        expect(result.success).toBe(true);
        expect(mockPptxInstance.author).toBe('Roo The AI');
        expect(mockPptxInstance.writeFile).toHaveBeenCalled();
      });

    test('should return error for unsupported propertyName for "set"', async () => {
      const params: ToolRequestParams = {
        filePath: 'new_props_unsupported.pptx',
        operation: 'set',
        propertyName: 'customNonStandardProp',
        propertyValue: 'Some Value',
        useComInterop: false,
      };
      const result = await propertiesToolHandler(params);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('INVALID_PARAM_LIB');
        expect(result.error.message).toMatch(/not directly supported/i);
      }
      expect(mockPptxInstance.writeFile).not.toHaveBeenCalled();
    });

    test('should require propertyName and propertyValue for "set"', async () => {
        const paramsNoName: ToolRequestParams = {
            filePath: 'props_set_noname.pptx',
            operation: 'set',
            propertyValue: 'A Value',
            useComInterop: false,
        };
        let res = await propertiesToolHandler(paramsNoName);
        expect(res.success).toBe(false);
        if(!res.success) expect(res.error.code).toBe('VALIDATION_ERROR_LIB');

        const paramsNoValue: ToolRequestParams = {
            filePath: 'props_set_noval.pptx',
            operation: 'set',
            propertyName: 'title',
            useComInterop: false,
        };
        res = await propertiesToolHandler(paramsNoValue);
        expect(res.success).toBe(false);
        if(!res.success) expect(res.error.code).toBe('VALIDATION_ERROR_LIB');
    });
  });

  // --- Configure Layout/Size Tests (Library Path) ---
  describe('propertiesTool: configure operation (Library Path)', () => {
    test('should configure slide layout to "16:9"', async () => {
      const params: ToolRequestParams = {
        filePath: 'new_layout_16x9.pptx',
        operation: 'configure',
        size: '16:9',
        useComInterop: false,
      };
      const result = await propertiesToolHandler(params);

      expect(result.success).toBe(true);
      expect(PptxGenJS).toHaveBeenCalledTimes(1);
      expect(mockPptxInstance.layout).toBe('LAYOUT_16x9'); // As mapped in the tool
      expect(mockPptxInstance.addSlide).toHaveBeenCalledTimes(1); // Default slide added
      expect(mockPptxInstance.writeFile).toHaveBeenCalledWith({ fileName: path.resolve(params.filePath) });
    });

    test('should configure slide layout to "LAYOUT_WIDE"', async () => {
        const params: ToolRequestParams = {
          filePath: 'new_layout_wide.pptx',
          operation: 'configure',
          size: 'LAYOUT_WIDE',
          useComInterop: false,
        };
        const result = await propertiesToolHandler(params);
  
        expect(result.success).toBe(true);
        expect(mockPptxInstance.layout).toBe('LAYOUT_WIDE');
        expect(mockPptxInstance.writeFile).toHaveBeenCalled();
      });


    test('should return error for unsupported layout size', async () => {
      const params: ToolRequestParams = {
        filePath: 'new_layout_unsupported.pptx',
        operation: 'configure',
        size: '21:9' as any, // Cast to any to bypass enum check for testing invalid value
        useComInterop: false,
      };
      const result = await propertiesToolHandler(params);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('INVALID_PARAM_LIB');
        expect(result.error.message).toMatch(/Unsupported layout/i);
      }
      expect(mockPptxInstance.writeFile).not.toHaveBeenCalled();
    });
  });

  // --- Add Property (Treated as Set) Tests (Library Path) ---
  describe('propertiesTool: add operation (Library Path)', () => {
    test('should "add" (set) "company" property', async () => {
      const params: ToolRequestParams = {
        filePath: 'new_props_add_company.pptx',
        operation: 'add',
        propertyName: 'company',
        propertyValue: 'Roo Inc.',
        useComInterop: false,
      };
      const result = await propertiesToolHandler(params);

      expect(result.success).toBe(true);
      expect(mockPptxInstance.company).toBe('Roo Inc.');
      expect(mockPptxInstance.writeFile).toHaveBeenCalled();
    });

    test('should return error for "add" with unsupported custom propertyName', async () => {
        const params: ToolRequestParams = {
          filePath: 'new_props_add_custom.pptx',
          operation: 'add',
          propertyName: 'MyCustomProp',
          propertyValue: 'MyCustomValue',
          useComInterop: false,
        };
        const result = await propertiesToolHandler(params);
  
        expect(result.success).toBe(false);
        if (!result.success) {
          expect(result.error.code).toBe('UNSUPPORTED_OPERATION_LIB'); // Or INVALID_PARAM_LIB depending on tool's exact message
          expect(result.error.message).toMatch(/Adding custom property .* is not supported/i);
        }
      });
  });

  // --- Get Property Tests (Library Path) ---
  describe('propertiesTool: get operation (Library Path)', () => {
    test('should report "not supported" for get operation', async () => {
      const params: ToolRequestParams = {
        filePath: 'existing_props.pptx',
        operation: 'get',
        propertyName: 'title',
        useComInterop: false,
      };
      (fs.pathExists as jest.Mock).mockResolvedValue(true); // File "exists"
      const result = await propertiesToolHandler(params);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('UNSUPPORTED_OPERATION_LIB');
        expect(result.error.message).toMatch(/'get' operation is not supported/i);
      }
    });
  });
});