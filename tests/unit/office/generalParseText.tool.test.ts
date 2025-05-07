import { jest, describe, expect, test, beforeEach } from '@jest/globals';
import { generalParseTextTool } from '../../../src/tools/office/generalParseText.tool';
import officeparser from 'officeparser';
import fs from 'fs-extra';
import { ZodError } from 'zod';

// Mock dependencies
// Mock dependencies
jest.mock('officeparser');
jest.mock('fs-extra');

// Use jest.fn() directly for mocks
const mockParseOfficeAsync = jest.fn();
const mockPathExists = jest.fn();

// Assign the mocks to the original functions (using unknown to bypass type checking here)
(officeparser.parseOfficeAsync as unknown) = mockParseOfficeAsync;
(fs.pathExists as unknown) = mockPathExists;


describe('generalParseTextTool', () => {
  const { handler, schema } = generalParseTextTool;

  beforeEach(() => {
    // Reset mocks before each test
    mockParseOfficeAsync.mockReset();
    mockPathExists.mockReset();
  });

  describe('Input Validation', () => {
    it('should return a validation error if filePath is missing', async () => {
      const result = await handler({} as any, undefined); // Pass undefined for context
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('VALIDATION_ERROR');
        expect(result.error.message).toContain('Input validation failed');
        // Check for ZodError details
        expect(Array.isArray(result.error.details)).toBe(true);
        const zodErrorDetails = result.error.details as ZodError['errors'];
        expect(zodErrorDetails.some(e => e.path.includes('filePath'))).toBe(true);
      }
    });

    it('should return a validation error if filePath is an empty string', async () => {
      const result = await handler({ filePath: '' } as any, undefined);
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('VALIDATION_ERROR');
      }
    });

    it('should pass validation with a valid filePath', () => {
      const validInput = { filePath: 'valid/path/to/file.docx' };
      expect(() => schema?.parse(validInput)).not.toThrow();
    });
  });

  describe('File Existence Check', () => {
    it('should return FILE_NOT_FOUND error if file does not exist', async () => {
      mockPathExists.mockResolvedValue(false);
      const input = { filePath: 'nonexistent/file.docx' };
      const result = await handler(input, undefined);

      expect(mockPathExists).toHaveBeenCalledWith(input.filePath);
      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('FILE_NOT_FOUND');
        expect(result.error.message).toContain('Error: File not found at path');
      }
    });
  });

  describe('Officeparser Interaction and Output', () => {
    it('should successfully parse a docx file and return extracted text', async () => {
      const filePath = 'valid/dummy.docx';
      const mockText = 'This is a test docx content.';
      mockPathExists.mockResolvedValue(true);
      mockParseOfficeAsync.mockResolvedValue(mockText);

      const result = await handler({ filePath }, undefined);

      expect(mockPathExists).toHaveBeenCalledWith(filePath);
      expect(mockParseOfficeAsync).toHaveBeenCalledWith(filePath);
      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data.extractedText).toBe(mockText);
        expect(result.data.message).toBe('Text extracted successfully.');
        expect(result.data.detectedFileType).toBeUndefined(); // No metadata in this mock
      }
    });

    it('should successfully parse an xlsx file and return extracted text and metadata if provided', async () => {
      const filePath = 'valid/dummy.xlsx';
      const mockText = 'Sheet1 data...';
      const mockMetadata = { type: 'xlsx', otherMeta: 'value' };
      mockPathExists.mockResolvedValue(true);
      mockParseOfficeAsync.mockResolvedValue({ content: mockText, meta: mockMetadata });

      const result = await handler({ filePath }, undefined);

      expect(result.success).toBe(true);
      if (result.success) {
        expect(result.data.extractedText).toBe(mockText);
        expect(result.data.detectedFileType).toBe('xlsx');
        expect(result.data.message).toBe('Text extracted successfully.');
      }
    });
    
    it('should handle officeparser returning only string content', async () => {
        const filePath = 'valid/dummy.pptx';
        const mockText = 'PowerPoint presentation text.';
        mockPathExists.mockResolvedValue(true);
        mockParseOfficeAsync.mockResolvedValue(mockText); // officeparser returns string directly

        const result = await handler({ filePath }, undefined);
        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data.extractedText).toBe(mockText);
            expect(result.data.message).toBe('Text extracted successfully.');
            expect(result.data.detectedFileType).toBeUndefined();
        }
    });

    it('should handle officeparser returning an object without meta.type', async () => {
        const filePath = 'valid/dummy.odt';
        const mockText = 'OpenOffice document text.';
        mockPathExists.mockResolvedValue(true);
        mockParseOfficeAsync.mockResolvedValue({ content: mockText, meta: { someOtherKey: 'value' } });

        const result = await handler({ filePath }, undefined);
        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data.extractedText).toBe(mockText);
            expect(result.data.message).toBe('Text extracted successfully.');
            expect(result.data.detectedFileType).toBeUndefined();
        }
    });
    
    it('should handle officeparser returning an object with null meta', async () => {
        const filePath = 'valid/dummy.ods';
        const mockText = 'OpenOffice spreadsheet text.';
        mockPathExists.mockResolvedValue(true);
        mockParseOfficeAsync.mockResolvedValue({ content: mockText, meta: null });

        const result = await handler({ filePath }, undefined);
        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data.extractedText).toBe(mockText);
            expect(result.data.message).toBe('Text extracted successfully.');
            expect(result.data.detectedFileType).toBeUndefined();
        }
    });

    it('should handle officeparser returning an unexpected object structure by stringifying it', async () => {
        const filePath = 'valid/dummy.unknown';
        const unexpectedResult = { data: "some data", info: "more info" };
        mockPathExists.mockResolvedValue(true);
        mockParseOfficeAsync.mockResolvedValue(unexpectedResult);

        const result = await handler({ filePath }, undefined);
        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data.extractedText).toBe(JSON.stringify(unexpectedResult));
            expect(result.data.message).toBe('Text extracted successfully.');
        }
    });
    
    it('should handle officeparser returning null or undefined by providing empty extractedText', async () => {
        const filePath = 'valid/dummy.empty';
        mockPathExists.mockResolvedValue(true);
        
        // Test with null
        mockParseOfficeAsync.mockResolvedValue(null);
        let result = await handler({ filePath }, undefined);
        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data.extractedText).toBe('');
            expect(result.data.message).toBe('Text extracted successfully.');
        }

        // Test with undefined
        mockParseOfficeAsync.mockResolvedValue(undefined);
        result = await handler({ filePath }, undefined);
        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data.extractedText).toBe('');
            expect(result.data.message).toBe('Text extracted successfully.');
        }
    });

    it('should return OFFICEPARSER_ERROR if officeparser.parseOfficeAsync throws an error', async () => {
      const filePath = 'valid/corrupted.docx';
      const errorMessage = 'Failed to parse corrupted file';
      mockPathExists.mockResolvedValue(true);
      mockParseOfficeAsync.mockRejectedValue(new Error(errorMessage));

      const result = await handler({ filePath }, undefined);

      expect(result.success).toBe(false);
      if (!result.success) {
        expect(result.error.code).toBe('OFFICEPARSER_ERROR'); // Or a more generic code if handleToolError defaults
        expect(result.error.message).toBe(errorMessage);
      }
    });

    it('should handle officeparser throwing a non-Error object', async () => {
        const filePath = 'valid/weird-error.docx';
        const errorObject = { code: 123, reason: "weird reason" };
        mockPathExists.mockResolvedValue(true);
        mockParseOfficeAsync.mockRejectedValue(errorObject);

        const result = await handler({ filePath }, undefined);
        expect(result.success).toBe(false);
        if(!result.success) {
            expect(result.error.code).toBe('OFFICEPARSER_ERROR');
            expect(result.error.message).toBe('An unexpected error occurred.'); // Default message for non-Error
            expect(result.error.details).toEqual(errorObject);
        }
    });
  });
});