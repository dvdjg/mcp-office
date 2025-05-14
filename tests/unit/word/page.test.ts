import { wordPageTool } from '../../../src/tools/word/page.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { handleToolError } from '../../../src/utils/errorHandler';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';

// Mock dependencies
jest.mock('../../../src/utils/officeInterop', () => ({
  getOfficeApplication: jest.fn(),
  releaseObject: jest.fn(),
}));

jest.mock('../../../src/utils/errorHandler', () => {
  class MockErrorResponse {
    public success: false = false;
    public error: { code: string; message: string; details: any | null };

    constructor(message: string, code: string, details: any | null = null) {
      this.error = { code, message, details };
    }
  }
  return {
    handleToolError: jest.fn((caughtError, code, details) => {
      const message = caughtError.message || 'Mocked error from handleToolError';
      // Return an instance of a class that has the 'success' and 'error' properties
      return new MockErrorResponse(message, code, details);
    }),
    createErrorResponse: jest.fn((code, message, details) => {
      return new MockErrorResponse(message, code, details);
    }),
  };
});

jest.mock('../../../src/utils/security', () => ({
  validateFilePath: jest.fn(path => true), // Assume valid path for these tests
}));

jest.mock('../../../src/utils/logger', () => ({
  info: jest.fn(),
  warn: jest.fn(),
  error: jest.fn(),
  debug: jest.fn(),
}));

const mockWordApp = {
  Documents: {
    Open: jest.fn(),
  },
  InchesToPoints: jest.fn(val => val * 72),
  CentimetersToPoints: jest.fn(val => val * 28.3465),
  MillimetersToPoints: jest.fn(val => val * 2.83465),
};

const mockDoc = {
  ComputeStatistics: jest.fn(),
  Sections: {
    Item: jest.fn(),
    Count: 0,
  },
  PageSetup: {},
  Save: jest.fn(),
  Close: jest.fn(),
};

const mockOfficeAppInstance = {
  app: mockWordApp,
  release: jest.fn(),
};

describe('wordPageTool - getPageCount', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    (getOfficeApplication as jest.Mock).mockResolvedValue(mockOfficeAppInstance);
    (mockWordApp.Documents.Open as jest.Mock).mockResolvedValue(mockDoc);
  });

  it('should successfully get the page count', async () => {
    const mockFilePath = 'C:/test/document.docx';
    const expectedPageCount = 5;
    (mockDoc.ComputeStatistics as jest.Mock).mockResolvedValue(expectedPageCount);

    const params = {
      operation: 'getPageCount',
      filePath: mockFilePath,
    };

    const result = await wordPageTool.handler(params as any);

    expect(getOfficeApplication).toHaveBeenCalledWith('Word.Application');
    expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(mockFilePath, false, true);
    expect(mockDoc.ComputeStatistics).toHaveBeenCalledWith(2); // wdStatisticPages
    expect(mockDoc.Close).toHaveBeenCalledWith(false);
    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data).toEqual({ pageCount: expectedPageCount });
    }
    expect(logger.info).toHaveBeenCalledWith(`Successfully retrieved page count (${expectedPageCount}) for ${mockFilePath}.`);
    expect(mockOfficeAppInstance.release).toHaveBeenCalled();
  });

  it('should return an error if COM object fails to get page count', async () => {
    const mockFilePath = 'C:/test/document.docx';
    const expectedErrorMessage = 'COM Error during ComputeStatistics';
    // Make ComputeStatistics reject with an object that looks like an ErrorResponse
    // that would have been produced by handleToolError(new Error(expectedErrorMessage), 'OFFICE_API_ERROR')
    const mockErrorResponse = {
      success: false,
      error: {
        code: 'OFFICE_API_ERROR',
        message: expectedErrorMessage,
        details: null,
      },
    };
    (mockDoc.ComputeStatistics as jest.Mock).mockRejectedValue(mockErrorResponse); // This is what's caught by getDocumentPageCount

    const params = {
      operation: 'getPageCount',
      filePath: mockFilePath,
    };

    const result = await wordPageTool.handler(params as any); // This calls the main handler

    // wordPageTool's main handler catches the error thrown by getDocumentPageCount.
    // getDocumentPageCount catches mockErrorResponse and then throws the result of handleToolError(mockErrorResponse, 'OFFICE_API_ERROR').
    // The mocked handleToolError (above) should return an object that satisfies the 'if' condition in the main handler.

    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('OFFICE_API_ERROR'); // This is the assertion failing
      expect(result.error?.message).toBe(expectedErrorMessage); // This should be the message from mockErrorResponse.error.message
    }
    
    // Logger in getDocumentPageCount is called with the originally caught error (mockErrorResponse)
    expect(logger.error).toHaveBeenCalledWith(
      `Error getting page count for ${mockFilePath}: ${expectedErrorMessage}`, // message from mockErrorResponse.error.message
      { error: mockErrorResponse } // the error object is mockErrorResponse
    );
    expect(mockDoc.Close).toHaveBeenCalledWith(false); // Ensure close is still attempted
    expect(mockOfficeAppInstance.release).toHaveBeenCalled();
  });

  it('should return a validation error for invalid operation parameters (e.g. extra param for getPageCount)', async () => {
    const params = {
      operation: 'getPageCount',
      filePath: 'C:/test.docx',
      sectionIndex: 1, // Invalid extra parameter for getPageCount
    };

    // The schema refinement should catch this.
    // We expect the handler's initial Zod parse to fail.
    // For this test, we'll simulate the ZodError that would be caught by the handler.
    // The actual Zod schema is tested implicitly by the tool's own parsing logic.
    // Here, we are testing how the handler reacts to a Zod-like error.

    const result = await wordPageTool.handler(params as any);
    // Based on the updated schema, this should now be caught by the refine function.
    // The error message comes from the refine block.
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('VALIDATION_ERROR');
      expect(result.error?.message).toContain("Input validation failed");
      // The exact message depends on Zod's formatting, but it should indicate the issue.
      // Example: "Input validation failed:  - Invalid input for the specified operation. For 'set', provide at least one config property and include pageWidth/pageHeight if size is wdPaperCustom. For 'modify', provide at least one config property. For 'getPageCount', only 'filePath' is allowed."
      // We can check for a substring
      expect(result.error?.message).toContain("For 'getPageCount', only 'filePath' is allowed.");
    }
  });

   it('should return a validation error if filePath is missing for getPageCount', async () => {
    const params = {
      operation: 'getPageCount',
      // filePath is missing
    };
    const result = await wordPageTool.handler(params as any);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error?.code).toBe('VALIDATION_ERROR');
      expect(result.error?.message).toContain("filePath - Required");
    }
  });
});