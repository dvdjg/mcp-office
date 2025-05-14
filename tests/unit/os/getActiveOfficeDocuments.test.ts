import getActiveOfficeDocumentsTool from '../../../src/tools/os/getActiveOfficeDocuments.tool';
import { exec } from 'child_process';

// Mock child_process.exec
jest.mock('child_process', () => ({
  ...jest.requireActual('child_process'), // Import and retain default behavior
  exec: jest.fn(),
}));

const mockedExec = exec as jest.MockedFunction<typeof exec>;

describe('os/getActiveOfficeDocuments Tool', () => {
  let originalPlatform: NodeJS.Platform;

  beforeEach(() => {
    // Reset mocks before each test
    mockedExec.mockReset();
    originalPlatform = process.platform;
  });

  afterEach(() => {
    // Restore original platform
    Object.defineProperty(process, 'platform', {
      value: originalPlatform,
      writable: true,
    });
  });

  it('should return an empty array if not on Windows', async () => {
    Object.defineProperty(process, 'platform', {
      value: 'darwin', // Simulate macOS
      writable: true,
    });
    const result = await getActiveOfficeDocumentsTool.handler({});
    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data.documents).toEqual([]);
    }
  });

  it('should return an empty array if no Office apps are running (Windows)', async () => {
    Object.defineProperty(process, 'platform', {
      value: 'win32',
      writable: true,
    });
    mockedExec.mockImplementation((command, callback: any) => {
      callback(null, '', ''); // Simulate no output (no apps found)
      return {} as any; // Return a dummy child process object
    });

    const result = await getActiveOfficeDocumentsTool.handler({});
    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data.documents).toEqual([]);
    }
    expect(mockedExec).toHaveBeenCalledTimes(3); // Word, Excel, PowerPoint
  });

  it('should correctly parse Word documents (Windows)', async () => {
    Object.defineProperty(process, 'platform', {
      value: 'win32',
      writable: true,
    });
    mockedExec.mockImplementation((command, callback: any) => {
      if (command.includes('Word.Application')) {
        callback(null, 'C:\\Users\\Test\\Documents\\Doc1.docx;D:\\Work\\Reports\\Report Q1.doc', '');
      } else {
        callback(null, '', '');
      }
      return {} as any;
    });

    const result = await getActiveOfficeDocumentsTool.handler({});
    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data.documents).toEqual([
        { filePath: 'C:\\Users\\Test\\Documents\\Doc1.docx', applicationType: 'Word' },
        { filePath: 'D:\\Work\\Reports\\Report Q1.doc', applicationType: 'Word' },
      ]);
    }
  });

  it('should correctly parse Excel documents (Windows)', async () => {
    Object.defineProperty(process, 'platform', {
      value: 'win32',
      writable: true,
    });
    mockedExec.mockImplementation((command, callback: any) => {
      if (command.includes('Excel.Application')) {
        callback(null, 'C:\\Spreadsheets\\Data.xlsx;C:\\Temp\\Book1.xlsb', '');
      } else {
        callback(null, '', '');
      }
      return {} as any;
    });

    const result = await getActiveOfficeDocumentsTool.handler({});
    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data.documents).toEqual([
        { filePath: 'C:\\Spreadsheets\\Data.xlsx', applicationType: 'Excel' },
        { filePath: 'C:\\Temp\\Book1.xlsb', applicationType: 'Excel' },
      ]);
    }
  });

  it('should correctly parse PowerPoint documents (Windows)', async () => {
    Object.defineProperty(process, 'platform', {
      value: 'win32',
      writable: true,
    });
    mockedExec.mockImplementation((command, callback: any) => {
      if (command.includes('PowerPoint.Application')) {
        callback(null, 'C:\\Presentations\\Meeting.pptx;C:\\Archive\\OldPres.ppt', '');
      } else {
        callback(null, '', '');
      }
      return {} as any;
    });

    const result = await getActiveOfficeDocumentsTool.handler({});
    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data.documents).toEqual([
        { filePath: 'C:\\Presentations\\Meeting.pptx', applicationType: 'PowerPoint' },
        { filePath: 'C:\\Archive\\OldPres.ppt', applicationType: 'PowerPoint' },
      ]);
    }
  });

  it('should handle mixed open documents (Windows)', async () => {
    Object.defineProperty(process, 'platform', {
      value: 'win32',
      writable: true,
    });
    mockedExec.mockImplementation((command, callback: any) => {
      if (command.includes('Word.Application')) {
        callback(null, 'C:\\Doc.docx', '');
      } else if (command.includes('Excel.Application')) {
        callback(null, 'C:\\Sheet.xlsx', '');
      } else if (command.includes('PowerPoint.Application')) {
        callback(null, 'C:\\Pres.pptx', '');
      } else {
        callback(null, '', '');
      }
      return {} as any;
    });

    const result = await getActiveOfficeDocumentsTool.handler({});
    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data.documents).toEqual(
        expect.arrayContaining([
          { filePath: 'C:\\Doc.docx', applicationType: 'Word' },
          { filePath: 'C:\\Sheet.xlsx', applicationType: 'Excel' },
          { filePath: 'C:\\Pres.pptx', applicationType: 'PowerPoint' },
        ])
      );
      expect(result.data.documents.length).toBe(3);
    }
  });

  it('should handle PowerShell script errors gracefully (Windows)', async () => {
    Object.defineProperty(process, 'platform', {
      value: 'win32',
      writable: true,
    });
    mockedExec.mockImplementation((command, callback: any) => {
      callback(new Error('PowerShell failed'), '', 'Some PS error');
      return {} as any;
    });

    const result = await getActiveOfficeDocumentsTool.handler({});
    expect(result.success).toBe(true); // The tool itself doesn't fail, just returns no docs
    if (result.success) {
      expect(result.data.documents).toEqual([]);
    }
  });

  it('should filter out invalid or empty paths (Windows)', async () => {
    Object.defineProperty(process, 'platform', {
      value: 'win32',
      writable: true,
    });
    mockedExec.mockImplementation((command, callback: any) => {
      if (command.includes('Word.Application')) {
        // Includes an empty path, a path without extension, a path with wrong extension, and a relative path
        callback(null, 'C:\\Valid\\Doc1.docx;;  ;C:\\NoExtDoc;C:\\WrongExt.txt;Relative\\Path.docx', '');
      } else {
        callback(null, '', '');
      }
      return {} as any;
    });

    const result = await getActiveOfficeDocumentsTool.handler({});
    expect(result.success).toBe(true);
    if (result.success) {
      expect(result.data.documents).toEqual([
        { filePath: 'C:\\Valid\\Doc1.docx', applicationType: 'Word' },
      ]);
    }
  });

  it('should handle COM errors for individual applications gracefully (Windows)', async () => {
    Object.defineProperty(process, 'platform', {
        value: 'win32',
        writable: true,
    });
    mockedExec.mockImplementation((command, callback: any) => {
        if (command.includes('Word.Application')) {
            // Simulate Word running with a document
            callback(null, 'C:\\MyWordFile.docx', '');
        } else if (command.includes('Excel.Application')) {
            // Simulate Excel COM object not found or error during query
            callback(null, '', 'Error: Excel not available'); // PowerShell script would output ""
        } else if (command.includes('PowerPoint.Application')) {
            // Simulate PowerPoint running with a document
            callback(null, 'D:\\MyPresentation.pptx', '');
        }
        return {} as any;
    });

    const result = await getActiveOfficeDocumentsTool.handler({});
    expect(result.success).toBe(true);
    if (result.success) {
        expect(result.data.documents).toEqual(
            expect.arrayContaining([
                { filePath: 'C:\\MyWordFile.docx', applicationType: 'Word' },
                { filePath: 'D:\\MyPresentation.pptx', applicationType: 'PowerPoint' },
            ])
        );
        expect(result.data.documents.length).toBe(2); // Only Word and PowerPoint docs
    }
  });

});