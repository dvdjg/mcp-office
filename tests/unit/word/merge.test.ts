import { mergeDocuments } from '../../../src/tools/word/merge.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import path from 'path';
import { z } from 'zod';
import { FastMCPContext } from '../../../src/types/common.types'; // Import FastMCPContext
import { saveResource } from '../../../src/tools/dynamic/resources.tool';

// Mock de las dependencias
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/tools/dynamic/resources.tool'); // Mock saveResource
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
    createErrorResponse: jest.fn((message, code, details) => ({ success: false, error: { code, message, details } })),
}));
jest.mock('../../../src/utils/logger');
jest.mock('path', () => ({
    ...jest.requireActual('path'), // Import and retain default behavior
    basename: jest.fn(), // Mock only basename
    // dirname will use actual implementation or can be mocked if needed for schema tests elsewhere
}));

const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;
const mockLoggerInfo = jest.spyOn(logger, 'info');
const mockLoggerWarn = jest.spyOn(logger, 'warn'); // Added mock for warn
const mockLoggerError = jest.spyOn(logger, 'error');
const mockPathBasename = path.basename as jest.Mock;
const mockSaveResource = saveResource as jest.Mock;


describe('word/merge unit tests', () => {
    let mockWordApp: any; // This will be officeAppInstance.app
    let mockTargetDoc: any;
    let mockSourceDoc1: any;
    let mockSourceDoc2: any;
    let mockTargetDocContentRange: any;
    let mockSourceDocContentRange: any;
    let mockOfficeAppInstance: any; // This is what getOfficeApplication returns
    let mockContext: FastMCPContext<undefined>;

    beforeEach(() => {
        jest.clearAllMocks();

        mockTargetDocContentRange = {
            End: 100, Collapse: jest.fn(), Paste: jest.fn(), InsertBreak: jest.fn(), Text: "Target Content",
            release: jest.fn(),
        };
        mockSourceDocContentRange = { Copy: jest.fn(), Text: "Source Content", release: jest.fn() };
        mockTargetDoc = {
            Activate: jest.fn(), Content: mockTargetDocContentRange, SaveAs2: jest.fn(), Close: jest.fn(), release: jest.fn(),
        };
        mockSourceDoc1 = { Content: mockSourceDocContentRange, Close: jest.fn(), release: jest.fn() };
        mockSourceDoc2 = { Content: mockSourceDocContentRange, Close: jest.fn(), release: jest.fn() };
        
        // mockWordApp is the .app property of mockOfficeAppInstance
        mockWordApp = {
            Documents: {
                Add: jest.fn().mockReturnValue(mockTargetDoc),
                Open: jest.fn(), // This will be used directly by the tool via officeAppInstance.app
            },
            Visible: false, DisplayAlerts: 0, release: jest.fn(),
        };
        // mockOfficeAppInstance is what getOfficeApplication is mocked to return
        mockOfficeAppInstance = {
            app: mockWordApp,
            release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockOfficeAppInstance);
        mockValidateFilePath.mockImplementation(fp => fp);
        mockPathBasename.mockImplementation(p => p.substring(p.lastIndexOf('/') + 1));

        // Configure mockWordApp.Documents.Open (which is officeAppInstance.app.Documents.Open)
        mockWordApp.Documents.Open.mockImplementation((filePath: string) => {
            if (filePath.includes('doc1.docx')) return mockSourceDoc1;
            if (filePath.includes('doc2.docx')) return mockSourceDoc2;
            return null;
        });

        mockContext = {
            log: { info: mockLoggerInfo, warn: mockLoggerWarn, error: mockLoggerError, debug: jest.fn() },
            reportProgress: jest.fn(),
            userId: 'testUser',
        } as unknown as FastMCPContext<undefined>;
        mockSaveResource.mockResolvedValue({ success: true, data: { permalink: "test-permalink"} });
    });

    describe('mergeDocuments (COM only tests)', () => {
        const baseParams = {
            docs: ['C:/test/doc1.docx', 'C:/test/doc2.docx'],
            output: 'C:/test/output.docx',
        };

        test('should correctly merge documents using COM interop', async () => {
            const params = { ...baseParams };
            const result = await mergeDocuments(params, mockContext);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.output);
            expect(mockValidateFilePath).toHaveBeenCalledWith(params.docs[0]);
            expect(mockValidateFilePath).toHaveBeenCalledWith(params.docs[1]);

            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.app.Documents.Add).toHaveBeenCalled();
            expect(mockOfficeAppInstance.app.Documents.Open).toHaveBeenCalledWith(params.docs[0], false, true);
            expect(mockOfficeAppInstance.app.Documents.Open).toHaveBeenCalledWith(params.docs[1], false, true);

            expect(mockSourceDocContentRange.Copy).toHaveBeenCalledTimes(2);
            expect(mockTargetDoc.Activate).toHaveBeenCalledTimes(2);
            expect(mockTargetDocContentRange.Collapse).toHaveBeenCalledWith(0);
            expect(mockTargetDocContentRange.Paste).toHaveBeenCalledTimes(2);
            expect(mockTargetDocContentRange.InsertBreak).toHaveBeenCalledTimes(1);

            expect(mockTargetDoc.SaveAs2).toHaveBeenCalledWith(params.output);
            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false);
            expect(mockSourceDoc1.Close).toHaveBeenCalledWith(false);
            expect(mockSourceDoc2.Close).toHaveBeenCalledWith(false);

            expect(mockReleaseObject).toHaveBeenCalledWith(mockTargetDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc1);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc2);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockOfficeAppInstance.app); // Word app itself
            expect(mockOfficeAppInstance.release).toHaveBeenCalled(); // The instance wrapper

            expect(mockContext.reportProgress).toHaveBeenCalledWith({ progress: 0, total: 2 });
            // ... other reportProgress checks

            expect(mockPathBasename).toHaveBeenCalledWith(params.output);
            expect(mockSaveResource).toHaveBeenCalledWith('word/merge', params.output.substring(params.output.lastIndexOf('/') + 1), "Target Content");

            expect(result.success).toBe(true);
            if(result.success) {
                expect(result.data.outputPath).toBe(params.output);
                expect(result.message).toContain('Successfully merged 2 documents');
            }
        });

        test('should throw ZodError for less than two documents', async () => {
            const params = { ...baseParams, docs: ['C:/test/doc1.docx'] };
            await expect(mergeDocuments(params, mockContext)).rejects.toThrow(z.ZodError);
             try {
                await mergeDocuments(params, mockContext);
            } catch (e: any) {
                expect(e.errors[0].message).toBe('At least two documents are required for merging.');
            }
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
        });

        test('should throw error if output path validation fails', async () => {
            mockValidateFilePath.mockImplementation((filePathArg: string) => {
                if (filePathArg === baseParams.output) {
                    throw new Error(`Output path validation failed for: ${filePathArg}`);
                }
                return filePathArg;
            });
            const params = { ...baseParams };
            await expect(mergeDocuments(params, mockContext)).rejects.toThrow(`Output path validation failed for: ${params.output}`);
        });

        test('should throw error if a source path validation fails', async () => {
            mockValidateFilePath.mockImplementation((filePathArg: string) => {
                if (filePathArg === baseParams.docs[0]) {
                    throw new Error(`Source path validation failed for: ${filePathArg}`);
                }
                return filePathArg;
            });
            const params = { ...baseParams };
            await expect(mergeDocuments(params, mockContext)).rejects.toThrow(`Source path validation failed for: ${params.docs[0]}`);
        });

        test('should handle error when getOfficeApplication fails', async () => {
            mockGetOfficeApplication.mockRejectedValue(new Error('Failed to get Word app'));
            const params = { ...baseParams };
            await expect(mergeDocuments(params, mockContext)).rejects.toThrow('Failed to get Word app');
        });

        test('should handle error when creating target document fails', async () => {
            mockOfficeAppInstance.app.Documents.Add.mockReturnValue(null);
            const params = { ...baseParams };
            await expect(mergeDocuments(params, mockContext)).rejects.toThrow('Failed to create target document.');
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('should handle error when opening a source document', async () => {
            mockOfficeAppInstance.app.Documents.Open.mockImplementation((filePath: string) => {
                if (filePath.includes('doc1.docx')) throw new Error('COM Error: Failed to open doc1');
                return mockSourceDoc2;
            });
            const params = { ...baseParams };
            await expect(mergeDocuments(params, mockContext)).rejects.toThrow('COM Error: Failed to open doc1');
            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('should handle error when saving the target document', async () => {
            mockTargetDoc.SaveAs2.mockImplementation(() => { throw new Error('COM Error: Failed to save target'); });
            const params = { ...baseParams };
            await expect(mergeDocuments(params, mockContext)).rejects.toThrow('COM Error: Failed to save target');
            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('should release orphaned source documents on error', async () => {
            mockOfficeAppInstance.app.Documents.Open.mockImplementation((filePath: string) => {
                if (filePath.includes('doc1.docx')) return mockSourceDoc1;
                if (filePath.includes('doc2.docx')) throw new Error('Failed to open doc2');
                return null;
            });
            const params = { ...baseParams };
            await expect(mergeDocuments(params, mockContext)).rejects.toThrow('Failed to open doc2');
            expect(mockSourceDoc1.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc1);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('should use context logger if available', async () => {
            const params = { ...baseParams };
            await mergeDocuments(params, mockContext);
            expect(mockContext.log.info).toHaveBeenCalled();
            expect(mockLoggerInfo).not.toHaveBeenCalled();
        });

        test('should use global logger if context is not provided', async () => {
            const params = { ...baseParams };
            await mergeDocuments(params); // No context
            expect(mockLoggerInfo).toHaveBeenCalled();
        });

        test('should throw error if context is incomplete (missing reportProgress)', async () => {
            const params = { ...baseParams };
            const incompleteContext = { log: mockContext.log } as FastMCPContext<undefined>;
            await expect(mergeDocuments(params, incompleteContext)).rejects.toThrow('Tool context is missing required properties (log, reportProgress).');
        });

        test('should throw error if context is incomplete (missing log)', async () => {
            const params = { ...baseParams };
            const incompleteContext = { reportProgress: jest.fn() } as unknown as FastMCPContext<undefined>;
            await expect(mergeDocuments(params, incompleteContext)).rejects.toThrow('Tool context is missing required properties (log, reportProgress).');
        });
    });
});