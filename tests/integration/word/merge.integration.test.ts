import { mergeDocuments } from '../../../src/tools/word/merge.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import path from 'path';
import { z } from 'zod';
import { FastMCPContext } from '../../../src/types/common.types'; // Import FastMCPContext

// Mock de las dependencias externas (simulando la interacción con Office y el sistema de archivos/seguridad)
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
    createErrorResponse: jest.fn((message, code, details) => ({ success: false, error: { code, message, details } })),
}));
jest.mock('../../../src/utils/logger');
jest.mock('path'); // Mock the path module

const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;
const mockLoggerInfo = jest.spyOn(logger, 'info');
const mockLoggerWarn = jest.spyOn(logger, 'warn');
const mockLoggerError = jest.spyOn(logger, 'error');
const mockPathBasename = path.basename as jest.Mock; // Mock specific path functions if needed

describe('word/merge integration tests', () => {
    let mockWordApp: any;
    let mockTargetDoc: any;
    let mockSourceDoc1: any;
    let mockSourceDoc2: any;
    let mockTargetDocContentRange: any;
    let mockSourceDocContentRange: any;
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
        mockWordApp = {
            Documents: {
                Add: jest.fn().mockReturnValue(mockTargetDoc),
                Open: jest.fn(), // Will be configured per test
            },
            Visible: false, DisplayAlerts: 0, release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockWordApp); // getOfficeApplication returns the app directly
        mockValidateFilePath.mockImplementation(fp => fp); // Assume valid paths
        mockPathBasename.mockImplementation(p => p.substring(p.lastIndexOf('/') + 1));


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
    });

    describe('mergeDocuments (COM only)', () => {
        const baseParams = {
            docs: ['C:/test/doc1.docx', 'C:/test/doc2.docx'],
            output: 'C:/test/output.docx',
            // No useComInterop flag as this tool is COM-only
        };

        test('should correctly merge documents using COM interop', async () => {
            const params = { ...baseParams };
            const result = await mergeDocuments(params, mockContext);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.output);
            expect(mockValidateFilePath).toHaveBeenCalledWith(params.docs[0]);
            expect(mockValidateFilePath).toHaveBeenCalledWith(params.docs[1]);

            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Add).toHaveBeenCalled();
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.docs[0], false, true);
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.docs[1], false, true);

            expect(mockSourceDocContentRange.Copy).toHaveBeenCalledTimes(2);
            expect(mockTargetDoc.Activate).toHaveBeenCalledTimes(2);
            expect(mockTargetDocContentRange.Collapse).toHaveBeenCalledWith(0); // wdCollapseEnd
            expect(mockTargetDocContentRange.Paste).toHaveBeenCalledTimes(2);
            expect(mockTargetDocContentRange.InsertBreak).toHaveBeenCalledTimes(1); // Between doc1 and doc2

            expect(mockTargetDoc.SaveAs2).toHaveBeenCalledWith(params.output);
            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false);
            expect(mockSourceDoc1.Close).toHaveBeenCalledWith(false);
            expect(mockSourceDoc2.Close).toHaveBeenCalledWith(false);

            expect(mockReleaseObject).toHaveBeenCalledWith(mockTargetDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc1);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc2);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp); // Word app itself

            expect(mockContext.reportProgress).toHaveBeenCalledWith({ progress: 0, total: 2 });
            expect(mockContext.reportProgress).toHaveBeenCalledWith({ progress: 1, total: 2 });
            expect(mockContext.reportProgress).toHaveBeenCalledWith({ progress: 2, total: 2 });

            expect(result.success).toBe(true);
            if (result.success) {
                expect(result.data.outputPath).toBe(params.output);
                expect(result.message).toContain('Successfully merged 2 documents');
            }
            expect(mockLoggerInfo).toHaveBeenCalledWith(expect.stringContaining('Starting merge process for 2 documents...'));
            expect(mockLoggerInfo).toHaveBeenCalledWith(expect.stringContaining('Merged document saved successfully.'));
        });

        test('should handle error when opening a source document', async () => {
            mockWordApp.Documents.Open.mockImplementation((filePath: string) => {
                if (filePath.includes('doc1.docx')) throw new Error('COM Error: Failed to open doc1');
                return mockSourceDoc2;
            });

            const params = { ...baseParams };
            await expect(mergeDocuments(params, mockContext)).rejects.toThrow('COM Error: Failed to open doc1');

            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Add).toHaveBeenCalled(); // Target doc still created
            expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(params.docs[0], false, true); // Attempted doc1
            expect(mockWordApp.Documents.Open).not.toHaveBeenCalledWith(params.docs[1], false, true); // Not doc2

            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false); // Target doc closed on error
            expect(mockReleaseObject).toHaveBeenCalledWith(mockTargetDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp); // Word app released
            expect(mockLoggerError).toHaveBeenCalledWith(expect.stringContaining('Error processing source document C:/test/doc1.docx: COM Error: Failed to open doc1'));
            expect(mockLoggerError).toHaveBeenCalledWith(expect.stringContaining('Error in mergeDocuments: COM Error: Failed to open doc1'));
        });

        test('should handle error when saving the target document', async () => {
            mockTargetDoc.SaveAs2.mockImplementation(() => { throw new Error('COM Error: Failed to save target'); });
            const params = { ...baseParams };

            await expect(mergeDocuments(params, mockContext)).rejects.toThrow('COM Error: Failed to save target');

            expect(mockWordApp.Documents.Open).toHaveBeenCalledTimes(2); // Both source docs opened
            expect(mockSourceDocContentRange.Copy).toHaveBeenCalledTimes(2);
            expect(mockTargetDocContentRange.Paste).toHaveBeenCalledTimes(2);
            expect(mockTargetDoc.SaveAs2).toHaveBeenCalledWith(params.output); // Save attempted

            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false); // Target doc closed on error
            expect(mockReleaseObject).toHaveBeenCalledWith(mockTargetDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc1);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc2);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            expect(mockLoggerError).toHaveBeenCalledWith(expect.stringContaining('Error in mergeDocuments: COM Error: Failed to save target'));
        });
    });
});