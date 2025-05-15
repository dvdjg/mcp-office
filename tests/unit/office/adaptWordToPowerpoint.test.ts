import adaptWordToPowerpointTool from '../../../src/tools/office/adaptWordToPowerpoint.tool';
import getActiveOfficeDocumentsTool from '../../../src/tools/os/getActiveOfficeDocuments.tool';
import wordToPowerpointTool from '../../../src/tools/office/wordToPowerpoint.tool';
import { ApiResponse } from '../../../src/types/common.types';
import path from 'path';

// Mock the dependencies
jest.mock('../../../src/tools/os/getActiveOfficeDocuments.tool');
jest.mock('../../../src/tools/office/wordToPowerpoint.tool');
jest.mock('../../../src/utils/logger'); // Simpler mock, will spy on methods later if needed


describe('adaptWordToPowerpointTool', () => {
    const mockGetActiveOfficeDocuments = getActiveOfficeDocumentsTool.handler as jest.Mock;
    const mockWordToPowerpoint = wordToPowerpointTool.handler as jest.Mock;

    beforeEach(() => {
        jest.clearAllMocks();
    });

    it('should successfully convert a Word document specified by inputWordPath', async () => {
        const inputWordPath = 'C:\\Users\\Test\\Documents\\TestDoc.docx';
        const outputPowerpointPath = 'C:\\Users\\Test\\Documents\\TestDoc.pptx';
        mockWordToPowerpoint.mockResolvedValueOnce({
            success: true,
            data: { message: 'Conversion successful', powerpointFilePath: outputPowerpointPath },
        } as ApiResponse<any>);

        const params = { inputWordPath, outputPowerpointPath, headingLevelForNewSlide: 1 };
        const result = await adaptWordToPowerpointTool.handler(params);

        expect(result.success).toBe(true);
        if (result.success) { // Type guard
            expect(result.data.powerpointFilePath).toBe(outputPowerpointPath);
        }
        expect(mockWordToPowerpoint).toHaveBeenCalledWith(
            expect.objectContaining({
                wordFilePath: inputWordPath,
                powerpointFilePath: outputPowerpointPath,
                operation: 'transfer',
                headingLevelForNewSlide: 1,
            }),
            undefined // Context
        );
        expect(mockGetActiveOfficeDocuments).not.toHaveBeenCalled();
    });

    it('should use active Word document if inputWordPath is not provided', async () => {
        const activeWordPath = 'C:\\Users\\Test\\Documents\\ActiveDoc.docx';
        const expectedOutputPowerpointPath = path.join(path.dirname(activeWordPath), `${path.basename(activeWordPath, '.docx')}.pptx`);

        mockGetActiveOfficeDocuments.mockResolvedValueOnce({
            success: true,
            data: { documents: [{ filePath: activeWordPath, applicationType: 'Word' }] },
        } as ApiResponse<any>);
        mockWordToPowerpoint.mockResolvedValueOnce({
            success: true,
            data: { message: 'Conversion successful', powerpointFilePath: expectedOutputPowerpointPath },
        } as ApiResponse<any>);

        const params = { headingLevelForNewSlide: 2 };
        const result = await adaptWordToPowerpointTool.handler(params);

        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data.powerpointFilePath).toBe(expectedOutputPowerpointPath);
        }
        expect(mockGetActiveOfficeDocuments).toHaveBeenCalledTimes(1);
        expect(mockWordToPowerpoint).toHaveBeenCalledWith(
            expect.objectContaining({
                wordFilePath: activeWordPath,
                powerpointFilePath: expectedOutputPowerpointPath,
                operation: 'transfer',
                headingLevelForNewSlide: 2,
            }),
            undefined
        );
    });

    it('should default outputPowerpointPath if not provided', async () => {
        const inputWordPath = 'C:\\Users\\Test\\MyReport.docx';
        const expectedOutputPowerpointPath = 'C:\\Users\\Test\\MyReport.pptx';

        mockWordToPowerpoint.mockResolvedValueOnce({
            success: true,
            data: { message: 'Conversion successful', powerpointFilePath: expectedOutputPowerpointPath },
        } as ApiResponse<any>);

        const params = { inputWordPath }; // outputPowerpointPath is omitted
        const result = await adaptWordToPowerpointTool.handler(params);

        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data.powerpointFilePath).toBe(expectedOutputPowerpointPath);
        }
        expect(mockWordToPowerpoint).toHaveBeenCalledWith(
            expect.objectContaining({
                powerpointFilePath: expectedOutputPowerpointPath,
            }),
            undefined
        );
    });


    it('should return error if no active Word document is found and no inputPath is given', async () => {
        mockGetActiveOfficeDocuments.mockResolvedValueOnce({
            success: true,
            data: { documents: [] }, // No Word documents
        } as ApiResponse<any>);

        const params = {};
        const result = await adaptWordToPowerpointTool.handler(params);

        expect(result.success).toBe(false);
        if (!result.success) { // Type guard
            expect(result.error.code).toBe('NO_WORD_DOC_ACTIVE');
        }
        expect(mockWordToPowerpoint).not.toHaveBeenCalled();
    });

    it('should return error if getActiveOfficeDocumentsTool fails', async () => {
        mockGetActiveOfficeDocuments.mockResolvedValueOnce({
            success: false,
            error: { code: 'SOME_OS_ERROR', message: 'OS tool failed' },
        } as ApiResponse<any>);

        const params = {};
        const result = await adaptWordToPowerpointTool.handler(params);

        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('ACTIVE_DOC_ERROR');
        }
        expect(mockWordToPowerpoint).not.toHaveBeenCalled();
    });

    it('should return error if wordToPowerpointTool fails', async () => {
        const inputWordPath = 'C:\\Users\\Test\\Documents\\TestDoc.docx';
        mockWordToPowerpoint.mockResolvedValueOnce({
            success: false,
            error: { code: 'CONVERSION_TOOL_ERROR', message: 'Conversion tool failed' },
        } as ApiResponse<any>);

        const params = { inputWordPath };
        const result = await adaptWordToPowerpointTool.handler(params);

        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('CONVERSION_FAILED');
        }
    });

    it('should handle validation error for input parameters', async () => {
        const params = { headingLevelForNewSlide: 'not-a-number' }; // Invalid type
        // @ts-ignore to test invalid param type
        const result = await adaptWordToPowerpointTool.handler(params);

        expect(result.success).toBe(false);
        if (!result.success) {
            expect(result.error.code).toBe('VALIDATION_ERROR');
        }
        expect(mockGetActiveOfficeDocuments).not.toHaveBeenCalled();
        expect(mockWordToPowerpoint).not.toHaveBeenCalled();
    });

    it('should use the first active Word document if multiple are open', async () => {
        const activeWordPath1 = 'C:\\Users\\Test\\Documents\\ActiveDoc1.docx';
        const activeWordPath2 = 'C:\\Users\\Test\\Documents\\ActiveDoc2.docx';
        const expectedOutputPowerpointPath = path.join(path.dirname(activeWordPath1), `${path.basename(activeWordPath1, '.docx')}.pptx`);

        mockGetActiveOfficeDocuments.mockResolvedValueOnce({
            success: true,
            data: { documents: [
                { filePath: activeWordPath1, applicationType: 'Word' },
                { filePath: activeWordPath2, applicationType: 'Word' }
            ]},
        } as ApiResponse<any>);
        mockWordToPowerpoint.mockResolvedValueOnce({
            success: true,
            data: { message: 'Conversion successful', powerpointFilePath: expectedOutputPowerpointPath },
        } as ApiResponse<any>);

        const params = {};
        const result = await adaptWordToPowerpointTool.handler(params);

        expect(result.success).toBe(true);
        if (result.success) {
            expect(result.data.powerpointFilePath).toBe(expectedOutputPowerpointPath);
        }
        expect(mockWordToPowerpoint).toHaveBeenCalledWith(
            expect.objectContaining({
                wordFilePath: activeWordPath1, // Should use the first one
            }),
            undefined
        );
    });

    it('should log AI enhancement placeholder if enabled', async () => {
        const inputWordPath = 'C:\\Users\\Test\\Documents\\AiTest.docx';
        const outputPowerpointPath = 'C:\\Users\\Test\\Documents\\AiTest.pptx';
        const logger = (await import('../../../src/utils/logger')).default; // Import and access default export

        mockWordToPowerpoint.mockResolvedValueOnce({
            success: true,
            data: { message: 'Conversion successful', powerpointFilePath: outputPowerpointPath },
        } as ApiResponse<any>);

        const params = { inputWordPath, outputPowerpointPath, enableAiEnhancement: true };
        await adaptWordToPowerpointTool.handler(params);

        expect(logger.info).toHaveBeenCalledWith('AI enhancement is enabled (placeholder). Future implementation needed.');
    });
});