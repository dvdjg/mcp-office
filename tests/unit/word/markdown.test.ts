import { exportToMarkdown, importFromMarkdown, exportSchema, importSchema } from '../../../src/tools/word/markdown.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { z } from 'zod';
import * as fs from 'fs-extra';
import mammoth from 'mammoth';
import archiver from 'archiver';
import * as path from 'path';
import { createErrorResponse as mockCreateErrorResponseUtil, handleToolError as mockHandleToolErrorUtil } from '../../../src/utils/errorHandler';
import { saveResource as mockSaveResourceUtil } from '../../../src/tools/dynamic/resources.tool';
import { applyMarkdownFormattingToWord as mockApplyMarkdownUtil } from '../../../src/utils/markdownToOffice';

// Mock dependencies
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler');
jest.mock('../../../src/utils/logger');
jest.mock('fs-extra');
jest.mock('mammoth');
jest.mock('archiver');
jest.mock('../../../src/tools/dynamic/resources.tool');
jest.mock('../../../src/utils/markdownToOffice');

// --- Mocks ---
const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;

const mockLoggerInfo = jest.spyOn(logger, 'info');
const mockLoggerWarn = jest.spyOn(logger, 'warn');
const mockLoggerError = jest.spyOn(logger, 'error');
const mockLoggerDebug = jest.spyOn(logger, 'debug');

const mockFsPathExists = fs.pathExists as unknown as jest.Mock;
const mockFsEnsureDir = fs.ensureDir as unknown as jest.Mock;
const mockFsWriteFile = fs.writeFile as unknown as jest.Mock;
const mockFsReadFile = fs.readFile as unknown as jest.Mock;
const mockFsCreateWriteStream = fs.createWriteStream as unknown as jest.Mock;
const mockFsReaddir = fs.readdir as unknown as jest.Mock;

const mockMammothConvertToMarkdown = (mammoth as any).convertToMarkdown as jest.Mock;
const mockMammothImagesImgElement = (mammoth.images as any).imgElement as jest.Mock;

const mockArchiverPipe = jest.fn();
const mockArchiverFile = jest.fn();
const mockArchiverDirectory = jest.fn();
const mockArchiverFinalize = jest.fn().mockResolvedValue(undefined);
const mockArchiver = archiver;
(mockArchiver as unknown as jest.Mock).mockReturnValue({
    pipe: mockArchiverPipe,
    file: mockArchiverFile,
    directory: mockArchiverDirectory,
    finalize: mockArchiverFinalize,
    on: jest.fn(), // For error handling if needed
    pointer: jest.fn().mockReturnValue(1024), // Mock pointer
} as any);


const mockCreateErrorResponse = mockCreateErrorResponseUtil as jest.Mock;
const mockHandleToolError = mockHandleToolErrorUtil as jest.Mock;
const mockSaveResource = mockSaveResourceUtil as jest.Mock;
const mockApplyMarkdownFormattingToWord = mockApplyMarkdownUtil as jest.Mock;


describe('word/markdown unit tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockStoryRange: any;
    let mockContentRange: any;

    beforeEach(() => {
        jest.clearAllMocks();

        mockStoryRange = { Text: 'Main story text', release: jest.fn() };
        mockContentRange = { Text: 'New doc content', Collapse: jest.fn(), release: jest.fn() };
        mockDoc = {
            StoryRanges: jest.fn().mockReturnValue(mockStoryRange), // For export
            Content: mockContentRange, // For import
            SaveAs2: jest.fn(),
            Close: jest.fn(),
            release: jest.fn(),
        };
        mockWordApp = {
            Documents: {
                Open: jest.fn().mockReturnValue(mockDoc),
                Add: jest.fn().mockReturnValue(mockDoc),
            },
            release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockWordApp);
        mockValidateFilePath.mockImplementation(fp => fp); // Pass through
        mockFsPathExists.mockResolvedValue(true);
        mockFsEnsureDir.mockResolvedValue(undefined);
        mockFsWriteFile.mockResolvedValue(undefined);
        mockFsReadFile.mockResolvedValue('## Markdown Content');
        mockFsCreateWriteStream.mockReturnValue({ on: jest.fn(), close: jest.fn() } as any);
        mockFsReaddir.mockResolvedValue(['image1.png'] as any);


        mockMammothConvertToMarkdown.mockResolvedValue({ value: 'Mammoth markdown output', messages: [] });
        mockMammothImagesImgElement.mockImplementation((converter: any) => converter); // Pass through for now

        mockSaveResource.mockResolvedValue({ success: true, data: { permalink: 'test-permalink' } });
        mockApplyMarkdownFormattingToWord.mockResolvedValue(undefined);
        mockHandleToolError.mockImplementation((err, code) => ({ success: false, error: { code, message: err.message } }));
        mockCreateErrorResponse.mockImplementation((message, code) => ({ success: false, error: { code, message } }));
    });

    const testModes = [
        { mode: 'COM', useComInterop: true, default: false },
        { mode: 'Library', useComInterop: false, default: true },
    ];

    describe.each(testModes)('exportToMarkdown (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParams = {
            filePath: 'C:/test/source.docx',
            output: 'C:/test/output.md',
        };
        const getParams = (options?: Partial<z.infer<typeof exportSchema>>) => ({
            ...baseParams,
            ...options,
            useComInterop,
        });

        if (useComInterop) {
            test('COM: should export successfully with basic parameters', async () => {
                const params = getParams();
                const result = await exportToMarkdown(params);

                expect(mockFsPathExists).toHaveBeenCalledWith(path.resolve(params.filePath));
                expect(mockFsEnsureDir).toHaveBeenCalledWith(path.join(path.dirname(path.resolve(params.output)), 'images'));
                expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
                expect(mockWordApp.Documents.Open).toHaveBeenCalledWith(path.resolve(params.filePath), false, true, false, "", "", true, "", "", "", 0, false, false);
                // Simplified check for COM traversal placeholder
                expect(mockDoc.StoryRanges).toHaveBeenCalledWith(1); // wdMainStory
                expect(mockFsWriteFile).toHaveBeenCalledWith(path.resolve(params.output), 'Main story text', 'utf8');
                expect(result.success).toBe(true);
                if (result.success) {
                    expect(result.data.outputPath).toBe(path.resolve(params.output));
                }
                expect(mockReleaseObject).toHaveBeenCalledWith(mockDoc);
                expect(mockReleaseObject).toHaveBeenCalledWith(mockWordApp);
            });

            test('COM: should handle zipOutput correctly', async () => {
                const params = getParams({ zipOutput: true, zipFileName: 'archive.zip' });
                const expectedZipPath = path.join(path.dirname(path.resolve(params.output)), 'archive.zip');
                const result = await exportToMarkdown(params);

                expect(mockArchiver).toHaveBeenCalledWith('zip', { zlib: { level: 9 } });
                expect(mockArchiverFile).toHaveBeenCalledWith(path.resolve(params.output), { name: path.basename(params.output) });
                // expect(mockArchiverDirectory).toHaveBeenCalled(); // Depends on image extraction
                expect(mockArchiverFinalize).toHaveBeenCalled();
                expect(result.success).toBe(true);
                if (result.success) {
                    expect(result.data.outputPath).toBe(expectedZipPath);
                }
            });

            test('COM: should return error if source file not found', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = getParams();
                const result = await exportToMarkdown(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('FILE_NOT_FOUND');
            });

        } else { // Library path
            test('Library: should export successfully with basic parameters (default useComInterop:false)', async () => {
                const params = { ...baseParams, output: 'C:/test/lib_output.md' }; // No useComInterop
                const result = await exportToMarkdown(params);

                expect(mockFsPathExists).toHaveBeenCalledWith(path.resolve(params.filePath));
                expect(mockFsEnsureDir).toHaveBeenCalledWith(path.join(path.dirname(path.resolve(params.output)), 'images'));
                expect(mockMammothConvertToMarkdown).toHaveBeenCalledWith({ path: path.resolve(params.filePath) }, expect.any(Object));
                expect(mockFsWriteFile).toHaveBeenCalledWith(path.resolve(params.output), 'Mammoth markdown output', 'utf8');
                expect(result.success).toBe(true);
                if (result.success) {
                    expect(result.data.outputPath).toBe(path.resolve(params.output));
                }
            });

            test('Library: should handle zipOutput correctly', async () => {
                const params = getParams({ zipOutput: true, zipFileName: 'lib_archive.zip', output: 'C:/test/lib_zip_output.md' });
                const expectedZipPath = path.join(path.dirname(path.resolve(params.output)), 'lib_archive.zip');
                const result = await exportToMarkdown(params);

                expect(mockArchiver).toHaveBeenCalledWith('zip', { zlib: { level: 9 } });
                expect(mockArchiverFile).toHaveBeenCalledWith(path.resolve(params.output), { name: path.basename(params.output) });
                expect(mockArchiverDirectory).toHaveBeenCalled();
                expect(mockArchiverFinalize).toHaveBeenCalled();
                expect(result.success).toBe(true);
                if (result.success) {
                    expect(result.data.outputPath).toBe(expectedZipPath);
                }
            });
            
            test('Library: should log warnings for mammoth messages', async () => {
                mockMammothConvertToMarkdown.mockResolvedValue({ value: 'output', messages: [{type: 'warning', message: 'A mammoth warning'}] });
                const params = getParams();
                await exportToMarkdown(params);
                expect(mockLoggerWarn).toHaveBeenCalledWith("Mammoth message (warning): A mammoth warning");
            });
        }
    });

    describe.each(testModes)('importFromMarkdown (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParams = {
            filePath: 'C:/test/source.md',
            output: 'C:/test/output.docx',
        };
        const getParams = (options?: Partial<z.infer<typeof importSchema>>) => ({
            ...baseParams,
            ...options,
            useComInterop,
        });

        if (useComInterop) {
            test('COM: should import successfully with basic parameters', async () => {
                const params = getParams();
                const result = await importFromMarkdown(params);

                expect(mockFsReadFile).toHaveBeenCalledWith(path.resolve(params.filePath), 'utf8');
                expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
                expect(mockWordApp.Documents.Add).toHaveBeenCalledWith(); // No template
                expect(mockApplyMarkdownFormattingToWord).toHaveBeenCalledWith(mockContentRange, '## Markdown Content', mockWordApp, mockDoc);
                expect(mockDoc.SaveAs2).toHaveBeenCalledWith(path.resolve(params.output), 16); // wdFormatDocumentDefault
                expect(result.success).toBe(true);
                if (result.success) {
                    expect(result.data.outputPath).toBe(path.resolve(params.output));
                }
            });

            test('COM: should use template if provided', async () => {
                const templatePath = 'C:/test/template.dotx';
                const params = getParams({ template: templatePath });
                await importFromMarkdown(params);
                expect(mockWordApp.Documents.Add).toHaveBeenCalledWith(path.resolve(templatePath));
            });

            test('COM: should return error if source file not found', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const params = getParams();
                const result = await importFromMarkdown(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('FILE_NOT_FOUND');
            });

        } else { // Library path
            test('Library: should return NOT_IMPLEMENTED_LIB (default useComInterop:false)', async () => {
                const params = { ...baseParams }; // No useComInterop
                const result = await importFromMarkdown(params);
                expect(result.success).toBe(false);
                if (!result.success) expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB');
                expect(mockHandleToolError).toHaveBeenCalledWith(expect.any(Error), 'NOT_IMPLEMENTED_LIB');
            });

            test('Library: should log warning if template is provided', async () => {
                const params = getParams({ template: 'C:/test/template.dotx' });
                await importFromMarkdown(params);
                expect(mockLoggerWarn).toHaveBeenCalledWith(expect.stringContaining("Template usage with .dotx files is not directly supported"));
            });
        }
    });
});