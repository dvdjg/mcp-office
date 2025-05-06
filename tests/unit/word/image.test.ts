import { handleExtractImage, handleInsertImage, ExtractImageSchema, InsertImageSchema } from '../../../src/tools/word/image.tool';
import { extractImageFromWord, insertImageIntoWord } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { z } from 'zod';
import * as fsPromises from 'fs/promises'; // Import fs/promises
import * as fsExtra from 'fs-extra'; // Import fs-extra
import mammoth from 'mammoth';
import { Packer, ImageRun, Paragraph, Document as DocxDocument } from 'docx';
import { UserError } from 'fastmcp';

// Mock dependencies
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/logger');
jest.mock('fs/promises', () => ({
    writeFile: jest.fn(),
}));
jest.mock('fs-extra', () => ({
    pathExists: jest.fn(),
}));
jest.mock('mammoth');
jest.mock('docx', () => {
    const originalDocx = jest.requireActual('docx');
    return {
        ...originalDocx,
        Packer: { toBuffer: jest.fn() },
        ImageRun: jest.fn().mockImplementation(props => ({ props, _data: props.data })),
        Paragraph: jest.fn().mockImplementation(props => ({ props })),
        Document: jest.fn().mockImplementation(props => ({ props })),
    };
});


// --- Mocks ---
const mockExtractImageFromWord = extractImageFromWord as jest.Mock;
const mockInsertImageIntoWord = insertImageIntoWord as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;

const mockLoggerInfo = jest.spyOn(logger, 'info');
const mockLoggerWarn = jest.spyOn(logger, 'warn');
const mockLoggerError = jest.spyOn(logger, 'error');

const mockFsWriteFile = fsPromises.writeFile as jest.Mock;
const mockFsPathExists = fsExtra.pathExists as jest.Mock;

const mockMammothConvertToHtml = mammoth.convertToHtml as jest.Mock;
const mockMammothImagesImgElement = (mammoth.images as any).imgElement as jest.Mock; // Cast to any if 'images' is problematic

const mockPackerToBuffer = Packer.toBuffer as jest.Mock;
const mockImageRun = ImageRun as jest.Mock;
const mockParagraph = Paragraph as jest.Mock;
const mockDocxDocument = DocxDocument as jest.Mock;


describe('word/image unit tests', () => {

    beforeEach(() => {
        jest.clearAllMocks();

        mockValidateFilePath.mockImplementation(fp => fp);
        mockFsPathExists.mockResolvedValue(true); // Default to file existing
        mockFsWriteFile.mockResolvedValue(undefined);
        mockPackerToBuffer.mockResolvedValue(Buffer.from("docx-image-buffer"));

        // Default for mammoth image extraction
        mockMammothImagesImgElement.mockImplementation((converter: any) => async (image: any) => {
            const buffer = await image.read();
            // This part is tricky as the tool itself pushes to an array.
            // We'll need to simulate this behavior in tests that rely on it.
            return { src: `data:${image.contentType};base64,${buffer.toString('base64')}` };
        });
        mockMammothConvertToHtml.mockResolvedValue({ value: '', messages: [] }); // Default, tests can override

        mockExtractImageFromWord.mockResolvedValue(Buffer.from("extracted-com-image-data"));
        mockInsertImageIntoWord.mockResolvedValue(undefined);
    });

    const testModes = [
        { mode: 'COM', useComInterop: true, default: false },
        { mode: 'Library', useComInterop: false, default: true },
    ];

    describe.each(testModes)('handleExtractImage (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParams: z.infer<typeof ExtractImageSchema> = {
            filePath: 'C:/test/doc_with_images.docx',
            identifier: 1,
            outputFormat: 'png',
            useComInterop,
        };

        if (useComInterop) {
            test('COM: should extract image successfully', async () => {
                const result = await handleExtractImage(baseParams);
                expect(mockExtractImageFromWord).toHaveBeenCalledWith(baseParams.filePath, baseParams.identifier);
                expect(result.success).toBe(true);
                if (result.success) {
                    expect(result.data).toEqual(Buffer.from("extracted-com-image-data"));
                }
            });

            test('COM: should throw UserError if extractImageFromWord throws UserError', async () => {
                const userErr = new UserError("Image not found by COM");
                mockExtractImageFromWord.mockRejectedValue(userErr);
                await expect(handleExtractImage(baseParams)).rejects.toThrow(userErr);
            });

            test('COM: should return error if extractImageFromWord throws generic error', async () => {
                mockExtractImageFromWord.mockRejectedValue(new Error("Generic COM failure"));
                const result = await handleExtractImage(baseParams);
                expect(result.success).toBe(false);
                if (!result.success) {
                    expect(result.error.code).toBe('IMAGE_EXTRACTION_FAILED_COM');
                    expect(result.error.message).toContain("Generic COM failure");
                }
            });
        } else { // Library path
            test('Library: should extract image by index if found', async () => {
                const mockImageBuffer1 = Buffer.from("lib-img-1");
                const mockImageBuffer2 = Buffer.from("lib-img-2");
                mockMammothConvertToHtml.mockImplementationOnce(async (options, mammothOptions) => {
                    // Simulate mammoth calling convertImage
                    await mammothOptions.convertImage({ read: async () => mockImageBuffer1, contentType: 'image/png' });
                    await mammothOptions.convertImage({ read: async () => mockImageBuffer2, contentType: 'image/jpeg' });
                    return { value: '', messages: [] };
                });

                const params = { ...baseParams, identifier: 2 };
                const result = await handleExtractImage(params);
                expect(result.success).toBe(true);
                if (result.success) {
                    expect(result.data).toEqual(mockImageBuffer2);
                }
            });

            test('Library: should return error if file not found', async () => {
                mockFsPathExists.mockResolvedValue(false);
                const result = await handleExtractImage(baseParams);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('FILE_NOT_FOUND_LIB');
            });

            test('Library: should throw UserError if no images found', async () => {
                mockMammothConvertToHtml.mockResolvedValue({ value: '', messages: [] }); // No images processed
                 await expect(handleExtractImage(baseParams)).rejects.toThrow(new UserError("Library: No images found in C:/test/doc_with_images.docx."));
            });
        }
    });


    describe.each(testModes)('handleInsertImage (mode: $mode, default: $default)', ({ useComInterop }) => {
        const baseParams: z.infer<typeof InsertImageSchema> = {
            filePath: 'C:/test/doc_for_image_insert.docx',
            imageDataBase64: Buffer.from("sample-image-data").toString('base64'),
            position: 'end',
            useComInterop,
        };

        if (useComInterop) {
            test('COM: should insert image successfully', async () => {
                const result = await handleInsertImage(baseParams);
                expect(mockInsertImageIntoWord).toHaveBeenCalledWith(
                    baseParams.filePath,
                    Buffer.from(baseParams.imageDataBase64, 'base64'),
                    baseParams.position,
                    { width: undefined, height: undefined, altText: undefined }
                );
                expect(result.success).toBe(true);
                if (result.success) {
                    expect(result.data).toContain("COM: Image successfully inserted");
                }
            });

            test('COM: should throw UserError if insertImageIntoWord throws UserError', async () => {
                const userErr = new UserError("Cannot insert at position by COM");
                mockInsertImageIntoWord.mockRejectedValue(userErr);
                await expect(handleInsertImage(baseParams)).rejects.toThrow(userErr);
            });
        } else { // Library path
            test('Library: should create new file and insert image', async () => {
                mockFsPathExists.mockResolvedValue(false); // New file
                const params = { ...baseParams, filePath: 'C:/test/new_doc_with_image.docx' };
                const result = await handleInsertImage(params);

                expect(mockFsPathExists).toHaveBeenCalledWith(params.filePath);
                expect(mockImageRun).toHaveBeenCalledWith({
                    data: Buffer.from(params.imageDataBase64, 'base64'),
                    transformation: { width: 200, height: 200 },
                });
                expect(mockParagraph).toHaveBeenCalledWith({ children: [expect.any(Object)] }); // Check if ImageRun instance is passed
                expect(mockDocxDocument).toHaveBeenCalled();
                expect(mockPackerToBuffer).toHaveBeenCalled();
                expect(mockFsWriteFile).toHaveBeenCalledWith(params.filePath, Buffer.from("docx-image-buffer"));
                expect(result.success).toBe(true);
                if(result.success) expect(result.data).toContain("Library (docx): Image operation completed");
            });

            test('Library: should return error if file exists (modification not supported)', async () => {
                mockFsPathExists.mockResolvedValue(true); // File exists
                const result = await handleInsertImage(baseParams);
                expect(result.success).toBe(false);
                if (!result.success) {
                    expect(result.error.code).toBe('NOT_IMPLEMENTED_LIB_MODIFY');
                }
            });
            
            test('Library: should return error for empty base64 data', async () => {
                const params = { ...baseParams, imageDataBase64: "" };
                const result = await handleInsertImage(params);
                expect(result.success).toBe(false);
                if(!result.success) expect(result.error.code).toBe('INVALID_IMAGE_DATA');
            });
        }
    });
});