/**
 * @file Implements tools for handling images within Word documents using COM Interop.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { Document, Packer, Media, ImageRun } from 'docx';
import * as mammoth from 'mammoth';
import { McpResource, ApiResponse, ErrorResponse, SuccessResponse, ToolRequestParams } from '../../types/common.types.js';
import { imageContent, TextContent, UserError, Context as FastMCPContext } from 'fastmcp';
import { extractImageFromWord, insertImageIntoWord } from '../../utils/officeInterop.js';
import { validateFilePath } from '../../utils/security.js';
import logger from '../../utils/logger.js';
import path from 'path';
import fs from 'fs/promises';
import { pathExists } from 'fs-extra'; // Import pathExists from fs-extra
import { Paragraph } from 'docx'; // Ensure Paragraph is imported

// --- Schema Definitions ---

/**
 * Zod schema for the input parameters of the 'word/image/extract' tool.
 */
export const ExtractImageSchema = z.object({
  /** The path to the Word file (relative to the current workspace directory). */
  filePath: z.string().min(1, "File path cannot be empty.").refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  /** Identifier for the image (e.g., 1-based index or placeholder text/bookmark). */
  identifier: z.union([z.number().int().positive("Image index must be a positive integer."), z.string().min(1, "Image identifier text cannot be empty.")], {
    description: "Identifier for the image (e.g., 1-based index or placeholder text/bookmark).",
  }),
  /** Desired output format for the extracted image. */
  outputFormat: z.enum(['png', 'jpeg', 'gif', 'bmp']).default('png').describe("Desired output format for the extracted image. Note: Library path may not fully support all format conversions."),
  useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
});

/**
 * Zod schema for the input parameters of the 'word/image/insert' tool.
 */
export const InsertImageSchema = z.object({
  /** The path to the Word file (relative to the current workspace directory). */
  filePath: z.string().min(1, "File path cannot be empty.").refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  /** Image data in base64 format. */
  imageDataBase64: z.string().min(1, "Image data (base64) cannot be empty."),
  /** Insertion position (e.g., 'end', 'bookmark:name', 'paragraph:N'). */
  position: z.string().min(1, "Insertion position cannot be empty (e.g., 'end', 'bookmark:name', 'paragraph:N')."),
  // Optional parameters for image sizing, etc.
  /** Optional width for the inserted image in points. */
  width: z.number().optional().describe("Optional width for the inserted image in points."),
  /** Optional height for the inserted image in points. */
  height: z.number().optional().describe("Optional height for the inserted image in points."),
  /** Optional alternative text for the image. */
  altText: z.string().optional().describe("Optional alternative text for the image."),
  useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
});

// --- Tool Handlers ---

/**
 * Handles the 'word/image/extract' tool request.
 * Extracts a specific image from a Word document.
 * @param params - The parameters for the tool, validated against `ExtractImageSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse containing the image buffer or an error.
 * @throws {UserError} If the image is not found or extraction fails.
 */
export async function handleExtractImage(
  params: ToolRequestParams,
  context?: FastMCPContext<undefined>
): Promise<ApiResponse<Buffer>> {
  const validationResult = ExtractImageSchema.safeParse(params);
  if (!validationResult.success) {
    logger.warn(`[word/image/extract] Invalid parameters`, { errors: validationResult.error.format() });
    return { success: false, error: { code: 'VALIDATION_ERROR', message: 'Invalid parameters provided.', details: validationResult.error.format() } };
  }
  const args = validationResult.data;
  const { filePath, identifier, outputFormat, useComInterop } = args;
  const safeFilePath = validateFilePath(filePath); // Already validated by schema, but good for direct use

  logger.info(`[word/image/extract] Request received`, { filePath: safeFilePath, identifier, outputFormat, useComInterop });

  if (useComInterop) {
    logger.info(`[word/image/extract] Using COM Interop path.`);
    try {
      const imageBuffer = await extractImageFromWord(safeFilePath, identifier);
      if (!imageBuffer || imageBuffer.length === 0) {
        throw new UserError(`COM: Image with identifier '${identifier}' not found or could not be extracted from ${safeFilePath}.`);
      }
      logger.info(`[word/image/extract] COM: Successfully extracted image`, { size: imageBuffer.length });
      return { success: true, data: imageBuffer };
    } catch (error: any) {
      logger.error(`[word/image/extract] COM Error: ${error.message}`, { error });
      if (error instanceof UserError) throw error;
      return { success: false, error: { code: 'IMAGE_EXTRACTION_FAILED_COM', message: error.message || 'COM: Unexpected error during image extraction.', details: error } };
    }
  } else {
    // Library path (Mammoth for image extraction is indirect)
    logger.info(`[word/image/extract] Using Library path.`);
    logger.warn(`[word/image/extract] Library path for precise image extraction by identifier ('${identifier}') is limited. Mammoth extracts images during full conversion. This tool will attempt a best-effort extraction.`);

    try {
        if (!await pathExists(safeFilePath)) { // Use pathExists from fs-extra
            return { success: false, error: { code: 'FILE_NOT_FOUND_LIB', message: `File not found: ${safeFilePath}`}};
        }

        let imageElementCount = 0;
        const images: { buffer: Buffer, contentType: string, altText?: string }[] = [];

        const options = {
            convertImage: mammoth.images.imgElement(async (image) => {
                imageElementCount++;
                const imageBuffer = await image.read();
                images.push({ buffer: imageBuffer, contentType: image.contentType, altText: (image as any).altText || (image as any).title }); // Attempt to get alt text
                // For mammoth, we don't directly control the output path here, just collect buffers
                return { src: `data:${image.contentType};base64,${imageBuffer.toString('base64')}` }; // Required by mammoth
            })
        };

        await mammoth.convertToHtml({ path: safeFilePath }, options); // Run conversion to trigger image extraction

        if (images.length === 0) {
            throw new UserError(`Library: No images found in ${safeFilePath}.`);
        }

        let targetImage: { buffer: Buffer, contentType: string, altText?: string } | undefined;

        if (typeof identifier === 'number') {
            if (identifier > 0 && identifier <= images.length) {
                targetImage = images[identifier - 1]; // 1-based index
            } else {
                throw new UserError(`Library: Image index ${identifier} is out of bounds. Found ${images.length} images.`);
            }
        } else if (typeof identifier === 'string') {
            // Attempt to match by alt text (case-insensitive)
            const lowerIdentifier = identifier.toLowerCase();
            targetImage = images.find(img => img.altText?.toLowerCase().includes(lowerIdentifier));
            if (!targetImage) {
                 // Fallback: if no alt text match, and identifier might be a 1-based index as string
                const potentialIndex = parseInt(identifier, 10);
                if (!isNaN(potentialIndex) && potentialIndex > 0 && potentialIndex <= images.length) {
                    targetImage = images[potentialIndex - 1];
                } else {
                    throw new UserError(`Library: Image with alt text similar to '${identifier}' not found, and identifier is not a valid index.`);
                }
            }
        }

        if (!targetImage) {
            throw new UserError(`Library: Could not identify target image with identifier '${identifier}'.`);
        }

        // Note: outputFormat is difficult to enforce with Mammoth's extraction. It gives original format.
        // Conversion would require an additional library. For now, we return what Mammoth gives.
        logger.info(`[word/image/extract] Library: Successfully extracted image (identifier: '${identifier}', found type: ${targetImage.contentType}, requested: ${outputFormat}). Format conversion not applied by library path.`);
        return { success: true, data: targetImage.buffer };

    } catch (error: any) {
        logger.error(`[word/image/extract] Library Error: ${error.message}`, { error });
        if (error instanceof UserError) throw error; // Re-throw UserError to be handled by FastMCP
        return { success: false, error: { code: 'IMAGE_EXTRACTION_FAILED_LIB', message: error.message || 'Library: Unexpected error during image extraction.', details: error } };
    }
  }
}

/**
 * Handles the 'word/image/insert' tool request.
 * Inserts an image into a Word document from base64 data.
 * @param params - The parameters for the tool, validated against `InsertImageSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse with a string confirmation or an error.
 * @throws {UserError} If the image data is invalid or insertion fails.
 */
export async function handleInsertImage(
  params: ToolRequestParams,
  context?: FastMCPContext<undefined>
): Promise<ApiResponse<string>> {
  const validationResult = InsertImageSchema.safeParse(params);
  if (!validationResult.success) {
    logger.warn(`[word/image/insert] Invalid parameters`, { errors: validationResult.error.format() });
    return { success: false, error: { code: 'VALIDATION_ERROR', message: 'Invalid parameters provided.', details: validationResult.error.format() } };
  }
  const args = validationResult.data;
  const { filePath, imageDataBase64, position, width, height, altText, useComInterop } = args;
  const safeFilePath = validateFilePath(filePath);

  logger.info(`[word/image/insert] Request received`, { filePath: safeFilePath, position, useComInterop }); // Avoid logging full base64

  const imageBuffer = Buffer.from(imageDataBase64, 'base64');
  if (imageBuffer.length === 0) {
    return { success: false, error: { code: 'INVALID_IMAGE_DATA', message: "Provided image data is empty or invalid base64." } };
  }

  if (useComInterop) {
    logger.info(`[word/image/insert] Using COM Interop path.`);
    try {
      await insertImageIntoWord(safeFilePath, imageBuffer, position, { width, height, altText });
      logger.info(`[word/image/insert] COM: Successfully inserted image`);
      return { success: true, data: `COM: Image successfully inserted into ${path.basename(safeFilePath)} at position '${position}'.` };
    } catch (error: any) {
      logger.error(`[word/image/insert] COM Error: ${error.message}`, { error });
      if (error instanceof UserError) throw error;
      return { success: false, error: { code: 'IMAGE_INSERTION_FAILED_COM', message: error.message || 'COM: Unexpected error during image insertion.', details: error } };
    }
  } else {
    // Library path (docx)
    logger.info(`[word/image/insert] Using Library (docx) path.`);
    try {
        let doc: Document;
        const imageRunProperties: any = {
            data: imageBuffer,
            transformation: {
                width: width || 200, // Default width if not provided
                height: height || 200, // Default height if not provided
            },
        };
        if (altText) {
            // docx.js doesn't have a direct altText on ImageRun in the same way COM does.
            // It's usually part of Drawing object properties, which is more complex.
            // For simplicity, we'll log this.
            logger.warn(`[word/image/insert] Library (docx) path: altText property is not directly supported on ImageRun in the same way as COM. It might be part of a more complex Drawing object. Alt text "${altText}" will be ignored for now.`);
        }
        const image = new ImageRun(imageRunProperties);


        if (await pathExists(safeFilePath)) { // Use pathExists from fs-extra
            // Modifying existing documents to insert at specific positions is complex with docx.js
            // For now, this path will primarily support creating new documents or simple appends.
            logger.warn(`[word/image/insert] Library (docx) path: Modifying existing file '${safeFilePath}' to insert image at specific position '${position}' is complex and not fully supported. Attempting to append image or create new if position is 'end' or simple.`);
            // For a simple append, one might read the existing doc, add a new section/paragraph with the image.
            // This is non-trivial. For now, let's restrict to new file creation or return error.
            return { success: false, error: { code: 'NOT_IMPLEMENTED_LIB_MODIFY', message: `Library (docx) path: Inserting images into existing documents at specific positions ('${position}') is not fully implemented. Try creating a new document or use COM Interop.`}};
        } else {
            // Create new document
            if (position.toLowerCase() !== 'end' && !position.toLowerCase().startsWith('paragraph:1')) { // crude check for start
                logger.warn(`[word/image/insert] Library (docx) path: Position '${position}' for new file ignored. Image will be added as a primary element.`);
            }
            doc = new Document({
                sections: [{
                    children: [new Paragraph({ children: [image] })],
                }],
            });
        }

        const buffer = await Packer.toBuffer(doc);
        await fs.writeFile(safeFilePath, buffer);
        logger.info(`[word/image/insert] Library (docx): Document with image operation completed at ${safeFilePath}`);
        return { success: true, data: `Library (docx): Image operation completed for ${path.basename(safeFilePath)}.` };

    } catch (error: any) {
        logger.error(`[word/image/insert] Library Error: ${error.message}`, { error });
        if (error instanceof UserError) throw error;
        return { success: false, error: { code: 'IMAGE_INSERTION_FAILED_LIB', message: error.message || 'Library: Unexpected error during image insertion.', details: error } };
    }
  }
}


// --- Resource Definitions ---

/**
 * McpResource definition for the 'word/image/extract' tool.
 * Extracts a specific image from a Word document.
 */
export const wordImageExtractTool: McpResource = {
  path: 'word/image/extract',
  description: 'Extracts a specific image from a Word document (.docx) based on index or identifier.',
  schema: ExtractImageSchema,
  handler: handleExtractImage,
};

/**
 * McpResource definition for the 'word/image/insert' tool.
 * Inserts an image into a Word document from base64 data.
 */
export const wordImageInsertTool: McpResource = {
  path: 'word/image/insert',
  description: 'Inserts an image (from base64 data) into a Word document (.docx) at a specified position.',
  schema: InsertImageSchema,
  handler: handleInsertImage,
};

// Export as a single object containing all tools
export const wordImageTools = {
    wordImageExtractTool,
    wordImageInsertTool,
    // Add other image/audio tools here later if needed
};

// No default export needed if importing the named export directly in tools/index.ts