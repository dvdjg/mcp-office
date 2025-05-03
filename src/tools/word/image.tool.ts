/**
 * @file Implements tools for handling images within Word documents using COM Interop.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ApiResponse, ErrorResponse, SuccessResponse, ToolRequestParams } from '../../types/common.types'; // Normalized relative path
import { imageContent, TextContent, UserError, Context as FastMCPContext } from 'fastmcp'; // Import TextContent and FastMCPContext
import { extractImageFromWord, insertImageIntoWord } from '../../utils/officeInterop'; // Normalized relative path
import { validateFilePath } from '../../utils/security'; // Normalized relative path
import logger from '../../utils/logger'; // Normalized relative path
import path from 'path';
import fs from 'fs/promises'; // For reading image data for insertion if needed

// --- Schema Definitions ---

/**
 * Zod schema for the input parameters of the 'word/image/extract' tool.
 */
const ExtractImageSchema = z.object({
  /** The path to the Word file (relative to the current workspace directory). */
  filePath: z.string().min(1, "File path cannot be empty.").refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  /** Identifier for the image (e.g., 1-based index or placeholder text/bookmark). */
  identifier: z.union([z.number().int().positive("Image index must be a positive integer."), z.string().min(1, "Image identifier text cannot be empty.")], {
    description: "Identifier for the image (e.g., 1-based index or placeholder text/bookmark).",
  }),
  /** Desired output format for the extracted image. */
  outputFormat: z.enum(['png', 'jpeg', 'gif', 'bmp']).default('png').describe("Desired output format for the extracted image."),
});

/**
 * Zod schema for the input parameters of the 'word/image/insert' tool.
 */
const InsertImageSchema = z.object({
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
async function handleExtractImage(
  params: ToolRequestParams,
  context?: FastMCPContext<undefined>
): Promise<ApiResponse<Buffer>> { // Return ApiResponse with Buffer
  const validationResult = ExtractImageSchema.safeParse(params);
  if (!validationResult.success) {
    logger.warn(`[word/image/extract] Invalid parameters`, { errors: validationResult.error.format() });
    return {
        success: false,
        error: {
            code: 'VALIDATION_ERROR',
            message: 'Invalid parameters provided.',
            details: validationResult.error.format(),
        }
    } as ErrorResponse;
  }
  const args = validationResult.data; // Use validated data
  logger.info(`[word/image/extract] Received request`, { args });

  try {
    // validateFilePath is already called by the schema refinement, but calling again here is harmless
    // await validateFilePath(args.filePath); // Ensure path is safe

    // Placeholder: Implement actual COM interop call
    // The outputFormat is handled by the server loop when wrapping the buffer in imageContent
    const imageBuffer = await extractImageFromWord(args.filePath, args.identifier);

    if (!imageBuffer || imageBuffer.length === 0) {
      throw new UserError(`Image with identifier '${args.identifier}' not found or could not be extracted.`);
    }

    logger.info(`[word/image/extract] Successfully extracted image`, { filePath: args.filePath, identifier: args.identifier, format: args.outputFormat, size: imageBuffer.length });

    // Return raw buffer in SuccessResponse
    return {
        success: true,
        data: imageBuffer // The server loop will wrap this in imageContent
    } as SuccessResponse<Buffer>;

  } catch (error: any) {
    logger.error(`[word/image/extract] Error: ${error.message}`, { error });
    // Re-throw FastMCP specific errors
    if (error instanceof UserError) {
        throw error;
    }
    // Construct standard ErrorResponse for other errors
    return {
        success: false,
        error: {
            code: 'IMAGE_EXTRACTION_FAILED',
            message: error.message || 'An unexpected error occurred during image extraction.',
            details: error, // Include the original error object as details
        }
    } as ErrorResponse;
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
async function handleInsertImage(
  params: ToolRequestParams,
  context?: FastMCPContext<undefined>
): Promise<ApiResponse<string>> { // Return ApiResponse with string confirmation
  const validationResult = InsertImageSchema.safeParse(params);
   if (!validationResult.success) {
    logger.warn(`[word/image/insert] Invalid parameters`, { errors: validationResult.error.format() });
    return {
        success: false,
        error: {
            code: 'VALIDATION_ERROR',
            message: 'Invalid parameters provided.',
            details: validationResult.error.format(),
        }
    } as ErrorResponse;
  }
  const args = validationResult.data; // Use validated data
  logger.info(`[word/image/insert] Received request`, { filePath: args.filePath, position: args.position }); // Avoid logging full base64

  try {
    // validateFilePath is already called by the schema refinement, but calling again here is harmless
    // await validateFilePath(args.filePath); // Ensure path is safe

    const imageBuffer = Buffer.from(args.imageDataBase64, 'base64');

    if (imageBuffer.length === 0) {
        throw new UserError("Provided image data is empty or invalid base64.");
    }

    // Placeholder: Implement actual COM interop call
    await insertImageIntoWord(args.filePath, imageBuffer, args.position, {
        width: args.width,
        height: args.height,
        altText: args.altText,
    });

    logger.info(`[word/image/insert] Successfully inserted image`, { filePath: args.filePath, position: args.position });

    // Return simple text confirmation in SuccessResponse
    const confirmationMessage = `Image successfully inserted into ${path.basename(args.filePath)} at position '${args.position}'.`;
    return {
        success: true,
        data: confirmationMessage // The server loop will return this string directly
    } as SuccessResponse<string>;

  } catch (error: any) {
    logger.error(`[word/image/insert] Error: ${error.message}`, { error });
    // Re-throw FastMCP specific errors
    if (error instanceof UserError) {
        throw error;
    }
    // Construct standard ErrorResponse for other errors
    return {
        success: false,
        error: {
            code: 'IMAGE_INSERTION_FAILED',
            message: error.message || 'An unexpected error occurred during image insertion.',
            details: error, // Include the original error object as details
        }
    } as ErrorResponse;
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