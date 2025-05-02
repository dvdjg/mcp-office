// /src/tools/word/image.tool.ts
// =============================================================================
/**
 * @file Defines tools for handling images within Word documents.
 */
import { z } from 'zod';
// Import necessary types from common.types and fastmcp
import { McpResource, ApiResponse, ErrorResponse, SuccessResponse, ToolRequestParams } from '@/types/common.types';
import { imageContent, TextContent, UserError, Context as FastMCPContext } from 'fastmcp'; // Import TextContent and FastMCPContext
import { extractImageFromWord, insertImageIntoWord } from '@/utils/officeInterop'; // Placeholder for actual interop functions
import { validateFilePath } from '@/utils/security'; // Assuming this utility exists for path validation
import logger from '@/utils/logger';
import path from 'path';
import fs from 'fs/promises'; // For reading image data for insertion if needed

// --- Schema Definitions ---

const ExtractImageSchema = z.object({
  filePath: z.string().min(1, "File path cannot be empty."),
  identifier: z.union([z.number().int().positive("Image index must be a positive integer."), z.string().min(1, "Image identifier text cannot be empty.")], {
    description: "Identifier for the image (e.g., 1-based index or placeholder text/bookmark).",
  }),
  outputFormat: z.enum(['png', 'jpeg', 'gif', 'bmp']).default('png').describe("Desired output format for the extracted image."),
});

const InsertImageSchema = z.object({
  filePath: z.string().min(1, "File path cannot be empty."),
  // Input image data could be provided directly (base64) or via a URI (future enhancement)
  imageDataBase64: z.string().min(1, "Image data (base64) cannot be empty."),
  position: z.string().min(1, "Insertion position cannot be empty (e.g., 'end', 'bookmark:name', 'paragraph:N')."),
  // Optional parameters for image sizing, etc.
  width: z.number().optional().describe("Optional width for the inserted image in points."),
  height: z.number().optional().describe("Optional height for the inserted image in points."),
  altText: z.string().optional().describe("Optional alternative text for the image."),
});

// --- Tool Handlers ---

// Adjust handler signature and perform validation inside
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
    await validateFilePath(args.filePath); // Ensure path is safe

    // Placeholder: Implement actual COM interop call
    const imageBuffer = await extractImageFromWord(args.filePath, args.identifier, args.outputFormat);

    if (!imageBuffer || imageBuffer.length === 0) {
      throw new UserError(`Image with identifier '${args.identifier}' not found or could not be extracted.`);
    }

    // const mimeType = `image/${args.outputFormat}`; // Mime type will be handled by server loop or imageContent wrapper later
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

// Adjust handler signature and perform validation inside
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
    await validateFilePath(args.filePath); // Ensure path is safe

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

export const wordImageTools: McpResource[] = [
  {
    path: 'word/image/extract',
    description: 'Extracts a specific image from a Word document (.docx) based on index or identifier.',
    schema: ExtractImageSchema,
    handler: handleExtractImage,
  },
  {
    path: 'word/image/insert',
    description: 'Inserts an image (from base64 data) into a Word document (.docx) at a specified position.',
    schema: InsertImageSchema,
    handler: handleInsertImage,
  },
  // Add audio tools here later if needed
];