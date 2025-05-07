import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, ErrorResponse, SuccessResponse } from '../../types/common.types.js';
import { Context as FastMCPContext } from 'fastmcp'; // Removed imageContent import as it's not a type
import logger from '../../utils/logger.js';
import { validateFilePath } from '../../utils/security.js';
import { extractImageFromWord, insertImageIntoWord } from '../../utils/officeInterop.js'; // Import office interop functions
import { generateText, analyzeImageWithText, generateImageFromText } from '../../utils/llmClient.js'; // Import new LLM client functions
import axios from 'axios'; // Import axios for fetching from URLs
import fs from 'fs/promises'; // Import fs for reading local files

// --- Schema Definitions ---

const ImageSourceSchema = z.union([
    z.object({ type: z.literal('document'), documentPath: z.string().min(1), imageIdentifier: z.union([z.number().int().positive(), z.string().min(1)]) }),
    z.object({ type: z.literal('dynamic_resource'), dynamicResourceId: z.string().min(1) }),
    z.object({ type: z.literal('filepath'), filePath: z.string().min(1).refine(validateFilePath) }),
    z.object({ type: z.literal('url'), url: z.string().url() }),
    z.object({ type: z.literal('previous'), previousImageIdentifier: z.string().min(1) }), // How to identify previous? Need to define
]);

const ImageAnalysisParamsSchema = z.object({
    imageSource: ImageSourceSchema,
    analysisPrompt: z.string().min(1),
});

const ImageGenerationParamsSchema = z.object({
    generationPrompt: z.string().min(1),
    targetDocumentPath: z.string().min(1).refine(validateFilePath),
    insertionPosition: z.string().min(1), // e.g., 'end', 'paragraph:N', 'bookmark:name'
});

const ImageAnalysisGenerationInputSchema = z.object({
    operation: z.enum(['analyze', 'generate']),
    analysis: ImageAnalysisParamsSchema.optional(),
    generation: ImageGenerationParamsSchema.optional(),
}).refine(data => {
    // Ensure that if operation is 'analyze', analysis params are provided, and vice versa
    if (data.operation === 'analyze' && !data.analysis) return false;
    if (data.operation === 'generate' && !data.generation) return false;
    return true;
}, {
    message: "Analysis parameters are required for 'analyze' operation, and generation parameters for 'generate'.",
});

type ImageAnalysisGenerationInput = z.infer<typeof ImageAnalysisGenerationInputSchema>;

// --- Tool Handler ---

async function handleImageAnalysisGeneration(
    params: ToolRequestParams,
    context?: FastMCPContext<undefined>
): Promise<ApiResponse<string>> { // Return type is string (analysis result or insertion confirmation)
    const validationResult = ImageAnalysisGenerationInputSchema.safeParse(params);
    if (!validationResult.success) {
        logger.warn(`[office/imageAnalysisGeneration] Invalid parameters`, { errors: validationResult.error.format() });
        return {
            success: false,
            error: {
                code: 'VALIDATION_ERROR',
                message: 'Invalid parameters provided.',
                details: validationResult.error.format(),
            }
        } as ErrorResponse;
    }
    const args = validationResult.data;
    logger.info(`[office/imageAnalysisGeneration] Received request`, { operation: args.operation });

    try {
        if (args.operation === 'analyze') {
            const analysisParams = args.analysis!;
            let imageDataBase64: string | undefined;

            // Step 2: Access/Extract Image Data
            if (analysisParams.imageSource.type === 'document') {
                const { documentPath, imageIdentifier } = analysisParams.imageSource;

                if (typeof imageIdentifier === 'number') {
                    // Identifier is an index, extract directly
                    logger.info(`Attempting to extract image by index: ${imageIdentifier}`);
                    try {
                        const imageBuffer = await extractImageFromWord(documentPath, imageIdentifier);
                        if (!imageBuffer || imageBuffer.length === 0) {
                            throw new Error(`Image index ${imageIdentifier} not found or extraction returned empty buffer.`);
                        }
                        imageDataBase64 = imageBuffer.toString('base64');
                    } catch (extractError: any) {
                        throw new Error(`Failed to extract image by index ${imageIdentifier}: ${extractError.message}`);
                    }
                } else {
                    // Identifier is a description, need to iterate and ask LLM
                    logger.info(`Attempting to identify image by description: "${imageIdentifier}"`);
                    // Placeholder: Need a way to get image count, e.g., from officeInterop
                    let imageCount: number;
                    // ** Placeholder: Replace with actual image count retrieval **
                    // Example: imageCount = await getImageCount(documentPath);
                    imageCount = 5; // Using placeholder value for now, but typed as number
                    if (imageCount === 0) {
                        throw new Error(`No images found in document ${documentPath}.`);
                    }

                    let foundImage = false;
                    for (let i = 1; i <= imageCount; i++) {
                        logger.debug(`Checking image index ${i} for description match...`);
                        let imageBuffer: Buffer | null = null;
                        try {
                            imageBuffer = await extractImageFromWord(documentPath, i);
                        } catch (loopExtractError: any) {
                            logger.warn(`Could not extract image index ${i} during description search: ${loopExtractError.message}`);
                            continue; // Skip if extraction fails for this index
                        }

                        if (imageBuffer && imageBuffer.length > 0) {
                            const tempBase64 = imageBuffer.toString('base64');
                            const matchPrompt = `Does the following image match this description: "${imageIdentifier}"? Please answer only with "Yes" or "No".`;
                            try {
                                // Use analyzeImageWithText to ask the LLM for a match
                                const matchResponse = await analyzeImageWithText(matchPrompt, tempBase64, 'image/png'); // Assuming PNG
                                logger.debug(`LLM match response for index ${i}: ${matchResponse}`);
                                if (matchResponse.trim().toLowerCase().startsWith('yes')) {
                                    logger.info(`Found matching image at index ${i} for description: "${imageIdentifier}"`);
                                    imageDataBase64 = tempBase64;
                                    foundImage = true;
                                    break; // Exit loop once match is found
                                }
                            } catch (llmError: any) {
                                logger.warn(`LLM check failed for image index ${i}: ${llmError.message}`);
                                // Decide whether to continue or stop if LLM fails repeatedly
                            }
                        }
                    }

                    if (!foundImage) {
                        throw new Error(`Could not find an image matching the description: "${imageIdentifier}"`);
                    }
                }

            } else if (analysisParams.imageSource.type === 'dynamic_resource') {
                 // Need to call the dynamic/resources tool handler directly or use its underlying logic
                 // Assuming dynamic/resources read handler returns ApiResponse<string> (base64)
                 // Placeholder for calling dynamic/resources read handler:
                 // const readResult: ApiResponse<string> = await dynamicResourcesReadHandler({ resourceId: analysisParams.imageSource.dynamicResourceId });
                 // if (!readResult.success) {
                 //     throw new Error(`Failed to read dynamic resource: ${readResult.error.error.message}`);
                 // }
                 // imageDataBase64 = readResult.data;
                 throw new Error("Reading image from dynamic resource is not yet implemented in this tool.");

            } else if (analysisParams.imageSource.type === 'filepath') {
                // Read file content as binary and convert to base64
                try {
                    const fileContent = await fs.readFile(analysisParams.imageSource.filePath);
                    imageDataBase64 = fileContent.toString('base64');
                } catch (fileError: any) {
                    throw new Error(`Failed to read image from local filepath: ${fileError.message}`);
                }

            } else if (analysisParams.imageSource.type === 'url') {
                // Fetch image from URL as arraybuffer and convert to base64
                try {
                    const response = await axios.get(analysisParams.imageSource.url, { responseType: 'arraybuffer' });
                    imageDataBase64 = Buffer.from(response.data, 'binary').toString('base64');
                } catch (urlError: any) {
                    throw new Error(`Failed to read image from URL: ${urlError.message}`);
                }

            } else if (analysisParams.imageSource.type === 'previous') {
                // Need a mechanism to store and retrieve previous images
                 throw new Error("Referring to previous images is not yet implemented.");
            }

            if (!imageDataBase64) {
                 throw new Error("Could not obtain image data for analysis.");
            }

            // Step 3: Attempt Image Analysis with Master AI
            // This is where the master AI's logic would go.
            // For now, we'll directly use the connected LLM as a placeholder for the fallback.
            logger.info("[office/imageAnalysisGeneration] Attempting analysis with connected LLM (as master AI placeholder)");
            // Assuming generateText can be extended or a new function created for multimodal input
            // and it returns a string analysis result.
            // The generateText function currently only takes prompt and options.
            // It needs modification to accept image data.
            // For now, calling it with a prompt that includes base64 data as a placeholder.
            // This will likely require changes in llmClient.ts and the actual LLM call format.
            // Use the new analyzeImageWithText function
            const analysisResult = await analyzeImageWithText(
                analysisParams.analysisPrompt,
                imageDataBase64,
                'image/png', // Assuming PNG, might need to determine actual type
                {} // Options
            );

            // Step 5: Present Analysis Result
            return { success: true, data: analysisResult } as SuccessResponse<string>;

        } else if (args.operation === 'generate') {
            const generationParams = args.generation!;

            // Step 1: Attempt Image Generation with Master AI
            // This is where the master AI's generation logic would go.
            // For now, we'll directly use the connected LLM as a placeholder for the fallback.
            logger.info("[office/imageAnalysisGeneration] Attempting generation with connected LLM (as master AI placeholder)");
            // Assuming generateText can be extended or a new function created for image generation
            // and it returns base64 image data.
            // The generateText function currently only returns text.
            // It needs modification to handle image generation and return image data.
            // For now, calling it with a prompt as a placeholder.
            // Use the new generateImageFromText function
            const generatedImageDataBase64 = await generateImageFromText(
                generationParams.generationPrompt,
                {} // Options
            );

            if (!generatedImageDataBase64) {
                 throw new Error("Image generation failed.");
            }

            // Step 4: Identify Target Document and Location (already done by schema)

            // Step 5: Insert Image into Document
            // The insertImageIntoWord function should return void or throw an error
            await insertImageIntoWord(
                generationParams.targetDocumentPath,
                Buffer.from(generatedImageDataBase64, 'base64'), // insertImageIntoWord expects Buffer
                generationParams.insertionPosition,
                {} // Options (width, height, altText)
            );

            // Step 6: Confirm Insertion
            return { success: true, data: "Image successfully inserted into document." } as SuccessResponse<string>;
        }

        // Should not reach here due to schema validation
        return { success: false, error: { code: 'INVALID_OPERATION', message: 'Invalid operation specified.' } };

    } catch (error: any) {
        logger.error(`[office/imageAnalysisGeneration] Error: ${error.message}`, { error });
        // Handle errors and indicate incapability if necessary
        // For now, just return a generic error response
        return {
            success: false,
            error: {
                code: 'PROCESSING_ERROR',
                message: error.message || 'An unexpected error occurred during image processing.',
                details: error,
            }
        } as ErrorResponse;
    }
}

// --- Resource Definition ---

export const officeImageAnalysisGenerationTool: McpResource = {
    path: 'office/imageAnalysisGeneration',
    description: 'Analyzes images from Office documents or other sources and inserts generated images into Office documents.',
    schema: ImageAnalysisGenerationInputSchema,
    handler: handleImageAnalysisGeneration,
};

// No default export needed if importing the named export directly in tools/index.ts