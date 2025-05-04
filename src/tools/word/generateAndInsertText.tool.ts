/**
 * @file Implements the 'word/generate-and-insert-text' tool using FastMCP sampling and COM Interop.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ApiResponse, FastMCPContext, ToolRequestParams } from '../../types/common.types'; // Normalized relative path
import { handleToolError, createErrorResponse } from '../../utils/errorHandler'; // Normalized relative path
import logger from '../../utils/logger'; // Normalized relative path
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Normalized relative path
import { validateFilePath } from '../../utils/security'; // Normalized relative path
import { UserError } from 'fastmcp'; // Import UserError for sampling errors

// Define the session data type expected by this tool's context
// If authentication is required and provides session data:
// import { AuthSessionData } from '@/server/index'; // Adjust path if needed
// type ToolContextType = FastMCPContext<AuthSessionData>;
// If no specific session data is needed beyond authentication being done:
/**
 * Type definition for the tool's context, allowing any session data.
 */
type ToolContextType = FastMCPContext<any>; // Use 'any' or a more specific type if available

// --- Schema ---
/**
 * Zod schema for the input parameters of the 'word/generate-and-insert-text' tool.
 */
const generateAndInsertSchema = z.object({
  /** The path to the Word file (relative to the current workspace directory). */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  /** The position within the document where the text should be inserted (e.g., "start", "end", "paragraph:N:start", "selection"). */
  position: z.string().min(1, 'Position specifier is required (e.g., "start", "end", "paragraph:N:start", "selection").'),
  /** The prompt to use for generating text with the LLM. */
  prompt: z.string().min(1, 'A prompt for text generation is required.'),
  // Optional sampling parameters (add more as needed from FastMCP spec)
  /** Maximum tokens for the generated text. */
  maxTokens: z.number().int().positive().optional().describe("Maximum tokens for the generated text."),
});

// --- Handler ---

/**
 * Generates text using LLM sampling via FastMCP context and inserts it into a Word document.
 * @param params - The parameters for the tool, validated against `generateAndInsertSchema`.
 * @param context - The FastMCP context, expected to contain a session for sampling.
 * @returns A promise resolving to an ApiResponse indicating the outcome of the operation.
 * @throws {Error} If validation fails, context/session is unavailable, sampling fails, the document cannot be opened, or insertion fails.
 */
async function generateAndInsertText(
    params: ToolRequestParams, // Change signature to accept ToolRequestParams
    context?: ToolContextType // Use the defined context type
): Promise<ApiResponse<{}>> {
    // Ensure context and session are available, as requestSampling requires it
    if (!context || !context.session || typeof context.session.requestSampling !== 'function') {
        logger.error('[word/generate-and-insert-text] FastMCP session or sampling function is unavailable.');
        return createErrorResponse(
            'LLM sampling is required for this tool but the FastMCP session or sampling function is not available. ' +
            'Please ensure the Office MCP server is configured with a valid LLM API key.',
            'LLM_SAMPLING_UNAVAILABLE'
        );
    }

    let officeAppInstance: any = null;
    let wordApp: any = null;
    let doc: any = null;
    let insertionRange: any = null;
    let paraRange: any = null; // For paragraph logic

    // Parse and validate params inside the function
    let validatedParams: z.infer<typeof generateAndInsertSchema>;
    try {
        validatedParams = generateAndInsertSchema.parse(params);
    } catch (error) {
         logger.error(`[word/generate-and-insert-text] Input validation failed: ${error instanceof Error ? error.message : String(error)}`, { params });
         // Use handleToolError which checks for ZodError
         return handleToolError(error, 'VALIDATION_ERROR');
    }

    logger.info(`Executing word/generate-and-insert-text for file: ${validatedParams.filePath}, position: ${validatedParams.position}`);

    try {
        // 1. Request text generation from the LLM via FastMCP context.session
        logger.debug(`Requesting sampling with prompt: "${validatedParams.prompt}"`);
        // Call requestSampling on context.session
        const samplingResult = await context.session.requestSampling({
            // Construct messages array if needed, or just use prompt directly if supported
            // messages: [{ role: 'user', content: { type: 'text', text: validatedParams.prompt } }],
            prompt: validatedParams.prompt, // Assuming direct prompt usage is supported
            maxTokens: validatedParams.maxTokens, // Pass optional parameters
            // includeContext: 'thisServer', // Optionally include server context if needed by LLM
        });

        // Check if sampling was successful and extract text
        // FastMCP's requestSampling likely throws on error or returns a specific structure.
        // Adjust based on actual FastMCP behavior. Assuming it returns content or throws.
        // Let's assume the primary text result is in the first content block.
        const generatedText = samplingResult?.content?.[0]?.text;

        if (!generatedText) {
            logger.warn('Sampling request did not return usable text content.', { result: samplingResult });
            return createErrorResponse('Failed to generate text: No content returned from sampling.', 'SAMPLING_FAILED');
        }
        logger.info(`Successfully generated text via sampling (length: ${generatedText.length})`);

        // 2. Open Word Document (Read/Write)
        // Use validatedParams.filePath
        officeAppInstance = await getOfficeApplication('Word.Application');
        wordApp = officeAppInstance.app;
        // Open the document in read/write mode (ReadOnly = false) and visible (Visible = true)
        doc = officeAppInstance.openDocument(validatedParams.filePath, false, true); // Open read/write, Visible = true
        if (!doc) {
            logger.error(`[word/generate-and-insert-text] Failed to open document: ${validatedParams.filePath}`);
            return createErrorResponse(`Failed to open document: ${validatedParams.filePath}`, 'FILE_OPEN_FAILED');
        }
        logger.debug(`Document opened successfully: ${validatedParams.filePath}`);


        // 3. Determine Insertion Range (similar logic to word/text/insert)
        const positionLower = validatedParams.position.toLowerCase(); // Use validatedParams
        logger.debug(`Determining insertion range for position: ${positionLower}`);

        // Simplified range logic - adapt from word/text/insert if needed
        if (positionLower === 'start') {
            insertionRange = doc.Range(0, 0);
            logger.debug('Insertion position set to start of document.');
        } else if (positionLower === 'end') {
            const endPos = doc.Content.End;
            insertionRange = doc.Range(endPos, endPos);
            logger.debug('Insertion position set to end of document.');
        } else if (positionLower === 'selection') {
             if (!wordApp.Selection) {
                 logger.warn('Cannot insert at selection: No selection found.');
                 return createErrorResponse("Cannot insert at selection: No selection found.", 'NO_SELECTION');
             }
             insertionRange = wordApp.Selection.Range;
             // Collapse if it's not an insertion point
             if (wordApp.Selection.Type !== 2 /* wdSelectionIP */) {
                  insertionRange.Collapse(1); // wdCollapseStart
                  logger.debug('Insertion position set to start of current selection.');
             } else {
                 logger.debug('Insertion position set to current insertion point.');
             }
        } else if (positionLower.startsWith('paragraph:')) {
            // Simplified - inserts at the start of the paragraph
            const parts = positionLower.split(':');
            const indexStr = parts[1];
            const paraIndex = parseInt(indexStr, 10);
            if (isNaN(paraIndex) || paraIndex <= 0 || paraIndex > doc.Paragraphs.Count) {
                logger.warn(`Invalid or out-of-bounds paragraph index: ${indexStr}`);
                return createErrorResponse(`Invalid or out-of-bounds paragraph index: ${indexStr}`, 'INVALID_PARAM');
            }
            paraRange = doc.Paragraphs(paraIndex).Range;
            insertionRange = doc.Range(paraRange.Start, paraRange.Start); // Insert at start
            logger.debug(`Insertion position set to start of paragraph ${paraIndex}.`);
        } else {
            logger.warn(`Unsupported position specifier: ${validatedParams.position}`);
            return createErrorResponse(`Unsupported position specifier: ${validatedParams.position}`, 'INVALID_POSITION'); // Use validatedParams
        }

        if (!insertionRange) {
             logger.error(`[word/generate-and-insert-text] Could not determine insertion range for position: ${validatedParams.position}`);
             return createErrorResponse(`Could not determine insertion range for position: ${validatedParams.position}`, 'RANGE_ERROR'); // Use validatedParams
        }

        // 4. Insert Generated Text
        logger.debug(`Inserting generated text (length: ${generatedText.length})`);
        insertionRange.Text = generatedText;
        logger.debug('Generated text inserted into document.');

        // 5. Save and Close
        doc.Save();
        logger.info(`Successfully inserted generated text at position "${validatedParams.position}" and saved ${validatedParams.filePath}`); // Use validatedParams
        return { success: true, data: {} };

    } catch (error: any) {
        logger.error(`Error in word/generate-and-insert-text: ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
         if (error instanceof z.ZodError) {
            return createErrorResponse('Input validation failed', 'VALIDATION_ERROR', error.errors);
        }
        // Handle potential UserError from requestSampling or other errors
        return handleToolError(error, 'TOOL_EXECUTION_ERROR'); // Use generic handler
    } finally {
        // 6. Release COM Objects
        releaseObject(insertionRange);
        releaseObject(paraRange);
        if (doc) {
            try { doc.Close(false); } catch (e: any) { logger.warn(`Error closing document ${validatedParams?.filePath || 'unknown'}: ${e.message}`); } // Log error message
            releaseObject(doc);
        }
        if (officeAppInstance) {
            officeAppInstance.release();
        }
        logger.debug('generateAndInsertText finished, COM objects released.');
    }
}

// --- Resource Definition ---
/**
 * McpResource definition for the 'word/generate-and-insert-text' tool.
 * Generates text based on a prompt using the LLM and inserts it into a Word document at a specified position.
 */
export const wordGenerateAndInsertTextTool: McpResource = { // Export as single object, not array
    path: 'word/generate-and-insert-text',
    handler: generateAndInsertText, // Pass the handler function directly
    schema: generateAndInsertSchema,
    description: 'Generates text based on a prompt using the LLM and inserts it into a Word document at a specified position.',
    // Add annotations if desired
    // annotations: {
    //     openWorldHint: true, // Interacts with LLM
    // },
};

// No default export needed if importing the named export directly in tools/index.ts