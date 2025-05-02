// /src/tools/word/generateAndInsertText.tool.ts
// =============================================================================
/**
 * @file Implements the 'word/generate-and-insert-text' tool using FastMCP sampling and COM Interop.
 */
import { z } from 'zod';
// Import ToolRequestParams as well
import { McpResource, ApiResponse, FastMCPContext, ToolRequestParams } from '@/types/common.types';
import { handleToolError, createErrorResponse } from '@/utils/errorHandler';
import logger from '@/utils/logger';
import { getOfficeApplication, releaseObject } from '@/utils/officeInterop'; // Assuming insertText logic might be here or called
import { validateFilePath } from '@/utils/security';
import { UserError } from 'fastmcp'; // Import UserError for sampling errors

// Define the session data type expected by this tool's context
// If authentication is required and provides session data:
// import { AuthSessionData } from '@/server/index'; // Adjust path if needed
// type ToolContextType = FastMCPContext<AuthSessionData>;
// If no specific session data is needed beyond authentication being done:
type ToolContextType = FastMCPContext<any>; // Use 'any' or a more specific type if available

// --- Schema ---
const generateAndInsertSchema = z.object({
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  position: z.string().min(1, 'Position specifier is required (e.g., "start", "end", "paragraph:N:start", "selection").'),
  prompt: z.string().min(1, 'A prompt for text generation is required.'),
  // Optional sampling parameters (add more as needed from FastMCP spec)
  maxTokens: z.number().int().positive().optional().describe("Maximum tokens for the generated text."),
});

// --- Handler ---

/**
 * Generates text using LLM sampling via FastMCP context and inserts it into a Word document.
 */
async function generateAndInsertText(
    params: ToolRequestParams, // Change signature to accept ToolRequestParams
    context?: ToolContextType // Use the defined context type
): Promise<ApiResponse<{}>> {
    // Ensure context and session are available, as requestSampling requires it
    if (!context || !context.session) { // Check for context.session as well
        return createErrorResponse('Context and session are required for requesting sampling.', 'CONTEXT_UNAVAILABLE');
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
            maxTokens: validatedParams.maxTokens, // Pass optional parameters (Corrected: removed duplicate)
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
        doc = officeAppInstance.openDocument(validatedParams.filePath, false, false); // Open read/write
        if (!doc) {
            return createErrorResponse(`Failed to open document: ${validatedParams.filePath}`, 'FILE_OPEN_FAILED');
        }

        // 3. Determine Insertion Range (similar logic to word/text/insert)
        const positionLower = validatedParams.position.toLowerCase(); // Use validatedParams
        logger.debug(`Determining insertion range for position: ${positionLower}`);

        // Simplified range logic - adapt from word/text/insert if needed
        if (positionLower === 'start') {
            insertionRange = doc.Range(0, 0);
        } else if (positionLower === 'end') {
            const endPos = doc.Content.End;
            insertionRange = doc.Range(endPos, endPos);
        } else if (positionLower === 'selection') {
             if (!wordApp.Selection) return createErrorResponse("Cannot insert at selection: No selection found.", 'NO_SELECTION');
             insertionRange = wordApp.Selection.Range;
             // Collapse if it's not an insertion point
             if (wordApp.Selection.Type !== 2 /* wdSelectionIP */) {
                  insertionRange.Collapse(1); // wdCollapseStart
             }
        } else if (positionLower.startsWith('paragraph:')) {
            // Simplified - inserts at the start of the paragraph
            const parts = positionLower.split(':');
            const indexStr = parts[1];
            const paraIndex = parseInt(indexStr, 10);
            if (isNaN(paraIndex) || paraIndex <= 0 || paraIndex > doc.Paragraphs.Count) {
                return createErrorResponse(`Invalid or out-of-bounds paragraph index: ${indexStr}`, 'INVALID_PARAM');
            }
            paraRange = doc.Paragraphs(paraIndex).Range;
            insertionRange = doc.Range(paraRange.Start, paraRange.Start); // Insert at start
        } else {
            return createErrorResponse(`Unsupported position specifier: ${validatedParams.position}`, 'INVALID_POSITION'); // Use validatedParams
        }

        if (!insertionRange) {
             return createErrorResponse(`Could not determine insertion range for position: ${validatedParams.position}`, 'RANGE_ERROR'); // Use validatedParams
        }

        // 4. Insert Generated Text
        logger.debug(`Inserting generated text (length: ${generatedText.length})`);
        insertionRange.Text = generatedText;

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
            try { doc.Close(false); } catch (e) { logger.warn('Error closing document after generate/insert', e); }
            releaseObject(doc);
        }
        if (officeAppInstance) {
            officeAppInstance.release();
        }
        logger.debug('generateAndInsertText finished, COM objects released.');
    }
}

// --- Resource Definition ---
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