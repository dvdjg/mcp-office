// /src/tools/word/styles.tool.ts
// =============================================================================
/**
 * @file Implements the 'word/styles' tool using Office JavaScript APIs.
 * NOTE: This implementation assumes execution within an Office Add-in context
 * or a similar environment where the Office JS APIs are available and initialized.
 * Running this directly from a standalone Node.js server requires bridging
 * (e.g., using COM Interop on Windows, or potentially Office Scripts via Graph API).
 * This example shows the Office JS API logic.
 */
import { z } from 'zod';
import { McpResource, ApiResponse, ToolContext, ToolRequestParams } from '@/types/common.types';
import { createErrorResponse, handleToolError } from '@/utils/errorHandler';
import logger from '@/utils/logger';

// --- Schemas ---
const styleSchema = z.object({
    documentReference: z.string().optional(), // How to reference the target document (e.g., active document, file path) - Needs clarification based on runtime context
    style: z.string().min(1),
    range: z.string().min(1), // e.g., 'paragraph:1', 'selection', 'range:A1:B2' (needs parsing)
});

const listStyleSchema = z.object({
     documentReference: z.string().optional(),
});

// --- Placeholder for Office JS Execution Context ---
// This function simulates how you might get the Word context.
// In a real Add-in, this is provided by the Office environment.
// In a server scenario, this needs complex bridging.
export async function getWordContext(documentReference?: string): Promise<Word.RequestContext> {
    // --- !! MAJOR CAVEAT !! ---
    // This is a placeholder. Getting a Word.RequestContext outside an Add-in
    // is non-trivial. It might involve:
    // 1. Office Scripts + Microsoft Graph API (Cloud-based, requires M365)
    // 2. COM Interop (Windows Desktop only, requires libraries like node-win32ole)
    // 3. Driving a hidden Word instance via command line/scripting (Less reliable)
    //
    // This mock throws an error to indicate it needs proper implementation.
    if (typeof Word === 'undefined' || !Word.run) {
         logger.error("Word JS API context is not available in this environment.");
         throw new Error("Word JS API context is not available in this environment. Requires Add-in context or server-side bridging (COM/Graph API).");
    }
    // Example of how it *would* look in an Add-in:
    // return Word.run(async context => context);

    // Placeholder throwing error:
    throw new Error("Word context simulation not implemented.");
}

function parseRange(rangeString: string, context: Word.RequestContext): Word.Range {
    // Basic range parser - Needs robust implementation
    // Example: 'paragraph:3', 'selection'
    if (rangeString.toLowerCase() === 'selection') {
        return context.document.getSelection();
    }
    if (rangeString.startsWith('paragraph:')) {
        const index = parseInt(rangeString.split(':')[1], 10);
        if (!isNaN(index) && index >= 0) {
             // Note: Paragraph indices might be 0-based or 1-based depending on API/context. Adjust as needed.
             // This example assumes 0-based access if possible, but Office JS often uses item collections.
             // A more robust way might involve getting all paragraphs and selecting by index.
             // return context.document.body.paragraphs.getItemAt(index); // This is conceptual
             logger.warn(`Paragraph range parsing ('${rangeString}') is simplified.`);
             // Fallback to body for this example
             return context.document.body;
        }
    }
    // Default/fallback or throw error
    logger.warn(`Unsupported range format: ${rangeString}. Falling back to document body.`);
    return context.document.body; // Fallback
}


// --- Handlers ---

/**
 * Applies a style to a specified range in a Word document.
 */
async function applyStyle(params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<{}>> {
    try {
        const validatedParams = styleSchema.parse(params);
        logger.info(`Attempting to apply style: ${validatedParams.style} to range: ${validatedParams.range}`);

        // --- Context Acquisition (Placeholder) ---
        const wordContext = await getWordContext(validatedParams.documentReference);
        // --- End Context Acquisition ---

        const targetRange = parseRange(validatedParams.range, wordContext);
        targetRange.load('style'); // Load the current style for potential comparison/logging
        await wordContext.sync();

        const currentStyle = targetRange.style;
        logger.debug(`Current style for range '${validatedParams.range}': ${currentStyle}`);

        targetRange.style = validatedParams.style;
        await wordContext.sync();

        logger.info(`Successfully applied style '${validatedParams.style}' to range '${validatedParams.range}'`);
        return { success: true, data: {} };

    } catch (error) {
         logger.error(`Error applying style: ${error}`, { params });
         // Ensure OfficeExtension.Error details are captured if available
         if (error instanceof Error && error.name === 'OfficeExtension.Error') {
             return handleToolError(error, 'OFFICE_API_ERROR');
         }
        return handleToolError(error, 'WORD_STYLE_ERROR');
    }
}

/**
 * Lists available styles in the document.
 */
async function listStyles(params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<string[]>> {
     try {
        const validatedParams = listStyleSchema.parse(params);
        logger.info(`Attempting to list styles for document: ${validatedParams.documentReference || 'active'}`);

        // --- Context Acquisition (Placeholder) ---
        const wordContext = await getWordContext(validatedParams.documentReference);
        // --- End Context Acquisition ---

        const styles = wordContext.document.styles;
        styles.load('items/nameLocal'); // Load only the names
        await wordContext.sync();

        const styleNames = styles.items.map(style => style.nameLocal);

        logger.info(`Found ${styleNames.length} styles.`);
        return { success: true, data: styleNames };

    } catch (error) {
         logger.error(`Error listing styles: ${error}`, { params });
         if (error instanceof Error && error.name === 'OfficeExtension.Error') {
             return handleToolError(error, 'OFFICE_API_ERROR');
         }
        return handleToolError(error, 'WORD_STYLE_ERROR');
    }
}


// --- Resource Definition ---
export const wordStylesTool: McpResource[] = [
    {
        path: 'word/styles/apply',
        handler: applyStyle,
        schema: styleSchema,
        description: 'Applies a named style to a specified range (e.g., paragraph, selection) in a Word document. Requires Office JS context.',
        completions: async () => ({ style: ['Normal', 'Heading 1', 'Heading 2', 'Title'], range: ['selection', 'paragraph:1'] }),
    },
     {
        path: 'word/styles/list',
        handler: listStyles,
        schema: listStyleSchema,
        description: 'Lists the names of available styles in a Word document. Requires Office JS context.',
    },
    // Add create, modify, delete operations here following similar patterns
];