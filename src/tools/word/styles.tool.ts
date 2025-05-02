// /src/tools/word/styles.tool.ts
// =============================================================================
/**
 * @file Implements the 'word/styles' tool using COM Interop (winax).
 */
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '@/types/common.types'; // Import FastMCPContext, remove ToolContext
import { handleToolError } from '@/utils/errorHandler';
import logger from '@/utils/logger';
import { getOfficeApplication, releaseObject } from '@/utils/officeInterop';
import { validateFilePath } from '@/utils/security';

// --- Schemas ---
const styleSchema = z.object({
    filePath: z.string().min(1).refine(validateFilePath, {
        message: "Invalid or potentially unsafe file path provided.",
    }),
    style: z.string().min(1),
    range: z.string().min(1), // e.g., 'paragraph:1', 'selection', 'document'
});

const listStyleSchema = z.object({
    filePath: z.string().min(1).refine(validateFilePath, {
        message: "Invalid or potentially unsafe file path provided.",
    }),
});


// --- Handlers ---

/**
 * Applies a style to a specified range in a Word document using COM Interop.
 */
export async function applyStyle(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> { // Use FastMCPContext<undefined>
    let wordApp: any = null;
    let doc: any = null;

    try {
        const validatedParams = styleSchema.parse(params);
        const safeFilePath = validatedParams.filePath; // Already validated by Zod refine
        logger.info(`Attempting to apply style: ${validatedParams.style} to range: ${validatedParams.range} in document: ${safeFilePath}`);

        wordApp = await getOfficeApplication('Word.Application');
        // Consider making Word visible for debugging: wordApp.Visible = true;

        doc = wordApp.Documents.Open(safeFilePath);
        if (!doc) {
            throw new Error(`Failed to open document: ${safeFilePath}`);
        }

        let selectedRange: any;
        const rangeStringLower = validatedParams.range.toLowerCase();

        if (rangeStringLower === 'selection') {
            // Note: 'selection' might be tricky if Word isn't visible or doesn't have focus.
            // It refers to the current selection in the Word UI.
            // If running headless, this might not be what the user expects.
            // Consider if 'document' or specific paragraphs are more reliable.
            selectedRange = wordApp.Selection.Range; // Get the Range object from the Selection
             if (!selectedRange) {
                 logger.warn("wordApp.Selection.Range was null or undefined. This might happen if there's no active selection or the app is headless.");
                 throw new Error("Could not get range from selection. Ensure the document is active and has a selection, or use a different range specifier.");
             }
        } else if (rangeStringLower === 'document') {
            selectedRange = doc.Content;
        } else if (rangeStringLower.startsWith('paragraph:')) {
            const indexStr = validatedParams.range.split(':')[1];
            const index = parseInt(indexStr, 10);
            if (!isNaN(index) && index > 0) { // COM indices are typically 1-based
                 if (index <= doc.Paragraphs.Count) {
                    selectedRange = doc.Paragraphs(index).Range;
                 } else {
                    throw new Error(`Paragraph index ${index} is out of bounds. Document has ${doc.Paragraphs.Count} paragraphs.`);
                 }
            } else {
                 throw new Error(`Invalid paragraph index format: '${indexStr}'. Use 'paragraph:N' where N is a positive integer.`);
            }
        } else {
            throw new Error(`Unsupported range format: '${validatedParams.range}'. Supported formats: 'selection', 'document', 'paragraph:N'.`);
        }

        // Apply the style
        logger.debug(`Applying style '${validatedParams.style}' to range type: ${rangeStringLower}`);
        selectedRange.Style = validatedParams.style;
        // Optionally save the document: doc.Save();
        // For this tool, we typically don't save automatically.

        logger.info(`Successfully applied style '${validatedParams.style}' to range '${validatedParams.range}' in document '${safeFilePath}'`);
        return { success: true, data: {} };

    } catch (error) {
         logger.error(`Error applying style via COM: ${error}`, { params });
         return handleToolError(error, 'WORD_STYLE_ERROR');
    } finally {
        // --- CRUCIAL: Release COM Objects ---
        if (doc) {
            try {
                doc.Close(false); // Close without saving changes (wdDoNotSaveChanges = 0)
                logger.debug(`Closed document: ${params.filePath}`);
            } catch (closeError) {
                logger.error(`Error closing document: ${closeError}`);
            }
            releaseObject(doc);
            doc = null;
        }
        if (wordApp) {
            // Only quit the application if we opened it and no other docs are open (tricky to determine reliably without more complex logic)
            // For simplicity now, we might leave Word running if it was already open.
            // A safer approach for background tasks might be to always quit.
            // wordApp.Quit(); // Consider the implications
            releaseObject(wordApp);
            wordApp = null;
            logger.debug("Released Word Application COM object.");
        }
    }
}

/**
 * Lists available styles in the document using COM Interop.
 */
export async function listStyles(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<string[]>> { // Use FastMCPContext<undefined>
    let wordApp: any = null;
    let doc: any = null;
     try {
        const validatedParams = listStyleSchema.parse(params);
        const safeFilePath = validatedParams.filePath; // Already validated
        logger.info(`Attempting to list styles for document: ${safeFilePath}`);

        wordApp = await getOfficeApplication('Word.Application');
        doc = wordApp.Documents.Open(safeFilePath);
        if (!doc) {
            throw new Error(`Failed to open document: ${safeFilePath}`);
        }

        const styles = doc.Styles;
        const styleCount = styles.Count;
        const styleNames: string[] = [];

        logger.debug(`Found ${styleCount} styles in the collection. Iterating...`);
        // COM collections are often 1-based
        for (let i = 1; i <= styleCount; i++) {
            let style = null;
            try {
                 style = styles(i); // Access item by 1-based index
                 if (style && style.NameLocal) {
                     styleNames.push(style.NameLocal);
                 } else {
                      logger.warn(`Style at index ${i} was null or had no NameLocal.`);
                 }
            } catch (itemError) {
                 logger.error(`Error accessing style at index ${i}: ${itemError}`);
                 // Continue to next item if possible
            } finally {
                 if (style) releaseObject(style); // Release the individual style object
            }
        }

        logger.info(`Successfully listed ${styleNames.length} styles from document '${safeFilePath}'.`);
        return { success: true, data: styleNames };

    } catch (error) {
         logger.error(`Error listing styles via COM: ${error}`, { params });
        return handleToolError(error, 'WORD_STYLE_ERROR');
    } finally {
        // --- CRUCIAL: Release COM Objects ---
        if (doc) {
            try {
                doc.Close(false); // Close without saving
                logger.debug(`Closed document: ${params.filePath}`);
            } catch (closeError) {
                logger.error(`Error closing document: ${closeError}`);
            }
            releaseObject(doc);
            doc = null;
        }
        if (wordApp) {
            releaseObject(wordApp);
            wordApp = null;
            logger.debug("Released Word Application COM object.");
        }
    }
}


// --- Resource Definition ---
export const wordStylesTool: McpResource[] = [
    {
        path: 'word/styles/apply',
        handler: applyStyle,
        schema: styleSchema,
        description: 'Applies a named style to a specified range (e.g., "paragraph:N", "selection", "document") in a Word document using COM Interop.',
        completions: async () => ({ style: ['Normal', 'Heading 1', 'Heading 2', 'Title'], range: ['selection', 'document', 'paragraph:1'] }),
    },
     {
        path: 'word/styles/list',
        handler: listStyles,
        schema: listStyleSchema,
        description: 'Lists the names of available styles in a Word document using COM Interop.',
    },
    // TODO: Add create, modify, delete operations using COM Interop
];