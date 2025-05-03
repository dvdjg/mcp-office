/**
 * @file Implements the 'word/styles' tool using COM Interop (winax) for managing styles in Word documents.
 * Provides functionality to apply and list styles.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '@/types/common.types'; // Import FastMCPContext from common types
import { handleToolError } from '@/utils/errorHandler';
import logger from '@/utils/logger';
import { getOfficeApplication, releaseObject } from '@/utils/officeInterop';
import { validateFilePath } from '@/utils/security';

// --- Schemas ---

/** Schema for the 'word/styles/apply' tool parameters. */
const styleSchema = z.object({
    /** The path of the Word document to modify (relative to the current workspace directory). */
    filePath: z.string().min(1).refine(validateFilePath, {
        message: "Invalid or potentially unsafe file path provided.",
    }),
    /** The name of the style to apply. */
    style: z.string().min(1),
    /** The range in the document to apply the style to (e.g., 'paragraph:N', 'selection', 'document'). */
    range: z.string().min(1), // e.g., 'paragraph:1', 'selection', 'document'
});

/** Schema for the 'word/styles/list' tool parameters. */
const listStyleSchema = z.object({
    /** The path of the Word document to analyze (relative to the current workspace directory). */
    filePath: z.string().min(1).refine(validateFilePath, {
        message: "Invalid or potentially unsafe file path provided.",
    }),
});


// --- Handlers ---

/**
 * Applies a style to a specified range in a Word document using COM Interop.
 * @param params - The parameters for the apply style operation, validated against `styleSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an empty ApiResponse indicating success.
 * @throws {Error} If the document fails to open, the range is invalid, or style application fails.
 */
export async function applyStyle(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> { // Use FastMCPContext<undefined>
    let wordApp: any = null;
    let doc: any = null;

    try {
        const validatedParams = styleSchema.parse(params);
        const safeFilePath = validatedParams.filePath; // Already validated by Zod refine
        logger.info(`[word/styles/apply] Attempting to apply style: ${validatedParams.style} to range: ${validatedParams.range} in document: ${safeFilePath}`);

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
                 logger.warn("[word/styles/apply] wordApp.Selection.Range was null or undefined. This might happen if there's no active selection or the app is headless.");
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
        logger.debug(`[word/styles/apply] Applying style '${validatedParams.style}' to range type: ${rangeStringLower}`);
        selectedRange.Style = validatedParams.style;
        // Optionally save the document: doc.Save();
        // For this tool, we typically don't save automatically.

        logger.info(`[word/styles/apply] Successfully applied style '${validatedParams.style}' to range '${validatedParams.range}' in document '${safeFilePath}'`);
        return { success: true, data: {} };

    } catch (error) {
         logger.error(`[word/styles/apply] Error applying style via COM: ${error}`, { params });
         return handleToolError(error, 'WORD_STYLE_ERROR');
    } finally {
        // --- CRUCIAL: Release COM Objects ---
        if (doc) {
            try {
                doc.Close(false); // Close without saving changes (wdDoNotSaveChanges = 0)
                logger.debug(`[word/styles/apply] Closed document: ${params.filePath}`);
            } catch (closeError) {
                logger.error(`[word/styles/apply] Error closing document: ${closeError}`);
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
            logger.debug("[word/styles/apply] Released Word Application COM object.");
        }
    }
}

/**
 * Lists available styles in the document using COM Interop.
 * @param params - The parameters for the list styles operation, validated against `listStyleSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse containing an array of style names (strings).
 * @throws {Error} If the document fails to open or style listing fails.
 */
export async function listStyles(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<string[]>> { // Use FastMCPContext<undefined>
    let wordApp: any = null;
    let doc: any = null;
     try {
        const validatedParams = listStyleSchema.parse(params);
        const safeFilePath = validatedParams.filePath; // Already validated
        logger.info(`[word/styles/list] Attempting to list styles for document: ${safeFilePath}`);

        wordApp = await getOfficeApplication('Word.Application');
        doc = wordApp.Documents.Open(safeFilePath);
        if (!doc) {
            throw new Error(`Failed to open document: ${safeFilePath}`);
        }

        const styles = doc.Styles;
        const styleCount = styles.Count;
        const styleNames: string[] = [];

        logger.debug(`[word/styles/list] Found ${styleCount} styles in the collection. Iterating...`);
        // COM collections are often 1-based
        for (let i = 1; i <= styleCount; i++) {
            let style = null;
            try {
                 style = styles(i); // Access item by 1-based index
                 if (style && style.NameLocal) {
                     styleNames.push(style.NameLocal);
                 } else {
                      logger.warn(`[word/styles/list] Style at index ${i} was null or had no NameLocal.`);
                 }
            } catch (itemError) {
                 logger.error(`[word/styles/list] Error accessing style at index ${i}: ${itemError}`);
                 // Continue to next item if possible
            } finally {
                 if (style) releaseObject(style); // Release the individual style object
            }
        }

        logger.info(`[word/styles/list] Successfully listed ${styleNames.length} styles from document '${safeFilePath}'.`);
        return { success: true, data: styleNames };

    } catch (error) {
         logger.error(`[word/styles/list] Error listing styles via COM: ${error}`, { params });
        return handleToolError(error, 'WORD_STYLE_ERROR');
    } finally {
        // --- CRUCIAL: Release COM Objects ---
        if (doc) {
            try {
                doc.Close(false); // Close without saving
                logger.debug(`[word/styles/list] Closed document: ${params.filePath}`);
            } catch (closeError) {
                logger.error(`[word/styles/list] Error closing document: ${closeError}`);
            }
            releaseObject(doc);
            doc = null;
        }
        if (wordApp) {
            releaseObject(wordApp);
            wordApp = null;
            logger.debug("[word/styles/list] Released Word Application COM object.");
        }
    }
}


// --- Resource Definition ---

/**
 * Array of McpResource definitions for Word style operations.
 */
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