/**
 * @file Implements the 'word/styles' tool using COM Interop for managing styles in Word documents.
 * Provides functionality to apply and list styles.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types'; // Normalized relative path
import { handleToolError, createErrorResponse } from '../../utils/errorHandler'; // Normalized relative path
import logger from '../../utils/logger'; // Normalized relative path
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Normalized relative path
import { validateFilePath } from '../../utils/security'; // Normalized relative path

// --- Schemas ---

/**
 * Zod schema for the input parameters of the 'word/styles/apply' tool.
 */
const styleSchema = z.object({
    /** The path of the Word document to modify (relative to the current workspace directory). */
    filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
        message: "Invalid or potentially unsafe file path provided.",
    }),
    /** The name of the style to apply. */
    style: z.string().min(1, 'Style name is required.'),
    /** The range in the document to apply the style to (e.g., 'paragraph:N', 'selection', 'document'). */
    range: z.string().min(1, 'Range specifier is required.'), // e.g., 'paragraph:1', 'selection', 'document'
});

/**
 * Zod schema for the input parameters of the 'word/styles/list' tool.
 */
const listStyleSchema = z.object({
    /** The path of the Word document to analyze (relative to the current workspace directory). */
    filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
        message: "Invalid or potentially unsafe file path provided.",
    }),
});


// --- Handlers ---

/**
 * Applies a style to a specified range in a Word document using COM Interop.
 * @param params - The parameters for the apply style operation, validated against `styleSchema`.
 * @param context - The FastMCP context (optional), providing logging.
 * @returns A promise resolving to an empty ApiResponse indicating success.
 * @throws {Error} If validation fails, the document fails to open, the range is invalid, or style application fails.
 */
export async function applyStyle(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> { // Use FastMCPContext<undefined>
    const log = context?.log ?? logger; // Use context logger or fallback

    let wordApp: any = null;
    let doc: any = null;
    let officeAppInstance: any = null; // To manage the application instance lifecycle

    try {
        const validatedParams = styleSchema.parse(params);
        const safeFilePath = validatedParams.filePath; // Already validated by Zod refine
        log.info(`[word/styles/apply] Attempting to apply style: ${validatedParams.style} to range: ${validatedParams.range} in document: ${safeFilePath}`);

        officeAppInstance = await getOfficeApplication('Word.Application');
        wordApp = officeAppInstance.app;
        wordApp.Visible = false; // Run in background
        wordApp.DisplayAlerts = 0; // wdAlertsNone = 0
        // Consider making Word visible for debugging: wordApp.Visible = true;

        doc = wordApp.Documents.Open(safeFilePath);
        if (!doc) {
            log.error(`[word/styles/apply] Failed to open document: ${safeFilePath}`);
            throw new Error(`Failed to open document: ${safeFilePath}`);
        }
        log.debug(`Document opened successfully.`);

        let selectedRange: any;
        const rangeStringLower = validatedParams.range.toLowerCase();

        if (rangeStringLower === 'selection') {
            // Note: 'selection' might be tricky if Word isn't visible or doesn't have focus.
            // It refers to the current selection in the Word UI.
            // If running headless, this might not be what the user expects.
            // Consider if 'document' or specific paragraphs are more reliable.
            selectedRange = wordApp.Selection.Range; // Get the Range object from the Selection
             if (!selectedRange) {
                 log.warn("[word/styles/apply] wordApp.Selection.Range was null or undefined. This might happen if there's no active selection or the app is headless.");
                 throw new Error("Could not get range from selection. Ensure the document is active and has a selection, or use a different range specifier.");
             }
             log.debug(`Applying style to current selection.`);
        } else if (rangeStringLower === 'document') {
            selectedRange = doc.Content;
            log.debug(`Applying style to entire document content.`);
        } else if (rangeStringLower.startsWith('paragraph:')) {
            const indexStr = validatedParams.range.split(':')[1];
            const index = parseInt(indexStr, 10);
            if (!isNaN(index) && index > 0) { // COM indices are typically 1-based
                 if (index <= doc.Paragraphs.Count) {
                    selectedRange = doc.Paragraphs(index).Range;
                    log.debug(`Applying style to paragraph ${index}.`);
                 } else {
                    log.warn(`Paragraph index ${index} is out of bounds. Document has ${doc.Paragraphs.Count} paragraphs.`);
                    throw new Error(`Paragraph index ${index} is out of bounds. Document has ${doc.Paragraphs.Count} paragraphs.`);
                 }
            } else {
                 log.warn(`Invalid paragraph index format: '${indexStr}'.`);
                 throw new Error(`Invalid paragraph index format: '${indexStr}'. Use 'paragraph:N' where N is a positive integer.`);
            }
        } else {
            log.warn(`Unsupported range format: '${validatedParams.range}'.`);
            throw new Error(`Unsupported range format: '${validatedParams.range}'. Supported formats: 'selection', 'document', 'paragraph:N'.`);
        }

        // Apply the style
        log.debug(`[word/styles/apply] Applying style '${validatedParams.style}' to range type: ${rangeStringLower}`);
        selectedRange.Style = validatedParams.style;
        // Optionally save the document: doc.Save();
        // For this tool, we typically don't save automatically.

        log.info(`[word/styles/apply] Successfully applied style '${validatedParams.style}' to range '${validatedParams.range}' in document '${safeFilePath}'`);
        return { success: true, data: {} };

    } catch (error: any) {
         log.error(`[word/styles/apply] Error applying style via COM: ${error.message}`, { error, params });
         // Ensure document is closed if it was opened
         if (doc) {
             try { doc.Close(false); } catch (e: any) { log.warn(`Error closing document during error handling: ${e.message}`); }
             releaseObject(doc);
         }
         return handleToolError(error, 'WORD_STYLE_ERROR');
    } finally {
        // --- CRUCIAL: Release COM Objects ---
        if (doc) { // Redundant if closed in catch, but safe
            releaseObject(doc);
        }
        if (officeAppInstance) {
            officeAppInstance.release(); // Release the application instance
            log.debug("[word/styles/apply] Office application instance released.");
        }
        if (wordApp) { // Release the app object reference
            releaseObject(wordApp);
        }
    }
}

/**
 * Lists available styles in the document using COM Interop.
 * @param params - The parameters for the list styles operation, validated against `listStyleSchema`.
 * @param context - The FastMCP context (optional), providing logging.
 * @returns A promise resolving to an ApiResponse containing an array of style names (strings).
 * @throws {Error} If validation fails, the document fails to open or style listing fails.
 */
export async function listStyles(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<string[]>> { // Use FastMCPContext<undefined>
    const log = context?.log ?? logger; // Use context logger or fallback

    let wordApp: any = null;
    let doc: any = null;
    let officeAppInstance: any = null; // To manage the application instance lifecycle

     try {
        const validatedParams = listStyleSchema.parse(params);
        const safeFilePath = validatedParams.filePath; // Already validated
        log.info(`[word/styles/list] Attempting to list styles for document: ${safeFilePath}`);

        officeAppInstance = await getOfficeApplication('Word.Application');
        wordApp = officeAppInstance.app;
        wordApp.Visible = false; // Run in background
        wordApp.DisplayAlerts = 0; // wdAlertsNone = 0

        doc = wordApp.Documents.Open(safeFilePath);
        if (!doc) {
            log.error(`[word/styles/list] Failed to open document: ${safeFilePath}`);
            throw new Error(`Failed to open document: ${safeFilePath}`);
        }
        log.debug(`Document opened successfully.`);

        const styles = doc.Styles;
        const styleCount = styles.Count;
        const styleNames: string[] = [];

        log.debug(`[word/styles/list] Found ${styleCount} styles in the collection. Iterating...`);
        // COM collections are often 1-based
        for (let i = 1; i <= styleCount; i++) {
            let style = null;
            try {
                 style = styles(i); // Access item by 1-based index
                 if (style && style.NameLocal) {
                     styleNames.push(style.NameLocal);
                 } else {
                      log.warn(`[word/styles/list] Style at index ${i} was null or had no NameLocal.`);
                 }
            } catch (itemError: any) {
                 log.error(`[word/styles/list] Error accessing style at index ${i}: ${itemError.message}`, { error: itemError });
                 // Continue to next item if possible
            } finally {
                 if (style) releaseObject(style); // Release the individual style object
            }
        }
        releaseObject(styles); // Release styles collection

        log.info(`[word/styles/list] Successfully listed ${styleNames.length} styles from document '${safeFilePath}'.`);
        return { success: true, data: styleNames };

    } catch (error: any) {
         log.error(`[word/styles/list] Error listing styles via COM: ${error.message}`, { error, params });
         // Ensure document is closed if it was opened
         if (doc) {
             try { doc.Close(false); } catch (e: any) { log.warn(`Error closing document during error handling: ${e.message}`); }
             releaseObject(doc);
         }
        return handleToolError(error, 'WORD_STYLE_ERROR');
    } finally {
        // --- CRUCIAL: Release COM Objects ---
        if (doc) { // Redundant if closed in catch, but safe
            releaseObject(doc);
        }
        if (officeAppInstance) {
            officeAppInstance.release(); // Release the application instance
            log.debug("[word/styles/list] Office application instance released.");
        }
        if (wordApp) { // Release the app object reference
            releaseObject(wordApp);
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