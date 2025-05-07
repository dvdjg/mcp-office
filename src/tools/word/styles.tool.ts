/**
 * @file Implements the 'word/styles' tool using COM Interop for managing styles in Word documents.
 * Provides functionality to apply and list styles.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import fs from 'fs-extra';
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types.js';
import { handleToolError, createErrorResponse } from '../../utils/errorHandler.js';
import logger from '../../utils/logger.js';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js';
import { validateFilePath } from '../../utils/security.js';

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
    useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
});

/**
 * Zod schema for the input parameters of the 'word/styles/list' tool.
 */
const listStyleSchema = z.object({
    /** The path of the Word document to analyze (relative to the current workspace directory). */
    filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
        message: "Invalid or potentially unsafe file path provided.",
    }),
    useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
});


// --- Handlers ---

/**
 * Applies a style to a specified range in a Word document using COM Interop.
 * @param params - The parameters for the apply style operation, validated against `styleSchema`.
 * @param context - The FastMCP context (optional), providing logging.
 * @returns A promise resolving to an empty ApiResponse indicating success.
 * @throws {Error} If validation fails, the document fails to open, the range is invalid, or style application fails.
 */
export async function applyStyle(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> {
    const log = context?.log ?? logger;
    const validatedParams = styleSchema.parse(params);
    const { filePath, style, range, useComInterop } = validatedParams;
    const safeFilePath = validateFilePath(filePath); // Already validated by Zod refine

    log.info(`[word/styles/apply] Request: style='${style}', range='${range}', file='${safeFilePath}', useComInterop=${useComInterop}`);

    if (useComInterop) {
        log.info(`[word/styles/apply] Using COM Interop path.`);
        let wordApp: any = null;
        let doc: any = null;
        // officeAppInstance was used to call release() on it, but getOfficeApplication now returns the app directly.
        // We'll release wordApp directly.

        try {
            wordApp = await getOfficeApplication('Word.Application');
            // wordApp.Visible = false; // Already handled by getOfficeApplication or default behavior
            // wordApp.DisplayAlerts = 0; // wdAlertsNone = 0 // Consider if this is globally desired

            doc = wordApp.Documents.Open(safeFilePath);
            if (!doc) {
                throw new Error(`COM: Failed to open document: ${safeFilePath}`);
            }
            log.debug(`COM: Document opened successfully: ${safeFilePath}`);

            let selectedRange: any;
            const rangeStringLower = range.toLowerCase();

            if (rangeStringLower === 'selection') {
                selectedRange = wordApp.Selection.Range;
                if (!selectedRange) {
                    throw new Error("COM: Could not get range from selection. Ensure Word is active and has a selection.");
                }
            } else if (rangeStringLower === 'document') {
                selectedRange = doc.Content;
            } else if (rangeStringLower.startsWith('paragraph:')) {
                const indexStr = range.split(':')[1];
                const index = parseInt(indexStr, 10);
                if (isNaN(index) || index <= 0) {
                    throw new Error(`COM: Invalid paragraph index format: '${indexStr}'.`);
                }
                if (index > doc.Paragraphs.Count) {
                     throw new Error(`COM: Paragraph index ${index} out of bounds. Document has ${doc.Paragraphs.Count} paragraphs.`);
                }
                selectedRange = doc.Paragraphs(index).Range;
            } else {
                throw new Error(`COM: Unsupported range format: '${range}'. Supported: 'selection', 'document', 'paragraph:N'.`);
            }

            log.debug(`COM: Applying style '${style}' to range type: ${rangeStringLower}`);
            selectedRange.Style = style;
            // Not saving automatically, as per original logic. User should call save if needed.
            // doc.Save();
            log.info(`COM: Successfully applied style '${style}' to range '${range}'.`);
            return { success: true, data: {} };

        } catch (error: any) {
            log.error(`[word/styles/apply] COM Error: ${error.message}`, { error });
            if (doc) {
                try { doc.Close(false); } catch (e: any) { log.warn(`COM: Error closing document during error handling: ${e.message}`); }
            }
            return handleToolError(error, 'WORD_STYLE_ERROR_COM');
        } finally {
            if (doc) releaseObject(doc);
            if (wordApp) releaseObject(wordApp); // Release the app object itself
            log.debug("[word/styles/apply] COM: Objects released.");
        }
    } else {
        // Library path (docx)
        log.info(`[word/styles/apply] Using Library (docx) path.`);
        log.warn(`[word/styles/apply] Library path: Applying arbitrary named styles ('${style}') is complex with 'docx'. Only limited, common styles like 'Heading1', 'Heading2' might be supported by re-creating paragraphs with specific formatting. Range '${range}' support is also limited.`);

        if (range.toLowerCase() === 'selection') {
            return createErrorResponse("Library path does not support 'selection' range for applying styles. Use COM Interop.", 'LIB_RANGE_NOT_SUPPORTED');
        }
        
        // Example for Heading styles - very limited
        // A full implementation would require parsing the doc, finding the range, and re-creating elements.
        // This is highly complex and error-prone for arbitrary styles and ranges.
        if (style.toLowerCase().startsWith('heading')) {
            // This is a placeholder for what a very limited implementation might look like.
            // It would involve reading the file, finding the paragraph(s), changing their properties, and re-saving.
            // This is beyond a simple diff for now.
            log.warn(`[word/styles/apply] Library path: Applying heading style '${style}' would require reading, modifying structure, and re-saving the document. This is not fully implemented.`);
             return createErrorResponse(
                `Library path for applying style '${style}' to range '${range}' is not fully implemented. For comprehensive style application, please use COM Interop (useComInterop: true).`,
                'NOT_IMPLEMENTED_LIB_STYLE_APPLY'
            );
        }

        return createErrorResponse(
            `Library path for applying style '${style}' is not implemented or the style is not supported directly by the 'docx' library. Use COM Interop for full style support.`,
            'NOT_IMPLEMENTED_LIB_STYLE_APPLY'
        );
    }
}

/**
 * Lists available styles in the document using COM Interop.
 * @param params - The parameters for the list styles operation, validated against `listStyleSchema`.
 * @param context - The FastMCP context (optional), providing logging.
 * @returns A promise resolving to an ApiResponse containing an array of style names (strings).
 * @throws {Error} If validation fails, the document fails to open or style listing fails.
 */
export async function listStyles(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<string[]>> {
    const log = context?.log ?? logger;
    const validatedParams = listStyleSchema.parse(params);
    const { filePath, useComInterop } = validatedParams;
    const safeFilePath = validateFilePath(filePath);

    log.info(`[word/styles/list] Request: file='${safeFilePath}', useComInterop=${useComInterop}`);

    if (useComInterop) {
        log.info(`[word/styles/list] Using COM Interop path.`);
        let wordApp: any = null;
        let doc: any = null;

        try {
            wordApp = await getOfficeApplication('Word.Application');
            // wordApp.Visible = false; // Handled by getOfficeApplication or default
            // wordApp.DisplayAlerts = 0; // Consider if needed

            doc = wordApp.Documents.Open(safeFilePath, false, true); // Open read-only
            if (!doc) {
                throw new Error(`COM: Failed to open document: ${safeFilePath}`);
            }
            log.debug(`COM: Document opened successfully: ${safeFilePath}`);

            const stylesCollection = doc.Styles;
            const styleCount = stylesCollection.Count;
            const styleNames: string[] = [];

            log.debug(`COM: Found ${styleCount} styles. Iterating...`);
            for (let i = 1; i <= styleCount; i++) { // COM collections are 1-based
                let style = null;
                try {
                    style = stylesCollection(i);
                    if (style && style.NameLocal) {
                        styleNames.push(style.NameLocal);
                    }
                } catch (itemError: any) {
                    log.error(`COM: Error accessing style at index ${i}: ${itemError.message}`);
                } finally {
                    if (style) releaseObject(style);
                }
            }
            releaseObject(stylesCollection);

            log.info(`COM: Successfully listed ${styleNames.length} styles.`);
            return { success: true, data: styleNames };

        } catch (error: any) {
            log.error(`[word/styles/list] COM Error: ${error.message}`, { error });
            if (doc) {
                try { doc.Close(false); } catch (e: any) { log.warn(`COM: Error closing document during error handling: ${e.message}`); }
            }
            return handleToolError(error, 'WORD_STYLE_ERROR_COM');
        } finally {
            if (doc) releaseObject(doc);
            if (wordApp) releaseObject(wordApp);
            log.debug("[word/styles/list] COM: Objects released.");
        }
    } else {
        // Library path (docx)
        log.info(`[word/styles/list] Using Library (docx) path.`);
        log.warn(`[word/styles/list] Library path: Listing all styles from a .docx file's definition is not directly supported by 'docx' library. It primarily generates documents, not parses existing style definitions in detail.`);
        
        // The 'docx' library does not provide a direct API to list all styles from an existing file.
        // It can define styles when creating a document, but not easily read them back.
        // We could return a predefined list of common styles that 'docx' can generate,
        // but this would not reflect the actual styles in the user's document.
        
        // For now, return an error or a very limited set of known styles.
        // Returning an error is more accurate regarding the capability.
        return createErrorResponse(
            "Library path: Listing all styles from an existing document is not supported. This feature relies on COM Interop to inspect the document's style gallery. Please use COM Interop (useComInterop: true).",
            'NOT_IMPLEMENTED_LIB_STYLE_LIST'
        );
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