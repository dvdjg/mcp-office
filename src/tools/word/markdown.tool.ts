/**
 * @file Implements Markdown import/export tools for Word using COM Interop.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import fs from 'fs-extra';
import path from 'path';
import MarkdownIt from 'markdown-it';
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams } from '../../types/common.types'; // Normalized relative path
import { saveResource } from '../dynamic/resources.tool'; // Normalized relative path
import { Context as FastMCPContext } from 'fastmcp'; // Import FastMCP Context
import { handleToolError } from '../../utils/errorHandler'; // Normalized relative path
import { validateFilePath } from '../../utils/security'; // Normalized relative path
import logger from '../../utils/logger'; // Normalized relative path
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Normalized relative path
import { applyMarkdownFormattingToWord } from '../../utils/markdownToOffice'; // Import the Markdown formatting utility

// --- Schemas ---
/**
 * Zod schema for the input parameters of the 'word/markdown/export' tool.
 */
const exportSchema = z.object({
    /** The path to the source Word document (.docx). */
    filePath: z.string().min(1).refine(validateFilePath, { // Renamed from documentReference
        message: "Invalid or potentially unsafe source file path provided.",
    }),
    /** The path where the output Markdown file (.md) will be saved. */
    output: z.string().min(1).refine(value => { // Validate output path and directory
        try {
            const dir = path.dirname(value);
            validateFilePath(dir); // Check directory safety/existence implicitly
            return true;
        } catch {
            return false;
        }
    }, {
        message: "Invalid or potentially unsafe output file path or directory.",
    }),
    /** How to handle comments during export ('ignore', 'append', 'inline'). */
    comments: z.enum(['ignore', 'append', 'inline']).default('ignore').describe("How to handle comments during export ('ignore', 'append', 'inline')."),
});

/**
 * Zod schema for the input parameters of the 'word/markdown/import' tool.
 */
const importSchema = z.object({
    /** The path to the source Markdown file (.md). */
    filePath: z.string().min(1).refine(validateFilePath, { // Renamed from path
        message: "Invalid or potentially unsafe source Markdown file path provided.",
    }),
    /** The path where the output Word document (.docx) will be saved. */
    output: z.string().min(1).refine(value => { // Validate output path and directory
        try {
            const dir = path.dirname(value);
            validateFilePath(dir);
            return true;
        } catch {
            return false;
        }
    }, {
        message: "Invalid or potentially unsafe output Word file path or directory.",
    }),
    /** Optional path to a Word template (.dotx) to use for the new document. */
    template: z.string().optional().refine(value => !value || validateFilePath(value), { // Validate template if provided
        message: "Invalid or potentially unsafe template file path provided.",
    }),
});

// --- Markdown Parser Setup (Kept as is, used for import logic) ---
const md = new MarkdownIt({
  html: false,
  xhtmlOut: false,
  breaks: true,
  linkify: true,
  typographer: true,
});
// md.use(require('markdown-it-footnote'));


// --- Handlers ---

/**
 * Exports a Word document to Markdown format using COM Interop (Basic Text Extraction).
 * NOTE: This implementation extracts plain text. Preserving formatting (headings, lists, bold, etc.)
 * requires complex iteration over the Word document structure via COM.
 * @param params - The parameters for the tool, validated against `exportSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse containing the output file path.
 * @throws {Error} If a COM error occurs, file access fails, or validation fails.
 */
async function exportToMarkdown(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{ outputPath: string }>> {
    const log = context?.log ?? logger; // Use context logger or fallback
    const reportProgress = context?.reportProgress; // Get reportProgress function if context exists
    const totalSteps = 3; // Define total steps for progress

    let wordApp: any = null;
    let doc: any = null;
    const safeOutputPath = path.resolve(params.output as string); // Already validated by Zod

    try {
        reportProgress?.({ progress: 0, total: totalSteps }); // Step 0: Start
        const validatedParams = exportSchema.parse(params);
        const safeInputPath = validatedParams.filePath; // Already validated

        log.info(`Attempting COM export Word doc '${safeInputPath}' to Markdown '${safeOutputPath}'`);

        const officeResult = await getOfficeApplication('Word.Application');
        wordApp = officeResult.app;
        // Open the document in read-only mode (ReadOnly = true) and non-visible (Visible = false)
        log.info(`Opening document: ${safeInputPath} (Read-Only)`);
        doc = wordApp.Documents.Open(safeInputPath, false, true, false, "", "", true, "", "", "", 0, false, false); // ReadOnly=true, Visible=false
        if (!doc) {
            throw new Error(`Failed to open document via COM: ${safeInputPath}`);
        }
        log.info(`Document opened successfully.`);

        // --- Basic Text Extraction using COM ---
        log.warn("Exporting using basic COM text extraction (doc.Content.Text). Formatting will be lost.");
        let extractedText = doc.Content.Text || ''; // Get the plain text content

        // --- Comment Handling (Basic COM Placeholder) ---
        if (validatedParams.comments !== 'ignore' && doc.Comments && doc.Comments.Count > 0) {
            log.info(`Extracting ${doc.Comments.Count} comments (basic)...`);
            extractedText += `\n\n## Comments (Extracted via COM)\n`;
            let commentsCollection = null;
            try {
                commentsCollection = doc.Comments;
                for (let i = 1; i <= commentsCollection.Count; i++) {
                    let comment: any = null;
                    try {
                        comment = commentsCollection(i);
                        const commentText = comment?.Range?.Text || 'Error reading comment text';
                        const author = comment?.Author || 'Unknown Author';
                        const scope = comment?.Scope?.Text ? ` (Scope: "${comment.Scope.Text.substring(0, 50)}...")` : '';
                        extractedText += `- **${author}**: ${commentText}${scope}\n`;
                    } catch (commentError: any) {
                        log.error(`Error reading comment at index ${i}: ${commentError.message}`);
                        extractedText += `- Error reading comment at index ${i}.\n`;
                    } finally {
                         if (comment) releaseObject(comment);
                    }
                }
            } catch (commentsError: any) {
                 log.error(`Error accessing comments collection: ${commentsError.message}`);
                 extractedText += `- Error accessing comments collection.\n`;
            } finally {
                 if (commentsCollection) releaseObject(commentsCollection);
            }
        }
        // --- End Comment Handling ---
        reportProgress?.({ progress: 1, total: totalSteps }); // Step 1: Text Extracted

        log.info(`Writing extracted text to ${safeOutputPath}`);
        await fs.writeFile(safeOutputPath, extractedText, 'utf8');
        log.info(`Successfully wrote extracted text via COM to ${safeOutputPath}`);
        reportProgress?.({ progress: 2, total: totalSteps }); // Step 2: File Written

        // Save the exported Markdown file as a dynamic resource
        try {
            const markdownFileContent = await fs.readFile(safeOutputPath, 'utf8');
            await saveResource('word/markdown/export', path.basename(safeOutputPath), markdownFileContent);
            log.info(`Saved ${safeOutputPath} as a dynamic resource.`);
        } catch (resourceSaveError: any) {
            log.error(`Failed to save ${safeOutputPath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continue execution even if resource saving fails
        }

        reportProgress?.({ progress: 3, total: totalSteps }); // Step 3: Complete
        return { success: true, data: { outputPath: safeOutputPath } };

    } catch (error: any) {
        // Convert error to string for logging
        log.error(`Error during Word to Markdown export: ${error.message}`, { error: String(error) });
        // Use handleToolError to create a standardized error response
        return handleToolError(error, 'WORD_MD_EXPORT_ERROR');
    } finally {
        // --- CRUCIAL: Release COM Objects ---
        if (doc) {
            try {
                doc.Close(false); // Close without saving
                log.debug(`Closed document: ${params.filePath}`);
            } catch (closeError: any) {
                log.error(`Error closing document: ${closeError.message}`);
            }
            releaseObject(doc);
            doc = null;
        }
        if (wordApp) {
            // Consider wordApp.Quit() if needed, but releasing is usually sufficient if obtained via getOfficeApplication
            releaseObject(wordApp);
            wordApp = null;
            log.debug("Released Word Application COM object for export.");
        }
    }
}

/**
 * Imports a Markdown file into a Word document using COM Interop (Basic Text Insertion).
 * NOTE: This implementation inserts the Markdown content as plain text.
 * Applying Word formatting based on Markdown syntax (headings, lists, bold, etc.)
 * requires complex parsing and interaction with Word's COM API.
 * @param params - The parameters for the tool, validated against `importSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse containing the output file path.
 * @throws {Error} If a COM error occurs, file access fails, validation fails, or document creation/saving fails.
 */
async function importFromMarkdown(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{ outputPath: string }>> {
    const log = context?.log ?? logger; // Use context logger or fallback
    const reportProgress = context?.reportProgress; // Get reportProgress function if context exists
    const totalSteps = 4; // Define total steps for progress

    let wordApp: any = null;
    let newDoc: any = null;
    const safeOutputPath = path.resolve(params.output as string); // Already validated by Zod

    try {
        reportProgress?.({ progress: 0, total: totalSteps }); // Step 0: Start
        const validatedParams = importSchema.parse(params);
        const safeInputPath = validatedParams.filePath; // Already validated
        const safeTemplatePath = validatedParams.template; // Already validated (if present)

        log.info(`Attempting COM import Markdown '${safeInputPath}' to Word '${safeOutputPath}'`);
        if (safeTemplatePath) {
            log.info(`Using template: ${safeTemplatePath}`);
        }

        log.info(`Reading Markdown file: ${safeInputPath}`);
        const markdownContent = await fs.readFile(safeInputPath, 'utf8');
        reportProgress?.({ progress: 1, total: totalSteps }); // Step 1: Markdown Read

        const officeResult = await getOfficeApplication('Word.Application');
        wordApp = officeResult.app;

        // Create new document
        log.info(`Creating new Word document (using template: ${!!safeTemplatePath})`);
        if (safeTemplatePath) {
            newDoc = wordApp.Documents.Add(safeTemplatePath);
        } else {
            newDoc = wordApp.Documents.Add();
        }
        if (!newDoc) {
            throw new Error("Failed to create new Word document via COM.");
        }
        reportProgress?.({ progress: 2, total: totalSteps }); // Step 2: Document Created

       // --- Apply Markdown Formatting using Utility ---
       log.info("Applying Markdown formatting using the utility.");
       const docContentRange = newDoc.Content;
       docContentRange.Collapse(1); // wdCollapseStart - Ensure we start from the beginning
       await applyMarkdownFormattingToWord(docContentRange, markdownContent, wordApp);
       reportProgress?.({ progress: 3, total: totalSteps }); // Step 3: Text Inserted and Formatted

       // Save the new document
        log.info(`Saving new Word document to: ${safeOutputPath}`);
        // Use WdSaveFormat enumeration for DOCX (value 16)
        const wdFormatDocumentDefault = 16; // .docx format
        newDoc.SaveAs2(safeOutputPath, wdFormatDocumentDefault);
        log.info(`Successfully saved new Word document via COM to ${safeOutputPath}`);
        reportProgress?.({ progress: 4, total: totalSteps }); // Step 4: Saved (Complete)

        // Read the content of the imported Word document before closing it
        let importedDocContent = '';
        try {
             // Read the text from the COM document object
             importedDocContent = newDoc.Content.Text;
             log.info(`Read content from imported document.`);
         } catch (readContentError: any) {
             log.error(`Failed to read content from imported document before closing: ${readContentError.message}`);
             // Do not throw error here, attempt to save the resource empty or with error
         }

        // Save the imported Word document content as a dynamic resource after closing it
        try {
            await saveResource('word/markdown/import', path.basename(safeOutputPath), importedDocContent);
            log.info(`Saved ${safeOutputPath} as a dynamic resource.`);
        } catch (resourceSaveError: any) {
            log.error(`Failed to save ${safeOutputPath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continue execution even if resource saving fails
        }

        return { success: true, data: { outputPath: safeOutputPath } };

    } catch (error: any) {
        // Convert error to string for logging
        log.error(`Error during Markdown to Word import: ${error.message}`, { error: String(error) });
        // Use handleToolError to create a standardized error response
        return handleToolError(error, 'WORD_MD_IMPORT_ERROR');
    } finally {
        // --- CRUCIAL: Release COM Objects ---
        if (newDoc) {
            try {
                // Close the *newly created* document. Saving already happened.
                // Pass false to SaveChanges parameter if you are sure no more changes are needed.
                newDoc.Close(false);
                log.debug(`Closed newly created document: ${safeOutputPath}`);
            } catch (closeError: any) {
                log.error(`Error closing newly created document: ${closeError.message}`);
            }
            releaseObject(newDoc);
            newDoc = null;
        }
        if (wordApp) {
            // Consider wordApp.Quit() if needed
            releaseObject(wordApp);
            wordApp = null;
            log.debug("Released Word Application COM object for import.");
        }
    }
}


// --- Resource Definition ---

/**
 * McpResource definition for the 'word/markdown/export' tool.
 * Exports a Word document to Markdown format.
 */
export const wordMarkdownExportTool: McpResource = {
    path: 'word/markdown/export',
    handler: exportToMarkdown,
    schema: exportSchema,
    description: 'Exports a Word document (.docx) to Markdown (.md) using COM Interop (basic text extraction, formatting lost). Includes basic comment extraction.',
};

/**
 * McpResource definition for the 'word/markdown/import' tool.
 * Imports a Markdown file into a Word document.
 */
export const wordMarkdownImportTool: McpResource = {
    path: 'word/markdown/import',
    handler: importFromMarkdown,
    schema: importSchema,
    description: 'Imports a Markdown (.md) file into a new Word document (.docx) using COM Interop (inserts as plain text, formatting lost), optionally using a template.',
};

// Export as an array of McpResource objects
export const wordMarkdownTool: McpResource[] = [
    wordMarkdownExportTool,
    wordMarkdownImportTool,
];