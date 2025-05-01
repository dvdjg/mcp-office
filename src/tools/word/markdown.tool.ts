// /src/tools/word/markdown.tool.ts
// =============================================================================
/**
 * @file Implements Markdown import/export tools for Word.
 */
import fs from 'fs-extra';
import path from 'path';
import MarkdownIt from 'markdown-it';
import { z } from 'zod';
import { McpResource, ApiResponse, ToolContext, ToolRequestParams } from '@/types/common.types';
import { handleToolError } from '@/utils/errorHandler';
import { validateFilePath } from '@/utils/security';
import logger from '@/utils/logger';
import { getOfficeApplication, releaseObject } from '@/utils/officeInterop'; // Use COM Interop

// --- Schemas ---
const exportSchema = z.object({
    filePath: z.string().min(1).refine(validateFilePath, { // Renamed from documentReference
        message: "Invalid or potentially unsafe source file path provided.",
    }),
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
    comments: z.enum(['ignore', 'append', 'inline']).default('ignore'),
});

const importSchema = z.object({
    filePath: z.string().min(1).refine(validateFilePath, { // Renamed from path
        message: "Invalid or potentially unsafe source Markdown file path provided.",
    }),
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
 */
async function exportToMarkdown(params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<{ outputPath: string }>> {
    let wordApp: any = null;
    let doc: any = null;
    const safeOutputPath = path.resolve(params.output as string); // Already validated by Zod

    try {
        const validatedParams = exportSchema.parse(params);
        const safeInputPath = validatedParams.filePath; // Already validated

        logger.info(`Attempting COM export Word doc '${safeInputPath}' to Markdown '${safeOutputPath}'`);

        wordApp = await getOfficeApplication('Word.Application');
        doc = wordApp.Documents.Open(safeInputPath);
        if (!doc) {
            throw new Error(`Failed to open document via COM: ${safeInputPath}`);
        }

        // --- Basic Text Extraction using COM ---
        logger.warn("Exporting using basic COM text extraction (doc.Content.Text). Formatting will be lost.");
        let extractedText = doc.Content.Text || ''; // Get the plain text content

        // --- Comment Handling (Basic COM Placeholder) ---
        if (validatedParams.comments !== 'ignore' && doc.Comments && doc.Comments.Count > 0) {
            logger.info(`Extracting ${doc.Comments.Count} comments (basic)...`);
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
                    } catch (commentError) {
                        logger.error(`Error reading comment at index ${i}: ${commentError}`);
                        extractedText += `- Error reading comment at index ${i}.\n`;
                    } finally {
                         if (comment) releaseObject(comment);
                    }
                }
            } catch (commentsError) {
                 logger.error(`Error accessing comments collection: ${commentsError}`);
                 extractedText += `- Error accessing comments collection.\n`;
            } finally {
                 if (commentsCollection) releaseObject(commentsCollection);
            }
        }
        // --- End Comment Handling ---

        await fs.writeFile(safeOutputPath, extractedText, 'utf8');
        logger.info(`Successfully wrote extracted text via COM to ${safeOutputPath}`);

        return { success: true, data: { outputPath: safeOutputPath } };

    } catch (error) {
        return handleToolError(error, 'WORD_MD_EXPORT_ERROR');
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
            // Consider wordApp.Quit() if needed, but releasing is usually sufficient if obtained via getOfficeApplication
            releaseObject(wordApp);
            wordApp = null;
            logger.debug("Released Word Application COM object for export.");
        }
    }
}

/**
 * Imports a Markdown file into a Word document using COM Interop (Basic Text Insertion).
 * NOTE: This implementation inserts the Markdown content as plain text.
 * Applying Word formatting based on Markdown syntax (headings, lists, bold, etc.)
 * requires complex parsing and interaction with Word's COM API.
 */
async function importFromMarkdown(params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<{ outputPath: string }>> {
    let wordApp: any = null;
    let newDoc: any = null;
    const safeOutputPath = path.resolve(params.output as string); // Already validated by Zod

    try {
        const validatedParams = importSchema.parse(params);
        const safeInputPath = validatedParams.filePath; // Already validated
        const safeTemplatePath = validatedParams.template; // Already validated (if present)

        logger.info(`Attempting COM import Markdown '${safeInputPath}' to Word '${safeOutputPath}'`);
        if (safeTemplatePath) {
            logger.info(`Using template: ${safeTemplatePath}`);
        }

        const markdownContent = await fs.readFile(safeInputPath, 'utf8');

        wordApp = await getOfficeApplication('Word.Application');

        // Create new document
        if (safeTemplatePath) {
            newDoc = wordApp.Documents.Add(safeTemplatePath);
        } else {
            newDoc = wordApp.Documents.Add();
        }
        if (!newDoc) {
            throw new Error("Failed to create new Word document via COM.");
        }

        // --- Basic Text Insertion using COM ---
        logger.warn("Importing Markdown as plain text using COM (newDoc.Content.Text). Formatting is lost.");
        newDoc.Content.Text = markdownContent; // Insert the entire Markdown as plain text

        // --- Complex Formatting (Placeholder Idea) ---
        // For real formatting, you would:
        // 1. Parse markdownContent using md.parse(markdownContent, {})
        // 2. Iterate through tokens:
        //    - If heading_open, insert text, apply style (e.g., newDoc.Paragraphs.Last.Style = "Heading 1")
        //    - If strong_open, turn on bold (e.g., wordApp.Selection.Font.Bold = true), insert text, turn off bold
        //    - If bullet_list_open, start applying list formatting... etc.
        // This is highly non-trivial.
        // logger.debug("Parsed Markdown Tokens (for potential future formatting):", md.parse(markdownContent, {}));
        // --- End Complex Formatting ---


        // Save the new document
        // Use WdSaveFormat enumeration for DOCX (value 16)
        const wdFormatDocumentDefault = 16; // .docx format
        newDoc.SaveAs2(safeOutputPath, wdFormatDocumentDefault);
        logger.info(`Successfully saved new Word document via COM to ${safeOutputPath}`);

        return { success: true, data: { outputPath: safeOutputPath } };

    } catch (error) {
        return handleToolError(error, 'WORD_MD_IMPORT_ERROR');
    } finally {
        // --- CRUCIAL: Release COM Objects ---
        if (newDoc) {
            try {
                // Close the *newly created* document. Saving already happened.
                // Pass false to SaveChanges parameter if you are sure no more changes are needed.
                newDoc.Close(false);
                logger.debug(`Closed newly created document: ${safeOutputPath}`);
            } catch (closeError) {
                logger.error(`Error closing newly created document: ${closeError}`);
            }
            releaseObject(newDoc);
            newDoc = null;
        }
        if (wordApp) {
            // Consider wordApp.Quit() if needed
            releaseObject(wordApp);
            wordApp = null;
            logger.debug("Released Word Application COM object for import.");
        }
    }
}


// --- Resource Definition ---
export const wordMarkdownTool: McpResource[] = [
    {
        path: 'word/markdown/export',
        handler: exportToMarkdown,
        schema: exportSchema,
        description: 'Exports a Word document (.docx) to Markdown (.md) using COM Interop (basic text extraction, formatting lost). Includes basic comment extraction.',
    },
    {
        path: 'word/markdown/import',
        handler: importFromMarkdown,
        schema: importSchema,
        description: 'Imports a Markdown (.md) file into a new Word document (.docx) using COM Interop (inserts as plain text, formatting lost), optionally using a template.',
    },
];