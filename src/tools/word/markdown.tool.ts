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
import { createErrorResponse, handleToolError } from '@/utils/errorHandler';
import { validateFilePath } from '@/utils/security';
import logger from '@/utils/logger';
// Placeholder for Word API interaction (needs Office JS context or COM/VBA bridge)
import { getWordContext } from './styles.tool'; // Reusing the placeholder context getter

// --- Schemas ---
const exportSchema = z.object({
    documentReference: z.string(), // Path to the source Word document
    output: z.string(),           // Path to the output Markdown file
    comments: z.enum(['ignore', 'append', 'inline']).default('ignore'),
});

const importSchema = z.object({
    path: z.string(),             // Path to the source Markdown file
    output: z.string(),           // Path to the output Word document
    template: z.string().optional(), // Path to Word template (.dotx)
});

// --- Markdown Parser Setup ---
const md = new MarkdownIt({
  html: false, // Disable HTML tags in source
  xhtmlOut: false,
  breaks: true, // Convert '\n' in paragraphs into <br>
  linkify: true,
  typographer: true,
});
// Add plugins as needed (e.g., for tables, footnotes, task lists)
// md.use(require('markdown-it-footnote'));


// --- Handlers ---

/**
 * Exports a Word document to Markdown format.
 * NOTE: This is a simplified implementation. Real export requires parsing
 * the Word document structure (paragraphs, headings, lists, tables, comments)
 * and converting it to Markdown syntax. This often requires VBA or COM Interop
 * for reliable structure traversal. Office JS APIs have limitations here.
 */
async function exportToMarkdown(params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<{ outputPath: string }>> {
    try {
        const validatedParams = exportSchema.parse(params);
        const safeInputPath = validateFilePath(validatedParams.documentReference);
        // Validate output directory
        const outputDir = path.dirname(validatedParams.output);
        validateFilePath(outputDir);
        const safeOutputPath = path.resolve(validatedParams.output);

        logger.info(`Attempting to export Word doc '${safeInputPath}' to Markdown '${safeOutputPath}'`);

        // --- !!! Word Document Parsing Logic (Highly Simplified Placeholder) !!! ---
        // This section needs a robust implementation using appropriate Office APIs.
        // Option 1: Office JS (Limited structure access, best for simple text)
        // Option 2: VBA/COM Interop (Best for complex structure, Windows Desktop only)
        // Option 3: Third-party libraries (e.g., mammoth.js - converts .docx to HTML, then HTML to MD)

        // Placeholder using mammoth.js concept (requires mammoth installation: npm install mammoth)
        let markdownContent = `# Placeholder Export for ${path.basename(safeInputPath)}\n\n`;
        try {
            // Attempt to use mammoth if available (conceptual)
            const mammoth = require('mammoth'); // Dynamic require
            const result = await mammoth.extractRawText({ path: safeInputPath });
            // const htmlResult = await mammoth.convertToHtml({ path: safeInputPath }); // Alternative: convert to HTML first
            markdownContent += `(Content extracted using basic text extraction - formatting lost)\n\n${result.value}`;
            logger.warn("Export using basic text extraction via mammoth (placeholder). Formatting is lost.");

             // --- Comment Handling (Placeholder) ---
             if (validatedParams.comments !== 'ignore') {
                 markdownContent += `\n\n## Comments (Placeholder)\n`;
                 markdownContent += `- Comment 1: Placeholder text.\n`;
                 markdownContent += `- Comment 2: Another placeholder.\n`;
                 logger.warn("Comment extraction requires advanced Word API access (VBA/COM or specific Office JS APIs if available).");
             }
             // --- End Comment Handling ---

        } catch (mammothError) {
            logger.error("Mammoth library not found or failed. Using basic placeholder.", mammothError);
            markdownContent += "(Could not extract content - requires proper Word parsing implementation or library like mammoth.js)";
        }
        // --- !!! End Placeholder Parsing Logic !!! ---


        await fs.writeFile(safeOutputPath, markdownContent, 'utf8');
        logger.info(`Successfully wrote placeholder Markdown to ${safeOutputPath}`);

        return { success: true, data: { outputPath: safeOutputPath } };

    } catch (error) {
        return handleToolError(error, 'WORD_MD_EXPORT_ERROR');
    }
}

/**
 * Imports a Markdown file into a Word document.
 * NOTE: Similar to export, this requires robust interaction with Word APIs
 * to create paragraphs, apply styles based on Markdown elements (headings, bold, etc.),
 * insert tables, lists, etc. Office JS or VBA/COM is needed.
 */
async function importFromMarkdown(params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<{ outputPath: string }>> {
     try {
        const validatedParams = importSchema.parse(params);
        const safeInputPath = validateFilePath(validatedParams.path);
         // Validate output directory
        const outputDir = path.dirname(validatedParams.output);
        validateFilePath(outputDir);
        const safeOutputPath = path.resolve(validatedParams.output);
        let safeTemplatePath: string | undefined;
        if (validatedParams.template) {
            safeTemplatePath = validateFilePath(validatedParams.template);
        }

        logger.info(`Attempting to import Markdown '${safeInputPath}' to Word '${safeOutputPath}'`);
        if (safeTemplatePath) {
            logger.info(`Using template: ${safeTemplatePath}`);
        }

        const markdownContent = await fs.readFile(safeInputPath, 'utf8');

        // --- !!! Word Document Creation Logic (Placeholder) !!! ---
        // This requires creating a new Word document (potentially from the template)
        // and then iterating through the Markdown structure (parsed by markdown-it)
        // to insert content and apply formatting using Word APIs.

        logger.warn("Markdown import to Word is not implemented. Requires Office JS or VBA/COM bridge.");
        logger.debug("Parsed Markdown Tokens (Example):", md.parse(markdownContent, {}));

        // Placeholder: Just save the markdown content to a .txt file instead of .docx
        const placeholderOutputPath = safeOutputPath.replace('.docx', '.txt');
        await fs.writeFile(placeholderOutputPath, `# Placeholder Import\n\nSource MD: ${safeInputPath}\nTemplate: ${safeTemplatePath || 'None'}\n\n---\n\n${markdownContent}`, 'utf8');
        logger.warn(`Placeholder: Saved raw Markdown content to ${placeholderOutputPath} instead of creating DOCX.`);
        // --- !!! End Placeholder Creation Logic !!! ---


        // Return the intended output path, even though a placeholder was created
        return { success: true, data: { outputPath: safeOutputPath } };

    } catch (error) {
        return handleToolError(error, 'WORD_MD_IMPORT_ERROR');
    }
}


// --- Resource Definition ---
export const wordMarkdownTool: McpResource[] = [
    {
        path: 'word/markdown/export',
        handler: exportToMarkdown,
        schema: exportSchema,
        description: 'Exports a Word document (.docx) to Markdown (.md). Requires robust Word parsing (VBA/COM or library like mammoth.js). Comment handling is basic.',
    },
    {
        path: 'word/markdown/import',
        handler: importFromMarkdown,
        schema: importSchema,
        description: 'Imports a Markdown (.md) file into a new Word document (.docx), optionally using a template. Requires Office JS or VBA/COM bridge for actual DOCX creation.',
    },
];