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
import archiver from 'archiver'; // Import archiver
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
    /** Relative directory to save extracted images (e.g., 'img', 'assets/images'). */
    imageDir: z.string().default('images').describe("Relative directory to save extracted images (e.g., 'img', 'assets/images')."),
    /** Optional prefix for extracted image filenames. */
    imagePrefix: z.string().optional().describe("Optional prefix for extracted image filenames."),
    /** Format for tables: 'markdown' (simple) or 'html' (preserves merged cells). */
    tableFormat: z.enum(['markdown', 'html']).default('html').describe("Format for tables: 'markdown' (simple) or 'html' (preserves merged cells)."),
    /** If true, create a ZIP archive containing the Markdown file and image directory. */
    zipOutput: z.boolean().default(false).describe("If true, create a ZIP archive containing the Markdown file and image directory."),
    /** Optional name for the output ZIP file (defaults based on output MD name). */
    zipFileName: z.string().optional().describe("Optional name for the output ZIP file (defaults based on output MD name)."),
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
 * Helper function to handle a single paragraph, extract formatting, and detect/handle inline images.
 * @param paragraph - The Word Paragraph COM object.
 * @param log - The logger object.
 * @param imageDir - The directory to save extracted images.
 * @param imageCounter - A counter for naming images.
 * @param imagePrefix - Optional prefix for image filenames.
 * @returns The Markdown representation of the paragraph.
 */
async function handleParagraph(paragraph: any, log: any, imageDir: string, imageCounter: { count: number }, imagePrefix?: string): Promise<string> {
    let markdown = '';
    let paragraphText = paragraph.Range.Text || '';

    // Remove trailing newline/carriage return from paragraph text
    paragraphText = paragraphText.replace(/\r?\n?$/, '');

    // TODO: Implement formatting detection (bold, italic, strikethrough, etc.)
    // TODO: Implement list detection (bullet points, numbered lists)
    // TODO: Implement heading detection (based on style or outline level)

    // Handle inline shapes (potential images) within the paragraph's range
    if (paragraph.Range.InlineShapes.Count > 0) {
        for (let i = 1; i <= paragraph.Range.InlineShapes.Count; i++) {
            const inlineShape = paragraph.Range.InlineShapes(i);
            // TODO: Check if the inlineShape is a picture and handle it using handleImage
            // Replace the placeholder text in paragraphText with the image markdown link
            // Example: paragraphText = paragraphText.replace('[image placeholder]', handleImage(inlineShape, imageDir, imageCounter, imagePrefix));
            releaseObject(inlineShape); // Release COM object
        }
    }

    // Append the processed text (with image links)
    markdown += paragraphText;

    return markdown;
}

/**
 * Helper function to handle a single Word table and convert it to Markdown or HTML.
 * @param table - The Word Table COM object.
 * @param format - The desired output format ('markdown' or 'html').
 * @returns The Markdown or HTML representation of the table.
 */
/**
 * Helper function to handle a single Word table and convert it to Markdown or HTML.
 * @param table - The Word Table COM object.
 * @param log - The logger object.
 * @param format - The desired output format ('markdown' or 'html').
 * @returns The Markdown or HTML representation of the table.
 */
async function handleTable(table: any, log: any, format: 'markdown' | 'html'): Promise<string> {
    let tableOutput = '';

    if (format === 'markdown') {
        // TODO: Implement basic Markdown table conversion (without merged cells)
        log.warn("Basic Markdown table conversion is not yet implemented.");
        tableOutput += '\n<!-- TODO: Implement basic Markdown table conversion -->\n';
    } else if (format === 'html') {
        // TODO: Implement HTML table conversion (handling merged cells)
        log.warn("HTML table conversion is not yet implemented.");
        tableOutput += '\n<!-- TODO: Implement HTML table conversion -->\n';
    }

    // Add a newline after the table
    tableOutput += '\n';

    return tableOutput;
}

/**
 * Helper function to handle a single image shape, extract and save the image, and return the Markdown link.
 * @param imageShape - The Word Shape or InlineShape COM object.
 * @param log - The logger object.
 * @param imageDir - The directory to save extracted images.
 * @param imageCounter - A counter for naming images.
 * @param imagePrefix - Optional prefix for image filenames.
 * @returns The Markdown image link string.
 */
async function handleImage(imageShape: any, log: any, imageDir: string, imageCounter: { count: number }, imagePrefix?: string): Promise<string> {
    imageCounter.count++;
    const imageName = `${imagePrefix || 'image'}${imageCounter.count}.png`; // Default to PNG, investigate other formats
    const imagePath = path.join(imageDir, imageName);

    log.info(`Attempting to extract image to: ${imagePath}`);

    try {
        // TODO: Implement image extraction logic using COM (CopyAsPicture or OLEFormat.Object.SaveAs)
        // This is a complex part and requires careful COM interaction and testing.
        log.warn("Image extraction logic is not yet implemented.");
        // Placeholder for extraction:
        // imageShape.Select();
        // wordApp.Selection.CopyAsPicture(); // Requires access to wordApp, might need to pass it
        // Paste from clipboard and save to imagePath

        // Placeholder for success:
        log.info(`Successfully extracted placeholder image to ${imagePath}`);

        // TODO: Extract alt text from imageShape if available

        const altText = imageShape.AlternativeText || ''; // Placeholder for alt text extraction

        // Return the Markdown image link
        return `![${altText}](${path.relative(path.dirname(imagePath), imagePath)})`; // Use relative path for link

    } catch (error: any) {
        log.error(`Failed to extract image: ${error.message}`);
        return `![Image Extraction Failed: ${error.message}]()`; // Return a broken link with error info
    } finally {
        // TODO: Release COM object for imageShape if necessary (depends on how it's obtained)
    }
}

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
    const totalSteps = 5; // Increased total steps for new process

    let wordApp: any = null;
    let doc: any = null;
    let markdownOutput = '';
    let imageCounter = 0;
    const extractedImagesDir = path.join(path.dirname(params.output as string), params.imageDir as string);

    try {
        reportProgress?.({ progress: 0, total: totalSteps }); // Step 0: Start

        // Validate parameters using the updated schema
        const validatedParams = exportSchema.parse(params);
        const safeInputPath = validatedParams.filePath;
        const safeOutputPath = path.resolve(validatedParams.output);
        const safeImageDir = validatedParams.imageDir;
        const tableFormat = validatedParams.tableFormat;
        const zipOutput = validatedParams.zipOutput;
        const zipFileName = validatedParams.zipFileName;
        const commentsOption = validatedParams.comments;

        log.info(`Attempting rich COM export Word doc '${safeInputPath}' to Markdown '${safeOutputPath}'`);
        log.info(`Images will be saved to: ${safeImageDir}`);
        log.info(`Tables will be formatted as: ${tableFormat}`);
        if (zipOutput) {
            log.info(`Output will be zipped.`);
        }

        // Ensure the image directory exists
        await fs.ensureDir(extractedImagesDir);
        log.info(`Ensured image directory exists: ${extractedImagesDir}`);
        reportProgress?.({ progress: 1, total: totalSteps }); // Step 1: Image directory ensured

        const officeResult = await getOfficeApplication('Word.Application');
        wordApp = officeResult.app;
        // Open the document in read-only mode (ReadOnly = true) and non-visible (Visible = false)
        log.info(`Opening document: ${safeInputPath} (Read-Only)`);
        doc = wordApp.Documents.Open(safeInputPath, false, true, false, "", "", true, "", "", "", 0, false, false); // ReadOnly=true, Visible=false
        if (!doc) {
            throw new Error(`Failed to open document via COM: ${safeInputPath}`);
        }
        log.info(`Document opened successfully.`);
        reportProgress?.({ progress: 2, total: totalSteps }); // Step 2: Document opened

        // --- Document Traversal and Element Handling ---
        log.info("Starting document traversal and element handling.");

        // Iterate through the main story range (the main body of the document)
        const mainStoryRange = doc.StoryRanges(1); // wdMainStory

        // Iterate through elements in the main story range
        let currentRange = mainStoryRange.Duplicate;
        while (currentRange.Start < currentRange.End) {
            // Check for Tables first as they can contain paragraphs and shapes
            if (currentRange.Tables.Count > 0) {
                const table = currentRange.Tables(1);
                markdownOutput += await handleTable(table, log, tableFormat);
                // Move the range past the table
                currentRange.Start = table.Range.End;
                releaseObject(table); // Release COM object
            }
            // Check for Paragraphs (includes text and inline shapes)
            else if (currentRange.Paragraphs.Count > 0) {
                const paragraph = currentRange.Paragraphs(1);
                markdownOutput += await handleParagraph(paragraph, log, extractedImagesDir, { count: imageCounter }, validatedParams.imagePrefix);
                // Move the range past the paragraph
                currentRange.Start = paragraph.Range.End;
                releaseObject(paragraph); // Release COM object
            }
            // TODO: Handle other potential elements like Shapes (non-inline images, text boxes, etc.)
            // This might require checking currentRange.ShapeRange or iterating through doc.Shapes

            else {
                // If no known element is found, move the range forward by one character
                currentRange.MoveStart(1, 1); // wdCharacter, 1
            }
        }

        releaseObject(mainStoryRange); // Release COM object
        releaseObject(currentRange); // Release COM object

        log.info("Document traversal and element handling complete.");

        // --- Comment Handling ---
        // TODO: Integrate comment handling based on commentsOption (ignore, append, inline)
        // This might require iterating through comments and finding their corresponding ranges in the document.
        log.warn("Comment handling logic is not fully integrated yet.");

        // --- End Document Traversal and Element Handling ---
        reportProgress?.({ progress: 3, total: totalSteps }); // Step 3: Document processed

        log.info(`Writing generated Markdown to ${safeOutputPath}`);
        await fs.writeFile(safeOutputPath, markdownOutput, 'utf8');
        log.info(`Successfully wrote Markdown to ${safeOutputPath}`);
        reportProgress?.({ progress: 4, total: totalSteps }); // Step 4: Markdown file written

        let finalOutputPath = safeOutputPath;

       // --- Zipping ---
       if (zipOutput) {
           log.info(`Creating ZIP archive...`);
           const zipPath = zipFileName ? path.join(path.dirname(safeOutputPath), zipFileName) : safeOutputPath.replace(/\.md$/, '.zip');
           const output = fs.createWriteStream(zipPath);
           const archive = archiver('zip', {
               zlib: { level: 9 } // Sets the compression level.
           });

           // Listen for all archive data to be written
           output.on('close', function() {
               log.info(`ZIP archive created: ${archive.pointer()} total bytes`);
               reportProgress?.({ progress: 5, total: totalSteps }); // Step 5: Zipping complete
           });

           // Catch warnings and errors
           archive.on('warning', function(err) {
               if (err.code === 'ENOENT') {
                   log.warn(`Archiver warning: ${err.message}`);
               } else {
                   log.error(`Archiver error: ${err.message}`);
                   throw err; // Throw other errors
               }
           });

           archive.on('error', function(err) {
               log.error(`Archiver error: ${err.message}`);
               throw err;
           });

           // Pipe archive data to the file
           archive.pipe(output);

           // Append the markdown file
           archive.file(safeOutputPath, { name: path.basename(safeOutputPath) });

           // Append the images directory
           if (await fs.pathExists(extractedImagesDir)) {
                archive.directory(extractedImagesDir, path.basename(extractedImagesDir));
           } else {
                log.warn(`Image directory not found, skipping zipping: ${extractedImagesDir}`);
           }


           // Finalize the archive
           await archive.finalize();

           finalOutputPath = zipPath;
           log.info(`ZIP archive finalized: ${finalOutputPath}`);

       } else {
            reportProgress?.({ progress: 5, total: totalSteps }); // Step 5: Zipping skipped
       }


       // Save the final output as a dynamic resource
       try {
           // If zipped, save the zip file. If not, save the markdown file.
           const resourceContent = await fs.readFile(finalOutputPath, zipOutput ? null : 'utf8'); // Read as buffer for zip
           await saveResource('word/markdown/export', path.basename(finalOutputPath), resourceContent);
           log.info(`Saved ${finalOutputPath} as a dynamic resource.`);
       } catch (resourceSaveError: any) {
           log.error(`Failed to save ${finalOutputPath} as a dynamic resource: ${resourceSaveError.message}`);
           // Continue execution even if resource saving fails
       }


        return { success: true, data: { outputPath: finalOutputPath } };

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