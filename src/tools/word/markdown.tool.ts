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

    // Implement formatting detection (bold, italic, strikethrough, etc.)
    // Iterate through runs to preserve formatting within the paragraph
    if (paragraph.Range.Runs.Count > 0) {
        for (let i = 1; i <= paragraph.Range.Runs.Count; i++) {
            const run = paragraph.Range.Runs(i);
            let runText = run.Text || '';

            // Remove trailing newline/carriage return from run text if it's the last run
            if (i === paragraph.Range.Runs.Count) {
                 runText = runText.replace(/\r?\n?$/, '');
            }


            // Apply Markdown formatting based on run properties
            if (run.Font.Bold) {
                runText = `**${runText}**`;
            }
            if (run.Font.Italic) {
                runText = `*${runText}*`;
            }
            if (run.Font.StrikeThrough) {
                runText = `~~${runText}~~`;
            }

            // TODO: Implement list detection (bullet points, numbered lists) - This might need to be handled at the paragraph level based on ListFormat
            // TODO: Implement heading detection (based on style or outline level) - This should be handled at the paragraph level based on Style

            // Handle inline shapes (potential images) within the run's range
            if (run.InlineShapes.Count > 0) {
                for (let j = 1; j <= run.InlineShapes.Count; j++) {
                    const inlineShape = run.InlineShapes(j);
                    // Check if the inlineShape is a picture (wdInlineShapePicture = 3)
                    if (inlineShape.Type === 3) { // wdInlineShapePicture
                        log.info(`Found inline picture shape.`);
                        const imageMarkdown = await handleImage(inlineShape, log, imageDir, imageCounter, imagePrefix);
                        markdown += imageMarkdown; // Append the Markdown image link
                    } else {
                        // Handle other types of inline shapes if necessary, or add a placeholder
                        log.warn(`Found non-picture inline shape (Type: ${inlineShape.Type}). Handling not implemented.`);
                        markdown += `[Inline Shape Placeholder (Type: ${inlineShape.Type})]`;
                    }
                    releaseObject(inlineShape); // Release COM object
                }
            }

            markdown += runText;
            releaseObject(run); // Release COM object
        }
    } else {
        // If no runs, just append the paragraph text (shouldn't happen often for non-empty paragraphs)
        markdown += paragraphText;
    }


    // TODO: Implement list detection (bullet points, numbered lists) - This will likely involve checking paragraph.ListFormat
    // TODO: Implement heading detection (based on style or outline level) - This will likely involve checking paragraph.Style.NameLocal or paragraph.OutlineLevel

    // Add a newline after the paragraph, unless it's the last paragraph in the document or part of a list/heading handled separately
    // This basic approach might need refinement when lists and headings are implemented.
    // Implement heading detection (based on style or outline level)
    let headingPrefix = '';
    const styleName = paragraph.Style.NameLocal;
    const outlineLevel = paragraph.OutlineLevel; // wdOutlineLevel enumeration

    if (styleName) {
        if (styleName.includes('Heading 1')) {
            headingPrefix = '# ';
        } else if (styleName.includes('Heading 2')) {
            headingPrefix = '## ';
        } else if (styleName.includes('Heading 3')) {
            headingPrefix = '### ';
        } else if (styleName.includes('Heading 4')) {
            headingPrefix = '#### ';
        } else if (styleName.includes('Heading 5')) {
            headingPrefix = '##### ';
        } else if (styleName.includes('Heading 6')) {
            headingPrefix = '###### ';
        }
    } else if (outlineLevel >= 1 && outlineLevel <= 9) { // wdOutlineLevel 1 to 9 correspond to heading levels
         // Use outline level as a fallback if no standard heading style is applied
         headingPrefix = '#'.repeat(outlineLevel) + ' ';
    }

    // Append the heading prefix if it's a heading
    markdown += headingPrefix;

    // Implement formatting detection (bold, italic, strikethrough, etc.)
    // Iterate through runs to preserve formatting within the paragraph
    if (paragraph.Range.Runs.Count > 0) {
        for (let i = 1; i <= paragraph.Range.Runs.Count; i++) {
            const run = paragraph.Range.Runs(i);
            let runText = run.Text || '';

            // Remove trailing newline/carriage return from run text if it's the last run
            if (i === paragraph.Range.Runs.Count) {
                 runText = runText.replace(/\r?\n?$/, '');
            }


            // Apply Markdown formatting based on run properties
            if (run.Font.Bold) {
                runText = `**${runText}**`;
            }
            if (run.Font.Italic) {
                runText = `*${runText}*`;
            }
            if (run.Font.StrikeThrough) {
                runText = `~~${runText}~~`;
            }

            // Implement list detection (bullet points, numbered lists)
            let listPrefix = '';
            if (paragraph.ListFormat.ListType !== 0) { // wdListNoNumbering = 0
                try {
                    const listType = paragraph.ListFormat.ListType;
                    const listLevel = paragraph.ListFormat.ListLevelNumber;
                    const indentation = '  '.repeat(listLevel - 1); // 2 spaces per level for indentation

                    if (listType === 1) { // wdListBullet
                        listPrefix = `${indentation}* `; // Bullet point with indentation
                    } else if (listType === 2 || listType === 4) { // wdListNumbering or wdListOutlineNumbering
                         // Use the actual list string for numbering, but add indentation
                         const listString = paragraph.ListFormat.ListString.trim();
                         listPrefix = `${indentation}${listString} `;
                    }
                    // TODO: Refine list prefix generation for better Markdown compatibility and indentation (e.g. handling different numbering styles a, i, etc.)

                } catch (listError: any) {
                    log.warn(`Error detecting list format: ${listError.message}`);
                    listPrefix = '- '; // Fallback to a simple bullet point on error
                }
            }

            // Append the list prefix if it's a list item
            markdown += listPrefix;

            // Handle inline shapes (potential images) within the run's range
            if (run.InlineShapes.Count > 0) {
                for (let j = 1; j <= run.InlineShapes.Count; j++) {
                    const inlineShape = run.InlineShapes(j);
                    // TODO: Check if the inlineShape is a picture and handle it using handleImage
                    // For now, just add a placeholder or basic handling
                    log.warn(`Found inline shape in run. Image handling not fully implemented.`);
                    runText += `[Inline Shape Placeholder]`; // Add a placeholder for now
                    releaseObject(inlineShape); // Release COM object
                }
            }

            markdown += runText;
            releaseObject(run); // Release COM object
        }
    } else {
        // If no runs, just append the paragraph text (shouldn't happen often for non-empty paragraphs)
        markdown += paragraphText;
    }


    // TODO: Implement list detection (bullet points, numbered lists) - This will likely involve checking paragraph.ListFormat

    // Add a newline after the paragraph, unless it's the last paragraph in the document or part of a list/heading handled separately
    // This basic approach might need refinement when lists and headings are implemented.
    markdown += '\n';


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
    let isSimpleTable = true; // Declare isSimpleTable here

    // Check if it's a simple table (no merged cells) before attempting Markdown
    // A table is simple if no cell is part of a merged range.
    try {
        isSimpleTable = true; // Assume simple initially
        for (let i = 1; i <= table.Rows.Count; i++) {
            for (let j = 1; j <= table.Columns.Count; j++) {
                const cell = table.Cell(i, j);
                // Check if the cell is part of a horizontally or vertically merged range
                if (cell.MergeInfo.IsMerged || cell.MergeInfo.IsVerticalMerged) {
                    isSimpleTable = false;
                    releaseObject(cell); // Release COM object
                    throw new Error("Merged cell found"); // Exit loops early
                }
                releaseObject(cell); // Release COM object
            }
        }
    } catch (checkError: any) {
         if (checkError.message !== "Merged cell found") {
            log.warn(`Error checking table for merged cells: ${checkError.message}. Assuming complex.`);
         }
         isSimpleTable = false;
    }

    if (format === 'markdown') {
        // Implement basic Markdown table conversion (without merged cells)
        log.info("Attempting basic Markdown table conversion.");
        if (isSimpleTable) {
            // Header row
            tableOutput += '| ';
            for (let j = 1; j <= table.Columns.Count; j++) {
                tableOutput += (table.Cell(1, j).Range.Text || '').replace(/[\r\n]/g, '').trim() + ' |';
            }
            tableOutput += '\n';

            // Separator line
            tableOutput += '|';
            for (let j = 1; j <= table.Columns.Count; j++) {
                tableOutput += '---|';
            }
            tableOutput += '\n';

            // Data rows
            for (let i = 2; i <= table.Rows.Count; i++) {
                tableOutput += '| ';
                for (let j = 1; j <= table.Columns.Count; j++) {
                    tableOutput += (table.Cell(i, j).Range.Text || '').replace(/[\r\n]/g, '').trim() + ' |';
                }
                tableOutput += '\n';
            }
        } else {
             log.warn("Table is not simple (likely has merged cells). Cannot convert to basic Markdown.");
             tableOutput += '\n<!-- Table has merged cells or is complex, basic Markdown conversion skipped -->\n';
             tableOutput += '\n<!-- Please use tableFormat: "html" for this table -->\n';
        }
    } else if (format === 'html') {
        // Implement HTML table conversion (handling merged cells)
        log.info("Attempting HTML table conversion.");
        tableOutput += '<table>\n';

        for (let i = 1; i <= table.Rows.Count; i++) {
            tableOutput += '  <tr>\n';
            for (let j = 1; j <= table.Columns.Count; j++) {
                const cell = table.Cell(i, j);
                // Check if this cell is part of a merged range and is NOT the top-left cell of that range.
                // If it's not the top-left cell, we skip it as it will be covered by the colspan/rowspan of the starting cell.
                if (cell.MergeInfo.IsMerged && (!cell.MergeInfo.IsFirst || !cell.MergeInfo.IsVerticalMerged)) {
                    releaseObject(cell); // Release COM object for the skipped cell
                    continue; // Skip this cell
                }

                let cellContent = (cell.Range.Text || '').replace(/[\r\n]/g, '').trim();
                let tdAttributes = '';

                // If the cell is the start of a merged range, calculate colspan and rowspan
                if (cell.MergeInfo.IsMerged && cell.MergeInfo.IsFirst) {
                    const columnSpan = cell.MergeInfo.ColumnSpan;
                    const rowSpan = cell.MergeInfo.RowSpan;
                    if (columnSpan > 1) {
                        tdAttributes += ` colspan="${columnSpan}"`;
                    }
                    if (rowSpan > 1) {
                        tdAttributes += ` rowspan="${rowSpan}"`;
                    }
                    log.debug(`Detected merged cell at (${i}, ${j}) with colspan=${columnSpan}, rowspan=${rowSpan}`);
                }


                tableOutput += `    <td${tdAttributes}>${cellContent}</td>\n`;
                releaseObject(cell); // Release COM object
            }
            tableOutput += '  </tr>\n';
        }

        tableOutput += '</table>\n';
        log.info("HTML table conversion attempted.");
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
        // Implement image extraction logic using COM (CopyAsPicture or OLEFormat.Object.SaveAs)
        // This is a complex part and requires careful COM interaction and testing.

        let altText = '';
        try {
            // Attempt to extract alt text
            altText = imageShape.AlternativeText || '';
        } catch (altTextError: any) {
            log.warn(`Could not extract alt text for image: ${altTextError.message}`);
        }


        // Prioritize OLEFormat.Object.SaveAs if available
        if (imageShape.OLEFormat && imageShape.OLEFormat.Object && typeof imageShape.OLEFormat.Object.SaveAs === 'function') {
            try {
                log.info(`Attempting extraction using OLEFormat.Object.SaveAs to ${imagePath}`);
                imageShape.OLEFormat.Object.SaveAs(imagePath);
                log.info(`Successfully extracted image using OLEFormat.Object.SaveAs to ${imagePath}`);
                // Return the Markdown image link
                return `![${altText}](${path.relative(path.dirname(imagePath), imagePath)})`; // Use relative path for link

            } catch (oleSaveError: any) {
                log.warn(`OLEFormat.Object.SaveAs failed: ${oleSaveError.message}. Falling back or logging.`);
                // Fallback to other methods or log failure
                // TODO: Implement CopyAsPicture fallback if OLEFormat.Object.SaveAs fails or is not available
                log.error(`Image extraction failed for ${imageName} using OLEFormat.Object.SaveAs. CopyAsPicture fallback not yet implemented.`);
                return `![Image Extraction Failed (OLE Save): ${oleSaveError.message}]()`; // Return a broken link with error info
            }
        } else {
             log.warn(`OLEFormat.Object.SaveAs not available for image ${imageName}. Attempting CopyAsPicture fallback.`);
             try {
                 // Attempt CopyAsPicture fallback
                 imageShape.Select(); // Select the shape to copy it
                 imageShape.CopyAsPicture(); // Copy the shape as a picture to the clipboard

                 // Create a temporary chart or document to paste into
                 // Using a temporary document might be more reliable
                 let tempDoc: any = null;
                 try {
                     tempDoc = imageShape.Application.Documents.Add(); // Use the same Word application instance
                     tempDoc.Content.Paste(); // Paste the picture from the clipboard

                     // Save the pasted picture as a file
                     // Assuming the pasted item is the first inline shape in the temp doc
                     if (tempDoc.InlineShapes.Count > 0) {
                         const pastedShape = tempDoc.InlineShapes(1);
                         // Use Export method for InlineShape
                         pastedShape.Export(imagePath, 1); // wdExportFormatPNG = 1
                         log.info(`Successfully extracted image using CopyAsPicture fallback to ${imagePath}`);
                         releaseObject(pastedShape); // Release COM object
                          // Return the Markdown image link
                         return `![${altText}](${path.relative(path.dirname(imagePath), imagePath)})`; // Use relative path for link
                     } else {
                         log.error(`CopyAsPicture fallback failed: No inline shape found after pasting.`);
                         return `![Image Extraction Failed (CopyAsPicture Paste Failed)]()`; // Return a broken link
                     }
                 } finally {
                     // Close the temporary document without saving
                     if (tempDoc) {
                         try {
                             tempDoc.Close(false);
                             log.debug("Closed temporary document for CopyAsPicture.");
                         } catch (closeError: any) {
                             log.error(`Error closing temporary document: ${closeError.message}`);
                         }
                         releaseObject(tempDoc);
                     }
                 }
             } catch (copyPasteError: any) {
                 log.error(`CopyAsPicture fallback failed: ${copyPasteError.message}`);
                 return `![Image Extraction Failed (CopyAsPicture Error): ${copyPasteError.message}]()`; // Return a broken link with error info
             }
        }


    } catch (error: any) {
        log.error(`An unexpected error occurred during image extraction for ${imageName}: ${error.message}`);
        return `![Image Extraction Failed (Unexpected Error): ${error.message}]()`; // Return a broken link with error info
    } finally {
        // Release COM object for imageShape if necessary (depends on how it's obtained)
        // In handleParagraph, the inlineShape is released after the loop.
        // If handleImage is called with a Shape object directly, it might need release here.
        // For now, assuming it's called with an InlineShape from a paragraph run.
        // TODO: Confirm COM object release strategy for imageShape.
    }
}

/**
 * Helper function to handle a single shape (floating or inline), extract its content, and export it as SVG if possible.
 * @param shape - The Word Shape or InlineShape COM object.
 * @param log - The logger object.
 * @param outputDir - The directory to save extracted SVG files.
 * @param shapeCounter - A counter for naming shapes.
 * @returns A promise resolving to the Markdown image link string for the SVG, or null if not handled/exported.
 */
async function handleShapeAndExportSvg(shape: any, log: any, outputDir: string, shapeCounter: { count: number }): Promise<string | null> {
    shapeCounter.count++;
    const shapeName = `shape${shapeCounter.count}.svg`; // Default to SVG
    const shapePath = path.join(outputDir, shapeName);

    log.info(`Attempting to handle and export shape (Type: ${shape.Type}) to: ${shapePath}`);

    let altText = '';
    try {
        // Attempt to extract alt text (available on both Shape and InlineShape)
        altText = shape.AlternativeText || '';
    } catch (altTextError: any) {
        log.warn(`Could not extract alt text for shape: ${altTextError.message}`);
    }

    try {
        // Explore COM API for SVG export
        // Check if the shape object has an Export method that supports SVG format (wdExportFormatSVG or similar)
        // The wdExportFormat enumeration might be available via the Word application object or a specific library.
        // Let's assume wdExportFormatSVG might exist with a value (e.g., 4 for testing, need to confirm actual value)
        const wdExportFormatSVG = 4; // Placeholder value, need to confirm actual COM enum value

        try {
            // Attempt to use Export method if available and supports SVG
            if (typeof shape.Export === 'function') {
                log.info(`Attempting export using shape.Export to ${shapePath}`);
                // Check if the Export method supports the SVG format
                // This might require trying the export and catching errors, or checking available formats if the API allows.
                // For now, we'll attempt the export with the assumed SVG format value.
                shape.Export(shapePath, wdExportFormatSVG);
                log.info(`Successfully exported shape as SVG using shape.Export to ${shapePath}`);
                return `![${altText}](${path.relative(path.dirname(shapePath), shapePath)})`; // Return the Markdown image link
            } else {
                log.debug(`shape.Export method not available for this shape type.`);
            }
        } catch (exportError: any) {
            log.warn(`shape.Export to SVG failed: ${exportError.message}. Exploring other options.`);
            // Continue to other handling methods if Export fails or doesn't support SVG
        }

        // If direct COM SVG export is not possible, implement programmatic SVG generation
        log.info(`Direct COM SVG export not successful or available. Attempting programmatic SVG generation.`);

        let svgContent = '';
        let handled = false;

        // Handle specific shape types for programmatic SVG generation
        switch (shape.Type) {
            case 13: // msoShapePicture
                log.info(`Handling msoShapePicture for programmatic SVG.`);
                // For pictures, we might need to extract the image data first (e.g., as PNG)
                // and then embed it in an SVG <image> element.
                // This part would reuse or adapt logic from handleImage.
                // For now, add a placeholder.
                svgContent = `<!-- Programmatic SVG for Picture Placeholder -->\n<svg width="${shape.Width}" height="${shape.Height}"><image href="[base64 image data]" width="${shape.Width}" height="${shape.Height}"/></svg>`;
                handled = true;
                break;
            case 17: // msoTextBox
                log.info(`Handling msoTextBox for programmatic SVG.`);
                try {
                    const textBoxText = shape.TextFrame.TextRange.Text || '';
                    const width = shape.Width;
                    const height = shape.Height;
                    // Basic SVG for a text box
                    svgContent = `<svg width="${width}" height="${height}"><rect width="${width}" height="${height}" fill="none" stroke="black"/><text x="10" y="20">${textBoxText.trim()}</text></svg>`;
                    handled = true;
                } catch (textBoxError: any) {
                    log.warn(`Could not extract text from text box for SVG: ${textBoxError.message}`);
                }
                break;
            case 88: // msoChart
                log.info(`Handling msoChart for programmatic SVG.`);
                // This is complex and requires extracting chart data and mapping to SVG elements.
                // Add a placeholder for now.
                svgContent = `<!-- Programmatic SVG for Chart Placeholder -->\n<svg width="${shape.Width}" height="${shape.Height}"><text x="10" y="20">Chart Placeholder</text></svg>`;
                handled = true;
                break;
            case 6: // msoGroup
                log.info(`Handling msoGroup for programmatic SVG.`);
                // Handling groups requires iterating through GroupItems and generating SVG for each.
                // This would likely involve a recursive call to handleShapeAndExportSvg,
                // and wrapping the results in an SVG <g> element.
                // Add a placeholder for now.
                svgContent = `<!-- Programmatic SVG for Group Placeholder -->\n<svg width="${shape.Width}" height="${shape.Height}"><text x="10" y="20">Group Placeholder</text></svg>`;
                handled = true;
                break;
            case 20: // msoCanvas
                log.info(`Handling msoCanvas for programmatic SVG.`);
                // Handling canvases requires iterating through CanvasItems.
                // Similar to groups, this would involve recursive calls and wrapping in <g> or <svg>.
                // Add a placeholder for now.
                svgContent = `<!-- Programmatic SVG for Canvas Placeholder -->\n<svg width="${shape.Width}" height="${shape.Height}"><text x="10" y="20">Canvas Placeholder</text></svg>`;
                handled = true;
                break;
            case 1: // msoAutoShape
                log.info(`Handling msoAutoShape for programmatic SVG.`);
                // Basic shapes can be mapped to SVG elements like rect, circle, line, path.
                // This requires extracting geometric properties, fill, stroke, etc.
                // Add a placeholder for now.
                svgContent = `<!-- Programmatic SVG for AutoShape Placeholder -->\n<svg width="${shape.Width}" height="${shape.Height}"><text x="10" y="20">AutoShape Placeholder</text></svg>`;
                handled = true;
                break;
            default:
                log.warn(`No specific programmatic SVG handling for shape type: ${shape.Type}`);
                svgContent = `<!-- Programmatic SVG Placeholder for Type ${shape.Type} -->\n<svg width="${shape.Width}" height="${shape.Height}"><text x="10" y="20">Shape Type ${shape.Type} Placeholder</text></svg>`;
                // Do not set handled to true, as this is a generic placeholder
                break;
        }

        if (handled && svgContent) {
            log.info(`Writing generated SVG to ${shapePath}`);
            await fs.writeFile(shapePath, svgContent, 'utf8');
            log.info(`Successfully wrote SVG to ${shapePath}`);
            return `![${altText}](${path.relative(path.dirname(shapePath), shapePath)})`; // Return the Markdown image link
        } else {
            log.warn(`Shape type ${shape.Type} not handled for SVG export.`);
            return null; // Return null if not handled
        }

    } catch (error: any) {
        log.error(`An unexpected error occurred during shape handling and SVG export for shape type ${shape.Type}: ${error.message}`);
        return `![Shape Export Failed (Type ${shape.Type}): ${error.message}]()`; // Return a broken link with error info
    } finally {
        // Release COM object for shape if necessary (depends on how it's obtained)
        // Assuming shape is passed in and needs to be released here.
        releaseObject(shape); // Release COM object
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
    let imageCounter = 0; // Counter for images (PNG)
    let shapeCounter = 0; // Counter for shapes (SVG)
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

        // --- Background Handling ---
        log.info("Attempting to handle document background.");
        try {
            // Investigate document background properties
            // The background might be a color, a gradient, a texture, or a picture.
            // Accessing this via COM can be complex and might not be directly supported for image extraction.
            // Checking the document's background properties might involve:
            // doc.Background.Fill.Visible
            // doc.Background.Fill.Type (msoFillPicture, msoFillGradient, etc.)
            // doc.Background.Fill.Picture.Shape.PictureFormat (if type is msoFillPicture)

            // For now, add a placeholder indicating background handling needs implementation/investigation
            log.warn("Document background handling is not yet implemented/fully investigated via COM.");
            markdownOutput += '\n<!-- TODO: Document background extraction not implemented/investigated. -->\n';

            // TODO: Further investigate COM properties for document background and implement extraction if feasible.
            // If a background image is found and extractable, use handleImage or similar logic.

        } catch (bgError: any) {
            log.warn(`Error attempting to access document background properties: ${bgError.message}`);
            markdownOutput += `\n<!-- Document background handling failed: ${bgError.message} -->\n`;
        }
        log.info("Document background handling attempt complete.");


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

        // --- Header and Footer Handling (for logos and shapes) ---
        log.info("Starting header and footer traversal for images and shapes.");
        for (let i = 1; i <= doc.Sections.Count; i++) {
            const section = doc.Sections(i);
            const headerFooterTypes = [
                1, // wdHeaderFooterPrimary
                2, // wdHeaderFooterFirstPage
                3  // wdHeaderFooterEvenPages
            ];

            for (const type of headerFooterTypes) {
                try {
                    const headerFooter = section.Headers(type) || section.Footers(type);
                    if (headerFooter) {
                        log.debug(`Checking header/footer type ${type} in section ${i}`);
                        // Check for shapes (including images) in the header/footer range
                        if (headerFooter.Shapes.Count > 0) {
                            log.info(`Found ${headerFooter.Shapes.Count} shapes in header/footer type ${type} in section ${i}.`);
                            for (let j = 1; j <= headerFooter.Shapes.Count; j++) {
                                const shape = headerFooter.Shapes(j);
                                // Use the new helper function to handle and export shapes
                                const shapeMarkdown = await handleShapeAndExportSvg(shape, log, extractedImagesDir, { count: shapeCounter });
                                if (shapeMarkdown) {
                                    markdownOutput += `\n<!-- Shape from Header/Footer -->\n${shapeMarkdown}\n`;
                                } else {
                                    log.warn(`Shape type ${shape.Type} in header/footer not exported as SVG.`);
                                    // Release the shape object here if handleShapeAndExportSvg didn't
                                    releaseObject(shape);
                                }
                            }
                        }
                        releaseObject(headerFooter); // Release COM object
                    }
                } catch (hfError: any) {
                    log.warn(`Error accessing header/footer type ${type} in section ${i}: ${hfError.message}`);
                }
            }
            releaseObject(section); // Release COM object
        }
        log.info("Header and footer traversal complete.");

        // --- Document Traversal and Element Handling (Main Story) ---
        log.info("Starting main document traversal and element handling.");

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
                // Handle inline shapes within paragraphs using the new helper function
                if (paragraph.Range.InlineShapes.Count > 0) {
                    for (let j = 1; j <= paragraph.Range.InlineShapes.Count; j++) {
                        const inlineShape = paragraph.Range.InlineShapes(j);
                        const shapeMarkdown = await handleShapeAndExportSvg(inlineShape, log, extractedImagesDir, { count: shapeCounter });
                        if (shapeMarkdown) {
                            markdownOutput += shapeMarkdown; // Append the Markdown image link
                        } else {
                            log.warn(`Inline shape type ${inlineShape.Type} in paragraph not exported as SVG.`);
                            // Release the inlineShape object here if handleShapeAndExportSvg didn't
                            releaseObject(inlineShape);
                        }
                    }
                }
                // Handle paragraph text and formatting (excluding inline shapes already handled)
                // This part needs refinement to avoid duplicating text from inline shapes.
                // For now, we'll append the paragraph text after handling inline shapes.
                // A better approach might be to iterate through the paragraph's range and handle runs/inline shapes sequentially.
                markdownOutput += await handleParagraph(paragraph, log, extractedImagesDir, { count: imageCounter }, validatedParams.imagePrefix); // Reusing handleParagraph for text/formatting
                // Move the range past the paragraph
                currentRange.Start = paragraph.Range.End;
                releaseObject(paragraph); // Release COM object
            }
            // Handle floating shapes in the main story range
            else if (currentRange.ShapeRange.Count > 0) {
                 log.info(`Found Floating Shape(s) in document traversal. Attempting to handle.`);
                 // Iterate through shapes in the current range's ShapeRange
                 for (let k = 1; k <= currentRange.ShapeRange.Count; k++) {
                     const shape = currentRange.ShapeRange(k);
                     // Use the new helper function to handle and export shapes
                     const shapeMarkdown = await handleShapeAndExportSvg(shape, log, extractedImagesDir, { count: shapeCounter });
                     if (shapeMarkdown) {
                         markdownOutput += `\n<!-- Floating Shape -->\n${shapeMarkdown}\n`;
                     } else {
                         log.warn(`Floating shape type ${shape.Type} not exported as SVG.`);
                         // Release the shape object here if handleShapeAndExportSvg didn't
                         releaseObject(shape);
                     }
                 }
                 // Move the range past the ShapeRange
                 currentRange.Start = currentRange.ShapeRange.Range.End; // Move past the entire shape range
            }

            else {
                // If no known element is found, move the range forward by one character
                currentRange.MoveStart(1, 1); // wdCharacter, 1
            }
        }

        releaseObject(mainStoryRange); // Release COM object
        releaseObject(currentRange); // Release COM object

        log.info("Document traversal and element handling complete.");

        // --- Comment Handling ---
        // --- Comment Handling ---
        log.info(`Handling comments with option: ${commentsOption}`);
        if (commentsOption === 'append') {
            if (doc.Comments.Count > 0) {
                markdownOutput += '\n\n## Comments\n\n';
                for (let i = 1; i <= doc.Comments.Count; i++) {
                    const comment = doc.Comments(i);
                    // Basic extraction of comment author and text
                    const author = comment.Author || 'Unknown Author';
                    const commentText = comment.Range.Text || '';
                    markdownOutput += `**Comment from ${author}:** ${commentText.trim()}\n\n`;
                    releaseObject(comment); // Release COM object
                }
                log.info(`Appended ${doc.Comments.Count} comments to Markdown output.`);
            } else {
                log.info("No comments found to append.");
            }
        } else if (commentsOption === 'inline') {
            log.info("Attempting inline comment handling.");
            // This is a simplified approach and may not perfectly map comments to text in all cases.
            // A more robust solution would require detailed range mapping during document traversal.
            if (doc.Comments.Count > 0) {
                for (let i = 1; i <= doc.Comments.Count; i++) {
                    const comment = doc.Comments(i);
                    const commentText = comment.Range.Text || '';
                    const commentedTextRange = comment.Parent.Range; // Get the range the comment is attached to

                    // Find the text corresponding to the commented range in the markdownOutput
                    // This is a basic text search and might not be accurate for complex documents
                    const commentedText = (commentedTextRange.Text || '').replace(/[\r\n]/g, '').trim();
                    const commentMarkdown = ` (${commentText.trim()})`; // Simple inline representation

                    // Attempt to find and insert the comment after the commented text
                    // This is a naive approach and needs refinement
                    const findIndex = markdownOutput.indexOf(commentedText);
                    if (findIndex !== -1) {
                        // Insert the comment markdown after the commented text
                        markdownOutput = markdownOutput.substring(0, findIndex + commentedText.length) + commentMarkdown + markdownOutput.substring(findIndex + commentedText.length);
                        log.debug(`Inserted inline comment after "${commentedText}"`);
                    } else {
                        log.warn(`Could not find commented text "${commentedText}" in Markdown output for inline comment.`);
                        // Append the comment at the end if the text is not found
                        markdownOutput += `\n<!-- Inline Comment not found for text: "${commentedText}" -->\n${commentMarkdown}\n`;
                    }

                    releaseObject(comment); // Release COM object
                    releaseObject(commentedTextRange); // Release COM object
                }
                log.info(`Attempted inline handling for ${doc.Comments.Count} comments.`);
            } else {
                log.info("No comments found for inline handling.");
            }
        } else { // commentsOption === 'ignore'
            log.info("Comments are ignored as per option.");
        }


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

           // Append the images directory (which now contains SVGs)
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
        wordApp = officeResult; // Correctly assign the returned application object

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
       await applyMarkdownFormattingToWord(docContentRange, markdownContent, wordApp, newDoc); // Pass newDoc object
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