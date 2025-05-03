/**
 * @file Implements the 'word/mermaid/import' tool using COM Interop for importing Mermaid diagrams into Word documents.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types'; // Normalized relative path
import { Context as FastMCPContext } from 'fastmcp'; // Import FastMCP Context
import mermaid from 'mermaid'; // Assuming mermaid is installed
import { handleToolError } from '../../utils/errorHandler'; // Normalized relative path
import * as fs from 'fs/promises'; // Import fs module for file handling
import * as path from 'path'; // Import path module
import os from 'os'; // Import os module for temporary directory
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Use COM Interop utilities
import { validateFilePath } from '../../utils/security'; // Import validateFilePath
import logger from '../../utils/logger'; // Normalized relative path

// Input schema for the word/mermaid/import tool
/**
 * Zod schema for the input parameters of the 'word/mermaid/import' tool.
 */
const mermaidImportInputSchema = z.object({
  /** Path to the Word document. */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  /** Mermaid diagram syntax. */
  mermaidSyntax: z.string().min(1, 'Mermaid syntax is required.').describe('Mermaid diagram syntax.'),
  /** Output format for the rendered diagram ('svg' or 'png'). */
  format: z.enum(['svg', 'png']).default('svg').describe('Output format for the rendered diagram.'),
  /** Location in the document to insert the diagram (e.g., bookmark, ContentControl). Not fully implemented. */
  location: z.string().optional().describe('Location in the document to insert the diagram (e.g., bookmark, ContentControl).'),
});

/**
 * @tool word/mermaid/import
 * @description Imports a Mermaid diagram into a Word document.
 *
 * @param {ToolRequestParams} params - Input parameters.
 * @param {string} params.filePath - Path to the Word document.
 * @param {string} params.mermaidSyntax - Mermaid diagram syntax.
 * @param {'svg' | 'png'} [params.format='svg'] - Output format for the rendered diagram.
 * @param {string} [params.location] - Location in the document to insert the diagram (e.g., bookmark, ContentControl).
 * @param {FastMCPContext} [context] - The FastMCP context (optional).
 *
 * @returns {Promise<ApiResponse<object>>} - Result of the operation.
 */
const mermaidImportTool: McpResource = {
  path: 'word/mermaid/import', // Define the path here
  description: 'Imports a Mermaid diagram into a Word document.',
  schema: mermaidImportInputSchema, // Use schema for input validation
 handler: async (params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<object>> => { // Use ToolRequestParams and ApiResponse
   const log = context?.log ?? logger; // Use context logger or fallback

   let wordApp: any = null;
   let doc: any = null;
   let officeAppInstance: any = null; // To manage the application instance lifecycle
   let tempFilePath: string | null = null;

   try {
     const validatedParams = mermaidImportInputSchema.parse(params);
     const { filePath, mermaidSyntax, format, location } = validatedParams;

     // Validate file path (already done by schema refinement, but explicit call is fine)
     const absoluteFilePath = validateFilePath(filePath);
     log.debug(`Validated absolute file path: ${absoluteFilePath}`);


     // Initialize Mermaid
     mermaid.initialize({ startOnLoad: false });
     log.debug('Mermaid initialized.');

     // Render the Mermaid syntax
     log.info(`Rendering Mermaid syntax (format: ${format})...`);
     const { svg } = await mermaid.render('mermaid-diagram', mermaidSyntax);
     log.info('Mermaid syntax rendered to SVG.');

     // Create temporary file for the image
     const tempDir = os.tmpdir();
     const fileExtension = format === 'svg' ? '.svg' : '.png'; // Assume if not SVG, it's PNG (although current rendering is only SVG)
     tempFilePath = path.join(tempDir, `mermaid_diagram_${Date.now()}${fileExtension}`);
     log.debug(`Temporary file path: ${tempFilePath}`);

     // If the format is PNG, we would need to convert the SVG to PNG here.
     // For now, we only save the SVG.
     if (format === 'svg') {
        await fs.writeFile(tempFilePath, svg);
        log.debug(`Rendered SVG saved to temporary file.`);
     } else if (format === 'png') {
         // TODO: Implement SVG to PNG conversion
         log.warn('PNG format requested, but SVG to PNG conversion is not implemented. Saving as SVG.');
         await fs.writeFile(tempFilePath, svg); // Save as SVG for now
     }


     // Open the Word document
     log.info(`Opening document: ${absoluteFilePath}`);
     officeAppInstance = await getOfficeApplication('Word.Application');
     wordApp = officeAppInstance.app;
     wordApp.Visible = false; // Run in background
     wordApp.DisplayAlerts = 0; // wdAlertsNone = 0
     doc = wordApp.Documents.Open(absoluteFilePath);
     log.info(`Document opened successfully.`);

     // Determine insertion location. For now, insert at the end of the document.
     // TODO: Implement logic to insert at a specific location (bookmark, ContentControl).
     log.warn(`Insertion location '${location || 'end of document'}' not fully implemented. Inserting at the end of the document.`);
     const selection = wordApp.Selection;
     selection.EndKey(6); // Move cursor to the end of the document (wdStory = 6)

     // Insert the image
     // Use AddPicture to insert from a file. LinkToFile and SaveWithDocument set to false and true respectively.
     // The range is the current selection.
     const range = selection.Range;
     const inlineShape = range.InlineShapes.AddPicture(
       tempFilePath,
       false, // LinkToFile
       true // SaveWithDocument
     );
     log.debug(`Image inserted from temporary file.`);


     // Optional: Adjust size or position if needed
     // inlineShape.Width = 300;
     // inlineShape.Height = 200;

     // Save and close the document
     doc.Save();
     log.info(`Document saved: ${absoluteFilePath}`);
     doc.Close();
     log.info(`Document closed.`);

     releaseObject(inlineShape); // Release inlineShape object
     releaseObject(range); // Release range object
     releaseObject(selection); // Release selection object
     releaseObject(doc); // Release document object


     log.info(`Mermaid diagram inserted into ${filePath}`);

     return { success: true, data: { message: `Mermaid diagram inserted into ${filePath}` } };
   } catch (error: any) {
     log.error(`Error importing Mermaid diagram: ${error.message}`, { error });
     // Ensure document is closed if it was opened
     if (doc) {
         try { doc.Close(false); } catch (e: any) { log.warn(`Error closing document during error handling: ${e.message}`); }
         releaseObject(doc);
     }
     return handleToolError(error, 'MERMAID_IMPORT_ERROR');
   } finally {
     // Clean up temporary file
     if (tempFilePath) {
       try {
         await fs.unlink(tempFilePath);
         log.debug(`Temporary file cleaned up: ${tempFilePath}`);
       } catch (cleanupError: any) {
         log.error(`Error cleaning up temporary file ${tempFilePath}: ${cleanupError.message}`);
       }
     }
     // Close the Word application if it was opened
     if (officeAppInstance) {
       officeAppInstance.release();
       log.debug("Word application instance released.");
     }
   }
 },
};

export default mermaidImportTool;