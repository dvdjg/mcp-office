/**
 * @file Implements the 'word/mermaid/export' tool using COM Interop for exporting Mermaid diagrams from Word documents.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types.js'; // Normalized relative path
import { Context as FastMCPContext } from 'fastmcp'; // Import FastMCP Context
import * as fs from 'fs/promises'; // For saving the file
import * as path from 'path'; // Import path module
import { handleToolError } from '../../utils/errorHandler.js'; // Normalized relative path
import { validateFilePath } from '../../utils/security.js'; // Normalized relative path
import logger from '../../utils/logger.js'; // Normalized relative path
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js'; // Use COM Interop utilities

// Input schema for the word/mermaid/export tool
/**
 * Zod schema for the input parameters of the 'word/mermaid/export' tool.
 */
const mermaidExportInputSchema = z.object({
  /** Path to the Word document. */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  /** Identifier of the diagram to export (e.g., ContentControl name). */
  diagramIdentifier: z.string().optional().describe('Identifier of the diagram to export (e.g., ContentControl name).'),
  /** Directory where the diagram file will be saved. */
  outputDirectory: z.string().min(1, 'Output directory is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe output directory path provided.",
  }),
  /** Output format for the exported diagram ('svg', 'png', or 'txt' for syntax). */
  format: z.enum(['svg', 'png', 'txt']).default('txt').describe('Output format for the exported diagram (svg, png, or txt for syntax).'),
});

/**
 * @tool word/mermaid/export
 * @description Exports a Mermaid diagram from a Word document.
 *
 * @param {ToolRequestParams} params - Input parameters.
 * @param {string} params.filePath - Path to the Word document.
 * @param {string} [params.diagramIdentifier] - Identifier of the diagram to export (e.g., ContentControl name).
 * @param {string} params.outputDirectory - Directory where the diagram file will be saved.
 * @param {'svg' | 'png' | 'txt'} [params.format='txt'] - Output format for the exported diagram (svg, png, or txt for the syntax).
 * @param {FastMCPContext} [context] - The FastMCP context (optional).
 *
 * @returns {Promise<ApiResponse<object>>} - Result of the operation.
 */
const mermaidExportTool: McpResource = {
  path: 'word/mermaid/export', // Define the path here
  description: 'Exports a Mermaid diagram from a Word document.',
  schema: mermaidExportInputSchema, // Use schema for input validation
  handler: async (params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<object>> => { // Use ToolRequestParams and ApiResponse
    const log = context?.log ?? logger; // Use context logger or fallback

    let wordApp: any = null;
    let doc: any = null;
    let officeAppInstance: any = null; // To manage the application instance lifecycle

    try {
      const validatedParams = mermaidExportInputSchema.parse(params);
      const { filePath, diagramIdentifier, outputDirectory, format } = validatedParams;

      // Validate output directory exists or create it
      const absoluteOutputDir = path.resolve(outputDirectory);
      try {
          await fs.access(absoluteOutputDir);
          const stats = await fs.stat(absoluteOutputDir);
          if (!stats.isDirectory()) {
               throw new Error(`Output path is not a directory: ${absoluteOutputDir}`);
          }
          log.debug(`Output directory exists: ${absoluteOutputDir}`);
      } catch (error: any) {
           if (error.code === 'ENOENT') {
               log.info(`Output directory does not exist, attempting to create: ${absoluteOutputDir}`);
               await fs.mkdir(absoluteOutputDir, { recursive: true });
               log.info(`Output directory created: ${absoluteOutputDir}`);
           } else {
               log.error(`Error accessing output directory ${absoluteOutputDir}: ${error.message}`);
               throw new Error(`Error accessing output directory: ${error.message}`);
           }
      }


      // Open the Word document
      log.info(`Opening document: ${filePath}`);
      officeAppInstance = await getOfficeApplication('Word.Application');
      wordApp = officeAppInstance.app;
      wordApp.Visible = false; // Run in background
      wordApp.DisplayAlerts = 0; // wdAlertsNone = 0
      doc = wordApp.Documents.Open(filePath);
      log.info(`Document opened successfully.`);

      let extractedContent: string | Buffer | null = null;
      let foundDiagram = false;
      let diagramName = 'mermaid_diagram';

      // Logic to find the diagram
      // Prioritize searching in ContentControls if they were used for import
      // This assumes the Mermaid syntax was saved in the Tag or Title of the ContentControl
      // or as text within it.
      if (doc.ContentControls && doc.ContentControls.Count > 0) {
        log.debug(`Searching in ${doc.ContentControls.Count} ContentControls...`);
        for (let i = 1; i <= doc.ContentControls.Count; i++) {
          const cc = doc.ContentControls.Item(i);
          // We could search by a specific Tag or Title if defined during import
          // For now, we check if the text within the CC looks like Mermaid syntax
          if (cc.Range && cc.Range.Text && (cc.Range.Text.includes('graph') || cc.Range.Text.includes('sequenceDiagram'))) {
             if (!diagramIdentifier || cc.Title === diagramIdentifier || cc.Tag === diagramIdentifier || cc.Range.Text.includes(diagramIdentifier)) {
                extractedContent = cc.Range.Text;
                diagramName = cc.Title || cc.Tag || diagramName; // Use Title or Tag as name if they exist
                foundDiagram = true;
                log.debug(`Found potential Mermaid syntax in ContentControl ${i}.`);
                releaseObject(cc); // Release ContentControl object
                break; // Found the first diagram that matches or seems to be Mermaid
             }
          }
           releaseObject(cc); // Release ContentControl object if not matched
        }
      }

      // If not found in ContentControls or identifier was not specified, search in InlineShapes (images)
      if (!foundDiagram && doc.InlineShapes && doc.InlineShapes.Count > 0) {
          log.debug(`Searching in ${doc.InlineShapes.Count} InlineShapes...`);
          // If an identifier was specified, this is more complex without ContentControls.
          // We could try searching by alternative text if it was saved there.
          // Or simply iterate and if no identifier, take the first image.
          for (let i = 1; i <= doc.InlineShapes.Count; i++) {
              const shape = doc.InlineShapes.Item(i);
              // Check if it's a picture (Type 1 is msoPicture)
              if (shape.Type === 1) {
                  // If no identifier, or if the alternative text matches the identifier
                  if (!diagramIdentifier || (shape.AlternativeText && shape.AlternativeText.includes(diagramIdentifier))) {
                      // Attempt to save the image. Word COM does not have a direct method to get image bytes.
                      // One way is to copy the image, paste it into a new temporary document, and save that document as an image.
                      // This is complex. For now, we will only indicate that an image was found.
                      // TODO: Implement actual image extraction.
                      log.warn(`Found an image (InlineShape) that could be a diagram. Direct image extraction is not implemented.`);
                      // We could try saving the document temporarily as HTML or filtering the DOCX to extract the image.
                      // For now, we only mark as found if no identifier was specified or if the alternative text matches.
                      if (!diagramIdentifier || (shape.AlternativeText && shape.AlternativeText.includes(diagramIdentifier))) {
                         foundDiagram = true;
                         diagramName = shape.AlternativeText || `image_${i}`; // Use alternative text as name if it exists
                         // We cannot easily extract the actual image content here.
                         // If the requested format is 'txt', and we only found an image, we cannot fulfill the request.
                         if (format !== 'txt') {
                             // The complex logic to extract binary image would go here.
                             // For now, we just indicate it was found.
                             extractedContent = 'IMAGE_PLACEHOLDER'; // Use a placeholder
                         } else {
                             log.warn(`Requested 'txt' format but only an image was found. Cannot extract syntax.`);
                             foundDiagram = false; // Cannot fulfill the txt request
                         }
                         if (foundDiagram) {
                             releaseObject(shape); // Release shape object
                             break; // Found an image that matches or is the first without identifier
                         }
                      }
                  }
              }
               releaseObject(shape); // Release shape object if not matched
          }
      }


      if (foundDiagram && extractedContent !== null) {
        const fileName = diagramIdentifier ? `${diagramIdentifier}.${format}` : `${diagramName}.${format}`;
        const outputPath = path.join(absoluteOutputDir, fileName);

        if (format === 'txt') {
          if (typeof extractedContent === 'string') {
             await fs.writeFile(outputPath, extractedContent, 'utf-8');
             log.info(`Mermaid syntax exported to ${outputPath}`);
          } else {
             log.error(`Error: Requested 'txt' format but extracted content is not text.`);
             // We could throw an error here or return a specific failure result
             return { success: false, error: { code: 'EXTRACTION_ERROR', message: 'Extracted content is not text.' } };
          }
        } else { // svg or png
          // If extractedContent is 'IMAGE_PLACEHOLDER', it means we found an image but didn't extract it.
          if (extractedContent === 'IMAGE_PLACEHOLDER') {
              log.error(`Error: Binary image extraction is not implemented.`);
              return { success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Binary image extraction is not implemented.' } };
          }
          // If extractedContent is the syntax (extracted from ContentControl), we need to render it and save it as an image.
          else if (typeof extractedContent === 'string') {
              try {
                  // Render the syntax to SVG
                  const mermaid = require('mermaid'); // Import mermaid here to use it only when necessary
                  mermaid.initialize({ startOnLoad: false });
                  const { svg } = await mermaid.render('mermaid-diagram-export', extractedContent);

                  // If the format is PNG, we would need to convert the SVG to PNG here.
                  // For now, we only save the SVG if SVG is requested, or indicate that PNG is not implemented.
                  if (format === 'svg') {
                      await fs.writeFile(outputPath, svg);
                      log.info(`Mermaid diagram rendered and exported as SVG to ${outputPath}`);
                  } else if (format === 'png') {
                      log.error(`Error: SVG to PNG conversion is not implemented.`);
                      return { success: false, error: { code: 'NOT_IMPLEMENTED', message: 'SVG to PNG conversion is not implemented.' } };
                  }

              } catch (renderError: any) {
                  log.error(`Error rendering diagram for export: ${renderError.message}`);
                  return handleToolError(renderError, 'MERMAID_RENDER_ERROR');
              }
          } else {
               log.error(`Error: Unknown extracted content type.`);
               return { success: false, error: { code: 'EXTRACTION_ERROR', message: 'Unknown extracted content type.' } };
          }
        }
      } else {
        log.warn(`Diagram with identifier not found: ${diagramIdentifier || 'none specified'}`);
        return { success: false, error: { code: 'NOT_FOUND', message: `Diagram with identifier not found: ${diagramIdentifier || 'none specified'}.` } };
      }

      doc.Close();
      releaseObject(doc); // Release document object

      return { success: true, data: { message: `Export operation completed for ${filePath}.` } };
    } catch (error: any) {
      log.error(`Error exporting Mermaid diagram: ${error.message}`, { error });
      // Ensure document is closed if it was opened
      if (doc) {
          try { doc.Close(false); } catch (e: any) { log.warn(`Error closing document during error handling: ${e.message}`); }
          releaseObject(doc);
      }
      return handleToolError(error, 'MERMAID_EXPORT_ERROR');
    } finally {
      // Close the Word application if it was opened
      if (officeAppInstance) {
        officeAppInstance.release();
        log.debug("Word application instance released.");
      }
    }
  },
};

export default mermaidExportTool;