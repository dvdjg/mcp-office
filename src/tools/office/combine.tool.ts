/**
 * @file Tool for combining multiple files from a directory into a single Word document.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams } from '../../types/common.types.js'; // Import McpResource and necessary types
import path from 'path';
import fs from 'fs/promises';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js'; // Import specific officeInterop functions
import pdfParseTool from './pdfParse.tool.js'; // Import the pdfParse tool by default
import logger from '../../utils/logger.js'; // Import logger

// Input schema for the office/combine tool
const CombineToolInputSchema = z.object({
  directoryPath: z.string().describe('Path to the directory with source files.'),
  outputFilePath: z.string().describe('Path where the combined Word document will be saved.'),
  operation: z.enum(['read', 'normalize', 'insert']).describe('Operation to perform: read, normalize, or insert.'),
  filePatterns: z.array(z.string()).optional().describe('Glob patterns to filter files.'),
  formatsToInclude: z.array(z.string()).optional().describe('File formats to include (e.g., ["pdf", "docx"]).'),
});

type CombineToolInput = z.infer<typeof CombineToolInputSchema>;

/**
 * @tool office/combine
 * @description Combines multiple files from a directory into a single Word document.
 * @operations read, normalize, insert
 * @input CombineToolInputSchema
 * @output string (status message or output file path)
 */
export const officeCombineTool: McpResource = { // Use McpResource
  path: 'office/combine', // Add the path property
  description: 'Combines multiple files from a directory into a single Word document.',
  schema: CombineToolInputSchema, // Change inputSchema to schema
  handler: async (params: ToolRequestParams, context?: any): Promise<ApiResponse<string>> => { // Adjust handler signature and return type
    const input = params as CombineToolInput; // Cast params to the expected type
    const { directoryPath, outputFilePath, operation, filePatterns, formatsToInclude } = input;

    try {
      switch (operation) {
        case 'read': {
          logger.info(`[office/combine] Executing 'read' operation on ${directoryPath}`);
          const files = await fs.readdir(directoryPath);
          let filteredFiles = files;

          if (filePatterns) {
            // TODO: Implement glob pattern filtering if necessary
            logger.warn('[office/combine] Filtering by filePatterns not implemented yet.');
          }

          if (formatsToInclude) {
            filteredFiles = filteredFiles.filter(file => {
              const ext = path.extname(file).toLowerCase().replace('.', '');
              return formatsToInclude.includes(ext);
            });
          }

          logger.info(`[office/combine] Found ${filteredFiles.length} files.`);
          return { success: true, data: `Files found in ${directoryPath}: ${filteredFiles.join(', ')}` };
        }

        case 'normalize': {
          logger.info(`[office/combine] Executing 'normalize' operation on ${directoryPath}`);
          const filesToNormalize = await fs.readdir(directoryPath);
          let filteredFilesToNormalize = filesToNormalize;

          if (filePatterns) {
             logger.warn('[office/combine] Filtering by filePatterns not implemented yet.');
          }

          if (formatsToInclude) {
            filteredFilesToNormalize = filteredFilesToNormalize.filter(file => {
              const ext = path.extname(file).toLowerCase().replace('.', '');
              return formatsToInclude.includes(ext);
            });
          }

          const normalizedContent: { fileName: string, content: string }[] = [];

          for (const file of filteredFilesToNormalize) {
            const fullPath = path.join(directoryPath, file);
            const ext = path.extname(file).toLowerCase().replace('.', '');

            try {
              if (ext === 'pdf') {
                logger.info(`[office/combine] Normalizing PDF: ${file}`);
                // Use the pdfParseTool handler
                const pdfParseResult = await pdfParseTool.handler({ filePath: fullPath, operation: 'parse' });
                if (pdfParseResult.success) {
                  normalizedContent.push({ fileName: file, content: pdfParseResult.data as string });
                } else {
                  logger.error(`[office/combine] Error parsing PDF ${file}: ${pdfParseResult.error?.message}`);
                  // Decide whether to throw an error or continue with the next file
                }
              } else if (ext === 'docx' || ext === 'doc') {
                 logger.info(`[office/combine] Normalizing DOCX/DOC: ${file}`);
                 // TODO: Implement text and possibly other element extraction from DOCX/DOC using COM Interop
                 // This is complex and requires interacting with the Word API via winax.
                 // For now, we will just add a placeholder.
                 normalizedContent.push({ fileName: file, content: `[Content of ${file} - DOCX/DOC normalization not fully implemented]` });
              } else {
                 logger.warn(`[office/combine] Unsupported format for normalization: ${file}`);
                 // Optional: read as plain text if it's a known text file
                 // try {
                 //    const textContent = await fs.readFile(fullPath, 'utf-8');
                 //    normalizedContent.push({ fileName: file, content: textContent });
                 // } catch (readError) {
                 //    logger.warn(`[office/combine] Could not read ${file} as plain text.`);
                 // }
              }
            } catch (fileError: any) {
               logger.error(`[office/combine] Error processing file ${file}: ${fileError.message}`);
               // Continue with the next file
            }
          }

          // Here you could temporarily save the normalized content if the 'insert' operation is called separately
          // For now, we just return a summary.
          logger.info(`[office/combine] Normalization completed for ${normalizedContent.length} files.`);
          return { success: true, data: `Normalization completed. Normalized content for: ${normalizedContent.map(item => item.fileName).join(', ')}` };
        }

        case 'insert': {
          logger.info(`[office/combine] Executing 'insert' operation into ${outputFilePath}`);
          // TODO: Implement insertion logic using COM Interop (winax)
          // This would involve:
          // 1. Creating or opening a Word document at outputFilePath.
          // 2. Iterating over the normalized content (we would need a way to pass it or normalize it here again).
          // 3. Inserting the content of each file into the Word document.
          // 4. Saving and closing the Word document.

          // For insertion, we probably need to normalize the files again or expect the normalized content to be passed as a parameter.
          // For initial simplicity, we will normalize here again (less efficient but simpler to start).
           const filesToInsert = await fs.readdir(directoryPath);
           let filteredFilesToInsert = filesToInsert;

          if (filePatterns) {
             logger.warn('[office/combine] Filtering by filePatterns not implemented yet.');
          }

          if (formatsToInclude) {
            filteredFilesToInsert = filteredFilesToInsert.filter(file => {
              const ext = path.extname(file).toLowerCase().replace('.', '');
              return formatsToInclude.includes(ext);
            });
          }

          let wordApp: any = null;
          let doc: any = null;

          try {
            wordApp = await getOfficeApplication('Word.Application');
            // Create a new document or open an existing one? The plan says "combine them into a single Word document",
            // which suggests creating a new one if outputFilePath does not exist, or adding to an existing one.
            // For now, we will create a new one.
            logger.debug('[office/combine] Creating new Word document.');
            doc = wordApp.Documents.Add();

            for (const file of filteredFilesToInsert) {
              const fullPath = path.join(directoryPath, file);
              const ext = path.extname(file).toLowerCase().replace('.', '');
              logger.info(`[office/combine] Inserting content from: ${file}`);

              // Move the cursor to the end of the document before inserting
              doc.Content.Collapse(0); // wdCollapseEnd = 0
              const endRange = doc.Content;
              endRange.Collapse(0); // Ensure the range is at the end

              try {
                 if (ext === 'pdf') {
                    logger.info(`[office/combine] Inserting PDF content: ${file}`);
                    // For PDFs, we could try to insert the parsed text or insert the PDF as an object.
                    // Inserting as an object can maintain formatting but depends on installed viewers.
                    // Inserting parsed text loses formatting.
                    // Option 1: Insert parsed text (requires prior normalization or here)
                    const pdfParseResult = await pdfParseTool.handler({ filePath: fullPath, operation: 'parse' });
                    if (pdfParseResult.success) {
                       endRange.InsertAfter(pdfParseResult.data as string + '\n\n'); // Add line breaks between files
                    } else {
                       logger.error(`[office/combine] Error parsing PDF for insertion ${file}: ${pdfParseResult.error?.message}`);
                    }

                    // Option 2: Insert the file as an object (maintains format but requires viewer)
                    // logger.debug(`[office/combine] Attempting to insert PDF as object: ${file}`);
                    // // wdInsertObject = 0, wdFloatOverText = 0 (for InlineShape)
                    // endRange.InlineShapes.AddOLEObject(fullPath, undefined, false, false, undefined, undefined, undefined, 0);
                    // endRange.InsertParagraphAfter(); // Add a paragraph break after the object
                    // logger.debug(`[office/combine] PDF inserted as object: ${file}`);

                 } else if (ext === 'docx' || ext === 'doc') {
                    logger.info(`[office/combine] Inserting DOCX/DOC content: ${file}`);
                    // Insert the content of another Word document
                    // wdInsertFile = 4
                    endRange.InsertFile(fullPath);
                    endRange.InsertParagraphAfter(); // Add a paragraph break after the inserted content
                    logger.debug(`[office/combine] DOCX/DOC content inserted: ${file}`);

                 } else if (ext === 'txt') {
                     logger.info(`[office/combine] Inserting TXT content: ${file}`);
                     const textContent = await fs.readFile(fullPath, 'utf-8');
                     endRange.InsertAfter(textContent + '\n\n'); // Insert plain text
                     logger.debug(`[office/combine] TXT content inserted: ${file}`);
                 }
                 // TODO: Add handling for other formats if necessary (e.g., .xlsx, .pptx - can be complex)

              } catch (fileInsertError: any) {
                 logger.error(`[office/combine] Error inserting content from ${file}: ${fileInsertError.message}`);
                 // Continue with the next file
              }
            }

            // Save the final document
            logger.debug(`[office/combine] Saving final document to: ${outputFilePath}`);
            const absoluteOutputFilePath = path.resolve(outputFilePath);
            // wdFormatDocumentDefault = 16 (for .docx)
            doc.SaveAs2(absoluteOutputFilePath, 16); // Save as .docx
            logger.info(`[office/combine] Combined document saved to: ${absoluteOutputFilePath}`);

            return { success: true, data: `Combined document saved to: ${absoluteOutputFilePath}` };

          } catch (insertError: any) {
            logger.error(`[office/combine] Error during 'insert' operation: ${insertError.message}`, { error: insertError });
            return { success: false, error: { code: 'INSERT_ERROR', message: `Error during insertion: ${insertError.message}` } };
          } finally {
            // Clean up COM objects
            if (doc) {
              try {
                doc.Close(0); // wdDoNotSaveChanges = 0
                logger.debug('[office/combine] Word document closed.');
              } catch (closeError: any) {
                logger.warn(`[office/combine] Error closing Word document: ${closeError.message}`);
              }
              releaseObject(doc);
            }
            if (wordApp) {
              // Consider whether to close the Word application or leave it open.
              // If we opened it ourselves, we should probably close it if there are no other documents open.
              // This is difficult to determine safely. For now, we will leave it open.
              // try {
              //   if (wordApp.Documents.Count === 0) {
              //     wordApp.Quit();
              //     logger.debug('[office/combine] Word application closed.');
              //   }
              // } catch (quitError: any) {
              //   logger.warn(`[office/combine] Error closing Word application: ${quitError.message}`);
              // }
              releaseObject(wordApp);
            }
          }
        }

        default:
          // This should not happen if the Zod schema works correctly
          return { success: false, error: { code: 'INVALID_OPERATION', message: `Unsupported operation: ${operation}` } };
      }
    } catch (error: any) {
      // Basic error handling for unexpected errors outside the operation cases
      logger.error(`[office/combine] Unexpected error: ${error.message}`, { error });
      return { success: false, error: { code: 'UNEXPECTED_ERROR', message: `Unexpected error executing office/combine tool: ${error.message}` } };
    }
  },
};