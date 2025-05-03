/**
 * @file Implements the 'word/merge' tool using COM Interop for merging multiple Word documents.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Normalized relative path
import { validateFilePath } from '../../utils/security'; // Normalized relative path
import { ApiResponse, McpResource, ToolRequestParams } from '../../types/common.types'; // Normalized relative path
import { saveResource } from '../dynamic/resources.tool'; // Normalized relative path
import { Context as FastMCPContext } from 'fastmcp'; // Import FastMCP Context
import { handleToolError, createErrorResponse } from '../../utils/errorHandler'; // Normalized relative path
import path from 'path';
import logger from '../../utils/logger'; // Normalized relative path

// 1. Define the Input Schema
/**
 * Zod schema for the input parameters of the 'word/merge' tool.
 */
const mergeSchema = z.object({
    /** An array of file paths to the Word documents (.docx) to be merged. Requires at least two documents. */
    docs: z.array(z.string().min(1, 'Document path cannot be empty.')).min(2, { message: 'At least two documents are required for merging.' }),
    /** The path where the merged output Word document (.docx) will be saved. Must end with .docx. */
    output: z.string().min(1, 'Output path cannot be empty.').refine(val => val.toLowerCase().endsWith('.docx'), {
        message: 'Output path must end with .docx',
    }).refine(validateFilePath, { // Validate the output path itself
        message: "Invalid or potentially unsafe output file path provided.",
    }),
});

/**
 * Infers the type for the validated merge parameters.
 */
type MergeParams = z.infer<typeof mergeSchema>;

// 2. Implement the Handler `mergeDocuments` accepting an optional FastMCPContext
/**
 * Merges multiple Word documents into a single new document using COM Interop.
 * @param params - The parameters for the tool, validated against `mergeSchema`.
 * @param context - The FastMCP context (optional), providing logging and progress reporting.
 * @returns A promise resolving to an ApiResponse containing the output file path.
 * @throws {Error} If validation fails, a COM error occurs, file access fails, or document processing fails.
 */
export async function mergeDocuments(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{ outputPath: string }>> {
    // Use context logger if available, otherwise fallback to global logger
    const log = context?.log ?? logger;
    const reportProgress = context?.reportProgress; // Get reportProgress function if context exists

    // Validate context has necessary functions if context is provided
    if (context && (!context.log || typeof context.reportProgress !== 'function')) {
        // Log error using the determined logger
        log.error('Merge tool received context but it is missing required properties (log, reportProgress).');
        // Avoid returning ApiResponse directly, throw error for FastMCP handler
        throw new Error('Tool context is missing required properties (log, reportProgress).');
        // return createErrorResponse('CONFIGURATION_ERROR', 'Tool context is missing required properties (log, reportProgress).');
    }

    let wordApp: any = null;
    let targetDoc: any = null;
    const sourceDocs: any[] = []; // To keep track of open source documents
    let officeAppInstance: any = null; // Declare officeAppInstance outside the try

    try {
        // Validate and parse parameters
        const validatedParams = mergeSchema.parse(params);
        log.info(`Validating input parameters for merge operation.`);

        // Validate file paths
        const safeOutputPath = validatedParams.output; // Already validated by Zod

        const safeSourcePaths: string[] = [];
        for (const docPath of validatedParams.docs) {
            const safePath = validateFilePath(docPath); // Validate each source path
            if (!safePath) {
                 // validateFilePath should throw an error if validation fails,
                 // so this check might be redundant, but kept for safety.
                throw new Error(`Source path validation failed for: ${docPath}`);
            }
            safeSourcePaths.push(safePath);
        }
        log.info(`All file paths validated successfully.`);

        // Get Word application instance
        log.info('Getting Word application instance...');
        officeAppInstance = await getOfficeApplication('Word.Application'); // Assign here
        wordApp = officeAppInstance.app;
        wordApp.Visible = false; // Run in background
        wordApp.DisplayAlerts = 0; // wdAlertsNone = 0

        // Create target document
        log.info('Creating target document...');
        targetDoc = wordApp.Documents.Add();

        // Iterate and merge documents
        const totalDocs = safeSourcePaths.length;
        log.info(`Starting merge process for ${totalDocs} documents...`);
        // Using correct { progress, total } signature from documentation, only if reportProgress is available
        reportProgress?.({ progress: 0, total: totalDocs });

        for (let i = 0; i < totalDocs; i++) {
            const sourcePath = safeSourcePaths[i];
            const currentDocNum = i + 1;
            log.info(`Processing document ${currentDocNum}/${totalDocs}: ${sourcePath}`);
            // Using correct { progress, total } signature, only if reportProgress is available
            reportProgress?.({ progress: i, total: totalDocs });

            let sourceDoc: any = null;
            try {
                sourceDoc = wordApp.Documents.Open(sourcePath, false, true); // Open(FileName, ConfirmConversions=false, ReadOnly=true)
                sourceDocs.push(sourceDoc); // Add to the list for later release

                // Select and copy content
                // sourceDoc.Content.Select(); // Select can be problematic sometimes, use Range
                const sourceRange = sourceDoc.Content;
                sourceRange.Copy();


                // Paste into the target document
                targetDoc.Activate();
                // Move to the end of the document before pasting
                const endOfDocRange = targetDoc.Content;
                endOfDocRange.Collapse(0); // 0 = wdCollapseEnd
                // wordApp.Selection.SetRange(endOfDocRange.End, endOfDocRange.End); // Use Range.Paste instead of Selection
                endOfDocRange.Paste();


                // Insert page break after each document except the last one
                if (i < totalDocs - 1) {
                    // wordApp.Selection.InsertBreak(7); // 7 = wdPageBreak - Use Range
                    const breakRange = targetDoc.Content;
                    breakRange.Collapse(0); // Collapse to end
                    breakRange.InsertBreak(7); // Insert page break at the end
                }

                // Close source document (without saving changes)
                sourceDoc.Close(false); // wdDoNotSaveChanges = 0
                // Remove from the list of open documents as it was closed successfully
                const indexToRemove = sourceDocs.indexOf(sourceDoc);
                if (indexToRemove > -1) {
                    sourceDocs.splice(indexToRemove, 1);
                }
                releaseObject(sourceDoc); // Release the COM object of the closed source document
                sourceDoc = null; // Ensure the variable is clean
                 log.info(`Document ${currentDocNum}/${totalDocs} processed and closed.`);
                 // Using correct { progress, total } signature, only if reportProgress is available
                 reportProgress?.({ progress: currentDocNum, total: totalDocs });

            } catch (sourceDocError: any) {
                 log.error(`Error processing source document ${sourcePath}: ${sourceDocError.message || sourceDocError}`);
                 // Attempt to close the source document if it's still open before re-throwing
                 if (sourceDoc) {
                     try {
                         sourceDoc.Close(false);
                         const indexToRemove = sourceDocs.indexOf(sourceDoc);
                         if (indexToRemove > -1) {
                             sourceDocs.splice(indexToRemove, 1);
                         }
                         releaseObject(sourceDoc);
                     } catch (closeError: any) {
                         log.error(`Failed to close source document ${sourcePath} after error: ${closeError.message || closeError}`);
                     }
                 }
                throw sourceDocError; // Re-throw the error to be caught by the main catch block
            }
        }

        // Save the merged document
        log.info(`Saving merged document to: ${safeOutputPath}`);
        // Using correct { progress, total } signature - representing the saving step, only if reportProgress is available
        reportProgress?.({ progress: totalDocs, total: totalDocs });
        // Ensure the directory exists before saving
        // Note: COM might handle this, but being explicit is safer if possible.
        // However, creating directories from here might require additional permissions
        // or logic outside the scope of officeInterop. Assume the allowed base directory exists.
        targetDoc.SaveAs2(safeOutputPath);
        log.info(`Merged document saved successfully.`);

        // Read the content of the merged document before closing it
        let mergedContent = '';
        try {
             // Read the text from the COM document object
             mergedContent = targetDoc.Content.Text;
             log.info(`Read content from merged document.`);
         } catch (readContentError: any) {
             log.error(`Failed to read content from merged document before closing: ${readContentError.message}`);
             // Do not throw error here, attempt to save the resource empty or with error
         }

        // Close the target document
        targetDoc.Close(false); // wdDoNotSaveChanges = 0
        // Explicitly release the targetDoc object now that it's closed and saved
        releaseObject(targetDoc);
        targetDoc = null;

        // Save the merged document content as a dynamic resource after closing it
        try {
            await saveResource('word/merge', path.basename(safeOutputPath), mergedContent);
            log.info(`Saved ${safeOutputPath} as a dynamic resource.`);
        } catch (resourceSaveError: any) {
            log.error(`Failed to save ${safeOutputPath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continue execution even if resource saving fails
        }


        // Using correct { progress, total } signature - representing completion, only if reportProgress is available
        reportProgress?.({ progress: totalDocs, total: totalDocs });

        return {
            success: true,
            data: { outputPath: safeOutputPath },
            message: `Successfully merged ${totalDocs} documents into ${safeOutputPath}.`
        };

    } catch (error: any) {
        // Use context.log for error logging
        log.error(`Error in mergeDocuments: ${error.message || error}`, { error }); // Log the full error object
        // Ensure the target document is closed if it was created and not closed/released before
        if (targetDoc) {
            try {
                targetDoc.Close(false);
            } catch (closeError: any) {
                 log.error(`Failed to close target document after error: ${closeError.message || closeError}`);
            }
        }
        // Let the main execute wrapper in server/index.ts handle the error conversion
        throw error;
        // return handleToolError(error); // Avoid returning ApiResponse directly
    } finally {
        // Use context.log for final logging
        log.info('Starting cleanup process...');
        // Release source documents that might have remained open due to error
        if (sourceDocs.length > 0) {
            log.warn(`Releasing ${sourceDocs.length} potentially orphaned source document objects.`);
            for (const doc of sourceDocs) {
                 try {
                     doc.Close(false); // Attempt to close just in case
                 } catch (e) { /* Ignore errors when closing here */ }
                releaseObject(doc);
            }
        }
        // Release target document if reference still exists (e.g., if there was an error before explicit release)
        if (targetDoc) {
            releaseObject(targetDoc);
             log.info('Target document object released.');
        }
        // Release Word application
         if (wordApp) {
              // Do not call Quit() directly here, rely on releaseObject and officeInterop
              releaseObject(wordApp);
              log.info('Word application object reference released.');
         }
         // Release officeAppInstance
         if (officeAppInstance) {
             officeAppInstance.release();
             log.info('Office application instance released.');
         }
    }
}

// 3. Define and Export the Resource
/**
 * McpResource definition for the 'word/merge' tool.
 * Merges multiple Word documents into a single new document.
 */
export const wordMergeTool: McpResource = { // Export as a single object
    path: 'word/merge',
    // Handler signature matches McpResource expectation (params, context)
    handler: mergeDocuments,
    schema: mergeSchema,
    description: 'Merges multiple Word documents (.docx) into a single new document using COM Interop.',
    // completions: ... // Add completions if useful
};