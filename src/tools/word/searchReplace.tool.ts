/**
 * @file Implements the 'word/search-replace' tool using COM Interop for searching and replacing text in Word documents.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js'; // Normalized relative path
import { validateFilePath } from '../../utils/security.js'; // Normalized relative path
import { handleToolError, createErrorResponse } from '../../utils/errorHandler.js'; // Normalized relative path
import { saveResource } from '../dynamic/resources.tool.js'; // Normalized relative path
import type {
    ApiResponse, // Type alias for SuccessResponse | ErrorResponse
    SuccessResponse, // Specific type for success
    ErrorResponse, // Specific type for error
    McpResource,
    ToolRequestParams,
} from '../../types/common.types.js'; // Normalized relative path
import { Context as FastMCPContext } from 'fastmcp'; // Import FastMCP Context
import path from 'path';
import logger from '../../utils/logger.js'; // Normalized relative path

// Define Input Schemas
/**
 * Zod schema for the input parameters of the 'word/search-replace' tool.
 */
const searchReplaceSchema = z.object({
    /** The path to the Word document. */
    filePath: z.string().min(1, 'File path cannot be empty.').refine(validateFilePath, {
        message: "Invalid or potentially unsafe file path provided.",
    }),
    /** The text or pattern to search for. */
    find: z.string().min(1, 'Search text cannot be empty.'),
    /** The text to replace matches with. Defaults to an empty string (deletion). */
    replace: z.string().default(''),
    /** Whether the search should be case-sensitive. Defaults to false. */
    matchCase: z.boolean().optional().default(false),
    /** Whether to match only whole words. Defaults to false. */
    matchWholeWord: z.boolean().optional().default(false),
    /** Whether to use wildcards in the search text. Defaults to false. */
    useWildcards: z.boolean().optional().default(false),
    /** Whether to replace all occurrences (true) or only the first one (false). Defaults to true. */
    replaceAll: z.boolean().optional().default(true),
});

/**
 * Infers the type for the validated search and replace parameters.
 */
type SearchReplaceParams = z.infer<typeof searchReplaceSchema>;

// Implement the Handler searchAndReplace accepting an optional FastMCPContext
/**
 * Handles the 'word/search-replace' tool request.
 * Searches for text in a Word document and replaces it using COM Interop.
 * @param params - The parameters for the tool, validated against `searchReplaceSchema`.
 * @param context - The FastMCP context (optional), providing logging and progress reporting.
 * @returns A promise resolving to an ApiResponse indicating if replacements were made.
 * @throws {Error} If validation fails, a COM error occurs, or file access fails.
 */
async function searchAndReplace(
    params: ToolRequestParams,
    context?: FastMCPContext<undefined>
): Promise<ApiResponse<{ replacementsMade: boolean }>> {
    // Use context logger if available, otherwise fallback to global logger
    const log = context?.log ?? logger;
    const reportProgress = context?.reportProgress; // Get reportProgress function if context exists
    const totalSteps = 4; // Define total steps for progress

    let wordApp: any = null;
    let doc: any = null;
    let findObject: any = null;
    let replacementObject: any = null;
    let replacementsMade: boolean = false;

    const filePathParam = (params as any)?.filePath || 'unknown'; // For logging in case of early error

    try {
        reportProgress?.({ progress: 0, total: totalSteps }); // Step 0: Start
        // Validate the params
        const validatedParams = searchReplaceSchema.parse(params);

        // Validate the document path (already done by schema refinement, but explicit call is fine)
        const safeFilePath = validatedParams.filePath; // Already validated by Zod
        log.info(`Validated file path: ${safeFilePath}`);

        // Get the Word instance
        log.info('Getting Word application instance...');
        const officeAppInstance = await getOfficeApplication('Word.Application');
        wordApp = officeAppInstance.app;
        wordApp.Visible = false; // Keep Word hidden
        wordApp.DisplayAlerts = 0; // wdAlertsNone = 0

        // Open the document
        log.info(`Opening document: ${safeFilePath}`);
        doc = await wordApp.Documents.Open(safeFilePath);
        log.debug(`Document opened successfully.`);
        reportProgress?.({ progress: 1, total: totalSteps }); // Step 1: Document Opened

        // Access the Find and Replacement objects
        findObject = await doc.Content.Find;
        replacementObject = await findObject.Replacement;

        // Clear previous formatting
        await findObject.ClearFormatting();
        await replacementObject.ClearFormatting();

        // Configure search and replace properties
        log.info(`Configuring search for "${validatedParams.find}" and replace with "${validatedParams.replace}"`);
        findObject.Text = validatedParams.find;
        replacementObject.Text = validatedParams.replace;
        findObject.MatchCase = validatedParams.matchCase;
        findObject.MatchWholeWord = validatedParams.matchWholeWord;
        findObject.MatchWildcards = validatedParams.useWildcards;
        findObject.Forward = true;
        findObject.Wrap = 1; // wdFindContinue
        log.debug(`Search and replace properties configured.`);
        reportProgress?.({ progress: 2, total: totalSteps }); // Step 2: Configured

        // Execute the operation
        const replaceOption = validatedParams.replaceAll ? 2 : 1; // wdReplaceAll = 2, wdReplaceOne = 1
        log.info(`Executing find/replace (replaceAll: ${validatedParams.replaceAll})...`);

        replacementsMade = await findObject.Execute(
            undefined, // FindText (already set on findObject.Text)
            validatedParams.matchCase,
            validatedParams.matchWholeWord,
            validatedParams.useWildcards,
            undefined, // MatchSoundsLike
            undefined, // MatchAllWordForms
            true,      // Forward
            1,         // Wrap
            undefined, // Format
            undefined, // ReplaceWith (already set on replacementObject.Text)
            replaceOption // Replace
        );
        log.info(`Find/replace executed. Replacements made: ${replacementsMade}`);
        reportProgress?.({ progress: 3, total: totalSteps }); // Step 3: Executed

        // Save if successful
        if (replacementsMade) {
            log.info(`Saving document: ${safeFilePath}`);
            await doc.Save();
            log.info(`Document saved.`);

            // Read the content of the modified document before closing it
            let modifiedContent = '';
            try {
                 // Read the text from the COM document object
                 modifiedContent = doc.Content.Text;
                 log.info(`Read content from modified document.`);
             } catch (readContentError: any) {
                 log.error(`Failed to read content from modified document before closing: ${readContentError.message}`, { error: readContentError });
                 // Do not throw error here, attempt to save the resource empty or with error
             }

            // Save the modified document content as a dynamic resource after closing it
            try {
                await saveResource('word/search-replace', path.basename(safeFilePath), modifiedContent);
                log.info(`Saved ${safeFilePath} as a dynamic resource.`);
            } catch (resourceSaveError: any) {
                log.error(`Failed to save ${safeFilePath} as a dynamic resource: ${resourceSaveError.message}`, { error: resourceSaveError });
                // Continue execution even if resource saving fails
            }

        } else {
            log.info(`No replacements made, document not saved.`);
        }

        // Build the success response manually
        const successResponse: SuccessResponse<{ replacementsMade: boolean }> = {
            success: true,
            message: `Search and replace operation completed on ${path.basename(
                safeFilePath
            )}. Replacements made: ${replacementsMade}`, // Adjusted message
            data: { replacementsMade },
        };
        reportProgress?.({ progress: 4, total: totalSteps }); // Step 4: Complete
        return successResponse;

    } catch (error: any) {
        // Handle errors and return ErrorResponse
        log.error(`Error during search/replace on ${filePathParam}: ${error.message}`, { error });
        // Attempt to close document and release objects in case of error
        if (doc) {
          try {
            doc.Close(0); // wdDoNotSaveChanges = 0
            log.debug(`Document closed in error handler: ${filePathParam}`);
          } catch (closeError: any) {
            log.warn(`Error closing document in error handler: ${closeError.message}`, { error: closeError });
          }
          releaseObject(doc);
        }
        // Quitting the app in the error handler might be too aggressive
        // if (wordApp) {
        //   try {
        //     await wordApp.Quit();
        //   } catch (quitError: any) {
        //     logger.error(`Error quitting Word application in error handler: ${quitError.message}`);
        //   }
        // }
        return handleToolError(error, 'WORD_SEARCH_REPLACE_ERROR'); // Use handleToolError
    } finally {
        // Release COM objects in reverse order of creation/acquisition
        log.debug('Starting COM object cleanup for search/replace...');
        if (replacementObject) {
            releaseObject(replacementObject);
        }
        if (findObject) {
            releaseObject(findObject);
        }
        if (doc) { // Redundant if closed in catch, but safe
            releaseObject(doc);
        }
        if (wordApp) {
            // Decide whether to quit the application. Quitting might close a user's open instance.
            // A safer approach might be to only quit if we know we started the instance.
            // For now, let's not quit the application automatically.
            // wordApp.Quit();
            releaseObject(wordApp); // Release the reference
        }
        log.debug('COM object cleanup finished for search/replace.');
    }
}

// Define and Export the Resource
/**
 * McpResource definition for the 'word/search-replace' tool.
 * Searches for text in a Word document and replaces it.
 */
export const wordSearchReplaceTool: McpResource = { // Export as a single object
    path: 'word/search-replace',
    handler: searchAndReplace, // Correct handler signature
    schema: searchReplaceSchema,
    description:
        'Searches for text in a Word document and replaces it using COM Interop. Supports options like match case, whole word, wildcards, and replace all.',
    // examples property removed
};