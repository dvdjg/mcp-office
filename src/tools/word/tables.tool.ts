import { z } from 'zod';
import path from 'path';
import {
    getOfficeApplication,
    releaseObject,
    // safeString, // Removed - Not found in common utils
} from '../../utils/officeInterop';
import { handleToolError } from '../../utils/errorHandler'; // Corrected path
import { validateFilePath } from '../../utils/security'; // Corrected path
import { ApiResponse, McpResource, FastMCPContext } from '../../types/common.types'; // Import FastMCPContext, remove ToolContext
import logger from '../../utils/logger'; // Corrected import style

// --- Zod Schema for Input Validation ---
const insertTableSchema = z.object({
    filePath: z.string().min(1, 'File path cannot be empty.'),
    rows: z.number().int().positive('Number of rows must be a positive integer.'),
    columns: z.number().int().positive('Number of columns must be a positive integer.'),
    position: z.string().optional().describe("Optional: 'end', 'selection', 'paragraph:N' (1-based index). Defaults to end."),
    style: z.string().optional().describe('Optional: Predefined Word table style name (e.g., "Table Grid").'),
});

type InsertTableParams = z.infer<typeof insertTableSchema>;

// --- Tool Handler Implementation ---
export async function insertTable(
    params: unknown,
    context?: FastMCPContext<undefined> // Use FastMCPContext<undefined>
): Promise<ApiResponse<{}>> {
    // Removed safeString from log message, using JSON.stringify for basic logging
    logger.info(`Executing word/tables/insert tool with params: ${JSON.stringify(params)}`);
    let wordApp: any = null;
    let doc: any = null;
    let table: any = null;
    let insertionRange: any = null;
    let errorOccurred = false; // Flag to track if an error happened

    try {
        // 1. Validate Input Parameters
        const validatedParams = insertTableSchema.parse(params);
        logger.debug('Parameters validated successfully.');

        // 2. Validate File Path
        // Note: validateFilePath in security.ts doesn't seem to take userId based on provided snippet.
        // Pass context?.userId or relevant identifier if your implementation differs and context exists.
        const safeFilePath = validateFilePath(validatedParams.filePath /*, context?.userId */);
        // validateFilePath now throws on error, so no need for `if (!safeFilePath)` check.
        logger.debug(`File path validated: ${safeFilePath}`);

        // 3. Get/Create Word Application Instance
        wordApp = await getOfficeApplication('Word.Application');
        logger.debug('Word application instance obtained.');

        // 4. Open the Document
        doc = wordApp.Documents.Open(safeFilePath);
        if (!doc) {
            throw new Error(`Failed to open document: ${safeFilePath}`);
        }
        logger.debug(`Document opened: ${safeFilePath}`);

        // 5. Determine Insertion Range
        const position = validatedParams.position?.toLowerCase() || 'end';
        logger.debug(`Determining insertion range for position: ${position}`);

        if (position === 'selection') {
            insertionRange = wordApp.Selection.Range;
            if (!insertionRange) {
                throw new Error("Cannot insert at selection: No selection found.");
            }
            logger.debug('Insertion range set to current selection.');
        } else if (position.startsWith('paragraph:')) {
            const parts = position.split(':');
            const paragraphIndex = parseInt(parts[1], 10);
            if (isNaN(paragraphIndex) || paragraphIndex <= 0) {
                throw new Error(`Invalid paragraph index: ${parts[1]}. Must be a positive integer.`);
            }
            if (paragraphIndex > doc.Paragraphs.Count) {
                 throw new Error(`Paragraph index ${paragraphIndex} out of bounds. Document has ${doc.Paragraphs.Count} paragraphs.`);
            }
            // Insert *before* the specified paragraph
            insertionRange = doc.Paragraphs(paragraphIndex).Range;
            insertionRange.Collapse(1); // Collapse to the start of the paragraph range (wdCollapseStart = 1)
            logger.debug(`Insertion range set before paragraph ${paragraphIndex}.`);
        } else { // Default to 'end'
            // Create a range at the very end of the document content
            insertionRange = doc.Range(doc.Content.End, doc.Content.End);
             // Ensure we are not inside the final paragraph mark if it exists
            if (insertionRange.Start > 0) {
                 // Move range just before the final paragraph mark
                 insertionRange.Start = insertionRange.Start -1;
                 insertionRange.End = insertionRange.Start;
                 // Add a paragraph break *after* this position (which becomes before the table)
                 insertionRange.InsertParagraphAfter();
                 // Move the range to the start of the new empty paragraph for table insertion
                 insertionRange.Collapse(0); // wdCollapseEnd = 0 (moves to end of selection, which is start of new para)
            } else {
                 // Document is empty, just use the start
                 insertionRange = doc.Range(0, 0);
            }
            logger.debug('Insertion range set to the end of the document.');
        }

        if (!insertionRange) {
             throw new Error("Could not determine a valid insertion range.");
        }

        // 6. Insert the Table
        logger.debug(`Attempting to add table with ${validatedParams.rows} rows and ${validatedParams.columns} columns.`);
        // DefaultTableBehavior = wdWord9TableBehavior (0), AutoFitBehavior = wdAutoFitFixed (0)
        table = doc.Tables.Add(insertionRange, validatedParams.rows, validatedParams.columns, 0, 0);
        if (!table) {
            throw new Error('Failed to insert table.');
        }
        // Accessing table.Index might fail if table creation truly failed, guard it.
        logger.debug(`Table inserted successfully. Table index: ${table?.Index ?? 'N/A'}`);

        // 7. Apply Style (Optional)
        if (validatedParams.style) {
            logger.debug(`Attempting to apply style: ${validatedParams.style}`);
            try {
                table.Style = validatedParams.style;
                logger.debug(`Style "${validatedParams.style}" applied successfully.`);
            } catch (styleError: any) {
                logger.warn(`Failed to apply style "${validatedParams.style}": ${styleError.message}. Table inserted without style.`);
                // Continue without style, maybe add warning to response later if needed
            }
        }

        // 8. Save the Document
        logger.debug('Saving document...');
        doc.Save();
        logger.debug('Document saved successfully.');

        return { success: true, data: { message: `Table (${validatedParams.rows}x${validatedParams.columns}) inserted successfully.` } };

    } catch (error: any) {
        errorOccurred = true; // Set flag on error
        logger.error(`Error in word/tables/insert: ${error.message}`, { stack: error.stack });
        // Ensure doc is closed without saving changes in case of error after opening
        if (doc) {
            try {
                // Attempt to close without saving only if it's a valid document object
                 // Check for a property that indicates it's a valid COM object (e.g., FullName)
                 // This check might need refinement based on the COM library used.
                 if (typeof doc.Close === 'function') {
                     doc.Close(false); // wdDoNotSaveChanges = 0
                     logger.debug('Document closed without saving changes due to error.');
                 } else {
                     logger.debug('Document object seems invalid or already closed, skipping close attempt.');
                 }
            } catch (closeError: any) {
                // Avoid logging errors if the primary error was about the doc object itself
                if (!error.message?.toLowerCase().includes('object invalid')) {
                   logger.error(`Error closing document after initial error: ${closeError.message}`);
                }
            }
        }
        // Release wordApp explicitly before returning the error
        if (wordApp) {
            releaseObject(wordApp);
            logger.debug('Word application released in catch block.');
            wordApp = null; // Prevent double release in finally
        }
        // Provide a more specific default error code, remove wordApp argument
        return handleToolError(error, 'WORD_TABLE_INSERT_FAILED');
    } finally {
        // 9. Release COM Objects
        releaseObject(table);
        // Only release doc here if it wasn't closed in the catch block and seems valid
        // Refined check for validity before attempting close/release
        if (doc && typeof doc.Close === 'function' && !errorOccurred) {
             try {
                 doc.Close(false); // Close without saving again on success path
                 logger.debug('Document closed in finally block (success path).');
             } catch (finalCloseError: any) {
                 logger.warn(`Error during final document close: ${finalCloseError.message}`);
             } finally {
                 releaseObject(doc);
             }
        } else if (doc) {
             // Release even if closing failed or error occurred (doc object might still hold resources)
             releaseObject(doc);
        }
        // wordApp should have been released in catch block on error.
        // Release here only on success path.
        if (!errorOccurred && wordApp) {
             releaseObject(wordApp);
        }
        logger.debug('COM objects released.');
    }
}


// --- Tool Definition ---
export const wordTablesTool: McpResource[] = [
    {
        path: 'word/tables/insert',
        handler: insertTable,
        schema: insertTableSchema,
        description: 'Inserts a new table into a Word document at a specified position using COM Interop.',
        // Add OpenAPI schema generation for better documentation if openapi-zod-converter is available
        // inputSchema: insertTableSchema.openapi('InsertTableInput'),
        // outputSchema: z.object({ message: z.string() }).openapi('InsertTableOutput'),
    },
    // --- Placeholders for other table operations ---
    {
        path: 'word/tables/modify',
        handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/tables/modify not implemented.' } }),
        description: 'Modify table properties (Not Implemented)',
        schema: z.object({}),
        // inputSchema: z.object({}).openapi('ModifyTableInput'),
        // outputSchema: z.object({}).openapi('ModifyTableOutput'),
    },
    {
        path: 'word/tables/add',
        handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/tables/add not implemented.' } }),
        description: 'Add rows/columns to a table (Not Implemented)',
        schema: z.object({}),
        // inputSchema: z.object({}).openapi('AddTableElementInput'),
        // outputSchema: z.object({}).openapi('AddTableElementOutput'),
    },
    {
        path: 'word/tables/delete',
        handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/tables/delete not implemented.' } }),
        description: 'Delete rows/columns/table (Not Implemented)',
        schema: z.object({}),
        // inputSchema: z.object({}).openapi('DeleteTableElementInput'),
        // outputSchema: z.object({}).openapi('DeleteTableElementOutput'),
    },
];

// --- Export ---
// Exporting the array directly as requested
// export default wordTablesTool; // CommonJS style if needed elsewhere