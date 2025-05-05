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

// Add new schema for insertTableFromArray
const insertTableFromArraySchema = z.object({
    filePath: z.string().min(1, 'File path cannot be empty.'),
    data: z.array(z.array(z.any())).min(1, 'Data array cannot be empty.').refine(data => data.every(row => row.length === data[0].length), 'All rows in the data array must have the same number of columns.'),
    position: z.string().optional().describe("Optional: 'end', 'selection', 'paragraph:N' (1-based index), 'bookmark:BookmarkName'. Defaults to end."),
    styleName: z.string().optional().describe('Optional: Predefined Word table style name (e.g., "Table Grid").'),
    styleOptions: z.object({
        headerRow: z.boolean().optional().default(true).describe('Apply distinct formatting defined by the style to the first row.'),
        firstColumn: z.boolean().optional().default(false).describe('Apply distinct formatting defined by the style to the first column.'),
        bandedRows: z.boolean().optional().default(false).describe('Apply alternating row shading (banding).'),
        bandedColumns: z.boolean().optional().default(false).describe('Apply alternating column shading (banding).'),
    }).optional().describe('Optional: Fine-tune the applied style\'s appearance.'),
});

type InsertTableFromArrayParams = z.infer<typeof insertTableFromArraySchema>;


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

// Add new tool handler implementation
export async function insertTableFromArray(
    params: unknown,
    context?: FastMCPContext<undefined>
): Promise<ApiResponse<{}>> {
    logger.info(`Executing word/tables/insertFromArray tool with params: ${JSON.stringify(params)}`);
    let wordApp: any = null;
    let doc: any = null;
    let table: any = null;
    let insertionRange: any = null;
    let errorOccurred = false;

    try {
        // 1. Validate Input Parameters
        const validatedParams = insertTableFromArraySchema.parse(params);
        logger.debug('Parameters validated successfully.');

        const safeFilePath = validateFilePath(validatedParams.filePath);
        logger.debug(`File path validated: ${safeFilePath}`);

        // 2. Get/Create Word Application Instance
        wordApp = await getOfficeApplication('Word.Application');
        logger.debug('Word application instance obtained.');

        // 3. Open the Document
        doc = wordApp.Documents.Open(safeFilePath);
        if (!doc) {
            throw new Error(`Failed to open document: ${safeFilePath}`);
        }
        logger.debug(`Document opened: ${safeFilePath}`);

        // 4. Determine Insertion Range
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
            insertionRange = doc.Paragraphs(paragraphIndex).Range;
            insertionRange.Collapse(1); // Collapse to the start of the paragraph range (wdCollapseStart = 1)
            logger.debug(`Insertion range set before paragraph ${paragraphIndex}.`);
        } else if (position.startsWith('bookmark:')) {
             const bookmarkName = position.substring('bookmark:'.length);
             try {
                 insertionRange = doc.Bookmarks(bookmarkName).Range;
                 insertionRange.Collapse(1); // Collapse to the start of the bookmark range
                 logger.debug(`Insertion range set at bookmark: ${bookmarkName}`);
             } catch (bookmarkError: any) {
                 throw new Error(`Bookmark "${bookmarkName}" not found.`);
             }
        }
        else { // Default to 'end'
            insertionRange = doc.Range(doc.Content.End, doc.Content.End);
             if (insertionRange.Start > 0) {
                 insertionRange.Start = insertionRange.Start -1;
                 insertionRange.End = insertionRange.Start;
                 insertionRange.InsertParagraphAfter();
                 insertionRange.Collapse(0);
             } else {
                 insertionRange = doc.Range(0, 0);
             }
            logger.debug('Insertion range set to the end of the document.');
        }

        if (!insertionRange) {
             throw new Error("Could not determine a valid insertion range.");
        }

        // 5. Insert the Table and Populate Data
        const rows = validatedParams.data.length;
        const columns = validatedParams.data[0].length;
        logger.debug(`Attempting to add table with ${rows} rows and ${columns} columns and populate data.`);

        // DefaultTableBehavior = wdWord9TableBehavior (0), AutoFitBehavior = wdAutoFitFixed (0)
        table = doc.Tables.Add(insertionRange, rows, columns, 0, 0);
        if (!table) {
            throw new Error('Failed to insert table.');
        }
        logger.debug(`Table inserted successfully. Table index: ${table?.Index ?? 'N/A'}`);

        // Populate table cells
        for (let i = 0; i < rows; i++) {
            for (let j = 0; j < columns; j++) {
                // Word table cells are 1-based index
                table.Cell(i + 1, j + 1).Range.Text = validatedParams.data[i][j]?.toString() ?? '';
            }
        }
        logger.debug('Table populated with data.');

        // 6. Apply Style (Optional)
        if (validatedParams.styleName) {
            logger.debug(`Attempting to apply style: ${validatedParams.styleName}`);
            try {
                table.Style = validatedParams.styleName;
                logger.debug(`Style "${validatedParams.styleName}" applied successfully.`);

                // Apply style options if style was applied
                if (validatedParams.styleOptions) {
                    logger.debug(`Applying style options: ${JSON.stringify(validatedParams.styleOptions)}`);
                    // These properties control which parts of the table the style formatting is applied to
                    table.AllowFormatting = true; // Ensure formatting is allowed
                    table.ApplyStyleHeadingRows = validatedParams.styleOptions.headerRow ?? true; // Default to true as per schema default
                    table.ApplyStyleFirstColumn = validatedParams.styleOptions.firstColumn ?? false; // Default to false
                    table.ApplyStyleLastRow = false; // Assuming no last row formatting needed by default
                    table.ApplyStyleLastColumn = false; // Assuming no last column formatting needed by default
                    table.ApplyStyleBandedRows = validatedParams.styleOptions.bandedRows ?? false; // Default to false
                    table.ApplyStyleBandedColumns = validatedParams.styleOptions.bandedColumns ?? false; // Default to false
                    logger.debug('Table style options applied.');
                }

            } catch (styleError: any) {
                logger.warn(`Failed to apply style "${validatedParams.styleName}" or style options: ${styleError.message}. Table inserted without style/options.`);
                // Continue without style/options
            }
        } else if (validatedParams.styleOptions) {
             logger.warn('styleOptions were provided but no styleName was specified. styleOptions will not be applied.');
        }


        // 7. Save the Document
        logger.debug('Saving document...');
        doc.Save();
        logger.debug('Document saved successfully.');

        return { success: true, data: { message: `Table (${rows}x${columns}) inserted and populated successfully.` } };

    } catch (error: any) {
        errorOccurred = true;
        logger.error(`Error in word/tables/insertFromArray: ${error.message}`, { stack: error.stack });
        if (doc) {
            try {
                 if (typeof doc.Close === 'function') {
                     doc.Close(false); // wdDoNotSaveChanges = 0
                     logger.debug('Document closed without saving changes due to error.');
                 } else {
                     logger.debug('Document object seems invalid or already closed, skipping close attempt.');
                 }
            } catch (closeError: any) {
                if (!error.message?.toLowerCase().includes('object invalid')) {
                   logger.error(`Error closing document after initial error: ${closeError.message}`);
                }
            }
        }
        if (wordApp) {
            releaseObject(wordApp);
            logger.debug('Word application released in catch block.');
            wordApp = null;
        }
        return handleToolError(error, 'WORD_TABLE_INSERT_FROM_ARRAY_FAILED');
    } finally {
        releaseObject(table);
        if (doc && typeof doc.Close === 'function' && !errorOccurred) {
             try {
                 doc.Close(false);
                 logger.debug('Document closed in finally block (success path).');
             } catch (finalCloseError: any) {
                 logger.warn(`Error during final document close: ${finalCloseError.message}`);
             } finally {
                 releaseObject(doc);
             }
        } else if (doc) {
             releaseObject(doc);
        }
        if (!errorOccurred && wordApp) {
             releaseObject(wordApp);
        }
        logger.debug('COM objects released.');
    }
}


// --- Tool Definition ---
// Add new schema for extractTableData
const extractTableDataSchema = z.object({
    filePath: z.string().min(1, 'File path cannot be empty.'),
    tableIndex: z.number().int().positive('Table index must be a positive integer.'),
});

type ExtractTableDataParams = z.infer<typeof extractTableDataSchema>;


// Add new tool handler implementation for extractTableData
export async function extractTableData(
    params: unknown,
    context?: FastMCPContext<undefined>
): Promise<ApiResponse<any[][]>> {
    logger.info(`Executing word/tables/extractData tool with params: ${JSON.stringify(params)}`);
    let wordApp: any = null;
    let doc: any = null;
    let table: any = null;
    let errorOccurred = false;

    try {
        // 1. Validate Input Parameters
        const validatedParams = extractTableDataSchema.parse(params);
        logger.debug('Parameters validated successfully.');

        const safeFilePath = validateFilePath(validatedParams.filePath);
        logger.debug(`File path validated: ${safeFilePath}`);

        // 2. Get/Create Word Application Instance
        wordApp = await getOfficeApplication('Word.Application');
        logger.debug('Word application instance obtained.');

        // 3. Open the Document
        doc = wordApp.Documents.Open(safeFilePath);
        if (!doc) {
            throw new Error(`Failed to open document: ${safeFilePath}`);
        }
        logger.debug(`Document opened: ${safeFilePath}`);

        // 4. Access the specified table
        const tableIndex = validatedParams.tableIndex;
        if (tableIndex <= 0 || tableIndex > doc.Tables.Count) {
            throw new Error(`Table index ${tableIndex} is out of bounds. Document has ${doc.Tables.Count} tables.`);
        }
        table = doc.Tables(tableIndex);
        logger.debug(`Accessed table with index: ${tableIndex}`);

        // 5. Extract table data
        const numRows = table.Rows.Count;
        const numCols = table.Columns.Count;
        const tableData: any[][] = Array(numRows).fill(null).map(() => Array(numCols).fill(null));

        for (let r = 1; r <= numRows; r++) {
            for (let c = 1; c <= numCols; c++) {
                let cell = null;
                try {
                    cell = table.Cell(r, c);
                    // Check if this cell is the top-left cell of a merged area
                    // This is a simplified check; a more robust approach might track merged cells
                    // as they are encountered. For now, rely on checking if the cell's
                    // RowIndex and ColumnIndex match the loop indices.
                    if (cell.RowIndex === r && cell.ColumnIndex === c) {
                         tableData[r - 1][c - 1] = cell.Range.Text.replace(/\r?\n|\r/g, '').trim(); // Extract text, remove newlines, trim whitespace
                    } else {
                         // This cell is part of a merge started by a previous cell
                         tableData[r - 1][c - 1] = null;
                    }
                } catch (cellError: any) {
                    logger.warn(`Could not access cell (${r}, ${c}): ${cellError.message}`);
                    tableData[r - 1][c - 1] = null; // Mark inaccessible cells as null
                } finally {
                    if (cell) releaseObject(cell);
                }
            }
        }
        logger.debug('Table data extracted.');

        return { success: true, data: tableData };

    } catch (error: any) {
        errorOccurred = true;
        logger.error(`Error in word/tables/extractData: ${error.message}`, { stack: error.stack });
        if (doc) {
            try {
                 if (typeof doc.Close === 'function') {
                     doc.Close(false); // wdDoNotSaveChanges = 0
                     logger.debug('Document closed without saving changes due to error.');
                 } else {
                     logger.debug('Document object seems invalid or already closed, skipping close attempt.');
                 }
            } catch (closeError: any) {
                if (!error.message?.toLowerCase().includes('object invalid')) {
                   logger.error(`Error closing document after initial error: ${closeError.message}`);
                }
            }
        }
        if (wordApp) {
            releaseObject(wordApp);
            logger.debug('Word application released in catch block.');
            wordApp = null;
        }
        return handleToolError(error, 'WORD_TABLE_EXTRACT_DATA_FAILED');
    } finally {
        releaseObject(table);
        if (doc && typeof doc.Close === 'function' && !errorOccurred) {
             try {
                 doc.Close(false);
                 logger.debug('Document closed in finally block (success path).');
             } catch (finalCloseError: any) {
                 logger.warn(`Error during final document close: ${finalCloseError.message}`);
             } finally {
                 releaseObject(doc);
             }
        } else if (doc) {
             releaseObject(doc);
        }
        if (!errorOccurred && wordApp) {
             releaseObject(wordApp);
        }
        logger.debug('COM objects released.');
    }
}
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
    {
        path: 'word/tables/insertFromArray',
        handler: insertTableFromArray,
        schema: insertTableFromArraySchema,
        description: 'Inserts a new table into a Word document from a 2D array, with optional styling.',
        // inputSchema: insertTableFromArraySchema.openapi('InsertTableFromArrayInput'),
        // outputSchema: z.object({ message: z.string() }).openapi('InsertTableFromArrayOutput'),
    },
    {
        path: 'word/tables/extractData',
        handler: extractTableData,
        schema: extractTableDataSchema,
        description: 'Extracts data from a specified Word table into a 2D array, handling merged cells.',
        // inputSchema: extractTableDataSchema.openapi('ExtractTableDataInput'),
        // outputSchema: z.array(z.array(z.any())).openapi('ExtractTableDataOutput'),
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