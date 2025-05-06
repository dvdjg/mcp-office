import { z } from 'zod';
import path from 'path';
import fs from 'fs-extra';
import {
    Document,
    Packer,
    Table,
    TableRow,
    TableCell,
    Paragraph,
    TextRun,
    WidthType,
    BorderStyle,
    AlignmentType,
} from 'docx';
import * as mammoth from 'mammoth';
import {
    getOfficeApplication,
    releaseObject,
} from '../../utils/officeInterop';
import { handleToolError } from '../../utils/errorHandler';
import { validateFilePath } from '../../utils/security';
import { ApiResponse, McpResource, FastMCPContext } from '../../types/common.types';
import logger from '../../utils/logger';

// --- Zod Schema for Input Validation ---
export const insertTableSchema = z.object({
    filePath: z.string().min(1, 'File path cannot be empty.'),
    rows: z.number().int().positive('Number of rows must be a positive integer.'),
    columns: z.number().int().positive('Number of columns must be a positive integer.'),
    position: z.string().optional().describe("Optional: 'end', 'selection', 'paragraph:N' (1-based index). Defaults to end."),
    style: z.string().optional().describe('Optional: Predefined Word table style name (e.g., "Table Grid"). Note: COM-path specific for exact style matching.'),
    useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
});

type InsertTableParams = z.infer<typeof insertTableSchema>;

// Add new schema for insertTableFromArray
export const insertTableFromArraySchema = z.object({
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
    useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
});

type InsertTableFromArrayParams = z.infer<typeof insertTableFromArraySchema>;


// --- Tool Handler Implementation ---
export async function insertTable(
    params: unknown,
    context?: FastMCPContext<undefined> // Use FastMCPContext<undefined>
): Promise<ApiResponse<{}>> {
    logger.info(`Executing word/tables/insert tool with params: ${JSON.stringify(params)}`);
    const validatedParams = insertTableSchema.parse(params);
    const { filePath, rows, columns, position, style, useComInterop } = validatedParams;
    const safeFilePath = validateFilePath(filePath); // validateFilePath throws on error

    if (useComInterop) {
        logger.info(`Using COM Interop path for insertTable: ${safeFilePath}`);
        let wordApp: any = null;
        let doc: any = null;
        let table: any = null;
        let insertionRange: any = null;
        let errorOccurred = false;

        try {
            wordApp = await getOfficeApplication('Word.Application');
            doc = wordApp.Documents.Open(safeFilePath);
            if (!doc) throw new Error(`COM: Failed to open document: ${safeFilePath}`);
            logger.debug(`COM: Document opened: ${safeFilePath}`);

            const comPosition = position?.toLowerCase() || 'end';
            if (comPosition === 'selection') {
                insertionRange = wordApp.Selection.Range;
                if (!insertionRange) throw new Error("COM: Cannot insert at selection: No selection found.");
            } else if (comPosition.startsWith('paragraph:')) {
                const parts = comPosition.split(':');
                const paragraphIndex = parseInt(parts[1], 10);
                if (isNaN(paragraphIndex) || paragraphIndex <= 0) throw new Error(`COM: Invalid paragraph index: ${parts[1]}.`);
                if (paragraphIndex > doc.Paragraphs.Count) throw new Error(`COM: Paragraph index ${paragraphIndex} out of bounds.`);
                insertionRange = doc.Paragraphs(paragraphIndex).Range;
                insertionRange.Collapse(1); // wdCollapseStart
            } else {
                insertionRange = doc.Range(doc.Content.End, doc.Content.End);
                if (insertionRange.Start > 0) {
                    insertionRange.Start = insertionRange.Start - 1;
                    insertionRange.End = insertionRange.Start;
                    insertionRange.InsertParagraphAfter();
                    insertionRange.Collapse(0); // wdCollapseEnd
                } else {
                    insertionRange = doc.Range(0, 0);
                }
            }
            if (!insertionRange) throw new Error("COM: Could not determine a valid insertion range.");

            table = doc.Tables.Add(insertionRange, rows, columns, 0, 0);
            if (!table) throw new Error('COM: Failed to insert table.');
            logger.debug(`COM: Table inserted. Index: ${table?.Index ?? 'N/A'}`);

            if (style) {
                try {
                    table.Style = style;
                    logger.debug(`COM: Style "${style}" applied.`);
                } catch (styleError: any) {
                    logger.warn(`COM: Failed to apply style "${style}": ${styleError.message}.`);
                }
            }
            doc.Save();
            logger.debug('COM: Document saved.');
            return { success: true, data: { message: `COM: Table (${rows}x${columns}) inserted.` } };
        } catch (error: any) {
            errorOccurred = true;
            logger.error(`Error in word/tables/insert (COM path): ${error.message}`, { stack: error.stack });
            if (doc && typeof doc.Close === 'function') {
                try { doc.Close(false); } catch (e) { logger.error(`COM: Error closing doc on error: ${(e as Error).message}`); }
            }
            return handleToolError(error, 'WORD_TABLE_INSERT_FAILED_COM');
        } finally {
            releaseObject(table);
            if (doc && !errorOccurred && typeof doc.Close === 'function') {
                try { doc.Close(false); } catch (e) { logger.warn(`COM: Error closing doc in finally: ${(e as Error).message}`); }
            }
            if (doc) releaseObject(doc);
            if (wordApp) releaseObject(wordApp);
            logger.debug('COM: Objects released in finally.');
        }
    } else {
        // Library path (docx) - Initially for NEW documents only
        logger.info(`Using Library (docx) path for insertTable: ${safeFilePath}`);
        try {
            if (await fs.pathExists(safeFilePath)) {
                logger.error(`Library path for insertTable currently only supports creating new files. File exists: ${safeFilePath}`);
                return handleToolError(new Error('Library path for insertTable currently only supports creating new files. File already exists.'), 'LIB_INSERT_EXISTING_FILE_NOT_SUPPORTED');
            }

            const tableRows = Array(rows).fill(0).map(() => {
                const cells = Array(columns).fill(0).map(() => {
                    return new TableCell({
                        children: [new Paragraph("")], // Add an empty paragraph to each cell
                        // borders: { ... }, // Optional: define cell borders
                        // width: { size: ..., type: WidthType.AUTO }, // Optional: define cell width
                    });
                });
                return new TableRow({ children: cells });
            });

            const docxTable = new Table({
                rows: tableRows,
                // width: { size: 100, type: WidthType.PERCENTAGE }, // Optional: define table width
            });

            // Style parameter for docx library is complex and not a direct string match to Word styles.
            // For now, we acknowledge it but don't apply a complex style.
            if (style) {
                logger.warn(`Library (docx) path: Style parameter '${style}' received. Applying specific docx styles requires custom mapping and is not fully implemented for this basic table. Table will be unstyled or use default docx styling.`);
            }
            // Position parameter is ignored for new files with docx library, table is added to main body.
            if (position && position !== 'end') {
                 logger.warn(`Library (docx) path: Position parameter '${position}' is ignored when creating a new document. Table will be added to the main body.`);
            }


            const doc = new Document({
                sections: [{
                    children: [docxTable],
                }],
            });

            const buffer = await Packer.toBuffer(doc);
            await fs.writeFile(safeFilePath, buffer);
            logger.info(`Library (docx) path: New document with table created successfully at ${safeFilePath}`);
            return { success: true, data: { message: `Library (docx): New document with table (${rows}x${columns}) created at ${safeFilePath}.` } };

        } catch (error: any) {
            logger.error(`Error in word/tables/insert (Library path): ${error.message}`, { stack: error.stack });
            return handleToolError(error, 'WORD_TABLE_INSERT_FAILED_LIB');
        }
    }
}

// Add new tool handler implementation
export async function insertTableFromArray(
    params: unknown,
    context?: FastMCPContext<undefined>
): Promise<ApiResponse<{}>> {
    logger.info(`Executing word/tables/insertFromArray tool with params: ${JSON.stringify(params)}`);
    const validatedParams = insertTableFromArraySchema.parse(params);
    const { filePath, data, position, styleName, styleOptions, useComInterop } = validatedParams;
    const safeFilePath = validateFilePath(filePath);

    if (useComInterop) {
        logger.info(`Using COM Interop path for insertTableFromArray: ${safeFilePath}`);
        let wordApp: any = null;
        let doc: any = null;
        let table: any = null;
        let insertionRange: any = null;
        let errorOccurred = false;

        try {
            wordApp = await getOfficeApplication('Word.Application');
            doc = wordApp.Documents.Open(safeFilePath);
            if (!doc) throw new Error(`COM: Failed to open document: ${safeFilePath}`);
            logger.debug(`COM: Document opened: ${safeFilePath}`);

            const comPosition = position?.toLowerCase() || 'end';
            // ... (COM position logic from original function - lines 244-287)
            if (comPosition === 'selection') {
                insertionRange = wordApp.Selection.Range;
                if (!insertionRange) throw new Error("COM: Cannot insert at selection: No selection found.");
            } else if (comPosition.startsWith('paragraph:')) {
                const parts = comPosition.split(':');
                const paragraphIndex = parseInt(parts[1], 10);
                if (isNaN(paragraphIndex) || paragraphIndex <= 0) throw new Error(`COM: Invalid paragraph index: ${parts[1]}.`);
                if (paragraphIndex > doc.Paragraphs.Count) throw new Error(`COM: Paragraph index ${paragraphIndex} out of bounds.`);
                insertionRange = doc.Paragraphs(paragraphIndex).Range;
                insertionRange.Collapse(1); // wdCollapseStart
            } else if (comPosition.startsWith('bookmark:')) {
                const bookmarkName = comPosition.substring('bookmark:'.length);
                try {
                    insertionRange = doc.Bookmarks(bookmarkName).Range;
                    insertionRange.Collapse(1);
                } catch (bookmarkError: any) {
                    throw new Error(`COM: Bookmark "${bookmarkName}" not found.`);
                }
            } else { // Default to 'end'
                insertionRange = doc.Range(doc.Content.End, doc.Content.End);
                if (insertionRange.Start > 0) {
                    insertionRange.Start = insertionRange.Start - 1;
                    insertionRange.End = insertionRange.Start;
                    insertionRange.InsertParagraphAfter();
                    insertionRange.Collapse(0); // wdCollapseEnd
                } else {
                    insertionRange = doc.Range(0, 0);
                }
            }
            if (!insertionRange) throw new Error("COM: Could not determine a valid insertion range.");


            const rows = data.length;
            const columns = data[0].length;
            table = doc.Tables.Add(insertionRange, rows, columns, 0, 0);
            if (!table) throw new Error('COM: Failed to insert table.');
            logger.debug(`COM: Table inserted. Index: ${table?.Index ?? 'N/A'}`);

            for (let i = 0; i < rows; i++) {
                for (let j = 0; j < columns; j++) {
                    table.Cell(i + 1, j + 1).Range.Text = data[i][j]?.toString() ?? '';
                }
            }
            logger.debug('COM: Table populated.');

            if (styleName) {
                try {
                    table.Style = styleName;
                    logger.debug(`COM: Style "${styleName}" applied.`);
                    if (styleOptions) {
                        table.AllowFormatting = true;
                        table.ApplyStyleHeadingRows = styleOptions.headerRow ?? true;
                        table.ApplyStyleFirstColumn = styleOptions.firstColumn ?? false;
                        table.ApplyStyleLastRow = false;
                        table.ApplyStyleLastColumn = false;
                        table.ApplyStyleBandedRows = styleOptions.bandedRows ?? false;
                        table.ApplyStyleBandedColumns = styleOptions.bandedColumns ?? false;
                        logger.debug('COM: Table style options applied.');
                    }
                } catch (styleError: any) {
                    logger.warn(`COM: Failed to apply style "${styleName}" or options: ${styleError.message}.`);
                }
            } else if (styleOptions) {
                logger.warn('COM: styleOptions provided but no styleName. Options not applied.');
            }

            doc.Save();
            logger.debug('COM: Document saved.');
            return { success: true, data: { message: `COM: Table (${rows}x${columns}) inserted and populated.` } };
        } catch (error: any) {
            errorOccurred = true;
            logger.error(`Error in word/tables/insertFromArray (COM path): ${error.message}`, { stack: error.stack });
            if (doc && typeof doc.Close === 'function') {
                try { doc.Close(false); } catch (e) { logger.error(`COM: Error closing doc on error: ${(e as Error).message}`); }
            }
            return handleToolError(error, 'WORD_TABLE_INSERT_FROM_ARRAY_FAILED_COM');
        } finally {
            releaseObject(table);
            if (doc && !errorOccurred && typeof doc.Close === 'function') {
                try { doc.Close(false); } catch (e) { logger.warn(`COM: Error closing doc in finally: ${(e as Error).message}`); }
            }
            if (doc) releaseObject(doc);
            if (wordApp) releaseObject(wordApp);
            logger.debug('COM: Objects released in finally.');
        }
    } else {
        // Library path (docx) - Initially for NEW documents only
        logger.info(`Using Library (docx) path for insertTableFromArray: ${safeFilePath}`);
        try {
            if (await fs.pathExists(safeFilePath)) {
                logger.error(`Library path for insertTableFromArray currently only supports creating new files. File exists: ${safeFilePath}`);
                return handleToolError(new Error('Library path for insertTableFromArray currently only supports creating new files. File already exists.'), 'LIB_INSERT_ARRAY_EXISTING_FILE_NOT_SUPPORTED');
            }

            const tableRows = data.map(rowData => {
                const cells = rowData.map(cellData => {
                    return new TableCell({
                        children: [new Paragraph({ children: [new TextRun(cellData?.toString() ?? '')] })],
                    });
                });
                return new TableRow({ children: cells });
            });

            const docxTableProperties: any = {}; // Add type for properties if more specific needed from docx
            if (styleOptions?.bandedRows) {
                // Note: `rowBands: true` is a common way to enable this in docx, but might need specific style.
                // For direct property, it's often part of a table style.
                // This is a simplification.
                // docxTableProperties.rowBands = true; // This might not be a direct property.
                logger.warn("Library (docx) path: 'bandedRows' direct property might not be available; usually part of a table style.");
            }
            if (styleOptions?.bandedColumns) {
                // docxTableProperties.columnBands = true; // Similar to rowBands
                logger.warn("Library (docx) path: 'bandedColumns' direct property might not be available; usually part of a table style.");
            }


            const docxTable = new Table({
                rows: tableRows,
                // properties: docxTableProperties, // Apply if properties are set
            });

            if (styleOptions?.headerRow && tableRows.length > 0) {
                // In `docx`, header property is on the TableRow.
                // tableRows[0].properties.isHeader = true; // This is not the correct API.
                // It's usually set like: new TableRow({ children: cells, isHeader: true })
                // Or by applying a style that defines a header row.
                // For simplicity, we'll log a warning.
                logger.warn("Library (docx) path: Setting 'headerRow' directly on TableRow properties is complex. It's typically part of a style or set during row creation if API supports.");
            }
            if (styleOptions?.firstColumn){
                logger.warn("Library (docx) path: 'firstColumn' formatting is typically part of a table style and not a direct property. It will be ignored for now.");
            }
            if (styleName) {
                logger.warn(`Library (docx) path: StyleName '${styleName}' is COM-specific. Applying named Word styles directly is complex with the 'docx' library. Table will use default styling.`);
            }
            if (position && position.toLowerCase() !== 'end') {
                logger.warn(`Library (docx) path: Position parameter '${position}' is ignored when creating a new document. Table will be added to the main body.`);
            }


            const doc = new Document({
                sections: [{
                    children: [docxTable],
                }],
            });

            const buffer = await Packer.toBuffer(doc);
            await fs.writeFile(safeFilePath, buffer);
            logger.info(`Library (docx) path: New document with table from array created successfully at ${safeFilePath}`);
            return { success: true, data: { message: `Library (docx): New document with table from array created at ${safeFilePath}.` } };

        } catch (error: any) {
            logger.error(`Error in word/tables/insertFromArray (Library path): ${error.message}`, { stack: error.stack });
            return handleToolError(error, 'WORD_TABLE_INSERT_FROM_ARRAY_FAILED_LIB');
        }
    }
}


// --- Tool Definition ---
// Add new schema for extractTableData
export const extractTableDataSchema = z.object({
    filePath: z.string().min(1, 'File path cannot be empty.'),
    tableIndex: z.number().int().positive('Table index must be a positive integer.'),
    useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
});

type ExtractTableDataParams = z.infer<typeof extractTableDataSchema>;


export async function extractTableData(
    params: unknown,
    context?: FastMCPContext<undefined>
): Promise<ApiResponse<any[][]>> {
    logger.info(`Executing word/tables/extractData tool with params: ${JSON.stringify(params)}`);
    const validatedParams = extractTableDataSchema.parse(params);
    const { filePath, tableIndex, useComInterop } = validatedParams;
    const safeFilePath = validateFilePath(filePath);

    if (useComInterop) {
        logger.info(`Using COM Interop path for extractTableData: ${safeFilePath}, tableIndex: ${tableIndex}`);
        let wordApp: any = null;
        let doc: any = null;
        let table: any = null;
        let errorOccurred = false;

        try {
            wordApp = await getOfficeApplication('Word.Application');
            doc = wordApp.Documents.Open(safeFilePath, false, true); // Open read-only
            if (!doc) throw new Error(`COM: Failed to open document: ${safeFilePath}`);
            logger.debug(`COM: Document opened: ${safeFilePath}`);

            if (tableIndex <= 0 || tableIndex > doc.Tables.Count) {
                throw new Error(`COM: Table index ${tableIndex} out of bounds. Document has ${doc.Tables.Count} tables.`);
            }
            table = doc.Tables(tableIndex);
            logger.debug(`COM: Accessed table with index: ${tableIndex}`);

            const numRows = table.Rows.Count;
            const numCols = table.Columns.Count;
            const tableData: any[][] = Array(numRows).fill(null).map(() => Array(numCols).fill(null));

            for (let r = 1; r <= numRows; r++) {
                for (let c = 1; c <= numCols; c++) {
                    let cell = null;
                    try {
                        cell = table.Cell(r, c);
                        if (cell.RowIndex === r && cell.ColumnIndex === c) { // Handle merged cells simply
                            tableData[r - 1][c - 1] = cell.Range.Text.replace(/\r?\n|\r/g, '').trim();
                        } else {
                            tableData[r - 1][c - 1] = null; // Part of a merged cell
                        }
                    } catch (cellError: any) {
                        logger.warn(`COM: Could not access cell (${r}, ${c}): ${cellError.message}`);
                        tableData[r - 1][c - 1] = null;
                    } finally {
                        if (cell) releaseObject(cell);
                    }
                }
            }
            logger.debug('COM: Table data extracted.');
            return { success: true, data: tableData };
        } catch (error: any) {
            errorOccurred = true;
            logger.error(`Error in word/tables/extractData (COM path): ${error.message}`, { stack: error.stack });
            if (doc && typeof doc.Close === 'function') {
                try { doc.Close(false); } catch (e) { logger.error(`COM: Error closing doc on error: ${(e as Error).message}`); }
            }
            return handleToolError(error, 'WORD_TABLE_EXTRACT_DATA_FAILED_COM');
        } finally {
            releaseObject(table);
            if (doc && !errorOccurred && typeof doc.Close === 'function') {
                try { doc.Close(false); } catch (e) { logger.warn(`COM: Error closing doc in finally: ${(e as Error).message}`); }
            }
            if (doc) releaseObject(doc);
            if (wordApp) releaseObject(wordApp);
            logger.debug('COM: Objects released in finally.');
        }
    } else {
        // Library path (mammoth)
        logger.info(`Using Library (Mammoth) path for extractTableData: ${safeFilePath}, tableIndex: ${tableIndex}`);
        try {
            if (!await fs.pathExists(safeFilePath)) {
                return handleToolError(new Error(`File not found: ${safeFilePath}`), 'FILE_NOT_FOUND_LIB');
            }

            // Using mammoth.convertToHtml and then parsing HTML for tables is complex
            // and error-prone without a proper HTML parser.
            // For this refactoring, we'll state it's not fully implemented.
            logger.warn("Library (Mammoth) path for extractTableData: Extracting structured table data via HTML conversion is complex and not fully implemented. This path may only work for very simple tables or return an error.");

            // Placeholder for actual HTML parsing logic:
            // const htmlResult = await mammoth.convertToHtml({ path: safeFilePath });
            // const tablesHtml = parseHtmlAndExtractTables(htmlResult.value); // This function would be complex
            // if (tableIndex > tablesHtml.length) throw new Error("Table index out of bounds for HTML tables.");
            // const tableData = convertHtmlTableToArray(tablesHtml[tableIndex - 1]); // Also complex

            return handleToolError(
                new Error('Library path for extractTableData is not fully implemented due to HTML parsing complexity. Use COM Interop (useComInterop: true) for reliable table data extraction.'),
                'NOT_IMPLEMENTED_LIB_TABLE_EXTRACTION'
            );

        } catch (error: any) {
            logger.error(`Error in word/tables/extractData (Library path): ${error.message}`, { stack: error.stack });
            return handleToolError(error, 'WORD_TABLE_EXTRACT_DATA_FAILED_LIB');
        }
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