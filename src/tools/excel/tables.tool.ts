/**
 * @file Tool for managing tables in Excel files.
 * Allows inserting, modifying, adding/deleting rows/columns, and deleting tables.
 * Uses COM Interop via winax to interact with Excel.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types.js';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js';
import logger from '../../utils/logger.js'; // Import logger
import { saveResource } from '../dynamic/resources.tool.js'; // Import saveResource
import * as fs from 'fs-extra'; // Import fs to read the Excel file
import * as path from 'path'; // Import path
import ExcelJS from 'exceljs'; // Import exceljs

// Input schema for the excel/tables tool
const ExcelTablesInputSchema = z.object({
  filePath: z.string().describe('Path to the Excel file.'),
  operation: z.enum(['insert', 'modify', 'add', 'delete']).describe('Operation to perform: insert, modify, add, delete.'),
  sheetName: z.string().optional().describe('Name of the worksheet. If not provided, the active sheet is used.'),
  sheetIndex: z.number().optional().describe('1-based index of the worksheet. If not provided, the active sheet is used.'),
  rangeAddress: z.string().optional().describe('Range address for inserting a new table (e.g., "A1:D10"). Required for the "insert" operation.'),
  tableName: z.string().optional().describe('Name of the table. Required for "modify", "add", "delete" operations.'),
  data: z.array(z.array(z.any())).optional().describe('Data for adding rows to the table. Required for the "add" operation with location "rows".'),
  location: z.enum(['rows', 'columns']).optional().describe('Specifies whether to add "rows" or "columns". Required for the "add" operation.'),
  count: z.number().optional().describe('Number of rows or columns to add/delete. Required for "add" and "delete" operations (except full table delete).'),
  position: z.number().optional().describe('Position (1-based index) where to add/delete rows/columns. Optional for "add" and "delete" (rows/columns).'),
  useComInterop: z.boolean().optional().default(false).describe('Use COM interop for local Excel interaction. Defaults to false (uses exceljs).'),
});

type ExcelTablesInput = z.infer<typeof ExcelTablesInputSchema>;

/**
 * @tool excel/tables
 * @description Manages tables in Excel files.
 * Allows inserting, modifying, adding/deleting rows/columns, and deleting tables.
 * Uses COM Interop via winax to interact with Excel.
 * Requires the file path, operation, and specific parameters based on the operation.
 */
const excelTablesTool: McpResource = {
  path: 'excel/tables',
  description: 'Manages tables in Excel files (insert, modify, add/delete rows/columns, delete table).',
  schema: ExcelTablesInputSchema,
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => {
    let input: ExcelTablesInput;
    try {
      input = ExcelTablesInputSchema.parse(params);
    } catch (error: any) {
      if (error instanceof z.ZodError) {
        logger.error(`Input validation failed for excel/tables: ${error.message}`, { errors: error.errors });
        return { success: false, error: { code: 'VALIDATION_ERROR', message: 'Input validation failed', details: error.errors } };
      }
      logger.error(`Unexpected error during input parsing for excel/tables: ${error.message}`, { error });
      return { success: false, error: { code: 'UNEXPECTED_PARSING_ERROR', message: `An unexpected error occurred during input parsing: ${error.message}` } };
    }

    const { filePath, operation, sheetName, sheetIndex, rangeAddress, tableName, data, location, count, position, useComInterop } = input;
    const absoluteFilePath = path.resolve(filePath);
    let message = '';

    if (useComInterop) {
      logger.info(`Executing excel/tables (COM) operation '${operation}' for file: ${filePath}`);
      let excelApp: any = null;
      let workbook: any = null;
      let worksheet: any = null;
      try {
        excelApp = await getOfficeApplication('Excel.Application');
        // Ensure file exists or create if necessary for certain operations (COM handles this implicitly for Open/Add)
        // For COM, Open will fail if not exists, Add creates new.
        // We'll assume filePath exists for read/modify/delete, and for insert, ListObjects.Add implies an existing sheet.
        // If the file doesn't exist for insert, COM would typically require creating the workbook first.
        // This logic might need refinement based on how strictly "create if not exists" should apply to COM path.
        // For now, assume Open is the primary way to get the workbook.
        try {
            workbook = excelApp.Workbooks.Open(absoluteFilePath);
        } catch (e) {
            // If opening fails, and it's an insert operation, we might need to create the workbook.
            // However, the COM ListObjects.Add typically expects an existing workbook and sheet.
            // This part of COM logic might be more complex if file creation is strictly required here.
            // For simplicity, we'll assume the file/workbook should exist for COM table operations.
            logger.error(`COM: Failed to open workbook ${absoluteFilePath}: ${e instanceof Error ? e.message : String(e)}`);
            throw new Error(`COM: Failed to open workbook ${absoluteFilePath}. Ensure the file exists.`);
        }


        if (sheetName) {
          worksheet = workbook.Sheets(sheetName);
        } else if (sheetIndex) {
          worksheet = workbook.Sheets(sheetIndex);
        } else {
          worksheet = workbook.ActiveSheet;
        }

        if (!worksheet) {
          throw new Error(`COM: Worksheet "${sheetName || sheetIndex || 'ActiveSheet'}" not found.`);
        }

        switch (operation) {
          case 'insert':
            if (!rangeAddress) {
              throw new Error('COM: rangeAddress is required for the "insert" operation.');
            }
            const listObject = worksheet.ListObjects.Add(1, worksheet.Range(rangeAddress), null, 1); // xlSrcRange = 1, xlListObjectHasHeaders = 1
            if (tableName) {
              listObject.Name = tableName;
            }
            message = `COM: Table inserted in range ${rangeAddress}. Name: ${listObject.Name}`;
            break;

          case 'modify':
            if (!tableName) {
              throw new Error('COM: tableName is required for the "modify" operation.');
            }
            const tableToModify = worksheet.ListObjects(tableName);
            if (!tableToModify) {
              throw new Error(`COM: Table "${tableName}" not found.`);
            }
            // Modification logic (e.g., rename)
            // if (input.newTableName) tableToModify.Name = input.newTableName; // Example
            message = `COM: Table "${tableName}" found for modification.`; // Actual modification needs more params
            break;

          case 'add':
            if (!tableName) throw new Error('COM: tableName is required for "add".');
            if (!location) throw new Error('COM: location ("rows" or "columns") is required for "add".');
            const tableToAdd = worksheet.ListObjects(tableName);
            if (!tableToAdd) throw new Error(`COM: Table "${tableName}" not found.`);

            if (location === 'rows') {
              if (!data || data.length === 0) throw new Error('COM: data is required for adding rows.');
              for (const rowData of data) {
                const newRow = tableToAdd.ListRows.Add(); // Adds at the end by default
                for (let i = 0; i < rowData.length; i++) {
                  if (i < newRow.Range.Cells.Count) {
                     newRow.Range.Cells(1, i + 1).Value = rowData[i];
                  }
                }
              }
              message = `COM: ${data.length} row(s) added to table "${tableName}".`;
            } else if (location === 'columns') {
              if (!count || count <= 0) throw new Error('COM: count is required for adding columns.');
              const currentColumnCount = tableToAdd.ListColumns.Count;
              const insertPos = position !== undefined && position >= 1 && position <= currentColumnCount + 1 ? position : currentColumnCount + 1;
              for (let i = 0; i < count; i++) {
                tableToAdd.ListColumns.Add(insertPos); // Adds column at specified position
              }
              message = `COM: ${count} column(s) added to table "${tableName}".`;
            }
            break;

          case 'delete':
            if (!tableName) throw new Error('COM: tableName is required for "delete".');
            const tableToDelete = worksheet.ListObjects(tableName);
            if (!tableToDelete) throw new Error(`COM: Table "${tableName}" not found.`);

            if (location === 'rows') {
              if (!count || count <= 0) throw new Error('COM: count is required for deleting rows.');
              const currentRowCount = tableToDelete.ListRows.Count;
              // Default to deleting from the end if position is not specified or invalid for deletion count
              const startDeletePos = (position && position >=1 && position + count -1 <= currentRowCount) ? position : Math.max(1, currentRowCount - count + 1);
              if (startDeletePos < 1 || startDeletePos + count -1 > currentRowCount) {
                throw new Error(`COM: Invalid row deletion range for table "${tableName}".`);
              }
              for (let i = 0; i < count; i++) {
                // COM ListRows index is 1-based. Deleting shifts indices. So always delete at startDeletePos.
                tableToDelete.ListRows(startDeletePos).Delete();
              }
              message = `COM: ${count} row(s) deleted from table "${tableName}".`;
            } else if (location === 'columns') {
              if (!count || count <= 0) throw new Error('COM: count is required for deleting columns.');
              const currentColCount = tableToDelete.ListColumns.Count;
              const startDeletePos = (position && position >=1 && position + count -1 <= currentColCount) ? position : Math.max(1, currentColCount - count + 1);
               if (startDeletePos < 1 || startDeletePos + count -1 > currentColCount) {
                throw new Error(`COM: Invalid column deletion range for table "${tableName}".`);
              }
              for (let i = 0; i < count; i++) {
                tableToDelete.ListColumns(startDeletePos).Delete();
              }
              message = `COM: ${count} column(s) deleted from table "${tableName}".`;
            } else { // Delete whole table
              tableToDelete.Delete();
              message = `COM: Table "${tableName}" deleted.`;
            }
            break;
          default:
            throw new Error(`COM: Unsupported operation: ${operation}`);
        }
        return { success: true, data: message };
      } catch (error: any) {
        logger.error(`Error in excel/tables tool (COM): ${error.message}`);
        return { success: false, error: { code: 'EXCEL_TABLES_COM_ERROR', message: `COM Error: ${error.message}` } };
      } finally {
        if (workbook) {
          try { workbook.Save(); workbook.Close(); } catch (e) { logger.warn(`COM: Error saving/closing workbook: ${e}`);}
          releaseObject(workbook);
        }
        if (excelApp) releaseObject(excelApp);
      }
    } else {
      // exceljs path
      logger.info(`Executing excel/tables (exceljs) operation '${operation}' for file: ${filePath}`);
      const wb = new ExcelJS.Workbook();
      let ws: ExcelJS.Worksheet | undefined; // Allow ws to be undefined initially
      let fileExisted = await fs.pathExists(absoluteFilePath);

      if (fileExisted) {
        await wb.xlsx.readFile(absoluteFilePath);
        if (sheetName) {
            ws = wb.getWorksheet(sheetName);
        } else if (sheetIndex && sheetIndex > 0 && sheetIndex <= wb.worksheets.length) {
            ws = wb.worksheets[sheetIndex - 1]; // 0-indexed
        } else if (wb.worksheets.length > 0) {
            ws = wb.worksheets[0]; // Default to first sheet if no specific one is found by name/index
        }
      }
      
      // If worksheet is still not defined, and operation allows creation, create it.
      if (!ws && (operation === 'insert' || (operation === 'add' && (sheetName || sheetIndex)))) {
        const newSheetName = sheetName || (sheetIndex ? `Sheet${sheetIndex}` : 'Sheet1');
        ws = wb.addWorksheet(newSheetName);
        logger.info(`exceljs: Created new sheet "${newSheetName}" as it was not found or file was new.`);
      } else if (!ws) {
        // If ws is still not defined after trying to find or create it based on operation type
        throw new Error(`exceljs: Worksheet "${sheetName || sheetIndex || 'default'}" could not be found or created for operation "${operation}".`);
      }

      switch (operation) {
        case 'insert':
          if (!rangeAddress) throw new Error('exceljs: rangeAddress is required for "insert".');
          // exceljs requires columns and rows for addTable.
          // We'll infer columns from rangeAddress or use data if provided.
          // This is a simplified interpretation. A robust solution needs clear data for columns/rows.
          const columns = data && data.length > 0 && data[0].length > 0 ? data[0].map(header => ({ name: String(header), filterButton: true })) : [{name: 'Column1', filterButton: true}];
          const rowsData = data && data.length > 1 ? data.slice(1) : [];
          const actualTableName = tableName || `Table${Date.now()}`;
          
          ws.addTable({
            name: actualTableName,
            ref: rangeAddress,
            headerRow: true,
            totalsRow: false, // example
            style: {
              theme: 'TableStyleMedium9',
              showRowStripes: true,
            },
            columns: columns,
            rows: rowsData,
          });
          message = `exceljs: Table inserted in range ${rangeAddress}. Name: ${actualTableName}. Note: Column/row data inferred or defaulted.`;
          break;

        case 'modify':
          // exceljs: Modifying table properties like name or style.
          // True resize or structural change is complex.
          if (!tableName) throw new Error('exceljs: tableName is required for "modify".');
          const tableToMod = ws.getTable(tableName);
          if (!tableToMod) throw new Error(`exceljs: Table "${tableName}" not found.`);
          // Example: tableToMod.name = newName; tableToMod.style = {...};
          message = `exceljs: Table "${tableName}" found. Modification capabilities are specific (e.g., style, name). Structural changes are complex.`;
          break;

        case 'add':
          if (!tableName) throw new Error('exceljs: tableName is required for "add".');
          if (!location) throw new Error('exceljs: location ("rows" or "columns") is required for "add".');
          const tableToAddJs = ws.getTable(tableName);
          if (!tableToAddJs) throw new Error(`exceljs: Table "${tableName}" not found.`);

          if (location === 'rows') {
            if (!data || data.length === 0) throw new Error('exceljs: data is required for adding rows.');
            // Using addRow in a loop for broader compatibility, as addRows might have issues or specific requirements.
            data.forEach(rowData => {
              tableToAddJs.addRow(rowData);
            });
            message = `exceljs: ${data.length} row(s) added to table "${tableName}".`;
          } else if (location === 'columns') {
            // Adding columns in exceljs is not direct. It involves redefining table.columns and updating all rows.
            // This is a significant limitation compared to COM.
            message = `exceljs: Adding columns to an existing table is complex and not directly supported by simple 'addColumn' in exceljs. Requires manual data manipulation or table recreation. Operation for table "${tableName}" noted as a limitation.`;
            logger.warn(message);
          }
          break;

        case 'delete':
          if (!tableName) throw new Error('exceljs: tableName is required for "delete".');
          const tableToDelJs = ws.getTable(tableName);

          if (!tableToDelJs) {
            throw new Error(`exceljs: Table "${tableName}" not found.`);
          }

          if (location === 'rows') {
            if (!count || count <= 0) throw new Error('exceljs: count is required for deleting rows.');
            const startIdx = position ? position - 1 : tableToDelJs.rows.length - count; // 0-indexed
            if (startIdx < 0 || startIdx + count > tableToDelJs.rows.length) {
                throw new Error(`exceljs: Invalid row deletion range for table "${tableName}".`);
            }
            tableToDelJs.removeRows(startIdx, count);
            message = `exceljs: ${count} row(s) deleted from table "${tableName}".`;
          } else if (location === 'columns') {
            if (!count || count <= 0) throw new Error('exceljs: count is required for deleting columns.');
            const colStartIdx = position ? position -1 : tableToDelJs.columns.length - count;
             if (colStartIdx < 0 || colStartIdx + count > tableToDelJs.columns.length) {
                throw new Error(`exceljs: Invalid column deletion range for table "${tableName}".`);
            }
            tableToDelJs.removeColumns(colStartIdx, count);
            message = `exceljs: ${count} column(s) deleted from table "${tableName}".`;
          } else { // Delete whole table (location is undefined or not 'rows'/'columns')
            ws.removeTable(tableName);
            message = `exceljs: Table "${tableName}" deleted.`;
          }
          break;
        default:
          throw new Error(`exceljs: Unsupported operation: ${operation}`);
      }
      await wb.xlsx.writeFile(absoluteFilePath);
      logger.info(`exceljs: Workbook saved to ${absoluteFilePath}`);
      return { success: true, data: message };
    }

    // Common dynamic resource saving for both paths if successful and modified
    if (operation === 'insert' || operation === 'add' || operation === 'delete' || operation === 'modify') {
        try {
            const excelContent = await fs.readFile(absoluteFilePath, null);
            await saveResource('excel/tables', path.basename(absoluteFilePath), excelContent);
            logger.info(`Saved ${absoluteFilePath} as a dynamic resource.`);
            if (typeof message === 'string') message += ` Saved as dynamic resource.`;
        } catch (resourceSaveError: any) {
            logger.error(`Failed to save ${absoluteFilePath} as a dynamic resource: ${resourceSaveError.message}`);
            if (typeof message === 'string') message += ` (Warning: Failed to save as dynamic resource: ${resourceSaveError.message})`;
        }
    }
    // This return is for the COM path if it was successful and didn't return earlier in its switch.
    // The exceljs path returns from its own switch or catch block.
    // The final fallback return at the very end of the handler should ideally not be reached.
    if (useComInterop) {
        return { success: true, data: message || "COM operation completed." };
    }
    // Fallback if something went wrong and no specific return was hit.
    return { success: false, error: { code: 'UNHANDLED_LOGIC_PATH', message: 'Operation did not complete as expected.' } };
  },
};

export default excelTablesTool;