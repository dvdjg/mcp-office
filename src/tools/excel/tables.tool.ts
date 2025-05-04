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
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import logger from '../../utils/logger'; // Import logger
import { saveResource } from '../dynamic/resources.tool'; // Import saveResource
import * as fs from 'fs-extra'; // Import fs to read the Excel file
import * as path from 'path'; // Import path

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
    let excelApp: any = null;
    let workbook: any = null;
    let worksheet: any = null;
    let filePath: string | undefined; // Declare filePath outside the try and allow undefined

    try {
      const input = ExcelTablesInputSchema.parse(params);
      filePath = input.filePath; // Assign filePath here
      const { operation, sheetName, sheetIndex, rangeAddress, tableName, data, location, count, position } = input;

      excelApp = await getOfficeApplication('Excel.Application'); // Use getOfficeApplication
      workbook = excelApp.Workbooks.Open(filePath);

      if (sheetName) {
        worksheet = workbook.Sheets(sheetName);
      } else if (sheetIndex) {
        worksheet = workbook.Sheets(sheetIndex);
      } else {
        worksheet = workbook.ActiveSheet;
      }

      if (!worksheet) {
        throw new Error(`Worksheet "${sheetName || sheetIndex}" not found.`);
      }

      switch (operation) {
        case 'insert':
          if (!rangeAddress) {
            throw new Error('rangeAddress is required for the "insert" operation.');
          }
          // Insert a new table in the specified range
          // The second argument (xlYes) indicates that the first row is the header
          const listObject = worksheet.ListObjects.Add(1, worksheet.Range(rangeAddress), null, 1); // xlSrcRange = 1, xlListObjectHasHeaders = 1
          if (tableName) {
            listObject.Name = tableName;
          }
          return { success: true, data: `Table inserted in range ${rangeAddress}. Name: ${listObject.Name}` }; // Adjusted return

        case 'modify':
          if (!tableName) {
            throw new Error('tableName is required for the "modify" operation.');
          }
          // Modify an existing table (e.g., change name, resize - not implemented in this simple example)
          const tableToModify = worksheet.ListObjects(tableName);
          if (!tableToModify) {
            throw new Error(`Table "${tableName}" not found.`);
          }
          // Implement modification logic here if necessary.
          // For now, we just confirm that the table exists.
          return { success: true, data: `Table "${tableName}" found for modification.` }; // Adjusted return

        case 'add':
          if (!tableName) {
            throw new Error('tableName is required for the "add" operation.');
          }
          if (!location) {
            throw new Error('location ("rows" or "columns") is required for the "add" operation.');
          }
          const tableToAdd = worksheet.ListObjects(tableName);
          if (!tableToAdd) {
            throw new Error(`Table "${tableName}" not found.`);
          }

          if (location === 'rows') {
            if (!data || data.length === 0) {
              throw new Error('data is required and cannot be empty for adding rows.');
            }
            // Add rows to the table
            for (const rowData of data) {
              // AddDataBoundRow adds an empty row at the end. Then we fill the data.
              const newRow = tableToAdd.ListRows.Add();
              for (let i = 0; i < rowData.length; i++) {
                if (i < newRow.Range.Cells.Count) {
                   newRow.Range.Cells(1, i + 1).Value = rowData[i];
                }
              }
            }
             return { success: true, data: `${data.length} row(s) added to table "${tableName}".` }; // Adjusted return

          } else if (location === 'columns') {
             if (!count || count <= 0) {
                throw new Error('count (number of columns to add) is required and must be positive for adding columns.');
             }
             // Add columns to the table
             // Add method for ListColumns adds a column to the left of the specified position
             const currentColumnCount = tableToAdd.ListColumns.Count;
             const insertPosition = position !== undefined && position >= 1 && position <= currentColumnCount + 1 ? position : currentColumnCount + 1;

             for (let i = 0; i < count; i++) {
                tableToAdd.ListColumns.Add(insertPosition);
             }
             return { success: true, data: `${count} column(s) added to table "${tableName}" at position ${insertPosition}.` }; // Adjusted return
           }
           // break; // Should not reach here - Removed redundant break

        case 'delete':
          if (!tableName) {
            throw new Error('tableName is required for the "delete" operation.');
          }
          const tableToDelete = worksheet.ListObjects(tableName);
          if (!tableToDelete) {
            throw new Error(`Table "${tableName}" not found.`);
          }

          if (location === 'rows') {
             if (!count || count <= 0) {
                throw new Error('count (number of rows to delete) is required and must be positive for deleting rows.');
             }
             // Delete rows from the table
             const currentRowCount = tableToDelete.ListRows.Count;
             const deletePosition = position !== undefined && position >= 1 && position <= currentRowCount ? position : currentRowCount - count + 1;

             if (deletePosition < 1 || deletePosition + count - 1 > currentRowCount) {
                 throw new Error(`Invalid row deletion range. Attempting to delete ${count} row(s) from position ${deletePosition} in a table with ${currentRowCount} row(s).`);
             }

             for (let i = 0; i < count; i++) {
                tableToDelete.ListRows(deletePosition).Delete();
             }
             return { success: true, data: `${count} row(s) deleted from table "${tableName}" starting from position ${deletePosition}.` }; // Adjusted return

          } else if (location === 'columns') {
             if (!count || count <= 0) {
                throw new Error('count (number of columns to delete) is required and must be positive for deleting columns.');
             }
             // Delete columns from the table
             const currentColumnCount = tableToDelete.ListColumns.Count;
             const deletePosition = position !== undefined && position >= 1 && position <= currentColumnCount ? position : currentColumnCount - count + 1;

             if (deletePosition < 1 || deletePosition + count - 1 > currentColumnCount) {
                 throw new Error(`Invalid column deletion range. Attempting to delete ${count} column(s) from position ${deletePosition} in a table with ${currentColumnCount} column(s).`);
             }

             for (let i = 0; i < count; i++) {
                tableToDelete.ListColumns(deletePosition).Delete();
             }
             return { success: true, data: `${count} column(s) deleted from table "${tableName}" starting from position ${deletePosition}.` }; // Adjusted return

          } else {
            // Delete the entire table
            tableToDelete.Delete();
            return { success: true, data: `Table "${tableName}" deleted.` }; // Adjusted return
          }
          // break; // Should not reach here - Removed redundant break
      }

    } catch (error: any) {
      logger.error(`Error in excel/tables tool: ${error.message}`); // Use logger
      return { success: false, error: { code: 'EXCEL_TABLES_ERROR', message: `Error managing tables in Excel: ${error.message}` } }; // Adjusted error return
    } finally {
      if (workbook) {
        try {
            workbook.Save();
            // Save the modified Excel file as a dynamic resource
            // This is done in the finally block because Save() happens here for all modification operations.
            // We don't need to check the specific operation here.
            // Ensure filePath has a value before attempting to read the file
            if (filePath) {
                try {
                    const excelContent = await fs.readFile(filePath, null); // Read as Buffer
                    await saveResource('excel/tables', path.basename(filePath), excelContent);
                    // logger.info(`Saved ${filePath} as a dynamic resource.`);
                } catch (resourceSaveError: any) {
                    // logger.error(`Failed to save ${filePath} as a dynamic resource: ${resourceSaveError.message}`);
                    // Continue execution even if resource saving fails
                }
            }
            workbook.Close();
        } catch (closeError: any) {
            logger.warn(`Error closing the workbook: ${closeError.message}`); // Use logger
        }
        releaseObject(workbook);
      }
      // The Excel application is managed externally, we don't close it here.
      releaseObject(excelApp);
    }
    // Add a return at the end to cover all possible cases
    // This will only be reached if no error was thrown or returned before.
    // In an ideal scenario, all switch cases should return.
    // But to satisfy the linter, we add this fallback return.
    return { success: false, error: { code: 'UNHANDLED_CASE', message: 'Excel table operation did not return an explicit result.' } };
  },
};

export default excelTablesTool;