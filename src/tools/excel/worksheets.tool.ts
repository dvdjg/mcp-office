/**
 * @file Tool for managing worksheets in Excel files.
 * Allows adding, deleting, renaming, and setting active worksheets.
 * Uses COM Interop via winax to interact with Excel.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { saveResource } from '../dynamic/resources.tool'; // Import saveResource
import * as fs from 'fs-extra'; // Import fs to read the Excel file
import * as path from 'path'; // Import path

// Define the input schema for the excel/worksheets tool
const ExcelWorksheetsInputSchema = z.object({
  filePath: z.string().describe('Path to the Excel file.'),
  operation: z.enum(['add', 'delete', 'rename', 'set']).describe('Operation to perform on the worksheets.'),
  sheetName: z.string().optional().describe('Name of the sheet (for delete, rename, set).'),
  newSheetName: z.string().optional().describe('New name for the sheet (for rename).'),
  sheetIndex: z.number().int().positive().optional().describe('1-based index of the sheet (optional to identify the sheet).'),
  beforeSheet: z.string().optional().describe('Name of the sheet before which to insert the new sheet (for add).'),
  afterSheet: z.string().optional().describe('Name of the sheet after which to insert the new sheet (for add).'),
});

type ExcelWorksheetsInput = z.infer<typeof ExcelWorksheetsInputSchema>;

/**
 * @tool excel/worksheets
 * @description Manages worksheets in Excel files.
 * Allows adding, deleting, renaming, and setting active worksheets.
 * Uses COM Interop via winax to interact with Excel.
 * Requires the file path and the operation to perform.
 * Delete, rename, and set operations require identifying the sheet by name or index.
 * The add operation allows specifying the insertion position.
 * @input ExcelWorksheetsInputSchema
 * @output string - A message indicating the result of the operation.
 */
export const excelWorksheetsTool: McpResource[] = [{
  path: 'excel/worksheets', // Add the tool path
  description: 'Manages worksheets in Excel files.',
  schema: ExcelWorksheetsInputSchema, // Change inputSchema to schema
  handler: async (params: ToolRequestParams) => { // Type params
    const { filePath, operation, sheetName, newSheetName, sheetIndex, beforeSheet, afterSheet } = params as ExcelWorksheetsInput; // Cast params
    let excelApp: any;
    let workbook: any;
    let sheets: any;
    let sheet: any;
    let resultData: string | object = ''; // Use a more flexible type for the result

    try {
      excelApp = await getOfficeApplication('Excel.Application');
      excelApp.Visible = false; // Keep Excel hidden

      try {
        // Attempt to open the existing workbook
        workbook = excelApp.Workbooks.Open(filePath);
      } catch (error: any) {
        // If it doesn't exist, create a new one (only for the 'add' operation)
        if (operation === 'add') {
          workbook = excelApp.Workbooks.Add();
          // Save the new workbook immediately to be able to add sheets
          workbook.SaveAs(filePath);
        } else {
          throw new Error(`The Excel file was not found at the specified path: ${filePath}`);
        }
      }

      sheets = workbook.Sheets;

      switch (operation) {
        case 'add': {
          let beforeSheetObj = undefined;
          let afterSheetObj = undefined;

          if (beforeSheet) {
            try {
              beforeSheetObj = sheets.Item(beforeSheet);
            } catch {
              throw new Error(`The reference sheet 'beforeSheet' was not found: ${beforeSheet}`);
            }
          } else if (afterSheet) {
             try {
              afterSheetObj = sheets.Item(afterSheet);
            } catch {
              throw new Error(`The reference sheet 'afterSheet' was not found: ${afterSheet}`);
            }
          }

          // Add the new sheet
          const newSheet = sheets.Add({ Before: beforeSheetObj, After: afterSheetObj });
          if (sheetName) {
            newSheet.Name = sheetName;
            resultData = { message: `Sheet '${sheetName}' added successfully.` };
          } else {
             resultData = { message: `Sheet added successfully with default name '${newSheet.Name}'.` };
          }
          break;
        }

        case 'delete': {
          if (!sheetName && sheetIndex === undefined) {
            throw new Error('sheetName or sheetIndex is required for the delete operation.');
          }
          try {
             // Disable alerts to avoid the deletion confirmation dialog
            excelApp.DisplayAlerts = false;
            if (sheetName) {
              sheet = sheets.Item(sheetName);
              sheet.Delete();
              resultData = { message: `Sheet '${sheetName}' deleted successfully.` };
            } else if (sheetIndex !== undefined) {
              sheet = sheets.Item(sheetIndex);
              const deletedSheetName = sheet.Name;
              sheet.Delete();
              resultData = { message: `Sheet at index ${sheetIndex} ('${deletedSheetName}') deleted successfully.` };
            }
          } catch (error: any) {
             throw new Error(`Could not delete the sheet. Verify the name or index. Error: ${error.message}`);
          } finally {
             // Re-enable alerts
             excelApp.DisplayAlerts = true;
          }
          break;
        }

        case 'rename': {
          if (!sheetName && sheetIndex === undefined) {
            throw new Error('sheetName or sheetIndex is required for the rename operation.');
          }
          if (!newSheetName) {
            throw new Error('newSheetName is required for the rename operation.');
          }
          try {
            if (sheetName) {
              sheet = sheets.Item(sheetName);
            } else if (sheetIndex !== undefined) {
              sheet = sheets.Item(sheetIndex);
            }
            const oldSheetName = sheet.Name;
            sheet.Name = newSheetName;
            resultData = { message: `Sheet '${oldSheetName}' renamed to '${newSheetName}' successfully.` };
          } catch (error: any) {
             throw new Error(`Could not rename the sheet. Verify the name or index and the new name. Error: ${error.message}`);
          }
          break;
        }

        case 'set': {
           if (!sheetName && sheetIndex === undefined) {
            throw new Error('sheetName or sheetIndex is required for the set operation.');
          }
          try {
            if (sheetName) {
              sheet = sheets.Item(sheetName);
            } else if (sheetIndex !== undefined) {
              sheet = sheets.Item(sheetIndex);
            }
            sheet.Activate();
            resultData = { message: `Sheet '${sheet.Name}' selected successfully.` };
          } catch (error: any) {
             throw new Error(`Could not select the sheet. Verify the name or index. Error: ${error.message}`);
          }
          break;
        }

        default:
          throw new Error(`Unsupported operation: ${operation}`);
      }

      // Save changes and close the workbook
      workbook.Save();
      workbook.Close();

      // Save the modified Excel file as a dynamic resource
      // Only if the operation modified the file (add, delete, rename)
      if (operation === 'add' || operation === 'delete' || operation === 'rename') {
          try {
              const excelContent = await fs.readFile(filePath, null); // Read as Buffer
              await saveResource('excel/worksheets', path.basename(filePath), excelContent);
              // Do not use context.log here, as the handler does not receive context
              // logger.info(`Saved ${filePath} as a dynamic resource.`);
          } catch (resourceSaveError: any) {
              // Do not use context.log here
              // logger.error(`Failed to save ${filePath} as a dynamic resource: ${resourceSaveError.message}`);
              // Continue execution even if resource saving fails
          }
      }


      return { success: true, data: resultData };

    } catch (error: any) {
      console.error('Error in excel/worksheets tool:', error);
      return {
        success: false,
        error: {
          code: 'EXCEL_WORKSHEETS_ERROR',
          message: error.message || 'An error occurred while managing Excel worksheets.',
          details: { filePath, operation, sheetName, sheetIndex }
        }
      }; // Removed explicit cast
    } finally {
      // Release COM objects
      if (sheet) releaseObject(sheet);
      if (sheets) releaseObject(sheets);
      if (workbook) releaseObject(workbook);
      if (excelApp) releaseObject(excelApp);
    }
    // Add a return at the end to cover all possible cases
    // This will only be reached if no error was thrown or returned before.
    // In an ideal scenario, all switch cases should return.
    // But to satisfy the linter, we add this fallback return.
    return { success: false, error: { code: 'UNHANDLED_CASE', message: 'Excel worksheets operation did not return an explicit result.' } };
  },
}];