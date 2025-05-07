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
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { saveResource } from '../dynamic/resources.tool'; // Import saveResource
import * as fs from 'fs-extra'; // Import fs to read the Excel file
import * as path from 'path'; // Import path
import ExcelJS from 'exceljs'; // Import exceljs

// Define the input schema for the excel/worksheets tool
const ExcelWorksheetsInputSchema = z.object({
  filePath: z.string().describe('Path to the Excel file.'),
  operation: z.enum(['add', 'delete', 'rename', 'set']).describe('Operation to perform on the worksheets.'),
  sheetName: z.string().optional().describe('Name of the sheet (for delete, rename, set).'),
  newSheetName: z.string().optional().describe('New name for the sheet (for rename).'),
  sheetIndex: z.number().int().positive().optional().describe('1-based index of the sheet (optional to identify the sheet).'),
  beforeSheet: z.string().optional().describe('Name of the sheet before which to insert the new sheet (for add).'),
  afterSheet: z.string().optional().describe('Name of the sheet after which to insert the new sheet (for add).'),
  useComInterop: z.boolean().optional().default(false).describe('Use COM interop for local Excel interaction. Defaults to false (uses exceljs).'),
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
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => {
    let input: ExcelWorksheetsInput;
    try {
      input = ExcelWorksheetsInputSchema.parse(params);
    } catch (error: any) {
      if (error instanceof z.ZodError) {
        return { success: false, error: { code: 'VALIDATION_ERROR', message: 'Input validation failed', details: error.errors } };
      }
      return { success: false, error: { code: 'UNEXPECTED_PARSING_ERROR', message: `An unexpected error occurred during input parsing: ${error.message}` } };
    }

    const { filePath, operation, sheetName, newSheetName, sheetIndex, beforeSheet, afterSheet, useComInterop } = input;
    let resultData: string | object = '';
    const absoluteFilePath = path.resolve(filePath); // Define absoluteFilePath

    try { // Outer try for the whole handler logic
      if (useComInterop) {
        let excelApp: any;
      let workbook: any;
      let sheets: any;
      let sheet: any;
      try {
        excelApp = await getOfficeApplication('Excel.Application');
        excelApp.Visible = false; // Keep Excel hidden

        try {
          // Attempt to open the existing workbook
          workbook = excelApp.Workbooks.Open(absoluteFilePath); // Use absoluteFilePath
        } catch (error: any) {
          // If it doesn't exist, create a new one (only for the 'add' operation)
          if (operation === 'add') {
            workbook = excelApp.Workbooks.Add();
            // Save the new workbook immediately to be able to add sheets
            workbook.SaveAs(absoluteFilePath); // Use absoluteFilePath
          } else {
            throw new Error(`File not found: ${absoluteFilePath}`);
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
              resultData = { message: `Sheet '${sheetName}' added successfully using COM Interop.` };
            } else {
               resultData = { message: `Sheet added successfully with default name '${newSheet.Name}' using COM Interop.` };
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
                resultData = { message: `Sheet '${sheetName}' deleted successfully using COM Interop.` };
              } else if (sheetIndex !== undefined) {
                sheet = sheets.Item(sheetIndex);
                const deletedSheetName = sheet.Name;
                sheet.Delete();
                resultData = { message: `Sheet at index ${sheetIndex} ('${deletedSheetName}') deleted successfully using COM Interop.` };
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
              resultData = { message: `Sheet '${oldSheetName}' renamed to '${newSheetName}' successfully using COM Interop.` };
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
              resultData = { message: `Sheet '${sheet.Name}' selected successfully using COM Interop.` };
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

      } catch (comError: any) { // Specific catch for COM block
        // console.error('Error in excel/worksheets tool (COM Interop):', comError); // console.error removed
        throw comError; // Re-throw to be caught by outer try-catch
      } finally {
        // Release COM objects
        if (sheet) releaseObject(sheet);
        if (sheets) releaseObject(sheets);
        if (workbook) releaseObject(workbook);
        if (excelApp) releaseObject(excelApp);
      }
    } else {
      // Use exceljs
      const excelWorkbook = new ExcelJS.Workbook();
      // try { // Inner try for exceljs specific logic, will be caught by outer handler try
        const fileExists = await fs.pathExists(absoluteFilePath); // Use absoluteFilePath
        if (fileExists) {
          await excelWorkbook.xlsx.readFile(absoluteFilePath); // Use absoluteFilePath
        } else if (operation !== 'add') {
          throw new Error(`File not found: ${absoluteFilePath}`);
        }
        // If operation is 'add' and file doesn't exist, a new workbook is already instantiated.

        switch (operation) {
          case 'add': {
            let newSheetNameActual = sheetName || `Sheet${excelWorkbook.worksheets.length + 1}`;
            // Ensure unique sheet name if default is used and conflicts
            let counter = 1;
            while (excelWorkbook.getWorksheet(newSheetNameActual)) {
                newSheetNameActual = sheetName ? `${sheetName} (${counter})` : `Sheet${excelWorkbook.worksheets.length + 1 + counter}`;
                counter++;
            }

            const newSheet = excelWorkbook.addWorksheet(newSheetNameActual);

            // exceljs doesn't directly support adding before/after an existing named sheet in the same way COM does.
            // Worksheets are added to the end by default. Reordering would require moving sheets.
            // For simplicity, we'll note this limitation if beforeSheet or afterSheet is used.
            let orderMessage = '';
            if (beforeSheet || afterSheet) {
                orderMessage = ' Note: `beforeSheet` and `afterSheet` are not fully supported with exceljs path; sheet is added to the end.';
            }
            resultData = { message: `Sheet '${newSheet.name}' added successfully using exceljs.${orderMessage}` };
            break;
          }
          case 'delete': {
            if (!sheetName && sheetIndex === undefined) {
              throw new Error('sheetName or sheetIndex is required for the delete operation.');
            }
            let sheetToDelete: ExcelJS.Worksheet | undefined;
            let identifier: string | number = '';

            if (sheetName) {
              sheetToDelete = excelWorkbook.getWorksheet(sheetName);
              identifier = sheetName;
            } else if (sheetIndex !== undefined) {
              // exceljs worksheet indices are 0-based in the array, but COM is 1-based.
              // The schema defines sheetIndex as 1-based.
              if (sheetIndex > 0 && sheetIndex <= excelWorkbook.worksheets.length) {
                sheetToDelete = excelWorkbook.worksheets[sheetIndex - 1]; // Adjust to 0-based
                identifier = sheetIndex;
              }
            }

            if (sheetToDelete) {
              const deletedSheetNameActual = sheetToDelete.name;
              excelWorkbook.removeWorksheet(sheetToDelete.id); // Use id to remove
              resultData = { message: `Sheet '${deletedSheetNameActual}' (identified by ${sheetName ? `name '${sheetName}'` : `index ${sheetIndex}`}) deleted successfully using exceljs.` };
            } else {
              const id = sheetName ? `'${sheetName}'` : `index ${sheetIndex}`;
              throw new Error(`Sheet ${id} not found.`);
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
            let sheetToRename: ExcelJS.Worksheet | undefined;
            if (sheetName) {
              sheetToRename = excelWorkbook.getWorksheet(sheetName);
            } else if (sheetIndex !== undefined) {
               if (sheetIndex > 0 && sheetIndex <= excelWorkbook.worksheets.length) {
                sheetToRename = excelWorkbook.worksheets[sheetIndex - 1]; // Adjust to 0-based
              }
            }

            if (sheetToRename) {
              const oldSheetNameActual = sheetToRename.name;
              sheetToRename.name = newSheetName;
              resultData = { message: `Sheet '${oldSheetNameActual}' renamed to '${newSheetName}' successfully using exceljs.` };
            } else {
              const id = sheetName ? `'${sheetName}'` : `index ${sheetIndex}`;
              throw new Error(`Sheet ${id} not found.`);
            }
            break;
          }
          case 'set': {
            if (!sheetName && sheetIndex === undefined) {
              throw new Error('sheetName or sheetIndex is required for the set operation.');
            }
            let sheetToActivate: ExcelJS.Worksheet | undefined;

            if (sheetName) {
              sheetToActivate = excelWorkbook.getWorksheet(sheetName);
            } else if (sheetIndex !== undefined) {
              if (sheetIndex > 0 && sheetIndex <= excelWorkbook.worksheets.length) {
                sheetToActivate = excelWorkbook.worksheets[sheetIndex - 1]; // Adjust to 0-based
              }
            }

            if (sheetToActivate) {
              const sheetIndexInWorkbook = excelWorkbook.worksheets.indexOf(sheetToActivate);
              if (sheetIndexInWorkbook !== -1) {
                excelWorkbook.views = [
                  {
                    x: 0, y: 0, width: 10000, height: 20000,
                    firstSheet: 0, activeTab: sheetIndexInWorkbook, visibility: 'visible'
                  }
                ];
              }
              resultData = { message: `Sheet '${sheetToActivate.name}' (identified by ${sheetName ? `name '${sheetName}'` : `index ${sheetIndex}`}) set as active (view) successfully using exceljs.` };
            } else {
              const id = sheetName ? `'${sheetName}'` : `index ${sheetIndex}`;
              throw new Error(`Sheet ${id} not found.`);
            }
            break;
          }
          default:
            throw new Error(`Unsupported operation: ${operation}`);
        }
        await excelWorkbook.xlsx.writeFile(absoluteFilePath); // Use absoluteFilePath
      // } catch (exceljsError: any) { // Inner try for exceljs specific logic
      //   throw exceljsError; // Re-throw to be caught by outer try-catch
      // } // End of inner try for exceljs specific logic
    } // End of if-else for COM/exceljs

    // Save the modified Excel file as a dynamic resource
    // This part is common for both COM and exceljs paths if successful
    if (operation === 'add' || operation === 'delete' || operation === 'rename' || operation === 'set') { // Added 'set'
        try {
            const excelContent = await fs.readFile(absoluteFilePath, null); // Read as Buffer
            await saveResource('excel/worksheets', path.basename(absoluteFilePath), excelContent);
        } catch (resourceSaveError: any) {
            // console.warn removed
            if (typeof resultData === 'object' && resultData !== null) {
                (resultData as any).warning = `File operation successful, but failed to save as dynamic resource: ${resourceSaveError.message}`;
            } else if (typeof resultData === 'string') {
                resultData += ` (Warning: Failed to save as dynamic resource: ${resourceSaveError.message})`;
            }
        }
    }
    return { success: true, data: resultData };

  } catch (error: any) { // Outer catch for the whole handler
    // console.error removed
    return {
      success: false,
      error: {
        code: useComInterop ? 'EXCEL_WORKSHEETS_COM_ERROR' : 'EXCEL_WORKSHEETS_EXCELJS_ERROR',
        message: error.message || `An error occurred while managing Excel worksheets.`,
        details: { filePath: absoluteFilePath, operation, sheetName, sheetIndex }
      }
    };
  }
},
}];