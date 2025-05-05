/**
 * @file Tool for manipulating cell ranges in Excel worksheets.
 * Allows reading, writing, formatting, and applying other operations to specific ranges.
 * Uses COM Interop via winax to interact with Excel.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse, FastMCPContext } from '../../types/common.types'; // Added ApiResponse, FastMCPContext
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Added releaseObject
import { saveResource } from '../dynamic/resources.tool';
import * as fs from 'fs-extra';
import * as path from 'path';
import logger from '../../utils/logger'; // Added logger import
import { validateFilePath } from '../../utils/security'; // Added security import

// Helper to create standard error responses
const createErrorResponse = (message: string, code = 'TOOL_EXECUTION_ERROR', details?: unknown): ApiResponse<never> => ({
    success: false,
    error: { code, message, details },
});

// Define the input schema for the excel/range tool
const ExcelRangeInputSchema = z.object({
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, { // Added validation
    message: "Invalid or potentially unsafe file path provided.",
  }),
  sheetName: z.string().optional().describe('Name of the worksheet. If not provided, the active sheet is used.'),
  sheetIndex: z.number().int().positive().optional().describe('1-based index of the worksheet. If provided, it overrides sheetName.'),
  rangeAddress: z.string().describe('Range address (e.g., "A1", "B2:C5").'),
  operation: z.enum(['read', 'write', 'format', 'apply']).describe('Operation to perform on the range.'),
  values: z.array(z.array(z.any())).optional().describe('Values to write to the range (for the "write" operation).'),
  formatProperties: z.record(z.any()).optional().describe('Format properties to apply (for the "format" operation).'),
  // You can add more properties for the 'apply' operation if needed
});

type ExcelRangeInput = z.infer<typeof ExcelRangeInputSchema>;

/**
 * @tool excel/range
 * @description Manipulates cell ranges in Excel worksheets.
 * Allows reading, writing, formatting, and applying other operations to specific ranges.
 * For 'write', 'format', and 'apply' operations, the file will be created if it does not exist.
 * Uses COM Interop via winax to interact with Excel.
 * Requires the file path, sheet name or index, range address, and operation.
 * For writing, values are provided as an array of arrays.
 * For formatting, format properties are provided as an object.
 */
const excelRangeTool: McpResource = {
  path: 'excel/range',
  description: "Manipulates cell ranges in Excel worksheets. Allows reading, writing, formatting, and applying other operations. Creates the file if it doesn't exist for write/format/apply operations.", // Updated description
  schema: ExcelRangeInputSchema,
  handler: async (params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<any>> => { // Added context and ApiResponse
    let excelApp: any = null;
    let workbook: any = null; // Define workbook here
    let fileCreated = false; // Flag to track file creation

    try {
      // Validate input parameters
      const input = ExcelRangeInputSchema.parse(params);
      logger.info(`Executing excel/range operation '${input.operation}' for file: ${input.filePath}`);

      excelApp = await getOfficeApplication('Excel.Application');
      const absoluteFilePath = path.resolve(input.filePath);

      // --- Create if not exists logic (for write/format/apply) ---
      if (input.operation === 'write' || input.operation === 'format' || input.operation === 'apply') {
          try {
              if (await fs.pathExists(absoluteFilePath)) {
                  logger.info(`Opening existing workbook: ${absoluteFilePath}`);
                  workbook = excelApp.Workbooks.Open(absoluteFilePath);
              } else {
                  logger.info(`File not found. Creating new workbook at: ${absoluteFilePath}`);
                  workbook = excelApp.Workbooks.Add(); // Create new workbook
                  // Save the new workbook immediately to the target path
                  // Determine format based on extension (default to xlsx)
                  // You might need a more robust way to determine the format constant
                  const fileExt = path.extname(absoluteFilePath).toLowerCase();
                  let saveFormat = 51; // xlOpenXMLWorkbook (.xlsx)
                  if (fileExt === '.xlsb') saveFormat = 50; // xlExcel12 (.xlsb)
                  else if (fileExt === '.xls') saveFormat = 56; // xlExcel8 (.xls)

                  workbook.SaveAs(absoluteFilePath, saveFormat);
                  fileCreated = true;
                  logger.info(`Successfully created and saved new workbook: ${absoluteFilePath}`);
              }
          } catch (fileError: any) {
               logger.error(`Error opening or creating workbook '${absoluteFilePath}': ${fileError.message}`, { error: fileError });
               return createErrorResponse(`Failed to open or create workbook: ${fileError.message}`, 'FILE_OPERATION_FAILED', fileError);
          }
      } else if (input.operation === 'read') {
          // For read operation, file must exist
          if (!await fs.pathExists(absoluteFilePath)) {
              logger.warn(`File not found for read operation: ${absoluteFilePath}`);
              return createErrorResponse(`File not found: ${input.filePath}`, 'FILE_NOT_FOUND');
          }
          logger.info(`Opening existing workbook for read: ${absoluteFilePath}`);
          workbook = excelApp.Workbooks.Open(absoluteFilePath); // Consider opening read-only if possible
      }
      // --- End create if not exists logic ---


      if (!workbook) {
           // This check might be redundant if the try/catch above handles errors, but good as a safeguard
           return createErrorResponse(`Failed to obtain workbook object for: ${input.filePath}`, 'FILE_OPEN_FAILED');
      }

      let worksheet;
      try {
          if (input.sheetIndex !== undefined) {
            worksheet = workbook.Sheets.Item(input.sheetIndex);
          } else if (input.sheetName) {
            worksheet = workbook.Sheets.Item(input.sheetName);
          } else {
            // If file was just created, it will have a default sheet (e.g., "Sheet1")
            // If opened, use ActiveSheet
            worksheet = workbook.ActiveSheet;
          }
      } catch (sheetError: any) {
           logger.error(`Error getting worksheet '${input.sheetName || input.sheetIndex || 'ActiveSheet'}': ${sheetError.message}`);
           return createErrorResponse(`Worksheet "${input.sheetName || input.sheetIndex || 'ActiveSheet'}" not found or could not be accessed.`, 'SHEET_NOT_FOUND', sheetError);
      }


      if (!worksheet) {
        // Should be caught by the try/catch above, but safeguard
        return createErrorResponse(`Worksheet "${input.sheetName || input.sheetIndex || 'ActiveSheet'}" could not be resolved.`, 'SHEET_NOT_FOUND');
      }

      const range = worksheet.Range(input.rangeAddress);

      if (!range) {
        // Release objects before throwing
        releaseObject(worksheet);
        return createErrorResponse(`Invalid range "${input.rangeAddress}".`, 'INVALID_RANGE');
      }

      let resultData: any = null;
      let message = '';

      switch (input.operation) {
        case 'read':
          // Read values from the range
          resultData = range.Value2; // Use Value2 for better type handling
          message = `Values read from range "${input.rangeAddress}".`;
          logger.info(message);
          break; // Added break

        case 'write':
          // Write values to the range
          if (!input.values) {
            releaseObject(range);
            releaseObject(worksheet);
            throw new Error('Values are required for the "write" operation.');
          }
          // Simplified write logic - assumes range is large enough or single cell
          // A more robust version would resize or check dimensions carefully
          range.Value = input.values; // Let COM handle potential type conversions
          message = `Values written to range "${input.rangeAddress}".`;
          logger.info(message);
          break; // Added break

        case 'format':
          // Apply format to the range
          if (!input.formatProperties) {
            releaseObject(range);
            releaseObject(worksheet);
            throw new Error('Format properties are required for the "format" operation.');
          }
          for (const prop in input.formatProperties) {
            if (Object.prototype.hasOwnProperty.call(input.formatProperties, prop)) {
              const value = input.formatProperties[prop];
              try {
                  if (typeof range[prop] === 'object' && range[prop] !== null && typeof value === 'object') {
                     // If the property is an object (like Font, Interior), apply nested properties
                     for (const subProp in value) {
                        if (Object.prototype.hasOwnProperty.call(value, subProp)) {
                           range[prop][subProp] = value[subProp];
                        }
                     }
                  } else {
                     // If the property is a direct value
                     range[prop] = value;
                  }
              } catch (formatError: any) {
                   logger.warn(`Could not apply format property '${prop}.${Object.keys(value).join(',')}' or '${prop}': ${formatError.message}`);
                   // Optionally continue or re-throw
              }
            }
          }
          message = `Format applied to range "${input.rangeAddress}".`;
          logger.info(message);
          break; // Added break

        case 'apply':
          // Generic operation to apply properties to the range
          if (!input.formatProperties) { // Reuse formatProperties for general properties
             releaseObject(range);
             releaseObject(worksheet);
             throw new Error('Properties are required for the "apply" operation.');
          }
           for (const prop in input.formatProperties) {
            if (Object.prototype.hasOwnProperty.call(input.formatProperties, prop)) {
               try {
                  range[prop] = input.formatProperties[prop];
               } catch (applyError: any) {
                   logger.warn(`Could not apply property '${prop}': ${applyError.message}`);
                   // Optionally continue or re-throw
               }
            }
          }
          message = `Properties applied to range "${input.rangeAddress}".`;
          logger.info(message);
          break; // Added break

        default:
           // Should not happen due to enum validation, but good practice
           releaseObject(range);
           releaseObject(worksheet);
           throw new Error(`Unsupported operation: "${input.operation}".`);
      }

      // Save the workbook if it was modified
      if (input.operation === 'write' || input.operation === 'format' || input.operation === 'apply') {
          if (!fileCreated) {
              workbook.Save();
              logger.info(`Workbook saved: ${absoluteFilePath}`);
          }
          // Save the modified Excel file as a dynamic resource
          try {
              // Ensure workbook is saved before reading
              const excelContent = await fs.readFile(absoluteFilePath, null); // Read as Buffer
              await saveResource('excel/range', path.basename(absoluteFilePath), excelContent);
              logger.info(`Saved ${absoluteFilePath} as a dynamic resource.`);
          } catch (resourceSaveError: any) {
              logger.error(`Failed to save ${absoluteFilePath} as a dynamic resource: ${resourceSaveError.message}`);
              // Continue execution even if resource saving fails
          }
      }

      // Release COM objects used within the switch
      releaseObject(range);
      releaseObject(worksheet);

      return { success: true, data: resultData, message }; // Return consistent structure

    } catch (error: any) {
      logger.error(`Error in excel/range tool: ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
       if (error instanceof z.ZodError) {
          return createErrorResponse('Input validation failed', 'VALIDATION_ERROR', error.errors);
      }
      return createErrorResponse(`Excel range operation failed: ${error.message}`, 'EXCEL_RANGE_ERROR', error);
    } finally {
      // Release workbook object if it was created/opened
      if (workbook) {
          // Don't close workbook here, let getOfficeApplication manage app lifecycle
          releaseObject(workbook);
      }
      // Release excelApp object
      if (excelApp) {
          releaseObject(excelApp);
      }
      logger.debug("Released Excel COM objects for excel/range operation.");
    }
  },
};

export default excelRangeTool;