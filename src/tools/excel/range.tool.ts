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
import ExcelJS from 'exceljs'; // Import exceljs
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
  useComInterop: z.boolean().optional().default(false).describe('Use COM interop for local Excel interaction. Defaults to false (uses exceljs).'),
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
    try {
      // Validate input parameters outside the conditional block
      const input = ExcelRangeInputSchema.parse(params);
      const absoluteFilePath = path.resolve(input.filePath);
      let message = '';
      let resultData: any = null;

      if (input.useComInterop) {
        logger.info(`Executing excel/range (COM) operation '${input.operation}' for file: ${input.filePath}`);
        let excelApp: any = null;
        let workbook: any = null;
        let fileCreated = false;

        try {
          excelApp = await getOfficeApplication('Excel.Application');

          if (input.operation === 'write' || input.operation === 'format' || input.operation === 'apply') {
              try {
                  if (await fs.pathExists(absoluteFilePath)) {
                      logger.info(`COM: Opening existing workbook: ${absoluteFilePath}`);
                      workbook = excelApp.Workbooks.Open(absoluteFilePath);
                  } else {
                      logger.info(`COM: File not found. Creating new workbook at: ${absoluteFilePath}`);
                      workbook = excelApp.Workbooks.Add();
                      const fileExt = path.extname(absoluteFilePath).toLowerCase();
                      let saveFormat = 51; // xlOpenXMLWorkbook (.xlsx)
                      if (fileExt === '.xlsb') saveFormat = 50;
                      else if (fileExt === '.xls') saveFormat = 56;
                      workbook.SaveAs(absoluteFilePath, saveFormat);
                      fileCreated = true;
                      logger.info(`COM: Successfully created and saved new workbook: ${absoluteFilePath}`);
                  }
              } catch (fileError: any) {
                   logger.error(`COM: Error opening or creating workbook '${absoluteFilePath}': ${fileError.message}`, { error: fileError });
                   return createErrorResponse(`COM: Failed to open or create workbook: ${fileError.message}`, 'FILE_OPERATION_FAILED_COM', fileError);
              }
          } else if (input.operation === 'read') {
              if (!await fs.pathExists(absoluteFilePath)) {
                  logger.warn(`COM: File not found for read operation: ${absoluteFilePath}`);
                  return createErrorResponse(`COM: File not found: ${input.filePath}`, 'FILE_NOT_FOUND_COM');
              }
              logger.info(`COM: Opening existing workbook for read: ${absoluteFilePath}`);
              workbook = excelApp.Workbooks.Open(absoluteFilePath);
          }

          if (!workbook) {
               return createErrorResponse(`COM: Failed to obtain workbook object for: ${input.filePath}`, 'FILE_OPEN_FAILED_COM');
          }

          let worksheet;
          try {
              if (input.sheetIndex !== undefined) {
                worksheet = workbook.Sheets.Item(input.sheetIndex);
              } else if (input.sheetName) {
                worksheet = workbook.Sheets.Item(input.sheetName);
              } else {
                worksheet = workbook.ActiveSheet;
              }
          } catch (sheetError: any) {
               logger.error(`COM: Error getting worksheet '${input.sheetName || input.sheetIndex || 'ActiveSheet'}': ${sheetError.message}`);
               return createErrorResponse(`COM: Worksheet "${input.sheetName || input.sheetIndex || 'ActiveSheet'}" not found or could not be accessed.`, 'SHEET_NOT_FOUND_COM', sheetError);
          }

          if (!worksheet) {
            return createErrorResponse(`COM: Worksheet "${input.sheetName || input.sheetIndex || 'ActiveSheet'}" could not be resolved.`, 'SHEET_NOT_FOUND_COM');
          }

          const range = worksheet.Range(input.rangeAddress);

          if (!range) {
            releaseObject(worksheet);
            return createErrorResponse(`COM: Invalid range "${input.rangeAddress}".`, 'INVALID_RANGE_COM');
          }

          switch (input.operation) {
            case 'read':
              resultData = range.Value2;
              message = `COM: Values read from range "${input.rangeAddress}".`;
              logger.info(message);
              break;

            case 'write':
              if (!input.values) {
                releaseObject(range);
                releaseObject(worksheet);
                throw new Error('COM: Values are required for the "write" operation.');
              }
              range.Value = input.values;
              message = `COM: Values written to range "${input.rangeAddress}".`;
              logger.info(message);
              break;

            case 'format':
              if (!input.formatProperties) {
                releaseObject(range);
                releaseObject(worksheet);
                throw new Error('COM: Format properties are required for the "format" operation.');
              }
              for (const prop in input.formatProperties) {
                if (Object.prototype.hasOwnProperty.call(input.formatProperties, prop)) {
                  const value = input.formatProperties[prop];
                  try {
                      if (typeof range[prop] === 'object' && range[prop] !== null && typeof value === 'object') {
                         for (const subProp in value) {
                            if (Object.prototype.hasOwnProperty.call(value, subProp)) {
                               range[prop][subProp] = value[subProp];
                            }
                         }
                      } else {
                         range[prop] = value;
                      }
                  } catch (formatError: any) {
                       logger.warn(`COM: Could not apply format property '${prop}': ${formatError.message}`);
                  }
                }
              }
              message = `COM: Format applied to range "${input.rangeAddress}".`;
              logger.info(message);
              break;

            case 'apply':
              if (!input.formatProperties) {
                 releaseObject(range);
                 releaseObject(worksheet);
                 throw new Error('COM: Properties are required for the "apply" operation.');
              }
               for (const prop in input.formatProperties) {
                if (Object.prototype.hasOwnProperty.call(input.formatProperties, prop)) {
                   try {
                      range[prop] = input.formatProperties[prop];
                   } catch (applyError: any) {
                       logger.warn(`COM: Could not apply property '${prop}': ${applyError.message}`);
                   }
                }
              }
              message = `COM: Properties applied to range "${input.rangeAddress}".`;
              logger.info(message);
              break;

            default:
               releaseObject(range);
               releaseObject(worksheet);
               throw new Error(`COM: Unsupported operation: "${input.operation}".`);
          }

          if (input.operation === 'write' || input.operation === 'format' || input.operation === 'apply') {
              if (!fileCreated) { // Only save if it wasn't a new file (SaveAs already saved it)
                  workbook.Save();
                  logger.info(`COM: Workbook saved: ${absoluteFilePath}`);
              }
          }
          releaseObject(range);
          releaseObject(worksheet);

        } catch (error: any) {
          logger.error(`Error in excel/range tool (COM): ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
           if (error instanceof z.ZodError) {
              return createErrorResponse('COM: Input validation failed', 'VALIDATION_ERROR_COM', error.errors);
          }
          return createErrorResponse(`COM: Excel range operation failed: ${error.message}`, 'EXCEL_RANGE_ERROR_COM', error);
        } finally {
          if (workbook) {
              releaseObject(workbook);
          }
          if (excelApp) {
              releaseObject(excelApp);
          }
          logger.debug("COM: Released Excel COM objects for excel/range operation.");
        }
      } else {
        // Use exceljs
        logger.info(`Executing excel/range (exceljs) operation '${input.operation}' for file: ${input.filePath}`);
        const excelWorkbook = new ExcelJS.Workbook();
        let fileExisted = false;

        try {
          if (await fs.pathExists(absoluteFilePath)) {
            await excelWorkbook.xlsx.readFile(absoluteFilePath);
            fileExisted = true;
            logger.info(`exceljs: Opened existing workbook: ${absoluteFilePath}`);
          } else if (input.operation === 'write' || input.operation === 'format' || input.operation === 'apply') {
            logger.info(`exceljs: File not found. New workbook will be created at: ${absoluteFilePath}`);
          } else if (input.operation === 'read') {
            logger.warn(`exceljs: File not found for read operation: ${absoluteFilePath}`);
            return createErrorResponse(`exceljs: File not found: ${input.filePath}`, 'FILE_NOT_FOUND_EXCELJS');
          }

          let worksheet: ExcelJS.Worksheet;
          if (input.sheetIndex !== undefined) {
            if (input.sheetIndex > 0 && input.sheetIndex <= excelWorkbook.worksheets.length) {
              worksheet = excelWorkbook.worksheets[input.sheetIndex - 1]; // 0-based
            } else if ((input.operation === 'write' || input.operation === 'format' || input.operation === 'apply') && ((fileExisted && excelWorkbook.worksheets.length === 0) || !fileExisted)) {
              const newSheetName = input.sheetName || `Sheet${input.sheetIndex}`;
              worksheet = excelWorkbook.addWorksheet(newSheetName);
              logger.info(`exceljs: Added new sheet "${worksheet.name}" as index ${input.sheetIndex} was out of bounds or file was new.`);
            } else {
              return createErrorResponse(`exceljs: Sheet index ${input.sheetIndex} is out of bounds.`, 'SHEET_NOT_FOUND_EXCELJS');
            }
          } else if (input.sheetName) {
            let foundSheet = excelWorkbook.getWorksheet(input.sheetName);
            if (foundSheet) {
              worksheet = foundSheet;
            } else if (input.operation === 'write' || input.operation === 'format' || input.operation === 'apply') {
              worksheet = excelWorkbook.addWorksheet(input.sheetName);
              logger.info(`exceljs: Added new sheet "${input.sheetName}" as it was not found.`);
            } else {
               return createErrorResponse(`exceljs: Worksheet "${input.sheetName}" not found.`, 'SHEET_NOT_FOUND_EXCELJS');
            }
          } else {
            if (excelWorkbook.worksheets.length > 0) {
              worksheet = excelWorkbook.worksheets[0];
            } else if (input.operation === 'write' || input.operation === 'format' || input.operation === 'apply') {
              worksheet = excelWorkbook.addWorksheet('Sheet1');
              logger.info(`exceljs: Added new default sheet "Sheet1".`);
            } else {
              return createErrorResponse(`exceljs: No worksheets found in the file and no specific sheet requested for read.`, 'SHEET_NOT_FOUND_EXCELJS');
            }
          }

          if (!worksheet) {
            return createErrorResponse(`exceljs: Worksheet "${input.sheetName || input.sheetIndex || 'default'}" could not be resolved. This indicates an unexpected state.`, 'SHEET_RESOLUTION_ERROR_EXCELJS');
          }
          
          switch (input.operation) {
            case 'read':
              const cellToRead = worksheet.getCell(input.rangeAddress.split(':')[0]);
              resultData = cellToRead.value;
              message = `exceljs: Value read from cell "${input.rangeAddress.split(':')[0]}" (Note: full range read needs specific parsing).`;
              logger.info(message);
              break;

            case 'write':
              if (!input.values) {
                throw new Error('exceljs: Values are required for the "write" operation.');
              }
              const cellToWrite = worksheet.getCell(input.rangeAddress.split(':')[0]);
              cellToWrite.value = input.values[0]?.[0] ?? null;
              message = `exceljs: Value written to cell "${input.rangeAddress.split(':')[0]}" (Note: full range write needs specific parsing and iteration).`;
              logger.info(message);
              break;

            case 'format':
              if (!input.formatProperties) {
                throw new Error('exceljs: Format properties are required for the "format" operation.');
              }
              const cellToFormat = worksheet.getCell(input.rangeAddress.split(':')[0]);
              Object.assign(cellToFormat, input.formatProperties);
              message = `exceljs: Format attempted on cell "${input.rangeAddress.split(':')[0]}" (Note: formatting needs detailed mapping from COM to exceljs style).`;
              logger.warn("exceljs: Format application is currently very basic and may not support complex COM-style properties directly.");
              break;
              
            case 'apply':
              if (!input.formatProperties) {
                   throw new Error('exceljs: Properties are required for the "apply" operation.');
              }
              const cellToApply = worksheet.getCell(input.rangeAddress.split(':')[0]);
              Object.assign(cellToApply, input.formatProperties);
              message = `exceljs: Properties applied to cell "${input.rangeAddress.split(':')[0]}" (Note: needs detailed mapping).`;
              logger.warn("exceljs: Apply operation is currently very basic.");
              break;

            default:
              throw new Error(`exceljs: Unsupported operation: "${input.operation}".`);
          }

          if (input.operation === 'write' || input.operation === 'format' || input.operation === 'apply') {
            await excelWorkbook.xlsx.writeFile(absoluteFilePath);
            logger.info(`exceljs: Workbook saved: ${absoluteFilePath}`);
          }
        } catch (error: any) {
          logger.error(`Error in excel/range tool (exceljs): ${error.message}`, { error });
          return createErrorResponse(`exceljs: Excel range operation failed: ${error.message}`, 'EXCEL_RANGE_ERROR_EXCELJS', error);
        }
      }

      // Save the modified Excel file as a dynamic resource (common for both paths if successful)
      if (input.operation === 'write' || input.operation === 'format' || input.operation === 'apply') {
          try {
              const excelContent = await fs.readFile(absoluteFilePath, null); // Read as Buffer
              await saveResource('excel/range', path.basename(absoluteFilePath), excelContent);
              logger.info(`Saved ${absoluteFilePath} as a dynamic resource.`);
              if (typeof message === 'string') message += ` Saved as dynamic resource.`;
              else if (resultData && typeof resultData === 'object') (resultData as any).dynamicResourceSaved = true;

          } catch (resourceSaveError: any) {
              logger.error(`Failed to save ${absoluteFilePath} as a dynamic resource: ${resourceSaveError.message}`);
              if (typeof message === 'string') message += ` (Warning: Failed to save as dynamic resource: ${resourceSaveError.message})`;
              else if (resultData && typeof resultData === 'object') (resultData as any).dynamicResourceError = resourceSaveError.message;
          }
      }
      return { success: true, data: resultData, message };

    } catch (error: any) {
      if (error instanceof z.ZodError) {
        logger.error(`Input validation failed for excel/range: ${error.message}`, { errors: error.errors });
        return createErrorResponse('Input validation failed', 'VALIDATION_ERROR', error.errors);
      }
      logger.error(`Unexpected error in excel/range handler: ${error.message}`, { error });
      return createErrorResponse(`An unexpected error occurred: ${error.message}`, 'UNEXPECTED_HANDLER_ERROR', error);
    }
  },
};

export default excelRangeTool;