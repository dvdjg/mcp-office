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
import { McpResource, ToolRequestParams } from '../../types/common.types';
import { getOfficeApplication } from '../../utils/officeInterop';
import { saveResource } from '../dynamic/resources.tool'; // Import saveResource
import * as fs from 'fs-extra'; // Import fs to read the Excel file
import * as path from 'path'; // Import path

// Define the input schema for the excel/range tool
const ExcelRangeInputSchema = z.object({
  filePath: z.string().describe('Path to the Excel file.'),
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
 * Uses COM Interop via winax to interact with Excel.
 * Requires the file path, sheet name or index, range address, and operation.
 * For writing, values are provided as an array of arrays.
 * For formatting, format properties are provided as an object.
 */
const excelRangeTool: McpResource = {
  path: 'excel/range',
  description: 'Manipulates cell ranges in Excel worksheets.',
  schema: ExcelRangeInputSchema, // Added schema
  handler: async (params: any) => {
    let excelApp;
    try {
      // Validate input parameters
      const input = ExcelRangeInputSchema.parse(params);

      excelApp = await getOfficeApplication('Excel.Application'); // Corrected application name
      const workbook = excelApp.Workbooks.Open(input.filePath);
      let worksheet;

      if (input.sheetIndex !== undefined) {
        worksheet = workbook.Sheets.Item(input.sheetIndex);
      } else if (input.sheetName) {
        worksheet = workbook.Sheets.Item(input.sheetName);
      } else {
        worksheet = workbook.ActiveSheet;
      }

      if (!worksheet) {
        throw new Error(`Worksheet "${input.sheetName || input.sheetIndex}" not found.`);
      }

      const range = worksheet.Range(input.rangeAddress);

      if (!range) {
        throw new Error(`Invalid range "${input.rangeAddress}".`);
      }

      switch (input.operation) {
        case 'read':
          // Read values from the range
          const values = range.Value2;
          return { success: true, data: values }; // Success return with data

        case 'write':
          // Write values to the range
          if (!input.values) {
            throw new Error('Values are required for the "write" operation.');
          }
          // Ensure the range size matches the values size
          // This is a simplification; a robust implementation should handle different sizes
          // or allow writing to a larger destination range.
          if (range.Rows.Count !== input.values.length || range.Columns.Count !== (input.values[0]?.length || 0)) {
             // If the destination range is a single cell, winax can handle writing an array of arrays
             if (range.Cells.Count === 1) {
                range.Value = input.values;
             } else {
                // For larger ranges, attempting to write directly may fail if dimensions don't match exactly.
                // An alternative would be to iterate over cells or use CopyFromRecordset if data comes from a compatible source.
                // For now, we throw an error if dimensions don't match for ranges > 1 cell.
                 throw new Error(`Dimensions of provided values do not match the range dimensions. Range: ${range.Rows.Count}x${range.Columns.Count}, Values: ${input.values.length}x${(input.values[0]?.length || 0)}`);
             }
          } else {
             range.Value = input.values;
          }
          workbook.Save();
          // Save the modified Excel file as a dynamic resource
          try {
              const excelContent = await fs.readFile(input.filePath, null); // Read as Buffer
              await saveResource('excel/range', path.basename(input.filePath), excelContent);
              // logger.info(`Saved ${input.filePath} as a dynamic resource.`);
          } catch (resourceSaveError: any) {
              // logger.error(`Failed to save ${input.filePath} as a dynamic resource: ${resourceSaveError.message}`);
              // Continue execution even if resource saving fails
          }
          return { success: true, data: null, message: `Values written to range "${input.rangeAddress}".` }; // Include data: null

        case 'format':
          // Apply format to the range
          if (!input.formatProperties) {
            throw new Error('Format properties are required for the "format" operation.');
          }
          // Basic format implementation. Can be expanded to support more properties.
          // Example: { Font: { Bold: true, Color: 255 }, Interior: { ColorIndex: 6 } }
          for (const prop in input.formatProperties) {
            if (Object.prototype.hasOwnProperty.call(input.formatProperties, prop)) {
              const value = input.formatProperties[prop];
              if (typeof range[prop] === 'object' && range[prop] !== null) {
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
            }
          }
          workbook.Save();
          // Save the modified Excel file as a dynamic resource
          try {
              const excelContent = await fs.readFile(input.filePath, null); // Read as Buffer
              await saveResource('excel/range', path.basename(input.filePath), excelContent);
              // logger.info(`Saved ${input.filePath} as a dynamic resource.`);
          } catch (resourceSaveError: any) {
              // logger.error(`Failed to save ${input.filePath} as a dynamic resource: ${resourceSaveError.message}`);
              // Continue execution even if resource saving fails
          }
          return { success: true, data: null, message: `Format applied to range "${input.rangeAddress}".` }; // Include data: null

        case 'apply':
          // Generic operation to apply properties to the range
          if (!input.formatProperties) { // Reuse formatProperties for general properties
             throw new Error('Properties are required for the "apply" operation.');
          }
           for (const prop in input.formatProperties) {
            if (Object.prototype.hasOwnProperty.call(input.formatProperties, prop)) {
              range[prop] = input.formatProperties[prop];
            }
          }
          workbook.Save();
          // Save the modified Excel file as a dynamic resource
          try {
              const excelContent = await fs.readFile(input.filePath, null); // Read as Buffer
              await saveResource('excel/range', path.basename(input.filePath), excelContent);
              // logger.info(`Saved ${input.filePath} as a dynamic resource.`);
          } catch (resourceSaveError: any) {
              // logger.error(`Failed to save ${input.filePath} as a dynamic resource: ${resourceSaveError.message}`);
              // Continue execution even if resource saving fails
          }
          return { success: true, data: null, message: `Properties applied to range "${input.rangeAddress}".` }; // Include data: null

        default:
          throw new Error(`Unsupported operation: "${input.operation}".`);
      }
    } catch (error: any) {
      // Basic error handling
      return { success: false, error: { code: 'OFFICE_API_ERROR', message: error.message } }; // Adjust error format
    } finally {
      // Consider whether to close Excel or leave it open.
      // For an MCP server, you probably want to leave it open for future operations.
      // If you decide to close it, make sure to save first if necessary.
      // if (excelApp) {
      //   excelApp.Quit();
      // }
    }
  },
};

export default excelRangeTool;