/**
 * @file Tool for managing charts in Excel files.
 * Allows inserting, modifying, deleting, and repositioning charts.
 * Uses COM Interop via winax.
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

// Define the input schema for the excel/charts tool
const ExcelChartsInputSchema = z.object({
  filePath: z.string().describe('The path to the Excel file.'),
  operation: z.enum(['insert', 'modify', 'delete', 'reposition', 'list']).describe('The operation to perform.'),
  sheetName: z.string().optional().describe('The name of the worksheet. If not provided, the active sheet is used.'),
  sheetIndex: z.number().int().positive().optional().describe('The 1-based index of the worksheet. If provided, it overrides sheetName.'),
  rangeAddress: z.string().optional().describe('The data range address for the chart (e.g., "A1:B10"). Required for the "insert" operation.'),
  chartType: z.string().optional().describe('The chart type (e.g., "xlColumnClustered", "xlLine"). Required for the "insert" operation.'),
  chartTitle: z.string().optional().describe('The title of the chart.'),
  chartIndex: z.number().int().positive().optional().describe('The 1-based index of the chart object on the sheet. Required for "modify", "delete", "reposition".'),
  chartName: z.string().optional().describe('The name of the chart object on the sheet. Can be used instead of chartIndex for "modify", "delete", "reposition".'),
  position: z.object({
    left: z.number().optional(),
    top: z.number().optional(),
    width: z.number().optional(),
    height: z.number().optional(),
  }).optional().describe('The position and size of the chart. Required for "reposition".'),
  // Additional properties for modify (example)
  newRangeAddress: z.string().optional().describe('New data range address for the "modify" operation.'),
});

type ExcelChartsInput = z.infer<typeof ExcelChartsInputSchema>;

/**
 * @tool excel/charts
 * @description Manages charts in Excel files.
 * Allows inserting, modifying, deleting, and repositioning charts.
 * Uses COM Interop via winax.
 * @input ExcelChartsInputSchema
 * @output object - Depends on the operation.
 */
const excelChartsTool: McpResource = {
  path: 'excel/charts',
  description: 'Manages charts in Excel files.',
  schema: ExcelChartsInputSchema,
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => {
    let excelApp: any = null;
    let workbook: any = null;
    let worksheet: any = null;
    let filePath: string | undefined; // Declare filePath outside the try and allow undefined

    try {
      const input = ExcelChartsInputSchema.parse(params);
      filePath = input.filePath; // Assign filePath here

      const { operation, sheetName, sheetIndex, rangeAddress, chartType, chartTitle, chartIndex, chartName, position, newRangeAddress } = input;

      excelApp = await getOfficeApplication('Excel.Application');
      workbook = excelApp.Workbooks.Open(filePath);

      if (input.sheetName) {
        worksheet = workbook.Sheets(input.sheetName);
      } else if (input.sheetIndex) {
        worksheet = workbook.Sheets(input.sheetIndex);
      } else {
        worksheet = workbook.ActiveSheet;
      }

      if (!worksheet) {
        throw new Error(`Worksheet "${input.sheetName || input.sheetIndex}" not found.`);
      }

      const chartObjects = worksheet.ChartObjects();

      switch (operation) {
        case 'insert': {
          if (!rangeAddress || !chartType) {
            throw new Error('For inserting a chart, rangeAddress and chartType are required.');
          }
          const range = worksheet.Range(rangeAddress);
          if (!range) {
            throw new Error(`Invalid data range "${rangeAddress}".`);
          }

          // winax may require the numeric value of the chart type
          // Here we use a simple approach, strings could be mapped to COM constants
          const chartTypeValue = excelApp.constants[chartType] || parseInt(chartType, 10);
          if (isNaN(chartTypeValue)) {
             throw new Error(`Unrecognized chart type "${chartType}".`);
          }

          const chartObject = chartObjects.Add(0, 0, 300, 200); // Initial position and size
          const chart = chartObject.Chart;
          chart.SetSourceData(range);
          chart.ChartType = chartTypeValue;

          if (chartTitle) {
            chart.HasTitle = true;
            chart.ChartTitle.Text = chartTitle;
          }

          // Reposition if specified
          if (position) {
            if (position.left !== undefined) chartObject.Left = position.left;
            if (position.top !== undefined) chartObject.Top = position.top;
            if (position.width !== undefined) chartObject.Width = position.width;
            if (position.height !== undefined) chartObject.Height = position.height;
          }

          return { success: true, data: 'Chart inserted successfully.' }; // Adjusted return
        }

        case 'modify': {
          if (!chartIndex && !chartName) {
            throw new Error('For modifying a chart, chartIndex or chartName is required.');
          }
          const chartObject = chartIndex ? chartObjects.Item(chartIndex) : chartObjects.Item(chartName);
          if (!chartObject) {
            throw new Error(`Chart with index ${chartIndex} or name "${chartName}" not found.`);
          }
          const chart = chartObject.Chart;

          if (newRangeAddress) {
             const newRange = worksheet.Range(newRangeAddress);
             if (!newRange) {
                throw new Error(`Invalid new data range "${newRangeAddress}".`);
             }
             chart.SetSourceData(newRange);
          }

          if (chartType) {
             const chartTypeValue = excelApp.constants[chartType] || parseInt(chartType, 10);
             if (isNaN(chartTypeValue)) {
                throw new Error(`Unrecognized chart type "${chartType}".`);
             }
             chart.ChartType = chartTypeValue;
          }

          if (chartTitle) {
            chart.HasTitle = true;
            chart.ChartTitle.Text = chartTitle;
          } else if (chartTitle === '') { // Allow deleting the title
             chart.HasTitle = false;
          }

          // Modify size if specified
          if (position) {
            if (position.width !== undefined) chartObject.Width = position.width;
            if (position.height !== undefined) chartObject.Height = position.height;
          }


          return { success: true, data: 'Chart modified successfully.' }; // Adjusted return
        }

        case 'delete': {
          if (!chartIndex && !chartName) {
            throw new Error('For deleting a chart, chartIndex or chartName is required.');
          }
          const chartObject = chartIndex ? chartObjects.Item(chartIndex) : chartObjects.Item(chartName);
          if (!chartObject) {
            throw new Error(`Chart with index ${chartIndex} or name "${chartName}" not found.`);
          }
          chartObject.Delete();
          return { success: true, data: 'Chart deleted successfully.' }; // Adjusted return
        }

        case 'reposition': {
          if (!chartIndex && !chartName) {
            throw new Error('For repositioning a chart, chartIndex or chartName is required.');
          }
           if (!position) {
             throw new Error('For repositioning a chart, the position property is required.');
           }
          const chartObject = chartIndex ? chartObjects.Item(chartIndex) : chartObjects.Item(chartName);
          if (!chartObject) {
            throw new Error(`Chart with index ${chartIndex} or name "${chartName}" not found.`);
          }

          if (position.left !== undefined) chartObject.Left = position.left;
          if (position.top !== undefined) chartObject.Top = position.top;
          if (position.width !== undefined) chartObject.Width = position.width; // Allow modifying size when repositioning
          if (position.height !== undefined) chartObject.Height = position.height; // Allow modifying size when repositioning


          return { success: true, data: 'Chart repositioned successfully.' }; // Adjusted return
        }

        case 'list': {
            const chartsList = [];
            for (let i = 1; i <= chartObjects.Count; i++) {
                const chartObject = chartObjects.Item(i);
                chartsList.push({
                    index: i,
                    name: chartObject.Name,
                    left: chartObject.Left,
                    top: chartObject.Top,
                    width: chartObject.Width,
                    height: chartObject.Height,
                    chartTitle: chartObject.Chart.HasTitle ? chartObject.Chart.ChartTitle.Text : null,
                    chartType: chartObject.Chart.ChartType, // This will return a number, could be mapped to string if needed
                });
            }
            return { success: true, data: chartsList }; // Adjusted return
        }

        default:
          throw new Error(`Unsupported operation: "${operation}".`);
      }
    } catch (error: any) {
      // Here you can use your centralized error handler if you have one
      // For now, we return a simple ErrorResponse
      return { success: false, error: { code: 'EXCEL_CHART_ERROR', message: error.message } };
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
                    await saveResource('excel/charts', path.basename(filePath), excelContent);
                    // logger.info(`Saved ${filePath} as a dynamic resource.`);
                } catch (resourceSaveError: any) {
                    // logger.error(`Failed to save ${filePath} as a dynamic resource: ${resourceSaveError.message}`);
                    // Continue execution even if resource saving fails
                }
            }
            workbook.Close();
        } catch (closeError: any) {
            // Ignore errors when closing if there was already a main error
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
    return { success: false, error: { code: 'UNHANDLED_CASE', message: 'Excel chart operation did not return an explicit result.' } };
  },
};

export default excelChartsTool;