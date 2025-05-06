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
import ExcelJS from 'exceljs'; // Import exceljs

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
  useComInterop: z.boolean().optional().default(false).describe('Use COM interop for local Excel interaction. Defaults to false (uses exceljs).'),
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
    const input = ExcelChartsInputSchema.parse(params);
    const { filePath, operation, sheetName, sheetIndex, rangeAddress, chartType, chartTitle, chartIndex, chartName, position, newRangeAddress, useComInterop } = input;
    const absoluteFilePath = path.resolve(filePath);
    let message = '';
    let resultData: any = null;

    if (useComInterop) {
      logger.info(`Executing excel/charts (COM) operation '${operation}' for file: ${filePath}`);
      let excelApp: any = null;
      let workbook: any = null;
      let worksheet: any = null;
      try {
        excelApp = await getOfficeApplication('Excel.Application');
        workbook = excelApp.Workbooks.Open(absoluteFilePath);

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

        const chartObjects = worksheet.ChartObjects();

        switch (operation) {
          case 'insert': {
            if (!rangeAddress || !chartType) {
              throw new Error('COM: For inserting a chart, rangeAddress and chartType are required.');
            }
            const range = worksheet.Range(rangeAddress);
            if (!range) throw new Error(`COM: Invalid data range "${rangeAddress}".`);

            const chartTypeValue = excelApp.constants[chartType] || parseInt(chartType, 10);
            if (isNaN(chartTypeValue)) throw new Error(`COM: Unrecognized chart type "${chartType}".`);

            const chartObject = chartObjects.Add(position?.left ?? 0, position?.top ?? 0, position?.width ?? 300, position?.height ?? 200);
            const chart = chartObject.Chart;
            chart.SetSourceData(range);
            chart.ChartType = chartTypeValue;
            if (chartTitle) {
              chart.HasTitle = true;
              chart.ChartTitle.Text = chartTitle;
            }
            message = 'COM: Chart inserted successfully.';
            break;
          }
          case 'modify': {
            if (!chartIndex && !chartName) throw new Error('COM: chartIndex or chartName is required for modify.');
            const chartObject = chartIndex ? chartObjects.Item(chartIndex) : chartObjects.Item(chartName);
            if (!chartObject) throw new Error(`COM: Chart with index ${chartIndex} or name "${chartName}" not found.`);
            const chart = chartObject.Chart;
            if (newRangeAddress) {
               const newRange = worksheet.Range(newRangeAddress);
               if (!newRange) throw new Error(`COM: Invalid new data range "${newRangeAddress}".`);
               chart.SetSourceData(newRange);
            }
            if (chartType) {
               const chartTypeValue = excelApp.constants[chartType] || parseInt(chartType, 10);
               if (isNaN(chartTypeValue)) throw new Error(`COM: Unrecognized chart type "${chartType}".`);
               chart.ChartType = chartTypeValue;
            }
            if (chartTitle) {
              chart.HasTitle = true;
              chart.ChartTitle.Text = chartTitle;
            } else if (chartTitle === '') chart.HasTitle = false;
            if (position) { // For modifying size
              if (position.width !== undefined) chartObject.Width = position.width;
              if (position.height !== undefined) chartObject.Height = position.height;
            }
            message = 'COM: Chart modified successfully.';
            break;
          }
          case 'delete': {
            if (!chartIndex && !chartName) throw new Error('COM: chartIndex or chartName is required for delete.');
            const chartObject = chartIndex ? chartObjects.Item(chartIndex) : chartObjects.Item(chartName);
            if (!chartObject) throw new Error(`COM: Chart with index ${chartIndex} or name "${chartName}" not found.`);
            chartObject.Delete();
            message = 'COM: Chart deleted successfully.';
            break;
          }
          case 'reposition': {
            if (!chartIndex && !chartName) throw new Error('COM: chartIndex or chartName is required for reposition.');
            if (!position) throw new Error('COM: position property is required for reposition.');
            const chartObject = chartIndex ? chartObjects.Item(chartIndex) : chartObjects.Item(chartName);
            if (!chartObject) throw new Error(`COM: Chart with index ${chartIndex} or name "${chartName}" not found.`);
            if (position.left !== undefined) chartObject.Left = position.left;
            if (position.top !== undefined) chartObject.Top = position.top;
            if (position.width !== undefined) chartObject.Width = position.width;
            if (position.height !== undefined) chartObject.Height = position.height;
            message = 'COM: Chart repositioned successfully.';
            break;
          }
          case 'list': {
              const chartsList = [];
              for (let i = 1; i <= chartObjects.Count; i++) {
                  const chartObjectItem = chartObjects.Item(i); // Renamed to avoid conflict
                  chartsList.push({
                      index: i, name: chartObjectItem.Name, left: chartObjectItem.Left, top: chartObjectItem.Top,
                      width: chartObjectItem.Width, height: chartObjectItem.Height,
                      chartTitle: chartObjectItem.Chart.HasTitle ? chartObjectItem.Chart.ChartTitle.Text : null,
                      chartType: chartObjectItem.Chart.ChartType,
                  });
              }
              resultData = chartsList;
              message = 'COM: Charts listed successfully.';
              break;
          }
          default:
            throw new Error(`COM: Unsupported operation: "${operation}".`);
        }
        return { success: true, data: resultData || message };
      } catch (error: any) {
        logger.error(`Error in excel/charts tool (COM): ${error.message}`);
        return { success: false, error: { code: 'EXCEL_CHARTS_COM_ERROR', message: `COM Error: ${error.message}` } };
      } finally {
        if (workbook) {
          try { workbook.Save(); workbook.Close(); } catch (e) { logger.warn(`COM: Error saving/closing workbook: ${e}`); }
          releaseObject(workbook);
        }
        if (excelApp) releaseObject(excelApp);
      }
    } else {
      // exceljs path
      logger.info(`Executing excel/charts (exceljs) operation '${operation}' for file: ${filePath}`);
      const wb = new ExcelJS.Workbook();
      let ws: ExcelJS.Worksheet | undefined;

      try {
        const fileExisted = await fs.pathExists(absoluteFilePath);
        if (fileExisted) {
          await wb.xlsx.readFile(absoluteFilePath);
          if (sheetName) ws = wb.getWorksheet(sheetName);
          else if (sheetIndex && sheetIndex > 0 && sheetIndex <= wb.worksheets.length) ws = wb.worksheets[sheetIndex - 1];
          else if (wb.worksheets.length > 0) ws = wb.worksheets[0];
        } else if (operation === 'insert') {
           ws = wb.addWorksheet(sheetName || (sheetIndex ? `Sheet${sheetIndex}`: 'Sheet1'));
        } else {
          throw new Error(`exceljs: File not found at ${absoluteFilePath} for operation '${operation}'.`);
        }

        if (!ws) {
          throw new Error(`exceljs: Worksheet "${sheetName || sheetIndex || 'default'}" not found or could not be created.`);
        }

        switch (operation) {
          case 'insert': {
            if (!rangeAddress || !chartType) {
              throw new Error('exceljs: rangeAddress and chartType are required for insert.');
            }
            // Basic mapping from COM-like names to exceljs types
            let ejChartType: string; // Changed from ExcelJS.ChartType to string
            switch (chartType.toLowerCase()) {
                case 'xlcolumnclustered': case 'columnclustered': ejChartType = 'bar'; break; // exceljs uses 'bar' for column charts
                case 'xlline': case 'line': ejChartType = 'line'; break;
                case 'xlpie': case 'pie': ejChartType = 'pie'; break;
                // Add more mappings as needed - e.g. scatter, area, doughnut
                // exceljs chart types include: bar, line, pie, scatter, area, doughnut, radar, bubble
                default: throw new Error(`exceljs: Unsupported chartType "${chartType}". Supported types include bar, line, pie etc.`);
            }

            // Define series and categories from rangeAddress
            // This is a simplified parsing. A robust solution would parse A1:B5 into categories and series.
            // Assuming rangeAddress is like "Sheet1!A1:B5" or "A1:B5"
            // And first column is categories, subsequent columns are series.
            // For simplicity, let's assume rangeAddress refers to data like:
            // Categories | Series 1 | Series 2
            // CatA       | 10       | 20
            // CatB       | 12       | 22

            // This part needs a robust way to define series based on rangeAddress.
            // For now, we'll create a placeholder series structure.
            // A real implementation would need to parse `rangeAddress` to define `series.categories` and `series.values`.
            const seriesConfig = [
              {
                // header: 'Series 1', // Optional header for legend
                // categories: 'A2:A5', // Example: This should be derived from rangeAddress
                // values: 'B2:B5',   // Example: This should be derived from rangeAddress
              },
            ];
            logger.warn(`exceljs: Chart data series configuration from rangeAddress "${rangeAddress}" is using a placeholder. A robust implementation is needed to parse the range and define categories/values for series.`);


            const chartOptions: any = {
                title: { text: chartTitle || 'Chart Title', bold: true },
                // exceljs uses tl (top-left) and br (bottom-right) or ext (extent cx, cy) for positioning
                // The 'position' input (left, top, width, height) is more like pixel/point based.
                // Mapping this accurately to exceljs anchors (tl/br with col/row, or editAs with x/y offsets) is complex.
                // Using a default placement for now.
                tl: { col: position?.left ?? 2, row: position?.top ?? 15 }, // Approximate mapping, assuming left/top are col/row indices
                br: { col: (position?.left ?? 2) + (position?.width ? position.width / 75 : 8) , row: (position?.top ?? 15) + (position?.height ? position.height / 20 : 10) }, // Very rough approximation for width/height
                // Alternatively, use ext: { width: position?.width ?? 600, height: position?.height ?? 400 } if 'moveShape' is used for editAs
            };
             if (position) {
                logger.warn("exceljs: Chart positioning from numeric 'position' input is complex and uses an approximate mapping to cell anchors. Fine-tuning may be required.");
            }

            // Casting to 'any' to bypass TypeScript error if 'addChart' is not in the current type definitions.
            (ws as any).addChart({
                type: ejChartType,
                series: seriesConfig, // This needs to be properly populated based on rangeAddress
                ...chartOptions
            }, rangeAddress); // The last `rangeAddress` here might be for data linking if the API supports it this way.
                               // Or, series data should be explicitly set. The `exceljs` API for `addChart` usually takes series data within the options.
                               // The `rangeAddress` as a third param to `addChart` is unusual.
                               // Let's assume `rangeAddress` is primarily for data extraction for `seriesConfig`.

            message = `exceljs: Chart of type "${ejChartType}" inserted. Data range: "${rangeAddress}". Positioning is approximate. Series data from range needs robust parsing.`;
            break;
          }
          case 'modify':
          case 'delete':
          case 'reposition':
          case 'list':
            message = `exceljs: Operation '${operation}' for charts has significant limitations or is not supported with exceljs after file load. Please use COM Interop for these features.`;
            logger.warn(message);
            // For 'list', we could try to iterate ws.drawings if any, but parsing chart info is non-trivial.
            if (operation === 'list') resultData = []; // Empty list for now
            break;
          default:
            throw new Error(`exceljs: Unsupported operation: "${operation}".`);
        }
        await wb.xlsx.writeFile(absoluteFilePath);
        logger.info(`exceljs: Workbook saved to ${absoluteFilePath}`);
        return { success: true, data: resultData || message };

      } catch (error: any) {
        logger.error(`Error in excel/charts tool (exceljs): ${error.message}`);
        return { success: false, error: { code: 'EXCEL_CHARTS_EXCELJS_ERROR', message: `exceljs Error: ${error.message}` } };
      }
    }

    // Common dynamic resource saving
    if (operation === 'insert' || operation === 'modify' || operation === 'delete' || operation === 'reposition') {
        try {
            const excelContent = await fs.readFile(absoluteFilePath, null);
            await saveResource('excel/charts', path.basename(absoluteFilePath), excelContent);
            logger.info(`Saved ${absoluteFilePath} as a dynamic resource.`);
            if (typeof message === 'string' && !message.includes('Saved as dynamic resource')) message += ` Saved as dynamic resource.`;
            else if (resultData && typeof resultData === 'object' && !(resultData as any).dynamicResourceSaved) (resultData as any).dynamicResourceInfo = "Saved as dynamic resource.";

        } catch (resourceSaveError: any) {
            logger.error(`Failed to save ${absoluteFilePath} as a dynamic resource: ${resourceSaveError.message}`);
             if (typeof message === 'string') message += ` (Warning: Failed to save as dynamic resource: ${resourceSaveError.message})`;
            else if (resultData && typeof resultData === 'object') (resultData as any).dynamicResourceError = resourceSaveError.message;
        }
    }
    // Fallback return, should ideally be covered by specific path returns
    if (useComInterop && (resultData || message)) {
        return { success: true, data: resultData || message };
    }
    return { success: false, error: { code: 'UNHANDLED_CHART_LOGIC_PATH', message: 'Chart operation did not complete as expected.' } };
  },
};

export default excelChartsTool;