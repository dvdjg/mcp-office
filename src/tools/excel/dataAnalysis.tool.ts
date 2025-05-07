/**
 * @file Tool for performing data analysis operations (sort, filter, pivot, calculate) on Excel ranges or tables.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z, ZodError } from 'zod';
import { McpResource, ApiResponse, SuccessResponse, ErrorResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types.js';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js'; // Import releaseObject
import logger from '../../utils/logger.js'; // Import logger
import ExcelJS from 'exceljs'; // Import exceljs
import * as fs from 'fs-extra'; // Import fs to read the Excel file
import * as path from 'path'; // Import path

// --- COM Constants (Examples, verify in Excel documentation) ---
// https://learn.microsoft.com/en-us/office/vba/api/excel.xlautofilteroperator
const XlAutoFilterOperator = {
    xlAnd: 1,
    xlBottom10Items: 4,
    xlBottom10Percent: 6,
    xlOr: 2,
    xlTop10Items: 3,
    xlTop10Percent: 5,
    xlFilterAutomaticDateGroups: 9,
    xlFilterDynamic: 11,
    xlFilterIcon: 10,
    xlFilterValues: 7,
};

// https://learn.microsoft.com/en-us/office/vba/api/excel.xlconsolidationfunction
const XlConsolidationFunction = {
    xlAverage: 1,
    xlCount: 2,
    xlCountNums: 3,
    xlMax: 4,
    xlMin: 5,
    xlProduct: 6,
    xlStDev: 7,
    xlStDevP: 8,
    xlSum: 9,
    xlVar: 10,
    xlVarP: 11,
};

// https://learn.microsoft.com/en-us/office/vba/api/excel.xlheadershow
const XlHeaderShow = {
    xlYes: 1,
    xlNo: 2,
    xlGuess: 3,
};


// --- Schemas ---

// Combined schema for all operations (using z.object)
const ExcelDataAnalysisInputSchema = z.object({
    operation: z.enum(['sort', 'filter', 'pivot', 'calculate']).describe('The operation to perform (sort, filter, pivot, or calculate).'),
    // Include all possible fields for the operations
    filePath: z.string().describe('Path to the Excel file.'),
    sheetName: z.string().optional().describe('Name of the worksheet. If not provided, the active sheet is used.'),
    sheetIndex: z.number().int().positive().optional().describe('1-based index of the worksheet. If provided, it overrides sheetName.'),
    rangeAddress: z.string().optional().describe('Range address (e.g., "A1:D10"). Required if tableName is not used.'),
    tableName: z.string().optional().describe('Name of the table. Required if rangeAddress is not used.'),

    // Specific fields for 'sort'
    sortCriteria: z.array(z.object({
        column: z.union([z.string(), z.number().int().positive()]).describe('Column to sort by (column name or 1-based index).'),
        order: z.enum(['Ascending', 'Descending']).default('Ascending').describe('Sort order.'),
    })).min(1).optional().describe('Sort criteria. Required for "sort".'),
    header: z.nativeEnum(XlHeaderShow).default(XlHeaderShow.xlGuess).optional().describe('Specifies if the range has headers. Used in "sort".'),

    // Specific fields for 'filter'
    filterCriteria: z.array(z.object({
        column: z.union([z.string(), z.number().int().positive()]).describe('Column to filter by (column name or 1-based index).'),
        criteria1: z.any().describe('First filter criterion.'),
        operator: z.nativeEnum(XlAutoFilterOperator).optional().describe('Filter operator (XlAutoFilterOperator constant).'),
        criteria2: z.any().optional().describe('Second filter criterion (for operators requiring two).'),
        visibleDropDown: z.boolean().default(true).optional().describe('Show or hide the AutoFilter drop-down button.'),
    })).min(1).optional().describe('Filter criteria. Required for "filter".'),

    // Specific fields for 'pivot'
    pivotTableParameters: z.object({
        pivotTableName: z.string().describe('Name for the new pivot table.'),
        destination: z.string().describe('Cell where the pivot table will be placed (e.g., "Sheet2!A1").'),
        rowFields: z.array(z.union([z.string(), z.number().int().positive()])).optional().describe('Row fields.'),
        columnFields: z.array(z.union([z.string(), z.number().int().positive()])).optional().describe('Column fields.'),
        dataFields: z.array(z.object({
            field: z.union([z.string(), z.number().int().positive()]).describe('Data field.'),
            function: z.nativeEnum(XlConsolidationFunction).optional().describe('Summary function (XlConsolidationFunction constant).'),
            name: z.string().optional().describe('Custom name for the data field.'),
        })).optional().describe('Data fields.'),
        filterFields: z.array(z.union([z.string(), z.number().int().positive()])).optional().describe('Filter fields.'),
    }).optional().describe('Parameters for pivot table creation. Required for "pivot".'),

    // Specific fields for 'calculate'
    formulaRange: z.string().optional().describe('Address of the range containing formulas to calculate (e.g., "A1:A10"). Required for "calculate".'),
    useComInterop: z.boolean().optional().default(false).describe('Use COM interop for local Excel interaction. Defaults to false (uses exceljs).'),
}).refine(data => {
    // Conditional validations based on the operation
    if (data.operation === 'sort') {
        return data.sortCriteria !== undefined && data.sortCriteria.length > 0;
    } else if (data.operation === 'filter') {
        return data.filterCriteria !== undefined && data.filterCriteria.length > 0;
    } else if (data.operation === 'pivot') {
        return data.pivotTableParameters !== undefined && data.pivotTableParameters.pivotTableName !== undefined && data.pivotTableParameters.destination !== undefined;
    } else if (data.operation === 'calculate') {
        return data.formulaRange !== undefined;
    }
    return true; // If the operation does not require specific fields, validation passes
}, {
    message: "Missing required fields for the specified operation.",
    path: [], // Apply error to the whole object
}).refine(data => {
    // Validate that either rangeAddress or tableName is provided
    return data.rangeAddress !== undefined || data.tableName !== undefined;
}, {
    message: "Either rangeAddress or tableName must be provided.",
    path: ['rangeAddress', 'tableName'], // Apply error to these fields
});


// Infer the combined type for use in the handler
type ExcelDataAnalysisInput = z.infer<typeof ExcelDataAnalysisInputSchema>;


/**
 * @tool excel/data-analysis
 * @description Performs data analysis operations (sort, filter, pivot, calculate) on Excel ranges or tables.
 * @inputSchema See `ExcelDataAnalysisInputSchema` (z.object). Uses combined properties from all operations.
 * @outputSchema Returns a success status with a message in `data`.
 * @dependencies Requires Microsoft Excel installed and accessible via COM Interop.
 * @security Input `filePath` is validated. Ensure Excel COM security settings are appropriate.
 * @errorHandling Uses standard error handling. Catches COM errors, validation errors, file access issues. Returns standardized `ErrorResponse`.
 * @example_sort
 * ```json
 * {
 *   "operation": "sort",
 *   "filePath": "C:/path/to/document.xlsx",
 *   "sheetName": "Sheet1",
 *   "rangeAddress": "A1:D10",
 *   "sortCriteria": [
 *     { "column": "B", "order": "Ascending" },
 *     { "column": 1, "order": "Descending" }
 *   ],
 *   "header": "xlYes"
 * }
 * ```
 * @example_filter
 * ```json
 * {
 *   "operation": "filter",
 *   "filePath": "C:/path/to/document.xlsx",
 *   "sheetIndex": 1,
 *   "tableName": "Table1",
 *   "filterCriteria": [
 *     { "column": "Status", "criteria1": "Completed" },
 *     { "column": "Date", "criteria1": ">=2023-01-01", "operator": 7 } // xlFilterValues
 *   ]
 * }
 * ```
 * @example_pivot
 * ```json
 * {
 *   "operation": "pivot",
 *   "filePath": "C:/path/to/document.xlsx",
 *   "sheetName": "Data",
 *   "rangeAddress": "A1:F100",
 *   "pivotTableParameters": {
 *     "pivotTableName": "SalesSummary",
 *     "destination": "Summary!A1",
 *     "rowFields": ["Region"],
 *     "columnFields": ["Year"],
 *     "dataFields": [
 *       { "field": "Sales", "function": 9, "name": "Total Sales" }, // xlSum
 *       { "field": "Units", "function": 2 } // xlCount
 *     ]
 *   }
 * }
 * ```
 * @example_calculate
 * ```json
 * {
 *   "operation": "calculate",
 *   "filePath": "C:/path/to/document.xlsx",
 *   "sheetName": "Calculations",
   "formulaRange": "A1:A10"
 * }
 * ```
 */
export const excelDataAnalysisTool: McpResource[] = [{
  path: 'excel/data-analysis',
  description: 'Performs data analysis operations (sort, filter, pivot, calculate) on Excel ranges or tables.',
  schema: ExcelDataAnalysisInputSchema, // Use the new z.object schema
  handler: async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<string>> => {
    const input = ExcelDataAnalysisInputSchema.parse(params);
    const { filePath, operation, sheetName, sheetIndex, rangeAddress, tableName, useComInterop, sortCriteria, filterCriteria, pivotTableParameters, formulaRange, header } = input;
    const absoluteFilePath = path.resolve(filePath);
    let resultMessage: string;

    if (useComInterop) {
      logger.info(`Executing excel/data-analysis (COM) operation '${operation}' for file: ${filePath}`);
      let excelApp: any = null;
      let workbook: any = null;
      let sheet: any = null;
      try {
        excelApp = await getOfficeApplication('Excel.Application');
        workbook = excelApp.Workbooks.Open(absoluteFilePath);
        sheet = sheetIndex
          ? workbook.Sheets(sheetIndex)
          : workbook.Sheets(sheetName || 1); // Default to first sheet if name not provided

        let targetRange: any = null;
        if (tableName) {
          try {
            targetRange = sheet.ListObjects(tableName).Range;
          } catch (e: any) {
            throw new Error(`COM: Table "${tableName}" not found: ${e.message}`);
          }
        } else if (rangeAddress) {
          targetRange = sheet.Range(rangeAddress);
        } else {
          throw new Error('COM: rangeAddress or tableName is required.');
        }

        switch (operation) {
          case 'sort': {
            if (!sortCriteria) throw new Error("COM: sortCriteria is required for 'sort' operation.");
            const sort = targetRange.Sort;
            sort.SortFields.Clear();
            sortCriteria.forEach(criteria => {
              const columnRange = targetRange.Columns(criteria.column);
              sort.SortFields.Add(columnRange, excelApp.constants.xlSortOnValues, criteria.order === 'Ascending' ? excelApp.constants.xlAscending : excelApp.constants.xlDescending);
            });
            sort.SetRange(targetRange);
            sort.Header = header !== undefined ? header : excelApp.constants.xlGuess;
            sort.MatchCase = false;
            sort.Orientation = excelApp.constants.xlSortColumns;
            sort.SortMethod = excelApp.constants.xlPinYin;
            sort.Apply();
            resultMessage = `COM: Operation 'sort' applied to "${filePath}".`;
            break;
          }
          case 'filter': {
            if (!filterCriteria) throw new Error("COM: filterCriteria is required for 'filter' operation.");
            targetRange.AutoFilter();
            filterCriteria.forEach(criteria => {
              targetRange.AutoFilter(
                targetRange.Columns(criteria.column).Column,
                criteria.criteria1,
                criteria.operator !== undefined ? criteria.operator : excelApp.constants.xlFilterValues,
                criteria.criteria2,
                criteria.visibleDropDown !== undefined ? criteria.visibleDropDown : true
              );
            });
            resultMessage = `COM: Operation 'filter' applied to "${filePath}".`;
            break;
          }
          case 'pivot': {
            if (!pivotTableParameters) throw new Error("COM: pivotTableParameters is required for 'pivot' operation.");
            const pivotCache = workbook.PivotCaches.Create(excelApp.constants.xlDatabase, targetRange);
            const pivotTable = pivotCache.CreatePivotTable(
              pivotTableParameters.destination,
              pivotTableParameters.pivotTableName
            );
            pivotTableParameters.rowFields?.forEach((field: string | number) => {
              pivotTable.PivotFields(field).Orientation = excelApp.constants.xlRowField;
            });
            pivotTableParameters.columnFields?.forEach((field: string | number) => {
              pivotTable.PivotFields(field).Orientation = excelApp.constants.xlColumnField;
            });
            pivotTableParameters.dataFields?.forEach(dataField => {
              const field = pivotTable.PivotFields(dataField.field);
              pivotTable.AddDataField(field, dataField.name, dataField.function !== undefined ? dataField.function : excelApp.constants.xlSum);
            });
            pivotTableParameters.filterFields?.forEach((field: string | number) => {
              pivotTable.PivotFields(field).Orientation = excelApp.constants.xlPageField;
            });
            resultMessage = `COM: Pivot table "${pivotTableParameters.pivotTableName}" created.`;
            break;
          }
          case 'calculate': {
            if (!formulaRange) throw new Error("COM: formulaRange is required for 'calculate' operation.");
            const calcRange = sheet.Range(formulaRange);
            calcRange.Calculate();
            resultMessage = `COM: Calculation executed on range "${formulaRange}" in "${filePath}".`;
            break;
          }
          default:
            throw new Error(`COM: Unsupported operation: "${operation}".`);
        }
        return { success: true, data: resultMessage };
      } catch (error: any) {
        logger.error(`[excel/data-analysis COM] Error: ${error.message}`, { error: String(error), params });
        return { success: false, error: { code: 'OFFICE_API_COM_ERROR', message: `COM Error: ${error.message}`, details: String(error) } };
      } finally {
        if (workbook) { try { workbook.Save(); workbook.Close(false); } catch(e){ /* ignore */ } }
        releaseObject(sheet);
        releaseObject(workbook);
        releaseObject(excelApp); // Release app if we specifically got it for this op.
      }
    } else {
      // exceljs path
      logger.info(`Executing excel/data-analysis (exceljs) operation '${operation}' for file: ${filePath}`);
      const wb = new ExcelJS.Workbook();
      try {
        const fileExisted = await fs.pathExists(absoluteFilePath);
        if (!fileExisted && operation !== 'pivot') { // Pivot might create a new file if destination is new sheet
             throw new Error(`exceljs: File not found at ${absoluteFilePath}.`);
        }
        if(fileExisted) {
            await wb.xlsx.readFile(absoluteFilePath);
        }

        let ws: ExcelJS.Worksheet | undefined;
        if (sheetName) ws = wb.getWorksheet(sheetName);
        else if (sheetIndex && sheetIndex > 0 && sheetIndex <= wb.worksheets.length) ws = wb.worksheets[sheetIndex - 1];
        else if (wb.worksheets.length > 0) ws = wb.worksheets[0];
        
        if (!ws && operation !== 'pivot') { // Pivot might create its own destination sheet
            throw new Error(`exceljs: Worksheet "${sheetName || sheetIndex || 'default'}" not found.`);
        }


        switch (operation) {
          case 'sort':
            resultMessage = "exceljs: Sorting data within Excel files is not directly supported. For sorting, please use COM Interop or implement manual data extraction, sorting in code, and writing back.";
            logger.warn(resultMessage);
            break;
          case 'filter':
            if (ws && (rangeAddress || tableName)) {
                // exceljs can set an autofilter range, but not apply criteria programmatically easily.
                // If tableName is given, find its range.
                let filterRange = rangeAddress;
                if (tableName) {
                    const table = ws.getTable(tableName);
                    if (table) {
                        filterRange = table.ref; // Get the range from the table
                    } else {
                        throw new Error(`exceljs: Table "${tableName}" not found on sheet "${ws.name}".`);
                    }
                }
                if (filterRange) {
                    ws.autoFilter = filterRange;
                    resultMessage = `exceljs: AutoFilter enabled on range "${filterRange}" in sheet "${ws.name}". Applying specific criteria programmatically is not supported; use COM Interop or manual filtering in Excel.`;
                } else {
                     throw new Error(`exceljs: rangeAddress or tableName (resolving to a range) is required to enable AutoFilter.`);
                }
            } else {
              resultMessage = "exceljs: To enable AutoFilter, a worksheet and either rangeAddress or tableName must be specified. Applying specific criteria is not supported; use COM Interop.";
              logger.warn(resultMessage);
            }
            break;
          case 'pivot':
            resultMessage = "exceljs: Creating or manipulating Pivot Tables is not supported. Please use COM Interop for Pivot Table operations.";
            logger.warn(resultMessage);
            break;
          case 'calculate':
            // exceljs calculates formulas on load/save.
            // If we just load and save, it might trigger calculations.
            if(fileExisted) await wb.xlsx.writeFile(absoluteFilePath); // Save to trigger calculations
            resultMessage = `exceljs: Formulas are generally calculated upon file load/save. Workbook at "${filePath}" processed. Specific range calculation is not applicable.`;
            break;
          default:
            throw new Error(`exceljs: Unsupported operation: "${operation}".`);
        }
        if (operation === 'filter' || operation === 'calculate' && fileExisted) { // Only save if an action was taken
             await wb.xlsx.writeFile(absoluteFilePath);
        }
        return { success: true, data: resultMessage };
      } catch (error: any) {
        logger.error(`[excel/data-analysis exceljs] Error: ${error.message}`, { error: String(error), params });
        return { success: false, error: { code: 'EXCELJS_DATA_ANALYSIS_ERROR', message: `exceljs Error: ${error.message}`, details: String(error) } };
      }
    }
  },
}];