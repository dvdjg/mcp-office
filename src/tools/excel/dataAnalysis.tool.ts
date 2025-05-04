/**
 * @file Tool for performing data analysis operations (sort, filter, pivot, calculate) on Excel ranges or tables.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z, ZodError } from 'zod';
import { McpResource, ApiResponse, SuccessResponse, ErrorResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Import releaseObject
import logger from '../../utils/logger'; // Import logger

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
    let excelApp: any = null;
    let workbook: any = null;
    let sheet: any = null;

    try {
      // Validate input parameters using the combined schema
      const input = ExcelDataAnalysisInputSchema.parse(params);

      excelApp = await getOfficeApplication('Excel.Application');
      workbook = excelApp.Workbooks.Open(input.filePath);
      sheet = input.sheetIndex
        ? workbook.Sheets(input.sheetIndex)
        : workbook.Sheets(input.sheetName || 1);

      let targetRange: any = null;
      if (input.tableName) {
        try {
          targetRange = sheet.ListObjects(input.tableName).Range;
        } catch (e: any) {
          throw new Error(`Table "${input.tableName}" not found: ${e.message}`);
        }
      } else if (input.rangeAddress) {
        targetRange = sheet.Range(input.rangeAddress);
      } else {
        // This case should be caught by the schema refine, but as a fallback:
        throw new Error('rangeAddress or tableName is required.');
      }

      let resultMessage: string;

      switch (input.operation) {
        case 'sort': {
          const sortInput = input; // input is already validated as ExcelDataAnalysisInput
          if (!sortInput.sortCriteria) throw new Error("sortCriteria is required for 'sort' operation.");

          const sort = targetRange.Sort;
          sort.SortFields.Clear();
          sortInput.sortCriteria.forEach(criteria => {
            const columnRange = targetRange.Columns(criteria.column);
            sort.SortFields.Add(columnRange, excelApp.constants.xlSortOnValues, criteria.order === 'Ascending' ? excelApp.constants.xlAscending : excelApp.constants.xlDescending);
          });
          sort.SetRange(targetRange);
          sort.Header = input.header !== undefined ? input.header : excelApp.constants.xlGuess;
          sort.MatchCase = false;
          sort.Orientation = excelApp.constants.xlSortColumns;
          sort.SortMethod = excelApp.constants.xlPinYin; // Or xlStroke
          sort.Apply();
          resultMessage = `Operation 'sort' applied to the range/table in "${input.filePath}".`;
          break;
        }
        case 'filter': {
          const filterInput = input; // input is already validated as ExcelDataAnalysisInput
          if (!filterInput.filterCriteria) throw new Error("filterCriteria is required for 'filter' operation.");

          targetRange.AutoFilter(); // Ensure AutoFilter is activated
          filterInput.filterCriteria.forEach(criteria => {
            targetRange.AutoFilter(
              targetRange.Columns(criteria.column).Column, // Field
              criteria.criteria1,
              criteria.operator !== undefined ? criteria.operator : excelApp.constants.xlFilterValues, // Default operator
              criteria.criteria2,
              criteria.visibleDropDown !== undefined ? criteria.visibleDropDown : true // Default visibleDropDown
            );
          });
          resultMessage = `Operation 'filter' applied to the range/table in "${input.filePath}".`;
          break;
        }
        case 'pivot': {
          const pivotInput = input; // input is already validated as ExcelDataAnalysisInput
          if (!pivotInput.pivotTableParameters) throw new Error("pivotTableParameters is required for 'pivot' operation.");

          const pivotCache = workbook.PivotCaches.Create(excelApp.constants.xlDatabase, targetRange);
          const pivotTable = pivotCache.CreatePivotTable(
            pivotInput.pivotTableParameters.destination,
            pivotInput.pivotTableParameters.pivotTableName
          );

          // Add row fields
          pivotInput.pivotTableParameters.rowFields?.forEach((field: string | number) => { // Add type for field
            pivotTable.PivotFields(field).Orientation = excelApp.constants.xlRowField;
          });

          // Add column fields
          pivotInput.pivotTableParameters.columnFields?.forEach((field: string | number) => { // Add type for field
            pivotTable.PivotFields(field).Orientation = excelApp.constants.xlColumnField;
          });

          // Add data fields
          pivotInput.pivotTableParameters.dataFields?.forEach(dataField => {
            const field = pivotTable.PivotFields(dataField.field);
            const dataFieldItem = pivotTable.AddDataField(field, dataField.name, dataField.function !== undefined ? dataField.function : excelApp.constants.xlSum); // Default function
          });

          // Add filter fields
          pivotInput.pivotTableParameters.filterFields?.forEach((field: string | number) => { // Add type for field
            pivotTable.PivotFields(field).Orientation = excelApp.constants.xlPageField; // Filter
          });

          resultMessage = `Operation 'pivot' completed. Pivot table "${pivotInput.pivotTableParameters.pivotTableName}" created at "${pivotInput.pivotTableParameters.destination}".`;
          break;
        }
        case 'calculate': {
          const calcInput = input; // input is already validated as ExcelDataAnalysisInput
          if (!calcInput.formulaRange) throw new Error("formulaRange is required for 'calculate' operation.");

          const formulaRange = sheet.Range(calcInput.formulaRange);
          formulaRange.Calculate();
          // Note: This operation does not return the calculation results, it only executes them.
          resultMessage = `Operation 'calculate' executed on range "${calcInput.formulaRange}" in "${input.filePath}".`;
          break;
        }
        default:
          // This case should theoretically not be reached due to discriminated union,
          // but adding a type assertion for safety and to satisfy TypeScript.
          // const exhaustiveCheck: never = input; // No longer needed with combined schema
          throw new Error(`Unsupported operation: "${(input as any).operation}".`); // Access operation directly

      }

      // Return success response
      return {
        success: true,
        data: resultMessage,
      };

    } catch (error: any) {
      // Handle Zod validation errors specifically
      if (error instanceof ZodError) {
        // Serialize error.errors to be serializable
        const errorDetails = JSON.stringify(error.errors, null, 2);
        logger.warn(`[excel/data-analysis] Input validation failed: ${error.message}`, { errors: errorDetails, params });
        return {
          success: false,
          error: {
            code: 'VALIDATION_ERROR',
            message: 'Input validation error.',
            details: errorDetails,
          },
        };
      }
      // Handle other errors
      logger.error(`[excel/data-analysis] Error executing tool: ${error.message}`, { error: String(error), params }); // Serialize error
      return {
        success: false,
        error: {
          code: 'OFFICE_API_ERROR', // Or a more specific code if possible
          message: `Error executing excel/data-analysis tool: ${error.message}`,
          details: String(error), // Serialize error
        },
      };
    } finally {
      // Do not close Excel here, the application should remain open for future operations.
      // Release workbook and sheet if obtained
      if (workbook) workbook.Close(false); // Close without saving
      releaseObject(sheet);
      releaseObject(workbook);
      // Do not release excelApp here
    }
    // Add a return at the end to cover all possible cases
    // This will only be reached if no error was thrown or returned before.
    // In an ideal scenario, all switch cases should return.
    // But to satisfy the linter, we add this fallback return.
    return { success: false, error: { code: 'UNHANDLED_CASE', message: 'Excel data analysis operation did not return an explicit result.' } };
  },
}];