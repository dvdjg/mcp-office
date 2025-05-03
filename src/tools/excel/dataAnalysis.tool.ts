import { z, ZodError } from 'zod';
import { McpResource, ApiResponse, SuccessResponse, ErrorResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Importar releaseObject
import logger from '../../utils/logger'; // Importar logger

// --- Constantes COM (Ejemplos, verificar en documentación de Excel) ---
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

// Esquema combinado para todas las operaciones (usando z.object)
const ExcelDataAnalysisInputSchema = z.object({
    operation: z.enum(['sort', 'filter', 'pivot', 'calculate']).describe('The operation to perform (sort, filter, pivot, or calculate).'),
    // Incluir todos los campos posibles de las operaciones
    filePath: z.string().describe('Ruta al archivo de Excel.'),
    sheetName: z.string().optional().describe('Nombre de la hoja de cálculo. Si no se proporciona, se usa la hoja activa.'),
    sheetIndex: z.number().int().positive().optional().describe('Índice de la hoja de cálculo (1-basado). Si se proporciona, anula sheetName.'),
    rangeAddress: z.string().optional().describe('Dirección del rango (p. ej., "A1:D10"). Requerido si no se usa tableName.'),
    tableName: z.string().optional().describe('Nombre de la tabla. Requerido si no se usa rangeAddress.'),

    // Campos específicos para 'sort'
    sortCriteria: z.array(z.object({
        column: z.union([z.string(), z.number().int().positive()]).describe('Columna por la que ordenar (nombre de columna o índice 1-basado).'),
        order: z.enum(['Ascending', 'Descending']).default('Ascending').describe('Orden de ordenación.'),
    })).min(1).optional().describe('Criterios de ordenación. Requerido para "sort".'),
    header: z.nativeEnum(XlHeaderShow).default(XlHeaderShow.xlGuess).optional().describe('Especifica si el rango tiene encabezados. Usado en "sort".'),

    // Campos específicos para 'filter'
    filterCriteria: z.array(z.object({
        column: z.union([z.string(), z.number().int().positive()]).describe('Columna por la que filtrar (nombre de columna o índice 1-basado).'),
        criteria1: z.any().describe('Primer criterio de filtro.'),
        operator: z.nativeEnum(XlAutoFilterOperator).optional().describe('Operador de filtro (constante XlAutoFilterOperator).'),
        criteria2: z.any().optional().describe('Segundo criterio de filtro (para operadores que requieren dos).'),
        visibleDropDown: z.boolean().default(true).optional().describe('Mostrar u ocultar el botón desplegable de autofiltro.'),
    })).min(1).optional().describe('Criterios de filtro. Requerido para "filter".'),

    // Campos específicos para 'pivot'
    pivotTableParameters: z.object({
        pivotTableName: z.string().describe('Nombre para la nueva tabla dinámica.'),
        destination: z.string().describe('Celda donde se colocará la tabla dinámica (p. ej., "Sheet2!A1").'),
        rowFields: z.array(z.union([z.string(), z.number().int().positive()])).optional().describe('Campos de fila.'),
        columnFields: z.array(z.union([z.string(), z.number().int().positive()])).optional().describe('Campos de columna.'),
        dataFields: z.array(z.object({
            field: z.union([z.string(), z.number().int().positive()]).describe('Campo de datos.'),
            function: z.nativeEnum(XlConsolidationFunction).optional().describe('Función de resumen (constante XlConsolidationFunction).'),
            name: z.string().optional().describe('Nombre personalizado para el campo de datos.'),
        })).optional().describe('Campos de datos.'),
        filterFields: z.array(z.union([z.string(), z.number().int().positive()])).optional().describe('Campos de filtro.'),
    }).optional().describe('Parámetros para la creación de la tabla dinámica. Requerido para "pivot".'),

    // Campos específicos para 'calculate'
    formulaRange: z.string().optional().describe('Dirección del rango que contiene las fórmulas a calcular (p. ej., "A1:A10"). Requerido para "calculate".'),

}).refine(data => {
    // Validaciones condicionales basadas en la operación
    if (data.operation === 'sort') {
        return data.sortCriteria !== undefined && data.sortCriteria.length > 0;
    } else if (data.operation === 'filter') {
        return data.filterCriteria !== undefined && data.filterCriteria.length > 0;
    } else if (data.operation === 'pivot') {
        return data.pivotTableParameters !== undefined && data.pivotTableParameters.pivotTableName !== undefined && data.pivotTableParameters.destination !== undefined;
    } else if (data.operation === 'calculate') {
        return data.formulaRange !== undefined;
    }
    return true; // Si la operación no requiere campos específicos, pasa la validación
}, {
    message: "Missing required fields for the specified operation.",
    path: [], // Apply error to the whole object
}).refine(data => {
    // Validar que se proporcione rangeAddress o tableName
    return data.rangeAddress !== undefined || data.tableName !== undefined;
}, {
    message: "Either rangeAddress or tableName must be provided.",
    path: ['rangeAddress', 'tableName'], // Apply error to these fields
});


// Inferir el tipo combinado para usar en el handler
type ExcelDataAnalysisInput = z.infer<typeof ExcelDataAnalysisInputSchema>;


/**
 * @tool excel/data-analysis
 * @description Realiza operaciones de análisis de datos (ordenar, filtrar, tabla dinámica, calcular) en rangos o tablas de Excel.
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
  description: 'Performs data analysis operations (sort, filter, pivot, calculate) on Excel ranges or tables.', // Translated description
  schema: ExcelDataAnalysisInputSchema, // Usar el nuevo esquema z.object
  handler: async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<string>> => {
    let excelApp: any = null;
    let workbook: any = null;
    let sheet: any = null;

    try {
      // Validar input parameters using the combined schema
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
          throw new Error(`Tabla "${input.tableName}" no encontrada: ${e.message}`);
        }
      } else if (input.rangeAddress) {
        targetRange = sheet.Range(input.rangeAddress);
      } else {
        // This case should be caught by the schema refine, but as a fallback:
        throw new Error('Se requiere rangeAddress o tableName.');
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
          resultMessage = `Operación 'sort' aplicada al rango/tabla en "${input.filePath}".`;
          break;
        }
        case 'filter': {
          const filterInput = input; // input is already validated as ExcelDataAnalysisInput
          if (!filterInput.filterCriteria) throw new Error("filterCriteria is required for 'filter' operation.");

          targetRange.AutoFilter(); // Asegura que AutoFilter esté activado
          filterInput.filterCriteria.forEach(criteria => {
            targetRange.AutoFilter(
              targetRange.Columns(criteria.column).Column, // Field
              criteria.criteria1,
              criteria.operator !== undefined ? criteria.operator : excelApp.constants.xlFilterValues, // Default operator
              criteria.criteria2,
              criteria.visibleDropDown !== undefined ? criteria.visibleDropDown : true // Default visibleDropDown
            );
          });
          resultMessage = `Operación 'filter' aplicada al rango/tabla en "${input.filePath}".`;
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

          // Añadir campos de fila
          pivotInput.pivotTableParameters.rowFields?.forEach((field: string | number) => { // Añadir tipo para field
            pivotTable.PivotFields(field).Orientation = excelApp.constants.xlRowField;
          });

          // Añadir campos de columna
          pivotInput.pivotTableParameters.columnFields?.forEach((field: string | number) => { // Añadir tipo para field
            pivotTable.PivotFields(field).Orientation = excelApp.constants.xlColumnField;
          });

          // Añadir campos de datos
          pivotInput.pivotTableParameters.dataFields?.forEach(dataField => {
            const field = pivotTable.PivotFields(dataField.field);
            const dataFieldItem = pivotTable.AddDataField(field, dataField.name, dataField.function !== undefined ? dataField.function : excelApp.constants.xlSum); // Default function
          });

          // Añadir campos de filtro
          pivotInput.pivotTableParameters.filterFields?.forEach((field: string | number) => { // Añadir tipo para field
            pivotTable.PivotFields(field).Orientation = excelApp.constants.xlPageField; // Filter
          });

          resultMessage = `Operación 'pivot' completada. Tabla dinámica "${pivotInput.pivotTableParameters.pivotTableName}" creada en "${pivotInput.pivotTableParameters.destination}".`;
          break;
        }
        case 'calculate': {
          const calcInput = input; // input is already validated as ExcelDataAnalysisInput
          if (!calcInput.formulaRange) throw new Error("formulaRange is required for 'calculate' operation.");

          const formulaRange = sheet.Range(calcInput.formulaRange);
          formulaRange.Calculate();
          // Nota: Esta operación no devuelve los resultados del cálculo, solo los ejecuta.
          resultMessage = `Operación 'calculate' ejecutada en el rango "${calcInput.formulaRange}" en "${input.filePath}".`;
          break;
        }
        default:
          // This case should theoretically not be reached due to discriminated union,
          // but adding a type assertion for safety and to satisfy TypeScript.
          // const exhaustiveCheck: never = input; // No longer needed with combined schema
          throw new Error(`Operación "${(input as any).operation}" no soportada.`); // Acceder a operation directamente

      }

      // Return success response
      return {
        success: true,
        data: resultMessage,
      };

    } catch (error: any) {
      // Handle Zod validation errors specifically
      if (error instanceof ZodError) {
        // Serializar error.errors para que sea serializable
        const errorDetails = JSON.stringify(error.errors, null, 2);
        logger.warn(`[excel/data-analysis] Input validation failed: ${error.message}`, { errors: errorDetails, params });
        return {
          success: false,
          error: {
            code: 'VALIDATION_ERROR',
            message: 'Error de validación de entrada.',
            details: errorDetails,
          },
        };
      }
      // Handle other errors
      logger.error(`[excel/data-analysis] Error al ejecutar la herramienta: ${error.message}`, { error: String(error), params }); // Serializar error
      return {
        success: false,
        error: {
          code: 'OFFICE_API_ERROR', // Or a more specific code if possible
          message: `Error al ejecutar la herramienta excel/data-analysis: ${error.message}`,
          details: String(error), // Serializar error
        },
      };
    } finally {
      // No cerrar Excel aquí, la aplicación debe permanecer abierta para futuras operaciones.
      // Liberar workbook y sheet si se obtuvieron
      if (workbook) workbook.Close(false); // Cerrar sin guardar
      releaseObject(sheet);
      releaseObject(workbook);
      // No liberar excelApp aquí
    }
  },
}];