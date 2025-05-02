import { z, ZodError } from 'zod';
import { McpResource, ApiResponse, SuccessResponse, ErrorResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types';
import { getOfficeApplication } from '../../utils/officeInterop';

// Esquema base para la entrada de la herramienta
const DataAnalysisInputSchema = z.object({
  filePath: z.string().describe('Ruta al archivo de Excel.'),
  sheetName: z.string().optional().describe('Nombre de la hoja de cálculo. Si no se proporciona, se usa la hoja activa.'),
  sheetIndex: z.number().int().positive().optional().describe('Índice de la hoja de cálculo (1-basado). Si se proporciona, anula sheetName.'),
  rangeAddress: z.string().optional().describe('Dirección del rango (p. ej., "A1:D10"). Requerido si no se usa tableName.'),
  tableName: z.string().optional().describe('Nombre de la tabla. Requerido si no se usa rangeAddress.'),
  operation: z.enum(['sort', 'filter', 'pivot', 'calculate']).describe('Operación a realizar.'),
});

// Esquema específico para la operación 'sort'
const SortOperationSchema = DataAnalysisInputSchema.extend({
  operation: z.literal('sort'),
  sortCriteria: z.array(z.object({
    column: z.union([z.string(), z.number().int().positive()]).describe('Columna por la que ordenar (nombre de columna o índice 1-basado).'),
    order: z.enum(['Ascending', 'Descending']).default('Ascending').describe('Orden de ordenación.'),
  })).min(1).describe('Criterios de ordenación.'),
  header: z.enum(['xlYes', 'xlNo', 'xlGuess']).default('xlGuess').describe('Especifica si el rango tiene encabezados.'),
});

// Esquema específico para la operación 'filter'
const FilterOperationSchema = DataAnalysisInputSchema.extend({
  operation: z.literal('filter'),
  filterCriteria: z.array(z.object({
    column: z.union([z.string(), z.number().int().positive()]).describe('Columna por la que filtrar (nombre de columna o índice 1-basado).'),
    criteria1: z.any().describe('Primer criterio de filtro.'),
    operator: z.number().int().optional().describe('Operador de filtro (constante XlAutoFilterOperator).'),
    criteria2: z.any().optional().describe('Segundo criterio de filtro (para operadores que requieren dos).'),
    visibleDropDown: z.boolean().default(true).optional().describe('Mostrar u ocultar el botón desplegable de autofiltro.'),
  })).min(1).describe('Criterios de filtro.'),
});

// Esquema específico para la operación 'pivot'
const PivotOperationSchema = DataAnalysisInputSchema.extend({
  operation: z.literal('pivot'),
  pivotTableParameters: z.object({
    pivotTableName: z.string().describe('Nombre para la nueva tabla dinámica.'),
    destination: z.string().describe('Celda donde se colocará la tabla dinámica (p. ej., "Sheet2!A1").'),
    rowFields: z.array(z.union([z.string(), z.number().int().positive()])).optional().describe('Campos de fila.'),
    columnFields: z.array(z.union([z.string(), z.number().int().positive()])).optional().describe('Campos de columna.'),
    dataFields: z.array(z.object({
      field: z.union([z.string(), z.number().int().positive()]).describe('Campo de datos.'),
      function: z.number().int().optional().describe('Función de resumen (constante XlConsolidationFunction).'),
      name: z.string().optional().describe('Nombre personalizado para el campo de datos.'),
    })).optional().describe('Campos de datos.'),
    filterFields: z.array(z.union([z.string(), z.number().int().positive()])).optional().describe('Campos de filtro.'),
  }).describe('Parámetros para la creación de la tabla dinámica.'),
});

// Esquema específico para la operación 'calculate'
const CalculateOperationSchema = DataAnalysisInputSchema.extend({
  operation: z.literal('calculate'),
  formulaRange: z.string().describe('Dirección del rango que contiene las fórmulas a calcular (p. ej., "A1:A10").'),
});

// Esquema de entrada combinado
const CombinedDataAnalysisInputSchema = z.discriminatedUnion('operation', [
  SortOperationSchema,
  FilterOperationSchema,
  PivotOperationSchema,
  CalculateOperationSchema,
]);

type DataAnalysisInput = z.infer<typeof CombinedDataAnalysisInputSchema>;

/**
 * @tool excel/data-analysis
 * @description Realiza operaciones de análisis de datos (ordenar, filtrar, tabla dinámica, calcular) en rangos o tablas de Excel.
 * @param {object} input - Parámetros de entrada.
 * @param {string} input.filePath - Ruta al archivo de Excel.
 * @param {string} [input.sheetName] - Nombre de la hoja de cálculo.
 * @param {number} [input.sheetIndex] - Índice de la hoja de cálculo (1-basado).
 * @param {string} [input.rangeAddress] - Dirección del rango (p. ej., "A1:D10").
 * @param {string} [input.tableName] - Nombre de la tabla.
 * @param {'sort'|'filter'|'pivot'|'calculate'} input.operation - Operación a realizar.
 * @param {object[]} [input.sortCriteria] - Criterios de ordenación para la operación 'sort'.
 * @param {object[]} [input.filterCriteria] - Criterios de filtro para la operación 'filter'.
 * @param {object} [input.pivotTableParameters] - Parámetros para la operación 'pivot'.
 * @param {string} [input.formulaRange] - Rango con fórmulas para la operación 'calculate'.
 * @returns {Promise<string>} - Un mensaje indicando el resultado de la operación.
 */

/**
 * @tool excel/data-analysis
 * @description Realiza operaciones de análisis de datos (ordenar, filtrar, tabla dinámica, calcular) en rangos o tablas de Excel.
 * @param {object} input - Parámetros de entrada.
 * @param {string} input.filePath - Ruta al archivo de Excel.
 * @param {string} [input.sheetName] - Nombre de la hoja de cálculo.
 * @param {number} [input.sheetIndex] - Índice de la hoja de cálculo (1-basado).
 * @param {string} [input.rangeAddress] - Dirección del rango (p. ej., "A1:D10").
 * @param {string} [input.tableName] - Nombre de la tabla.
 * @param {'sort'|'filter'|'pivot'|'calculate'} input.operation - Operación a realizar.
 * @param {object[]} [input.sortCriteria] - Criterios de ordenación para la operación 'sort'.
 * @param {object[]} [input.filterCriteria] - Criterios de filtro para la operación 'filter'.
 * @param {object} [input.pivotTableParameters] - Parámetros para la operación 'pivot'.
 * @param {string} [input.formulaRange] - Rango con fórmulas para la operación 'calculate'.
 * @returns {Promise<string>} - Un mensaje indicando el resultado de la operación.
 */
const handler = async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<string>> => {
  let excelApp;
  try {
    // Validate input parameters
    const input = CombinedDataAnalysisInputSchema.parse(params);

    excelApp = await getOfficeApplication('Excel.Application');
    const workbook = excelApp.Workbooks.Open(input.filePath);
    const sheet = input.sheetIndex
      ? workbook.Sheets(input.sheetIndex)
      : workbook.Sheets(input.sheetName || 1);

    let targetRange;
    if (input.tableName) {
      try {
        targetRange = sheet.ListObjects(input.tableName).Range;
      } catch (e) {
        throw new Error(`Tabla "${input.tableName}" no encontrada.`);
      }
    } else if (input.rangeAddress) {
      targetRange = sheet.Range(input.rangeAddress);
    } else {
      throw new Error('Se requiere rangeAddress o tableName.');
    }

    let resultMessage: string;

    switch (input.operation) {
      case 'sort': {
        const sortInput = input as z.infer<typeof SortOperationSchema>;
        const sort = targetRange.Sort;
        sort.SortFields.Clear();
        sortInput.sortCriteria.forEach(criteria => {
          const columnRange = targetRange.Columns(criteria.column);
          sort.SortFields.Add(columnRange, 1, criteria.order === 'Ascending' ? 1 : 2); // 1=xlSortOnValues, 1=xlAscending, 2=xlDescending
        });
        sort.SetRange(targetRange);
        sort.Header = sortInput.header === 'xlYes' ? 1 : sortInput.header === 'xlNo' ? 2 : 3; // 1=xlYes, 2=xlNo, 3=xlGuess
        sort.MatchCase = false;
        sort.Orientation = 2; // 2=xlSortColumns
        sort.SortMethod = 1; // 1=xlPinYin
        sort.Apply();
        resultMessage = `Operación 'sort' aplicada al rango/tabla en "${input.filePath}".`;
        break;
      }
      case 'filter': {
        const filterInput = input as z.infer<typeof FilterOperationSchema>;
        targetRange.AutoFilter(); // Asegura que AutoFilter esté activado
        filterInput.filterCriteria.forEach(criteria => {
          targetRange.AutoFilter(
            targetRange.Columns(criteria.column).Column, // Field
            criteria.criteria1,
            criteria.operator,
            criteria.criteria2,
            criteria.visibleDropDown
          );
        });
        resultMessage = `Operación 'filter' aplicada al rango/tabla en "${input.filePath}".`;
        break;
      }
      case 'pivot': {
        const pivotInput = input as z.infer<typeof PivotOperationSchema>;
        const pivotCache = workbook.PivotCaches.Create(1, targetRange); // 1=xlDatabase
        const pivotTable = pivotCache.CreatePivotTable(
          pivotInput.pivotTableParameters.destination,
          pivotInput.pivotTableParameters.pivotTableName
        );

        // Añadir campos de fila
        pivotInput.pivotTableParameters.rowFields?.forEach(field => {
          pivotTable.PivotFields(field).Orientation = 1; // 1=xlRowField
        });

        // Añadir campos de columna
        pivotInput.pivotTableParameters.columnFields?.forEach(field => {
          pivotTable.PivotFields(field).Orientation = 2; // 2=xlColumnField
        });

        // Añadir campos de datos
        pivotInput.pivotTableParameters.dataFields?.forEach(dataField => {
          const field = pivotTable.PivotFields(dataField.field);
          const dataFieldItem = pivotTable.AddDataField(field, dataField.name, dataField.function);
        });

        // Añadir campos de filtro
        pivotInput.pivotTableParameters.filterFields?.forEach(field => {
          pivotTable.PivotFields(field).Orientation = 3; // 3=xlPageField (Filter)
        });

        resultMessage = `Operación 'pivot' completada. Tabla dinámica "${pivotInput.pivotTableParameters.pivotTableName}" creada en "${pivotInput.pivotTableParameters.destination}".`;
        break;
      }
      case 'calculate': {
        const calcInput = input as z.infer<typeof CalculateOperationSchema>;
        const formulaRange = sheet.Range(calcInput.formulaRange);
        formulaRange.Calculate();
        // Nota: Esta operación no devuelve los resultados del cálculo, solo los ejecuta.
        resultMessage = `Operación 'calculate' ejecutada en el rango "${calcInput.formulaRange}" en "${input.filePath}".`;
        break;
      }
      default:
        // This case should theoretically not be reached due to discriminated union,
        // but adding a type assertion for safety and to satisfy TypeScript.
        const exhaustiveCheck: never = input;
        throw new Error(`Operación "${(exhaustiveCheck as any).operation}" no soportada.`);
    }

    // Return success response
    return {
      success: true,
      data: resultMessage,
    };

  } catch (error: any) {
    // Handle Zod validation errors specifically
    if (error instanceof ZodError) {
      return {
        success: false,
        error: {
          code: 'VALIDATION_ERROR',
          message: 'Error de validación de entrada.',
          details: error.errors,
        },
      };
    }
    // Handle other errors
    return {
      success: false,
      error: {
        code: 'OFFICE_API_ERROR', // Or a more specific code if possible
        message: `Error al ejecutar la herramienta excel/data-analysis: ${error.message}`,
        details: error,
      },
    };
  } finally {
    // No cerrar Excel aquí, la aplicación debe permanecer abierta para futuras operaciones.
  }
};

export const excelDataAnalysisTool: McpResource[] = [{
  path: 'excel/data-analysis',
  description: 'Performs data analysis operations (sort, filter, pivot, calculate) on Excel ranges or tables.', // Translated description
  schema: CombinedDataAnalysisInputSchema,
  handler,
}];