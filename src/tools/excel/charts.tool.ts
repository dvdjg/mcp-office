import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import logger from '../../utils/logger'; // Importar logger

// Definir el esquema de entrada para la herramienta excel/charts
const ExcelChartsInputSchema = z.object({
  filePath: z.string().describe('La ruta al archivo de Excel.'),
  operation: z.enum(['insert', 'modify', 'delete', 'reposition', 'list']).describe('La operación a realizar.'),
  sheetName: z.string().optional().describe('El nombre de la hoja de cálculo. Si no se proporciona, se usa la hoja activa.'),
  sheetIndex: z.number().int().positive().optional().describe('El índice de la hoja de cálculo (1-basado). Si no se proporciona, se usa la hoja activa.'),
  rangeAddress: z.string().optional().describe('La dirección del rango de datos para el gráfico (e.g., "A1:B10"). Requerido para la operación "insert".'),
  chartType: z.string().optional().describe('El tipo de gráfico (e.g., "xlColumnClustered", "xlLine"). Requerido para la operación "insert".'),
  chartTitle: z.string().optional().describe('El título del gráfico.'),
  chartIndex: z.number().int().positive().optional().describe('El índice del objeto gráfico en la hoja (1-basado). Requerido para "modify", "delete", "reposition".'),
  chartName: z.string().optional().describe('El nombre del objeto gráfico en la hoja. Puede usarse en lugar de chartIndex para "modify", "delete", "reposition".'),
  position: z.object({
    left: z.number().optional(),
    top: z.number().optional(),
    width: z.number().optional(),
    height: z.number().optional(),
  }).optional().describe('La posición y tamaño del gráfico. Requerido para "reposition".'),
  // Propiedades adicionales para modify (ejemplo)
  newRangeAddress: z.string().optional().describe('Nueva dirección del rango de datos para la operación "modify".'),
});

type ExcelChartsInput = z.infer<typeof ExcelChartsInputSchema>;

/**
 * @tool excel/charts
 * @description Herramienta para gestionar gráficos en archivos de Excel.
 * Permite insertar, modificar, eliminar y reposicionar gráficos.
 * Utiliza COM Interop a través de winax.
 * @input ExcelChartsInputSchema
 * @output object - Depende de la operación.
 */
const excelChartsTool: McpResource = {
  path: 'excel/charts',
  description: 'Gestiona gráficos en archivos de Excel.',
  schema: ExcelChartsInputSchema, // Corregido de inputSchema a schema
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => {
    let excelApp: any = null;
    let workbook: any = null;
    let worksheet: any = null;

    try {
      const input = ExcelChartsInputSchema.parse(params);

      excelApp = await getOfficeApplication('Excel.Application');
      workbook = excelApp.Workbooks.Open(input.filePath);

      if (input.sheetName) {
        worksheet = workbook.Sheets(input.sheetName);
      } else if (input.sheetIndex) {
        worksheet = workbook.Sheets(input.sheetIndex);
      } else {
        worksheet = workbook.ActiveSheet;
      }

      if (!worksheet) {
        throw new Error(`Hoja de cálculo "${input.sheetName || input.sheetIndex}" no encontrada.`);
      }

      const chartObjects = worksheet.ChartObjects();

      switch (input.operation) {
        case 'insert': {
          if (!input.rangeAddress || !input.chartType) {
            throw new Error('Para insertar un gráfico, se requieren rangeAddress y chartType.');
          }
          const range = worksheet.Range(input.rangeAddress);
          if (!range) {
            throw new Error(`Rango de datos "${input.rangeAddress}" no válido.`);
          }

          // winax puede requerir el valor numérico del tipo de gráfico
          // Aquí usamos un enfoque simple, se podría mapear strings a constantes COM
          const chartType = excelApp.constants[input.chartType] || parseInt(input.chartType, 10);
          if (isNaN(chartType)) {
             throw new Error(`Tipo de gráfico "${input.chartType}" no reconocido.`);
          }

          const chartObject = chartObjects.Add(0, 0, 300, 200); // Posición y tamaño inicial
          const chart = chartObject.Chart;
          chart.SetSourceData(range);
          chart.ChartType = chartType;

          if (input.chartTitle) {
            chart.HasTitle = true;
            chart.ChartTitle.Text = input.chartTitle;
          }

          // Reposicionar si se especifica
          if (input.position) {
            if (input.position.left !== undefined) chartObject.Left = input.position.left;
            if (input.position.top !== undefined) chartObject.Top = input.position.top;
            if (input.position.width !== undefined) chartObject.Width = input.position.width;
            if (input.position.height !== undefined) chartObject.Height = input.position.height;
          }

          return { success: true, data: 'Gráfico insertado correctamente.' }; // Retorno ajustado
        }

        case 'modify': {
          if (!input.chartIndex && !input.chartName) {
            throw new Error('Para modificar un gráfico, se requiere chartIndex o chartName.');
          }
          const chartObject = input.chartIndex ? chartObjects.Item(input.chartIndex) : chartObjects.Item(input.chartName);
          if (!chartObject) {
            throw new Error(`Gráfico con índice ${input.chartIndex} o nombre "${input.chartName}" no encontrado.`);
          }
          const chart = chartObject.Chart;

          if (input.newRangeAddress) {
             const newRange = worksheet.Range(input.newRangeAddress);
             if (!newRange) {
                throw new Error(`Nuevo rango de datos "${input.newRangeAddress}" no válido.`);
             }
             chart.SetSourceData(newRange);
          }

          if (input.chartType) {
             const chartType = excelApp.constants[input.chartType] || parseInt(input.chartType, 10);
             if (isNaN(chartType)) {
                throw new Error(`Tipo de gráfico "${input.chartType}" no reconocido.`);
             }
             chart.ChartType = chartType;
          }

          if (input.chartTitle) {
            chart.HasTitle = true;
            chart.ChartTitle.Text = input.chartTitle;
          } else if (input.chartTitle === '') { // Permitir eliminar el título
             chart.HasTitle = false;
          }

          // Modificar tamaño si se especifica
          if (input.position) {
            if (input.position.width !== undefined) chartObject.Width = input.position.width;
            if (input.position.height !== undefined) chartObject.Height = input.position.height;
          }


          return { success: true, data: 'Gráfico modificado correctamente.' }; // Retorno ajustado
        }

        case 'delete': {
          if (!input.chartIndex && !input.chartName) {
            throw new Error('Para eliminar un gráfico, se requiere chartIndex o chartName.');
          }
          const chartObject = input.chartIndex ? chartObjects.Item(input.chartIndex) : chartObjects.Item(input.chartName);
          if (!chartObject) {
            throw new Error(`Gráfico con índice ${input.chartIndex} o nombre "${input.chartName}" no encontrado.`);
          }
          chartObject.Delete();
          return { success: true, data: 'Gráfico eliminado correctamente.' }; // Retorno ajustado
        }

        case 'reposition': {
          if (!input.chartIndex && !input.chartName) {
            throw new Error('Para reposicionar un gráfico, se requiere chartIndex o chartName.');
          }
           if (!input.position) {
             throw new Error('Para reposicionar un gráfico, se requiere la propiedad position.');
           }
          const chartObject = input.chartIndex ? chartObjects.Item(input.chartIndex) : chartObjects.Item(input.chartName);
          if (!chartObject) {
            throw new Error(`Gráfico con índice ${input.chartIndex} o nombre "${input.chartName}" no encontrado.`);
          }

          if (input.position.left !== undefined) chartObject.Left = input.position.left;
          if (input.position.top !== undefined) chartObject.Top = input.position.top;
          if (input.position.width !== undefined) chartObject.Width = input.position.width; // Permitir modificar tamaño al reposicionar
          if (input.position.height !== undefined) chartObject.Height = input.position.height; // Permitir modificar tamaño al reposicionar


          return { success: true, data: 'Gráfico reposicionado correctamente.' }; // Retorno ajustado
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
                    chartType: chartObject.Chart.ChartType, // Esto devolverá un número, se podría mapear a string si es necesario
                });
            }
            return { success: true, data: chartsList }; // Retorno ajustado
        }

        default:
          throw new Error(`Operación "${input.operation}" no soportada.`);
      }
    } catch (error: any) {
      // Aquí puedes usar tu manejador de errores si tienes uno centralizado
      // Por ahora, devolvemos un ErrorResponse simple
      return { success: false, error: { code: 'EXCEL_CHART_ERROR', message: error.message } };
    } finally {
      if (workbook) {
        try {
            workbook.Save();
            workbook.Close();
        } catch (closeError: any) {
            // Ignorar errores al cerrar si ya hubo un error principal
            logger.warn(`Error al cerrar el libro de trabajo: ${closeError.message}`); // Usar logger
        }
        releaseObject(workbook);
      }
      // La aplicación de Excel se gestiona externamente, no la cerramos aquí.
      releaseObject(excelApp);
    }
  },
};

export default excelChartsTool;