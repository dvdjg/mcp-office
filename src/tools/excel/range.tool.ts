import { z } from 'zod';
import { McpResource, ToolRequestParams } from '../../types/common.types';
import { getOfficeApplication } from '../../utils/officeInterop';
import { saveResource } from '../dynamic/resources.tool'; // Importar saveResource
import * as fs from 'fs-extra'; // Importar fs para leer el archivo Excel
import * as path from 'path'; // Importar path

// Define el esquema de entrada para la herramienta excel/range
const ExcelRangeInputSchema = z.object({
  filePath: z.string().describe('Ruta al archivo Excel.'),
  sheetName: z.string().optional().describe('Nombre de la hoja de cálculo. Si no se proporciona, se usa la hoja activa.'),
  sheetIndex: z.number().int().positive().optional().describe('Índice de la hoja de cálculo (1-basado). Si se proporciona, anula sheetName.'),
  rangeAddress: z.string().describe('Dirección del rango (e.g., "A1", "B2:C5").'),
  operation: z.enum(['read', 'write', 'format', 'apply']).describe('Operación a realizar en el rango.'),
  values: z.array(z.array(z.any())).optional().describe('Valores para escribir en el rango (para la operación "write").'),
  formatProperties: z.record(z.any()).optional().describe('Propiedades de formato a aplicar (para la operación "format").'),
  // Puedes añadir más propiedades para la operación 'apply' si es necesario
});

type ExcelRangeInput = z.infer<typeof ExcelRangeInputSchema>;

/**
 * @tool excel/range
 * @description Manipula rangos de celdas en hojas de cálculo de Excel.
 * Permite leer, escribir, formatear y aplicar otras operaciones a rangos específicos.
 * Utiliza COM Interop a través de winax para interactuar con Excel.
 * Requiere la ruta del archivo, el nombre o índice de la hoja, la dirección del rango y la operación.
 * Para escribir, se proporcionan los valores como un array de arrays.
 * Para formatear, se proporcionan las propiedades de formato como un objeto.
 */
const excelRangeTool: McpResource = {
  path: 'excel/range',
  description: 'Manipulates cell ranges in Excel worksheets.', // Translated description
  schema: ExcelRangeInputSchema, // Added schema
  handler: async (params: any) => {
    let excelApp;
    try {
      // Validar los parámetros de entrada
      const input = ExcelRangeInputSchema.parse(params);

      excelApp = await getOfficeApplication('Excel.Application'); // Corregido el nombre de la aplicación
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
        throw new Error(`Hoja de cálculo "${input.sheetName || input.sheetIndex}" no encontrada.`);
      }

      const range = worksheet.Range(input.rangeAddress);

      if (!range) {
        throw new Error(`Rango "${input.rangeAddress}" inválido.`);
      }

      switch (input.operation) {
        case 'read':
          // Leer valores del rango
          const values = range.Value2;
          return { success: true, data: values }; // Retorno de éxito con data

        case 'write':
          // Escribir valores en el rango
          if (!input.values) {
            throw new Error('Se requieren valores para la operación "write".');
          }
          // Asegurarse de que el tamaño del rango coincide con el de los valores
          // Esto es una simplificación; una implementación robusta debería manejar tamaños diferentes
          // o permitir escribir en un rango de destino más grande.
          if (range.Rows.Count !== input.values.length || range.Columns.Count !== (input.values[0]?.length || 0)) {
             // Si el rango de destino es una sola celda, winax puede manejar la escritura de un array de arrays
             if (range.Cells.Count === 1) {
                range.Value = input.values;
             } else {
                // Para rangos más grandes, intentar escribir directamente puede fallar si las dimensiones no coinciden exactamente.
                // Una alternativa sería iterar sobre las celdas o usar CopyFromRecordset si los datos provienen de una fuente compatible.
                // Por ahora, lanzamos un error si las dimensiones no coinciden para rangos > 1 celda.
                 throw new Error(`Las dimensiones de los valores proporcionados no coinciden con las del rango. Rango: ${range.Rows.Count}x${range.Columns.Count}, Valores: ${input.values.length}x${(input.values[0]?.length || 0)}`);
             }
          } else {
             range.Value = input.values;
          }
          workbook.Save();
          // Guardar el archivo Excel modificado como un recurso dinámico
          try {
              const excelContent = await fs.readFile(input.filePath, null); // Leer como Buffer
              await saveResource('excel/range', path.basename(input.filePath), excelContent);
              // logger.info(`Saved ${input.filePath} as a dynamic resource.`);
          } catch (resourceSaveError: any) {
              // logger.error(`Failed to save ${input.filePath} as a dynamic resource: ${resourceSaveError.message}`);
              // Continuar la ejecución aunque falle el guardado del recurso
          }
          return { success: true, data: null, message: `Valores escritos en el rango "${input.rangeAddress}".` }; // Incluir data: null

        case 'format':
          // Aplicar formato al rango
          if (!input.formatProperties) {
            throw new Error('Se requieren propiedades de formato para la operación "format".');
          }
          // Implementación básica de formato. Se puede expandir para soportar más propiedades.
          // Ejemplo: { Font: { Bold: true, Color: 255 }, Interior: { ColorIndex: 6 } }
          for (const prop in input.formatProperties) {
            if (Object.prototype.hasOwnProperty.call(input.formatProperties, prop)) {
              const value = input.formatProperties[prop];
              if (typeof range[prop] === 'object' && range[prop] !== null) {
                 // Si la propiedad es un objeto (como Font, Interior), aplicar propiedades anidadas
                 for (const subProp in value) {
                    if (Object.prototype.hasOwnProperty.call(value, subProp)) {
                       range[prop][subProp] = value[subProp];
                    }
                 }
              } else {
                 // Si la propiedad es un valor directo
                 range[prop] = value;
              }
            }
          }
          workbook.Save();
          // Guardar el archivo Excel modificado como un recurso dinámico
          try {
              const excelContent = await fs.readFile(input.filePath, null); // Leer como Buffer
              await saveResource('excel/range', path.basename(input.filePath), excelContent);
              // logger.info(`Saved ${input.filePath} as a dynamic resource.`);
          } catch (resourceSaveError: any) {
              // logger.error(`Failed to save ${input.filePath} as a dynamic resource: ${resourceSaveError.message}`);
              // Continuar la ejecución aunque falle el guardado del recurso
          }
          return { success: true, data: null, message: `Formato aplicado al rango "${input.rangeAddress}".` }; // Incluir data: null

        case 'apply':
          // Operación genérica para aplicar propiedades al rango
          if (!input.formatProperties) { // Reutilizamos formatProperties para propiedades generales
             throw new Error('Se requieren propiedades para la operación "apply".');
          }
           for (const prop in input.formatProperties) {
            if (Object.prototype.hasOwnProperty.call(input.formatProperties, prop)) {
              range[prop] = input.formatProperties[prop];
            }
          }
          workbook.Save();
          // Guardar el archivo Excel modificado como un recurso dinámico
          try {
              const excelContent = await fs.readFile(input.filePath, null); // Leer como Buffer
              await saveResource('excel/range', path.basename(input.filePath), excelContent);
              // logger.info(`Saved ${input.filePath} as a dynamic resource.`);
          } catch (resourceSaveError: any) {
              // logger.error(`Failed to save ${input.filePath} as a dynamic resource: ${resourceSaveError.message}`);
              // Continuar la ejecución aunque falle el guardado del recurso
          }
          return { success: true, data: null, message: `Propiedades aplicadas al rango "${input.rangeAddress}".` }; // Incluir data: null

        default:
          throw new Error(`Operación "${input.operation}" no soportada.`);
      }
    } catch (error: any) {
      // Manejo básico de errores
      return { success: false, error: { code: 'OFFICE_API_ERROR', message: error.message } }; // Ajustar formato de error
    } finally {
      // Considerar si cerrar Excel o dejarlo abierto.
      // Para un servidor MCP, probablemente quieras dejarlo abierto para futuras operaciones.
      // Si decides cerrarlo, asegúrate de guardar primero si es necesario.
      // if (excelApp) {
      //   excelApp.Quit();
      // }
    }
  },
};

export default excelRangeTool;