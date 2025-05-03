import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import logger from '../../utils/logger'; // Importar logger
import { saveResource } from '../dynamic/resources.tool'; // Importar saveResource
import * as fs from 'fs-extra'; // Importar fs para leer el archivo Excel
import * as path from 'path'; // Importar path

// Esquema de entrada para la herramienta excel/tables
const ExcelTablesInputSchema = z.object({
  filePath: z.string().describe('Ruta al archivo Excel.'),
  operation: z.enum(['insert', 'modify', 'add', 'delete']).describe('Operación a realizar: insert, modify, add, delete.'),
  sheetName: z.string().optional().describe('Nombre de la hoja de cálculo. Si no se proporciona, se usa la hoja activa.'),
  sheetIndex: z.number().optional().describe('Índice de la hoja de cálculo (1-basado). Si no se proporciona, se usa la hoja activa.'),
  rangeAddress: z.string().optional().describe('Dirección del rango para insertar una nueva tabla (ej: "A1:D10"). Requerido para la operación "insert".'),
  tableName: z.string().optional().describe('Nombre de la tabla. Requerido para las operaciones "modify", "add", "delete".'),
  data: z.array(z.array(z.any())).optional().describe('Datos para añadir filas a la tabla. Requerido para la operación "add" con location "rows".'),
  location: z.enum(['rows', 'columns']).optional().describe('Especifica si se añaden "rows" o "columns". Requerido para la operación "add".'),
  count: z.number().optional().describe('Número de filas o columnas a añadir/eliminar. Requerido para las operaciones "add" y "delete" (excepto delete de tabla completa).'),
  position: z.number().optional().describe('Posición (índice 1-basado) donde añadir/eliminar filas/columnas. Opcional para "add" y "delete" (filas/columnas).'),
});

type ExcelTablesInput = z.infer<typeof ExcelTablesInputSchema>;

/**
 * @tool excel/tables
 * @description Herramienta para gestionar tablas en archivos Excel.
 * Permite insertar, modificar, añadir filas/columnas y eliminar tablas.
 * Utiliza COM Interop a través de winax para interactuar con Excel.
 * Requiere la ruta del archivo, la operación y parámetros específicos según la operación.
 * @input ExcelTablesInputSchema
 */
const excelTablesTool: McpResource = {
  path: 'excel/tables',
  description: 'Gestiona tablas en archivos Excel (insertar, modificar, añadir/eliminar filas/columnas, eliminar tabla).',
  schema: ExcelTablesInputSchema, // Corregido de inputSchema a schema
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => {
    let excelApp: any = null;
    let workbook: any = null;
    let worksheet: any = null;
    let filePath: string | undefined; // Declarar filePath fuera del try y permitir undefined

    try {
      const input = ExcelTablesInputSchema.parse(params);
      filePath = input.filePath; // Asignar filePath aquí
      const { operation, sheetName, sheetIndex, rangeAddress, tableName, data, location, count, position } = input;

      excelApp = await getOfficeApplication('Excel.Application'); // Usar getOfficeApplication
      workbook = excelApp.Workbooks.Open(filePath);

      if (sheetName) {
        worksheet = workbook.Sheets(sheetName);
      } else if (sheetIndex) {
        worksheet = workbook.Sheets(sheetIndex);
      } else {
        worksheet = workbook.ActiveSheet;
      }

      if (!worksheet) {
        throw new Error(`Hoja de cálculo "${sheetName || sheetIndex}" no encontrada.`);
      }

      switch (operation) {
        case 'insert':
          if (!rangeAddress) {
            throw new Error('rangeAddress es requerido para la operación "insert".');
          }
          // Insertar una nueva tabla en el rango especificado
          // El segundo argumento (xlYes) indica que la primera fila es el encabezado
          const listObject = worksheet.ListObjects.Add(1, worksheet.Range(rangeAddress), null, 1); // xlSrcRange = 1, xlListObjectHasHeaders = 1
          if (tableName) {
            listObject.Name = tableName;
          }
          return { success: true, data: `Tabla insertada en el rango ${rangeAddress}. Nombre: ${listObject.Name}` }; // Retorno ajustado

        case 'modify':
          if (!tableName) {
            throw new Error('tableName es requerido para la operación "modify".');
          }
          // Modificar una tabla existente (ej: cambiar nombre, redimensionar - no implementado en este ejemplo simple)
          const tableToModify = worksheet.ListObjects(tableName);
          if (!tableToModify) {
            throw new Error(`Tabla "${tableName}" no encontrada.`);
          }
          // Implementar lógica de modificación aquí si es necesario.
          // Por ahora, solo confirmamos que la tabla existe.
          return { success: true, data: `Tabla "${tableName}" encontrada para modificación.` }; // Retorno ajustado

        case 'add':
          if (!tableName) {
            throw new Error('tableName es requerido para la operación "add".');
          }
          if (!location) {
            throw new Error('location ("rows" o "columns") es requerido para la operación "add".');
          }
          const tableToAdd = worksheet.ListObjects(tableName);
          if (!tableToAdd) {
            throw new Error(`Tabla "${tableName}" no encontrada.`);
          }

          if (location === 'rows') {
            if (!data || data.length === 0) {
              throw new Error('data es requerido y no puede estar vacío para añadir filas.');
            }
            // Añadir filas a la tabla
            for (const rowData of data) {
              // AddDataBoundRow añade una fila vacía al final. Luego llenamos los datos.
              const newRow = tableToAdd.ListRows.Add();
              for (let i = 0; i < rowData.length; i++) {
                if (i < newRow.Range.Cells.Count) {
                   newRow.Range.Cells(1, i + 1).Value = rowData[i];
                }
              }
            }
             return { success: true, data: `${data.length} fila(s) añadida(s) a la tabla "${tableName}".` }; // Retorno ajustado

          } else if (location === 'columns') {
             if (!count || count <= 0) {
                throw new Error('count (número de columnas a añadir) es requerido y debe ser positivo para añadir columnas.');
             }
             // Añadir columnas a la tabla
             // Add method for ListColumns adds a column to the left of the specified position
             const currentColumnCount = tableToAdd.ListColumns.Count;
             const insertPosition = position !== undefined && position >= 1 && position <= currentColumnCount + 1 ? position : currentColumnCount + 1;

             for (let i = 0; i < count; i++) {
                tableToAdd.ListColumns.Add(insertPosition);
             }
             return { success: true, data: `${count} columna(s) añadida(s) a la tabla "${tableName}" en la posición ${insertPosition}.` }; // Retorno ajustado
           }
           // break; // Should not reach here - Eliminado break redundante

        case 'delete':
          if (!tableName) {
            throw new Error('tableName es requerido para la operación "delete".');
          }
          const tableToDelete = worksheet.ListObjects(tableName);
          if (!tableToDelete) {
            throw new Error(`Tabla "${tableName}" no encontrada.`);
          }

          if (location === 'rows') {
             if (!count || count <= 0) {
                throw new Error('count (número de filas a eliminar) es requerido y debe ser positivo para eliminar filas.');
             }
             // Eliminar filas de la tabla
             const currentRowCount = tableToDelete.ListRows.Count;
             const deletePosition = position !== undefined && position >= 1 && position <= currentRowCount ? position : currentRowCount - count + 1;

             if (deletePosition < 1 || deletePosition + count - 1 > currentRowCount) {
                 throw new Error(`Rango de eliminación de filas inválido. Intentando eliminar ${count} fila(s) desde la posición ${deletePosition} en una tabla con ${currentRowCount} fila(s).`);
             }

             for (let i = 0; i < count; i++) {
                tableToDelete.ListRows(deletePosition).Delete();
             }
             return { success: true, data: `${count} fila(s) eliminada(s) de la tabla "${tableName}" desde la posición ${deletePosition}.` }; // Retorno ajustado

          } else if (location === 'columns') {
             if (!count || count <= 0) {
                throw new Error('count (número de columnas a eliminar) es requerido y debe ser positivo para eliminar columnas.');
             }
             // Eliminar columnas de la tabla
             const currentColumnCount = tableToDelete.ListColumns.Count;
             const deletePosition = position !== undefined && position >= 1 && position <= currentColumnCount ? position : currentColumnCount - count + 1;

             if (deletePosition < 1 || deletePosition + count - 1 > currentColumnCount) {
                 throw new Error(`Rango de eliminación de columnas inválido. Intentando eliminar ${count} columna(s) desde la posición ${deletePosition} en una tabla con ${currentColumnCount} columna(s).`);
             }

             for (let i = 0; i < count; i++) {
                tableToDelete.ListColumns(deletePosition).Delete();
             }
             return { success: true, data: `${count} columna(s) eliminada(s) de la tabla "${tableName}" desde la posición ${deletePosition}.` }; // Retorno ajustado

          } else {
            // Eliminar la tabla completa
            tableToDelete.Delete();
            return { success: true, data: `Tabla "${tableName}" eliminada.` }; // Retorno ajustado
          }
          // break; // Should not reach here - Eliminado break redundante
      }

    } catch (error: any) {
      logger.error(`Error in excel/tables tool: ${error.message}`); // Usar logger
      return { success: false, error: { code: 'EXCEL_TABLES_ERROR', message: `Error al gestionar tablas en Excel: ${error.message}` } }; // Retorno de error ajustado
    } finally {
      if (workbook) {
        try {
            workbook.Save();
            // Guardar el archivo Excel modificado como un recurso dinámico
            // Esto se hace en el finally porque Save() ocurre aquí para todas las operaciones de modificación.
            // No necesitamos verificar la operación específica aquí.
            // Asegurarse de que filePath tiene un valor antes de intentar leer el archivo
            if (filePath) {
                try {
                    const excelContent = await fs.readFile(filePath, null); // Leer como Buffer
                    await saveResource('excel/tables', path.basename(filePath), excelContent);
                    // logger.info(`Saved ${filePath} as a dynamic resource.`);
                } catch (resourceSaveError: any) {
                    // logger.error(`Failed to save ${filePath} as a dynamic resource: ${resourceSaveError.message}`);
                    // Continuar la ejecución aunque falle el guardado del recurso
                }
            }
            workbook.Close();
        } catch (closeError: any) {
            logger.warn(`Error al cerrar el libro de trabajo: ${closeError.message}`); // Usar logger
        }
        releaseObject(workbook);
      }
      // La aplicación de Excel se gestiona externamente, no la cerramos aquí.
      releaseObject(excelApp);
    }
    // Añadir un retorno al final para cubrir todos los casos posibles
    // Esto solo se alcanzará si no se lanzó un error o se retornó antes.
    // En un escenario ideal, todos los casos del switch deberían retornar.
    // Pero para satisfacer al linter, añadimos este retorno de fallback.
    return { success: false, error: { code: 'UNHANDLED_CASE', message: 'La operación de tabla de Excel no retornó un resultado explícito.' } };
  },
};

export default excelTablesTool;