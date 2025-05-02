import { z } from 'zod';
import { McpResource, ToolRequestParams } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';

// Define el esquema de entrada para la herramienta excel/worksheets
const ExcelWorksheetsInputSchema = z.object({
  filePath: z.string().describe('Ruta al archivo Excel.'),
  operation: z.enum(['add', 'delete', 'rename', 'set']).describe('Operación a realizar en las hojas de cálculo.'),
  sheetName: z.string().optional().describe('Nombre de la hoja (para delete, rename, set).'),
  newSheetName: z.string().optional().describe('Nuevo nombre para la hoja (para rename).'),
  sheetIndex: z.number().int().positive().optional().describe('Índice de la hoja (1-basado, opcional para identificar la hoja).'),
  beforeSheet: z.string().optional().describe('Nombre de la hoja antes de la cual insertar la nueva hoja (para add).'),
  afterSheet: z.string().optional().describe('Nombre de la hoja después de la cual insertar la nueva hoja (para add).'),
});

type ExcelWorksheetsInput = z.infer<typeof ExcelWorksheetsInputSchema>;

/**
 * @tool excel/worksheets
 * @description Gestiona hojas de cálculo en archivos de Excel.
 * Permite añadir, eliminar, renombrar y seleccionar hojas.
 * Utiliza COM Interop a través de winax para interactuar con Excel.
 * Requiere la ruta del archivo y la operación a realizar.
 * Las operaciones delete, rename y set requieren identificar la hoja por nombre o índice.
 * La operación add permite especificar la posición de inserción.
 * @input ExcelWorksheetsInputSchema
 * @output string - Un mensaje indicando el resultado de la operación.
 */
export const excelWorksheetsTool: McpResource[] = [{
  path: 'excel/worksheets', // Añadir la ruta de la herramienta
  description: 'Manages worksheets in Excel files.',
  schema: ExcelWorksheetsInputSchema, // Cambiar inputSchema a schema
  handler: async (params: ToolRequestParams) => { // Tipificar params
    const { filePath, operation, sheetName, newSheetName, sheetIndex, beforeSheet, afterSheet } = params as ExcelWorksheetsInput; // Castear params
    let excelApp: any;
    let workbook: any;
    let sheets: any;
    let sheet: any;
    let resultData: string | object = ''; // Usar un tipo más flexible para el resultado

    try {
      excelApp = await getOfficeApplication('Excel.Application');
      excelApp.Visible = false; // Mantener Excel oculto

      try {
        // Intentar abrir el libro existente
        workbook = excelApp.Workbooks.Open(filePath);
      } catch (error: any) {
        // Si no existe, crear uno nuevo (solo para la operación 'add')
        if (operation === 'add') {
          workbook = excelApp.Workbooks.Add();
          // Guardar el nuevo libro inmediatamente para poder añadir hojas
          workbook.SaveAs(filePath);
        } else {
          throw new Error(`El archivo Excel no se encontró en la ruta especificada: ${filePath}`);
        }
      }

      sheets = workbook.Sheets;

      switch (operation) {
        case 'add': {
          let beforeSheetObj = undefined;
          let afterSheetObj = undefined;

          if (beforeSheet) {
            try {
              beforeSheetObj = sheets.Item(beforeSheet);
            } catch {
              throw new Error(`La hoja de referencia 'beforeSheet' no se encontró: ${beforeSheet}`);
            }
          } else if (afterSheet) {
             try {
              afterSheetObj = sheets.Item(afterSheet);
            } catch {
              throw new Error(`La hoja de referencia 'afterSheet' no se encontró: ${afterSheet}`);
            }
          }

          // Añadir la nueva hoja
          const newSheet = sheets.Add({ Before: beforeSheetObj, After: afterSheetObj });
          if (sheetName) {
            newSheet.Name = sheetName;
            resultData = { message: `Hoja '${sheetName}' añadida exitosamente.` };
          } else {
             resultData = { message: `Hoja añadida exitosamente con nombre por defecto '${newSheet.Name}'.` };
          }
          break;
        }

        case 'delete': {
          if (!sheetName && sheetIndex === undefined) {
            throw new Error('Se requiere sheetName o sheetIndex para la operación delete.');
          }
          try {
             // Desactivar alertas para evitar el cuadro de diálogo de confirmación de eliminación
            excelApp.DisplayAlerts = false;
            if (sheetName) {
              sheet = sheets.Item(sheetName);
              sheet.Delete();
              resultData = { message: `Hoja '${sheetName}' eliminada exitosamente.` };
            } else if (sheetIndex !== undefined) {
              sheet = sheets.Item(sheetIndex);
              const deletedSheetName = sheet.Name;
              sheet.Delete();
              resultData = { message: `Hoja en el índice ${sheetIndex} ('${deletedSheetName}') eliminada exitosamente.` };
            }
          } catch (error: any) {
             throw new Error(`No se pudo eliminar la hoja. Verifique el nombre o índice. Error: ${error.message}`);
          } finally {
             // Volver a activar las alertas
             excelApp.DisplayAlerts = true;
          }
          break;
        }

        case 'rename': {
          if (!sheetName && sheetIndex === undefined) {
            throw new Error('Se requiere sheetName o sheetIndex para la operación rename.');
          }
          if (!newSheetName) {
            throw new Error('Se requiere newSheetName para la operación rename.');
          }
          try {
            if (sheetName) {
              sheet = sheets.Item(sheetName);
            } else if (sheetIndex !== undefined) {
              sheet = sheets.Item(sheetIndex);
            }
            const oldSheetName = sheet.Name;
            sheet.Name = newSheetName;
            resultData = { message: `Hoja '${oldSheetName}' renombrada a '${newSheetName}' exitosamente.` };
          } catch (error: any) {
             throw new Error(`No se pudo renombrar la hoja. Verifique el nombre o índice y el nuevo nombre. Error: ${error.message}`);
          }
          break;
        }

        case 'set': {
           if (!sheetName && sheetIndex === undefined) {
            throw new Error('Se requiere sheetName o sheetIndex para la operación set.');
          }
          try {
            if (sheetName) {
              sheet = sheets.Item(sheetName);
            } else if (sheetIndex !== undefined) {
              sheet = sheets.Item(sheetIndex);
            }
            sheet.Activate();
            resultData = { message: `Hoja '${sheet.Name}' seleccionada exitosamente.` };
          } catch (error: any) {
             throw new Error(`No se pudo seleccionar la hoja. Verifique el nombre o índice. Error: ${error.message}`);
          }
          break;
        }

        default:
          throw new Error(`Operación no soportada: ${operation}`);
      }

      // Guardar los cambios y cerrar el libro
      workbook.Save();
      workbook.Close();

      return { success: true, data: resultData };

    } catch (error: any) {
      console.error('Error en excel/worksheets tool:', error);
      return {
        success: false,
        error: {
          code: 'EXCEL_WORKSHEETS_ERROR',
          message: error.message || 'Ocurrió un error al gestionar las hojas de cálculo de Excel.',
          details: { filePath, operation, sheetName, sheetIndex }
        }
      };
    } finally {
      // Liberar los objetos COM
      if (sheet) releaseObject(sheet);
      if (sheets) releaseObject(sheets);
      if (workbook) releaseObject(workbook);
      if (excelApp) releaseObject(excelApp);
    }
  },
}];