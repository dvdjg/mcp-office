import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types';
import { OfficeAppName, getOfficeApplication } from '../../utils/officeInterop'; // Importar los nombres correctos
import { saveResource } from '../dynamic/resources.tool'; // Importar saveResource
import * as fs from 'fs-extra'; // Importar fs para leer el archivo PDF
import * as path from 'path'; // Importar path

// Esquema de entrada para la herramienta office/pdf/export
const PdfExportInputSchema = z.object({
  filePath: z.string().describe('Ruta al archivo de Office de origen (Word, Excel, o PowerPoint).'),
  outputFilePath: z.string().describe('Ruta donde guardar el archivo PDF de salida.'),
  application: z.enum(['Word', 'Excel', 'PowerPoint']).describe('La aplicación de Office a utilizar.'),
  configuration: z.record(z.any()).optional().describe('Configuración opcional para la exportación a PDF (e.g., rango de páginas, calidad).'),
});

type PdfExportInput = z.infer<typeof PdfExportInputSchema>;

/**
 * @tool office/pdf/export
 * @description Herramienta para exportar documentos de Office (Word, Excel, PowerPoint) a formato PDF.
 * Utiliza COM Interop a través de winax para interactuar con las aplicaciones de Office.
 * Soporta operaciones para convertir, configurar (opcional) y guardar el archivo PDF.
 * @operation convert - Convierte el archivo de Office especificado a PDF.
 * @operation configure - (Opcional) Permite configurar opciones de exportación antes de convertir. (Implementación básica)
 * @operation save - Guarda el archivo convertido a la ruta especificada. (Integrado en convert)
 * @param filePath - Ruta al archivo de Office de origen.
 * @param outputFilePath - Ruta donde guardar el archivo PDF de salida.
 * @param application - La aplicación de Office a utilizar ('Word', 'Excel', 'PowerPoint').
 * @param configuration - Configuración opcional para la exportación a PDF.
 */
export const pdfExportTool: McpResource = {
  path: 'office/pdf/export', // Usar 'path' en lugar de 'name'
  description: 'Exporta documentos de Office a PDF.',
  schema: PdfExportInputSchema, // Usar 'schema' en lugar de 'inputSchema'
  handler: async (params: ToolRequestParams, context: any): Promise<ApiResponse<any>> => { // Ajustar tipo de retorno
    let officeApp: any | null = null;
    let doc = null;

    try {
      // Validar los parámetros de entrada
      const validationResult = PdfExportInputSchema.safeParse(params); // Validar params directamente

      if (!validationResult.success) {
        return {
          success: false,
          error: { // Ajustar estructura del error
            code: 'VALIDATION_ERROR',
            message: 'Error de validación de entrada',
            details: validationResult.error.errors,
          },
        };
      }

      const { filePath, outputFilePath, application, configuration } = validationResult.data;

      // Obtener la aplicación de Office
      // Mapear el nombre de la aplicación a su ProgID
      let appProgId: OfficeAppName;
      if (application === 'Word') {
        appProgId = 'Word.Application';
      } else if (application === 'Excel') {
        appProgId = 'Excel.Application';
      } else if (application === 'PowerPoint') {
        appProgId = 'PowerPoint.Application';
      } else {
         // Although schema validation should catch this, keep as a fallback
         return { // Retornar como ErrorResponse
           success: false,
           error: {
             code: 'UNSUPPORTED_APPLICATION',
             message: `Aplicación de Office no soportada: ${application}`,
           },
         };
      }

      officeApp = await getOfficeApplication(appProgId);
      if (!officeApp) {
        return { // Retornar como ErrorResponse
           success: false,
           error: {
             code: 'OFFICE_APP_ERROR',
             message: `No se pudo obtener la instancia de la aplicación de Office: ${application}`,
           },
         };
      }

      // Abrir el documento/presentación
      if (application === 'Word') {
        doc = officeApp.Documents.Open(filePath);
      } else if (application === 'Excel') {
        doc = officeApp.Workbooks.Open(filePath);
      } else if (application === 'PowerPoint') {
        doc = officeApp.Presentations.Open(filePath); // Corregido 'Presenations' a 'Presentations'
      }

      if (!doc) {
        return { // Retornar como ErrorResponse
           success: false,
           error: {
             code: 'FILE_OPEN_ERROR',
             message: `No se pudo abrir el archivo: ${filePath}`,
           },
         };
      }

      // Configurar opciones de exportación (implementación básica, se puede expandir)
      const exportConfig = configuration || {};
      const exportFormat = 17; // wdExportFormatPDF, xlTypePDF, ppSaveAsPDF (generalmente 17)

      // Realizar la exportación a PDF
      if (application === 'Word') {
        // Word usa ExportAsFixedFormat
        doc.ExportAsFixedFormat(
          outputFilePath,
          exportFormat,
          false, // OpenAfterExport
          0, // OptimizeFor (0=Print, 1=Screen)
          0, // Range (0=Whole document)
          0, // From
          0, // To
          0, // Item (0=wdExportDocumentContent)
          true, // IncludeDocProperties
          true, // KeepIRM
          0, // CreateBookmarks (0=wdExportCreateNoBookmarks)
          true, // DocStructureTags
          false, // BitmapMissingFonts
          false, // UseISO19005_1
          exportConfig.OptimizeFor || 0,
          exportConfig.BitmapMissingFonts || false,
          exportConfig.IncludeDocProperties || true,
          exportConfig.KeepIRM || true,
          exportConfig.CreateBookmarks || 0,
          exportConfig.DocStructureTags || true,
          exportConfig.UseISO19005_1 || false,
          exportConfig.Range || 0,
          exportConfig.From || 0,
          exportConfig.To || 0,
          exportConfig.Item || 0,
          exportConfig.OpenAfterExport || false
        );
        // Guardar el archivo PDF exportado como un recurso dinámico
        try {
            const pdfContent = await fs.readFile(outputFilePath, null); // Leer como Buffer
            await saveResource('office/pdf/export', path.basename(outputFilePath), pdfContent);
            context?.log.info(`Saved ${outputFilePath} as a dynamic resource.`);
        } catch (resourceSaveError: any) {
            context?.log.error(`Failed to save ${outputFilePath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continuar la ejecución aunque falle el guardado del recurso
        }
      } else if (application === 'Excel') {
        // Excel usa ExportAsFixedFormat
        doc.ExportAsFixedFormat(
          exportFormat,
          outputFilePath,
          exportConfig.Quality || 0, // xlQualityStandard (0) or xlQualityMinimum (1)
          exportConfig.IncludeDocProperties || true,
          exportConfig.IgnorePrintAreas || false,
          exportConfig.From || 0,
          exportConfig.To || 0,
          exportConfig.OpenAfterPublish || false,
          exportConfig.FixedFormatExtClassPtr || null
        );
        // Guardar el archivo PDF exportado como un recurso dinámico
        try {
            const pdfContent = await fs.readFile(outputFilePath, null); // Leer como Buffer
            await saveResource('office/pdf/export', path.basename(outputFilePath), pdfContent);
            context?.log.info(`Saved ${outputFilePath} as a dynamic resource.`);
        } catch (resourceSaveError: any) {
            context?.log.error(`Failed to save ${outputFilePath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continuar la ejecución aunque falle el guardado del recurso
        }
      } else if (application === 'PowerPoint') {
        // PowerPoint usa ExportAsFixedFormat
        doc.ExportAsFixedFormat(
          outputFilePath,
          exportFormat,
          exportConfig.Intent || 1, // ppExportIntentPrint (1) or ppExportIntentScreen (2)
          exportConfig.FrameSlides || false,
          exportConfig.IncludeHiddenSlides || false,
          exportConfig.PrintRange || null, // ppPrintAll (1), ppPrintSelection (2), ppPrintCurrent (3), ppPrintSlideRange (4)
          exportConfig.From || 0,
          exportConfig.To || 0,
          exportConfig.OutputType || 1, // ppPrintOutputSlides (1), ppPrintOutputNotesPages (2), ppPrintOutputOutline (3), ppPrintOutputHandouts (4)
          exportConfig.PrintHandoutOrder || 1, // ppPrintHandoutHorizontal (1), ppPrintHandoutVertical (2)
          exportConfig.FitToPage || false,
          exportConfig.Orientation || 1, // ppPrintOrientationPortrait (1), ppPrintOrientationLandscape (2)
          exportConfig.ViewSlideDetails || false,
          exportConfig.ShowComments || false,
          exportConfig.ShowInking || false,
          exportConfig.CreateHiddenSlides || false,
          exportConfig.OpenAfterExport || false
        );
        // Guardar el archivo PDF exportado como un recurso dinámico
        try {
            const pdfContent = await fs.readFile(outputFilePath, null); // Leer como Buffer
            await saveResource('office/pdf/export', path.basename(outputFilePath), pdfContent);
            context?.log.info(`Saved ${outputFilePath} as a dynamic resource.`);
        } catch (resourceSaveError: any) {
            context?.log.error(`Failed to save ${outputFilePath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continuar la ejecución aunque falle el guardado del recurso
        }
      }


      return { // Ajustar estructura de éxito
        success: true,
        data: {
          outputFilePath: outputFilePath,
          message: `Archivo exportado a PDF exitosamente: ${outputFilePath}`,
        },
      };

    } catch (error: any) {
      return { // Ajustar estructura del error
        success: false,
        error: {
          code: 'EXPORT_ERROR',
          message: `Error al exportar a PDF: ${error.message}`,
          details: error,
        },
      };
    } finally {
      // Cerrar el documento y liberar objetos COM
      if (doc) {
        try {
          doc.Close();
        } catch (e) {
          console.error('Error al cerrar el documento:', e);
        }
      }
      // No cerrar la aplicación de Office aquí, ya que puede haber otros documentos abiertos
      // La gestión de la instancia de la aplicación de Office debe ser manejada externamente si es necesario.
    }
  },
};

// Helper function to get Office application instance (assuming it exists in officeInterop.ts)
// You might need to implement or adjust this function based on your officeInterop.ts
/*
import * as winax from 'winax';

export function getOfficeApp(appName: 'Word' | 'Excel' | 'PowerPoint'): any {
  try {
    return new winax.Object.create(`${appName}.Application`);
  } catch (e) {
    console.error(`Error al crear instancia de ${appName}:`, e);
    return null;
  }
}
*/