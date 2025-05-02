import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '@/types/common.types'; // Import McpResource, ToolRequestParams, ApiResponse
import * as winax from 'winax';
import mermaid from 'mermaid'; // Asumo que mermaid está instalado
import { handleToolError } from '@/utils/errorHandler'; // Import handleToolError
import * as fs from 'fs/promises'; // Importar módulo fs para manejo de archivos
import * as path from 'path'; // Importar módulo path
import os from 'os'; // Importar módulo os para directorio temporal

// Esquema de entrada para la herramienta word/mermaid/import
const mermaidImportInputSchema = z.object({
  filePath: z.string().describe('Ruta al documento Word.'),
  mermaidSyntax: z.string().describe('Sintaxis del diagrama Mermaid.'),
  format: z.enum(['svg', 'png']).default('svg').describe('Formato de salida para el diagrama renderizado.'),
  location: z.string().optional().describe('Ubicación en el documento para insertar el diagrama (p. ej., marcador, ContentControl).'),
});

/**
 * @tool word/mermaid/import
 * @description Importa un diagrama Mermaid en un documento Word.
 *
 * @param {object} params - Parámetros de entrada.
 * @param {string} params.filePath - Ruta al documento Word.
 * @param {string} params.mermaidSyntax - Sintaxis del diagrama Mermaid.
 * @param {'svg' | 'png'} [params.format='svg'] - Formato de salida para el diagrama renderizado.
 * @param {string} [params.location] - Ubicación en el documento para insertar el diagrama (p.ej., marcador, ContentControl).
 *
 * @returns {Promise<ApiResponse<object>>} - Resultado de la operación.
 */
const mermaidImportTool: McpResource = {
  path: 'word/mermaid/import', // Define the path here
  description: 'Importa un diagrama Mermaid en un documento Word.',
  schema: mermaidImportInputSchema, // Use schema for input validation
 handler: async (params: ToolRequestParams): Promise<ApiResponse<object>> => { // Use ToolRequestParams and ApiResponse
   let wordApp: any = null;
   let tempFilePath: string | null = null;

   try {
     const validatedParams = mermaidImportInputSchema.parse(params);
     const { filePath, mermaidSyntax, format, location } = validatedParams;

     // Inicializar Mermaid
     mermaid.initialize({ startOnLoad: false });

     // Renderizar la sintaxis Mermaid
     const { svg } = await mermaid.render('mermaid-diagram', mermaidSyntax);

     // Crear archivo temporal para la imagen
     const tempDir = os.tmpdir();
     const fileExtension = format === 'svg' ? '.svg' : '.png'; // Asumimos que si no es SVG, es PNG (aunque la renderización actual es solo SVG)
     tempFilePath = path.join(tempDir, `mermaid_diagram_${Date.now()}${fileExtension}`);

     // Si el formato es PNG, necesitaríamos convertir el SVG a PNG aquí.
     // Por ahora, solo guardamos el SVG.
     await fs.writeFile(tempFilePath, svg);

     // Abrir el documento Word
     wordApp = new winax.Object('Word.Application');
     wordApp.Visible = true; // Opcional: hacer visible la aplicación Word
     const doc = wordApp.Documents.Open(filePath);

     // Determinar la ubicación de inserción. Por ahora, insertamos al final del documento.
     // TODO: Implementar lógica para insertar en una ubicación específica (marcador, ContentControl).
     const selection = wordApp.Selection;
     selection.EndKey(6); // Mover cursor al final del documento

     // Insertar la imagen
     // Usamos AddPicture para insertar desde un archivo. LinkToFile y SaveWithDocument en false.
     // El rango es la selección actual.
     const range = selection.Range;
     const inlineShape = range.InlineShapes.AddPicture(
       tempFilePath,
       false, // LinkToFile
       true // SaveWithDocument
     );

     // Opcional: Ajustar tamaño o posición si es necesario
     // inlineShape.Width = 300;
     // inlineShape.Height = 200;

     // Guardar y cerrar el documento
     doc.Save();
     doc.Close();

     console.log(`Diagrama Mermaid insertado en ${filePath}`);

     return { success: true, data: { message: `Diagrama Mermaid insertado en ${filePath}` } };
   } catch (error: any) {
     console.error(`Error al importar diagrama Mermaid: ${error.message}`);
     return handleToolError(error, 'MERMAID_IMPORT_ERROR');
   } finally {
     // Limpiar archivo temporal
     if (tempFilePath) {
       try {
         await fs.unlink(tempFilePath);
       } catch (cleanupError: any) {
         console.error(`Error al limpiar archivo temporal ${tempFilePath}: ${cleanupError.message}`);
       }
     }
     // Cerrar la aplicación Word si se abrió
     if (wordApp) {
       wordApp.Quit();
     }
   }
 },
};

export default mermaidImportTool;