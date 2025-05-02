import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '@/types/common.types'; // Import McpResource, ToolRequestParams, ApiResponse
import * as winax from 'winax';
import mermaid from 'mermaid'; // Asumo que mermaid está instalado
import { handleToolError } from '@/utils/errorHandler'; // Import handleToolError

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
    try {
      const validatedParams = mermaidImportInputSchema.parse(params);
      const { filePath, mermaidSyntax, format, location } = validatedParams;

      // Inicializar Mermaid
      mermaid.initialize({ startOnLoad: false });

      // Renderizar la sintaxis Mermaid
      const { svg } = await mermaid.render('mermaid-diagram', mermaidSyntax);

      // Aquí iría la lógica de COM Interop con winax para abrir el documento Word
      // e insertar el SVG o convertirlo a PNG e insertarlo.
      // Esto requiere acceso a la instancia de Word y manipulación del documento.
      // Ejemplo básico (requiere adaptación):
      // let wordApp;
      // try {
      //   wordApp = new winax.Object('Word.Application');
      //   wordApp.Visible = true;
      //   const doc = wordApp.Documents.Open(filePath);
      //
      //   // Lógica para insertar 'svg' en la ubicación especificada
      //   // Esto es complejo y depende de cómo se quiera insertar (inline, shape, content control)
      //   // y cómo manejar el formato (SVG directo si Word lo soporta bien, o convertir a imagen)
      //
      //   doc.Save();
      //   doc.Close();
      // } finally {
      //   if (wordApp) {
      //     wordApp.Quit();
      //   }
      // }

      console.log(`Diagrama Mermaid renderizado (${format}). Lógica de inserción en Word pendiente.`);

      return { success: true, data: {} }; // Devolver un resultado exitoso vacío por ahora
    } catch (error: any) {
      console.error(`Error al importar diagrama Mermaid: ${error.message}`);
      return handleToolError(error, 'MERMAID_IMPORT_ERROR'); // Usar handleToolError
    }
  },
};

export default mermaidImportTool;