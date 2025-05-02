import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '@/types/common.types'; // Import McpResource, ToolRequestParams, ApiResponse
import * as winax from 'winax';
import * as fs from 'fs/promises'; // Para guardar el archivo
import * as path from 'path'; // Importar módulo path
import { handleToolError } from '@/utils/errorHandler'; // Import handleToolError

// Esquema de entrada para la herramienta word/mermaid/export
const mermaidExportInputSchema = z.object({
  filePath: z.string().describe('Ruta al documento Word.'),
  diagramIdentifier: z.string().optional().describe('Identificador del diagrama a exportar (p. ej., nombre del ContentControl).'),
  outputDirectory: z.string().describe('Directorio donde guardar el archivo del diagrama.'),
  format: z.enum(['svg', 'png', 'txt']).default('txt').describe('Formato de salida para el diagrama exportado (svg, png, o txt para la sintaxis).'),
});

/**
 * @tool word/mermaid/export
 * @description Exporta un diagrama Mermaid de un documento Word.
 *
 * @param {object} params - Parámetros de entrada.
 * @param {string} params.filePath - Ruta al documento Word.
 * @param {string} [params.diagramIdentifier] - Identificador del diagrama a exportar (p. ej., nombre del ContentControl).
 * @param {string} params.outputDirectory - Directorio donde guardar el archivo del diagrama.
 * @param {'svg' | 'png' | 'txt'} [params.format='txt'] - Formato de salida para el diagrama exportado (svg, png, o txt para la sintaxis).
 *
 * @returns {Promise<ApiResponse<object>>} - Resultado de la operación.
 */
const mermaidExportTool: McpResource = {
  path: 'word/mermaid/export', // Define the path here
  description: 'Exporta un diagrama Mermaid de un documento Word.',
  schema: mermaidExportInputSchema, // Use schema for input validation
  handler: async (params: ToolRequestParams): Promise<ApiResponse<object>> => { // Use ToolRequestParams and ApiResponse
    let wordApp: any = null;

    try {
      const validatedParams = mermaidExportInputSchema.parse(params);
      const { filePath, diagramIdentifier, outputDirectory, format } = validatedParams;

      // Abrir el documento Word
      wordApp = new winax.Object('Word.Application');
      wordApp.Visible = true; // Opcional: hacer visible la aplicación Word
      const doc = wordApp.Documents.Open(filePath);

      let extractedContent: string | Buffer | null = null;
      let foundDiagram = false;
      let diagramName = 'mermaid_diagram';

      // Lógica para encontrar el diagrama
      // Priorizamos buscar en ContentControls si se usaron para la importación
      // Esto asume que la sintaxis Mermaid se guardó en el Tag o Title del ContentControl
      // o como texto dentro de él.
      if (doc.ContentControls && doc.ContentControls.Count > 0) {
        for (let i = 1; i <= doc.ContentControls.Count; i++) {
          const cc = doc.ContentControls.Item(i);
          // Podríamos buscar por un Tag o Title específico si se definió en la importación
          // Por ahora, buscamos si el texto dentro del CC parece sintaxis Mermaid
          if (cc.Range && cc.Range.Text && cc.Range.Text.includes('graph') || cc.Range.Text.includes('sequenceDiagram')) {
             if (!diagramIdentifier || cc.Title === diagramIdentifier || cc.Tag === diagramIdentifier || cc.Range.Text.includes(diagramIdentifier)) {
                extractedContent = cc.Range.Text;
                diagramName = cc.Title || cc.Tag || diagramName; // Usar Title o Tag como nombre si existen
                foundDiagram = true;
                break; // Encontramos el primer diagrama que coincide o parece ser Mermaid
             }
          }
        }
      }

      // Si no se encontró en ContentControls o no se especificó identifier, buscar en InlineShapes (imágenes)
      if (!foundDiagram && doc.InlineShapes && doc.InlineShapes.Count > 0) {
          // Si se especificó un identifier, esto es más complejo sin ContentControls.
          // Podríamos intentar buscar por texto alternativo si se guardó allí.
          // O simplemente iterar y si no hay identifier, tomar la primera imagen.
          for (let i = 1; i <= doc.InlineShapes.Count; i++) {
              const shape = doc.InlineShapes.Item(i);
              // Verificar si es una imagen (Type 1 es msoPicture)
              if (shape.Type === 1) {
                  // Si no hay identifier, o si el texto alternativo coincide con el identifier
                  if (!diagramIdentifier || (shape.AlternativeText && shape.AlternativeText.includes(diagramIdentifier))) {
                      // Intentar guardar la imagen. Word COM no tiene un método directo para obtener los bytes de la imagen.
                      // Una forma es copiar la imagen, pegarla en un nuevo documento temporal, y guardar ese documento como imagen.
                      // Esto es complejo. Por ahora, solo indicaremos que se encontró una imagen.
                      // TODO: Implementar la extracción real de la imagen.
                      console.warn(`Se encontró una imagen (InlineShape) que podría ser un diagrama. La extracción de imagen directa no está implementada.`);
                      // Podríamos intentar guardar el documento temporalmente como HTML o filtrar el DOCX para extraer la imagen.
                      // Por ahora, solo marcamos como encontrado si no se especificó identifier o si el texto alternativo coincide.
                      if (!diagramIdentifier || (shape.AlternativeText && shape.AlternativeText.includes(diagramIdentifier))) {
                         foundDiagram = true;
                         diagramName = shape.AlternativeText || `image_${i}`; // Usar texto alternativo como nombre si existe
                         // No podemos extraer el contenido real de la imagen fácilmente aquí.
                         // Si el formato solicitado es 'txt', y solo encontramos una imagen, no podemos cumplir la solicitud.
                         if (format !== 'txt') {
                             // Aquí iría la lógica compleja para extraer la imagen binaria.
                             // Por ahora, solo indicamos que se encontró.
                             extractedContent = 'IMAGE_PLACEHOLDER'; // Usar un placeholder
                         } else {
                             console.warn(`Se solicitó formato 'txt' pero solo se encontró una imagen. No se puede extraer la sintaxis.`);
                             foundDiagram = false; // No podemos cumplir la solicitud de txt
                         }
                         if (foundDiagram) break; // Encontramos una imagen que coincide o es la primera sin identifier
                      }
                  }
              }
          }
      }


      if (foundDiagram && extractedContent !== null) {
        const fileName = diagramIdentifier ? `${diagramIdentifier}.${format}` : `${diagramName}.${format}`;
        const outputPath = path.join(outputDirectory, fileName);

        // Asegurarse de que el directorio de salida existe
        await fs.mkdir(outputDirectory, { recursive: true });

        if (format === 'txt') {
          if (typeof extractedContent === 'string') {
             await fs.writeFile(outputPath, extractedContent, 'utf-8');
             console.log(`Sintaxis Mermaid exportada a ${outputPath}`);
          } else {
             console.error(`Error: Se solicitó formato 'txt' pero el contenido extraído no es texto.`);
             // Podríamos lanzar un error aquí o devolver un resultado de fallo específico
             return { success: false, error: { code: 'EXTRACTION_ERROR', message: 'El contenido extraído no es texto.' } };
          }
        } else { // svg o png
          // Si extractedContent es 'IMAGE_PLACEHOLDER', significa que encontramos una imagen pero no la extrajimos.
          if (extractedContent === 'IMAGE_PLACEHOLDER') {
              console.error(`Error: La extracción de imagen binaria no está implementada.`);
              return { success: false, error: { code: 'NOT_IMPLEMENTED', message: 'La extracción de imagen binaria no está implementada.' } };
          }
          // Si extractedContent es la sintaxis (extraída de ContentControl), necesitamos renderizarla y guardarla como imagen.
          else if (typeof extractedContent === 'string') {
              try {
                  // Renderizar la sintaxis a SVG
                  const mermaid = require('mermaid'); // Importar mermaid aquí para usarlo solo cuando sea necesario
                  mermaid.initialize({ startOnLoad: false });
                  const { svg } = await mermaid.render('mermaid-diagram-export', extractedContent);

                  // Si el formato es PNG, necesitaríamos convertir el SVG a PNG aquí.
                  // Por ahora, solo guardamos el SVG si se solicita SVG, o indicamos que PNG no está implementado.
                  if (format === 'svg') {
                      await fs.writeFile(outputPath, svg);
                      console.log(`Diagrama Mermaid renderizado y exportado como SVG a ${outputPath}`);
                  } else if (format === 'png') {
                      console.error(`Error: La conversión de SVG a PNG no está implementada.`);
                      return { success: false, error: { code: 'NOT_IMPLEMENTED', message: 'La conversión de SVG a PNG no está implementada.' } };
                  }

              } catch (renderError: any) {
                  console.error(`Error al renderizar diagrama para exportación: ${renderError.message}`);
                  return handleToolError(renderError, 'MERMAID_RENDER_ERROR');
              }
          } else {
               console.error(`Error: Tipo de contenido extraído desconocido.`);
               return { success: false, error: { code: 'EXTRACTION_ERROR', message: 'Tipo de contenido extraído desconocido.' } };
          }
        }
      } else {
        console.warn(`No se encontró el diagrama con identificador: ${diagramIdentifier || 'ninguno especificado'}`);
        return { success: false, error: { code: 'NOT_FOUND', message: `No se encontró el diagrama con identificador: ${diagramIdentifier || 'ninguno especificado'}.` } };
      }

      doc.Close();

      return { success: true, data: { message: `Operación de exportación completada para ${filePath}.` } };
    } catch (error: any) {
      console.error(`Error al exportar diagrama Mermaid: ${error.message}`);
      return handleToolError(error, 'MERMAID_EXPORT_ERROR');
    } finally {
      // Cerrar la aplicación Word si se abrió
      if (wordApp) {
        wordApp.Quit();
      }
    }
  },
};

export default mermaidExportTool;