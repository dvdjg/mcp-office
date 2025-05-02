import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse, FastMCPContext } from '@/types/common.types'; // Importar McpResource, ToolRequestParams, ApiResponse y FastMCPContext
import { getOfficeApplication, OfficeAppName } from '../../utils/officeInterop'; // Importar OfficeAppName
import { TextContent, ContentResult } from 'fastmcp'; // Importar TextContent y ContentResult

// Define el esquema de entrada para la herramienta office/ai-suggest
const AiSuggestInputSchema = z.object({
  application: z.enum(['Word.Application', 'Excel.Application', 'PowerPoint.Application']), // Usar nombres completos
  operation: z.enum(['format', 'search', 'chart']),
  contextRange: z.string().optional(), // Podría ser un rango de Excel, un identificador de párrafo, etc.
  filePath: z.string().optional(), // Ruta del archivo si es relevante
});

type AiSuggestInput = z.infer<typeof AiSuggestInputSchema>;

/**
 * @tool office/ai-suggest
 * @description Proporciona sugerencias impulsadas por IA basadas en el contexto actual de una aplicación de Office.
 * Permite sugerir formatos, términos de búsqueda o tipos de gráficos.
 * @input AiSuggestInputSchema
 * @output z.string() // La sugerencia generada por la IA
 */
export const aiSuggestTool: McpResource = { // Usar McpResource
  path: 'office/ai-suggest', // Definir el path
  description: 'Proporciona sugerencias impulsadas por IA basadas en el contexto actual de una aplicación de Office.',
  schema: AiSuggestInputSchema, // Usar schema en lugar de inputSchema
  handler: async (params: ToolRequestParams, context?: FastMCPContext<any>) => { // Usar ToolRequestParams y FastMCPContext<any>
    // Acceder a la función de sampling a través del contexto de la sesión
    const requestSampling = context?.session?.requestSampling;

    try {
      // Validar los parámetros de entrada
      const input = AiSuggestInputSchema.parse(params);

      if (!requestSampling) {
          throw new Error('La función requestSampling no está disponible en el contexto de la sesión.');
      }

      const { application, operation, contextRange, filePath } = input;

      // Obtener la aplicación de Office activa
      const officeApp = await getOfficeApplication(application as OfficeAppName); // Usar await y OfficeAppName
      if (!officeApp) {
        throw new Error(`Aplicación de Office no encontrada o no soportada: ${application}`);
      }

      let contextText = ''; // Renombrar para evitar conflicto con el parámetro context
      try {
        // Intentar obtener el texto seleccionado a través del objeto COM
        // La forma de acceder a la selección varía ligeramente entre aplicaciones
        if (application === 'Word.Application') {
            contextText = officeApp.Selection.Text;
        } else if (application === 'Excel.Application') {
            contextText = officeApp.Selection.Text; // O Value, dependiendo del tipo de dato
        } else if (application === 'PowerPoint.Application') {
             // PowerPoint selection is more complex, might need to check ActiveWindow.Selection
             // For simplicity, we might skip selection context for now or get slide text
             contextText = 'No se pudo obtener contexto de selección en PowerPoint.';
        }

      } catch (selectErr) {
        // Si falla la selección, intentar obtener contexto del documento activo
        try {
           const activeDoc = officeApp.ActiveDocument || officeApp.ActivePresentation || officeApp.ActiveWorkbook; // Propiedad varía por aplicación
           if (activeDoc) {
               if (application === 'Word.Application' && activeDoc.Content) {
                   contextText = activeDoc.Content.Text;
                   // Limitar el tamaño del contexto para evitar prompts demasiado largos
                   if (contextText.length > 1000) {
                       contextText = contextText.substring(0, 1000) + '...';
                   }
               } else if (application === 'Excel.Application' && activeDoc.ActiveSheet) {
                   // Podríamos intentar obtener datos de la hoja activa o un rango específico si contextRange está definido
                   // Por ahora, solo indicamos que no se pudo obtener contexto específico
                   contextText = `Documento activo: ${activeDoc.Name}. No se pudo obtener contexto detallado.`;
               } else if (application === 'PowerPoint.Application' && activeDoc.Slides) {
                   // Podríamos intentar obtener texto de la diapositiva actual
                   contextText = `Presentación activa: ${activeDoc.Name}. No se pudo obtener contexto detallado.`;
               } else {
                   contextText = `No se pudo obtener contexto detallado del documento activo en ${application}.`;
               }
           } else {
               contextText = `No se pudo obtener contexto del documento activo en ${application}.`;
           }
        } catch (docErr) {
            contextText = `Error al intentar obtener contexto del documento o selección en ${application}.`;
            console.error("Error al obtener contexto:", docErr);
        }
      }


      let prompt = '';
      switch (operation) {
        case 'format':
          prompt = `Basado en el siguiente texto o contexto de un documento de ${application}, sugiere opciones de formato (estilos, negrita, cursiva, alineación, etc.) que serían apropiadas. Contexto: "${contextText}"`;
          break;
        case 'search':
          prompt = `Basado en el siguiente texto o contexto de un documento de ${application}, sugiere términos de búsqueda relevantes o posibles ubicaciones dentro del documento para encontrar información relacionada. Contexto: "${contextText}"`;
          break;
        case 'chart':
          if (application !== 'Excel.Application') { // Usar nombre completo
              throw new Error(`La operación 'chart' solo es soportada para Excel. Aplicación actual: ${application}`);
          }
          prompt = `Basado en el siguiente contexto de una hoja de cálculo de Excel, sugiere tipos de gráficos apropiados y posibles rangos de datos para visualizar. Contexto: "${contextText}"`;
          break;
        default:
          throw new Error(`Operación no soportada: ${operation}`);
      }

      // Generar sugerencia usando FastMCP sampling
      const samplingResult = await requestSampling({
          prompt: prompt, // Usar el prompt generado
          maxTokens: 500, // Limitar la longitud de la sugerencia (ajustar según necesidad)
          // Otros parámetros de sampling pueden ser añadidos aquí si son relevantes
      });

      // Procesar el resultado del sampling
      let suggestion = '';
      if (samplingResult && samplingResult.content && samplingResult.content.length > 0) {
          // Asumir que el primer bloque de contenido de texto es la sugerencia
          // Usar type assertion para acceder a 'type' y 'text' debido a error de tipo
          const textContent = samplingResult.content.find((c: any) => c.type === 'text') as TextContent | undefined;
          if (textContent) {
              suggestion = textContent.text;
          }
      }

      if (!suggestion) {
          // Si no se obtuvo texto del sampling, lanzar un error o devolver un mensaje por defecto
          throw new Error('FastMCP sampling no generó ninguna sugerencia de texto.');
      }

      // Devolver una respuesta de éxito con la sugerencia
      return { success: true, data: suggestion };

    } catch (error: any) {
      console.error(`Error en la herramienta office/ai-suggest: ${error.message}`);
      // Devolver una respuesta de error
      return { success: false, error: { code: 'AI_SUGGEST_ERROR', message: `Error al procesar la sugerencia de IA: ${error.message}` } };
    }
  },
};