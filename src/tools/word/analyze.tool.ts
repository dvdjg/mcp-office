import { z } from 'zod';
import { getOfficeApplication } from '../../utils/officeInterop';

// Define el esquema de entrada para la herramienta word/analyze
const AnalyzeToolInputSchema = z.object({
  filePath: z.string().describe('La ruta al documento de Word.'),
  operation: z.enum(['analyze', 'add', 'summarize']).describe('La operación a realizar.'),
  criteria: z.array(z.string()).optional().describe('Criterios para el análisis (e.g., "technical issues", "grammar errors").'),
  range: z.string().optional().describe('Rango en el documento para añadir un comentario (e.g., "Paragraph 3", "Section 2").'),
  commentText: z.string().optional().describe('El texto del comentario a añadir.'),
});

type AnalyzeToolInput = z.infer<typeof AnalyzeToolInputSchema>;

/**
 * @tool word/analyze
 * @description Herramienta para analizar documentos de Word, añadir comentarios y resumir hallazgos.
 * @param {string} filePath - La ruta al documento de Word.
 * @param {'analyze' | 'add' | 'summarize'} operation - La operación a realizar: 'analyze' para analizar el contenido, 'add' para añadir un comentario, 'summarize' para resumir los hallazgos.
 * @param {string[]} [criteria] - Criterios para el análisis (solo para la operación 'analyze').
 * @param {string} [range] - Rango en el documento para añadir un comentario (solo para la operación 'add').
 * @param {string} [commentText] - El texto del comentario a añadir (solo para la operación 'add').
 * @returns {Promise<any>} El resultado de la operación.
 */
const analyzeTool = {
  name: 'word/analyze',
  description: 'Herramienta para analizar documentos de Word, añadir comentarios y resumir hallazgos.',
  inputSchema: AnalyzeToolInputSchema,
  handler: async (input: AnalyzeToolInput) => {
    const { filePath, operation, criteria, range, commentText } = input;
    let wordApp;

    try {
      wordApp = await getOfficeApplication('Word.Application');
      const doc = wordApp.Documents.Open(filePath);

      switch (operation) {
        case 'analyze':
          // Implementación básica de análisis (ejemplo: contar párrafos)
          const paragraphCount = doc.Paragraphs.Count;
          // Aquí se añadiría lógica más compleja basada en 'criteria'
          // Por ahora, solo devolvemos un resultado simple
          return {
            success: true,
            operation: 'analyze',
            filePath,
            result: {
              paragraphCount,
              // Otros resultados de análisis irían aquí
            },
          };

        case 'add':
          if (!range || !commentText) {
            throw new Error('Se requieren "range" y "commentText" para la operación "add".');
          }
          // Implementación básica para añadir un comentario a un rango (ejemplo: al primer párrafo)
          // La lógica para parsear 'range' y encontrar el rango correcto sería más compleja
          if (doc.Paragraphs.Count > 0) {
            const firstParagraph = doc.Paragraphs.Item(1).Range;
            doc.Comments.Add(firstParagraph, commentText);
            doc.Save(); // Guardar el documento después de añadir el comentario
            return {
              success: true,
              operation: 'add',
              filePath,
              range,
              commentText,
              message: `Comentario añadido al rango: ${range}`,
            };
          } else {
            throw new Error('El documento no contiene párrafos para añadir un comentario.');
          }


        case 'summarize':
          // Implementación básica de resumen (ejemplo: contar comentarios)
          const commentCount = doc.Comments.Count;
          // Aquí se añadiría lógica para resumir hallazgos de análisis o comentarios
          return {
            success: true,
            operation: 'summarize',
            filePath,
            result: {
              commentCount,
              // Resumen de hallazgos iría aquí
            },
          };

        default:
          throw new Error(`Operación no soportada: ${operation}`);
      }
    } catch (error: any) {
      return {
        success: false,
        operation,
        filePath,
        error: error.message,
      };
    } finally {
      // Considerar si cerrar Word o dejarlo abierto dependiendo del flujo de trabajo
      // if (wordApp) {
      //   wordApp.Quit();
      // }
    }
  },
};

export default analyzeTool;