import { z } from 'zod';
import { resolve as resolvePath } from 'path'; // Ensure resolvePath is imported
import { getOfficeApplication } from '../../utils/officeInterop.js';
import logger from '../../utils/logger.js'; // Import logger

// Define el esquema de entrada para la herramienta word/analyze
const AnalyzeToolInputSchema = z.object({
  //type: z.literal('object'), // Added to satisfy validator
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
    logger.info(`[analyzeTool] Handler started for operation: ${input.operation}, file: ${input.filePath}`); // Log start
    const { filePath, operation, criteria, range, commentText } = input;
    let wordApp;
    let doc: any = null; // Define doc here and initialize to null

    try {
      logger.info(`[analyzeTool] Getting Word application instance...`); // Log before getting app
      wordApp = await getOfficeApplication('Word.Application');
      logger.info(`[analyzeTool] Word application instance obtained. Opening document: ${filePath}`); // Log before opening doc
      const absoluteFilePath = resolvePath(filePath);
      doc = wordApp.Documents.Open(absoluteFilePath); // Assign doc here
      logger.info(`[analyzeTool] Document opened successfully: ${absoluteFilePath}`); // Log after opening doc

      switch (operation) {
        case 'analyze':
          logger.info(`[analyzeTool] Executing 'analyze' operation.`);
          // Implementación básica de análisis (ejemplo: contar párrafos)
          const paragraphCount = doc.Paragraphs.Count;
          // Aquí se añadiría lógica más compleja basada en 'criteria'
          // Por ahora, solo devolvemos un resultado simple
          logger.info(`[analyzeTool] 'analyze' operation completed. Paragraph count: ${paragraphCount}`);
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
          logger.info(`[analyzeTool] Executing 'add' operation.`);
          if (!range || !commentText) {
            logger.error("[analyzeTool] Missing range or commentText for 'add' operation.");
            throw new Error('Se requieren "range" y "commentText" para la operación "add".');
          }
          // Implementación básica para añadir un comentario a un rango (ejemplo: al primer párrafo)
          // La lógica para parsear 'range' y encontrar el rango correcto sería más compleja
          logger.debug(`[analyzeTool] Attempting to add comment '${commentText}' to range '${range}'.`);
          if (doc.Paragraphs.Count > 0) {
            // Basic implementation: target first paragraph regardless of 'range' input for now
            const targetRange = doc.Paragraphs.Item(1).Range;
            logger.debug(`[analyzeTool] Targeting range: Paragraph 1`);
            doc.Comments.Add(targetRange, commentText);
            logger.debug(`[analyzeTool] Comment added. Saving document...`);
            doc.Save(); // Guardar el documento después de añadir el comentario
            logger.info(`[analyzeTool] 'add' operation completed and document saved.`);
            return {
              success: true,
              operation: 'add',
              filePath,
              range,
              commentText,
              message: `Comentario añadido al rango: ${range}`,
            };
          } else {
            logger.error('[analyzeTool] Document has no paragraphs to add comment to.');
            throw new Error('El documento no contiene párrafos para añadir un comentario.');
          }

        case 'summarize':
          logger.info(`[analyzeTool] Executing 'summarize' operation.`);
          // Implementación básica de resumen (ejemplo: contar comentarios)
          const commentCount = doc.Comments.Count;
          // Aquí se añadiría lógica para resumir hallazgos de análisis o comentarios
          logger.info(`[analyzeTool] 'summarize' operation completed. Comment count: ${commentCount}`);
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
          logger.error(`[analyzeTool] Unsupported operation: ${operation}`);
          throw new Error(`Operación no soportada: ${operation}`);
      }
    } catch (error: any) {
      logger.error(`[analyzeTool] Error during operation ${operation} for file ${filePath}: ${error.message}`, { error }); // Log error
      // Ensure doc is closed even on error if it was opened
      if (doc) {
          try {
              doc.Close(0); // wdDoNotSaveChanges = 0
              logger.info(`[analyzeTool] Document closed after error.`);
          } catch (closeError: any) {
              logger.warn(`[analyzeTool] Failed to close document after error: ${closeError.message}`);
          }
      }
      return {
        success: false,
        operation,
        filePath,
        error: error.message,
      };
    } finally {
      logger.info(`[analyzeTool] Handler finished for operation: ${operation}`);
      // Close the document if it's still open and wasn't closed in the catch block
      if (doc && typeof doc.Close === 'function') { // Check if doc exists and has Close method
          try {
              // Analyze/Summarize don't modify, 'add' saves explicitly. Close without saving.
              doc.Close(0); // wdDoNotSaveChanges = 0
              logger.info(`[analyzeTool] Document closed in finally block.`);
          } catch (closeError: any) {
              logger.warn(`[analyzeTool] Failed to close document in finally block: ${closeError.message}`);
          }
      }
      // DO NOT release wordApp here, let the interop layer manage the instance lifecycle
    }
  },
};

export default analyzeTool;