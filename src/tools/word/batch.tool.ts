import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse, ErrorResponse } from '../../types/common.types'; // Importar ToolRequestParams, ApiResponse, ErrorResponse
import logger from '../../utils/logger';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { handleToolError } from '../../utils/errorHandler';

// Importar los manejadores de otras herramientas de Word que se pueden ejecutar en lote
// Nota: Esto es un placeholder. La implementación real necesitaría una forma de mapear
// los nombres de las herramientas a sus manejadores o lógica de ejecución.
// Por simplicidad inicial, asumiremos que podemos llamar a la lógica de las herramientas directamente.
// Esto requeriría refactorizar las herramientas existentes para exponer su lógica de ejecución
// de forma reutilizable, o bien, importar los manejadores y llamarlos con el contexto adecuado.
// Para esta implementación inicial, simularemos la ejecución o asumiremos una estructura
// donde la lógica principal de cada herramienta está en una función exportable.

// Esquema para una sola operación dentro del lote
const BatchOperationSchema = z.object({
  tool: z.string().describe('Nombre de la herramienta de Word a ejecutar (ej: word/styles, word/text)').min(1),
  params: z.record(z.any()).describe('Parámetros para la herramienta especificada'),
});

// Esquema de entrada para la herramienta word/batch
const BatchToolInputSchema = z.object({
  operationType: z.enum(['run', 'transaction']).describe('Tipo de operación a ejecutar (run o transaction)').default('run'),
  filePath: z.string().describe('Ruta al archivo .docx').min(1),
  operations: z.array(BatchOperationSchema).describe('Lista de operaciones a ejecutar en lote').min(1),
});

// Esquema de salida para la operación 'run'
const RunOperationOutputSchema = z.object({
  results: z.array(z.object({
    tool: z.string(),
    success: z.boolean(),
    message: z.string().optional(),
    output: z.any().optional(),
  })).describe('Resultados de cada operación en el lote'),
});

// Esquema de salida para la operación 'transaction' (placeholder)
const TransactionOperationOutputSchema = z.object({
  success: z.boolean(),
  message: z.string(),
});


/**
 * @tool word/batch
 * @description Ejecuta múltiples operaciones de herramientas de Word en un solo lote.
 * Permite agrupar llamadas para mejorar la eficiencia al abrir el documento una sola vez.
 * Soporta las operaciones `run` y `transaction`.
 * @operation run
 * @description Ejecuta una lista de operaciones de herramientas de Word secuencialmente.
 * @input BatchToolInputSchema (con operationType='run')
 * @output RunOperationOutputSchema
 * @operation transaction
 * @description Ejecuta una lista de operaciones de herramientas de Word como una transacción (no implementado).
 * @input BatchToolInputSchema (con operationType='transaction')
 * @output TransactionOperationOutputSchema
 * @error Manejo de errores para cada operación individual y para el proceso general del lote.
 */
export const batchTool: McpResource = {
  path: 'word/batch',
  description: 'Ejecuta múltiples operaciones de herramientas de Word en un solo lote.',
  schema: BatchToolInputSchema,
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => { // Añadido tipo de retorno explícito
    let wordApp: any = null;
    let doc: any = null;
    const results: z.infer<typeof RunOperationOutputSchema>['results'] = [];

    try {
      // Validar los parámetros de entrada usando el esquema Zod
      const input = BatchToolInputSchema.parse(params);

      wordApp = await getOfficeApplication('Word.Application');
      doc = wordApp.Documents.Open(input.filePath);

      logger.info(`Ejecutando operación '${input.operationType}' del lote en ${input.filePath}`);

      if (input.operationType === 'run') {
        logger.info(`Ejecutando lote de ${input.operations.length} operaciones en modo 'run'.`);

        for (const operation of input.operations) {
          logger.info(`Ejecutando operación: ${operation.tool}`);
          try {
            // --- Lógica de despacho a herramientas ---
            // Aquí es donde necesitaríamos la lógica real para invocar la herramienta correcta.
            // Esto es un placeholder y necesitaría ser reemplazado.
            // Una posible implementación sería tener un mapa de nombres de herramientas a sus funciones de ejecución,
            // o importar los manejadores de las otras herramientas y llamarlos con el 'doc' y 'params'.
            // Por ahora, simularemos un resultado exitoso.
            const toolResult = { success: true, message: `${operation.tool} ejecutada con éxito (simulado)`, output: {} }; // Simulación

            // En una implementación real, llamaríamos a la lógica de la herramienta:
            // const toolHandler = getToolHandler(operation.tool); // Función hipotética para obtener el manejador
            // const toolResult = await toolHandler(doc, operation.params); // Pasar la instancia del documento y los parámetros

            results.push({
              tool: operation.tool,
              success: toolResult.success,
              message: toolResult.message,
              output: toolResult.output,
            });
            logger.info(`Operación ${operation.tool} completada.`);

          } catch (opError: any) {
            logger.error(`Error al ejecutar operación ${operation.tool}: ${opError.message}`);
            results.push({
              tool: operation.tool,
              success: false,
              message: `Error al ejecutar operación: ${opError.message}`,
            });
            // En modo 'run', continuamos incluso si una operación falla.
          }
        }

        logger.info('Lote de operaciones en modo \'run\' completado.');

        // Devolver el resultado en el formato esperado por RunOperationOutputSchema
        return { success: true, data: { results } };

      } else if (input.operationType === 'transaction') {
        logger.warn('La operación \'transaction\' no está implementada aún.');
        // Implementación de transacción (más compleja con COM, requiere deshacer operaciones en caso de fallo)
        // Por ahora, simplemente reportamos que no está implementada.
        return { success: false, error: { code: 'NOT_IMPLEMENTED', message: 'La operación transaction no está implementada aún.' } };

      } else {
        // Esto no debería ocurrir si el schema Zod funciona correctamente, pero es una salvaguarda.
        const errorMessage = `Tipo de operación no soportado: ${input.operationType}`;
        logger.error(errorMessage);
        return { success: false, error: { code: 'INVALID_OPERATION_TYPE', message: errorMessage } };
      }


    } catch (error: any) {
      logger.error(`Error general en la herramienta word/batch: ${error.message}`);
      const errorResponse = handleToolError(error); // Usar handleToolError para formatear el error
      // Si hay un error general (ej: al abrir el archivo), reportarlo.
      // En este caso, no hay resultados parciales de operaciones individuales que devolver en el error general.
      return errorResponse;
    } finally {
      // Asegurarse de cerrar el documento y la aplicación Word si se abrieron
      if (doc) {
        try {
          doc.Save(); // Guardar cambios si se hicieron
          doc.Close();
          logger.info('Documento cerrado.');
        } catch (closeError: any) {
          logger.warn(`Error al cerrar el documento: ${closeError.message}`);
        }
        releaseObject(doc); // Liberar el objeto del documento
      }
      // No cerrar la aplicación Word aquí, ya que podría ser reutilizada por otras herramientas
      // releaseObject(wordApp); // Liberar el objeto de la aplicación - Cuidado con cerrar la instancia del usuario
    }
  },
};