import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types'; // Importar tipos relevantes del proyecto
import logger from '../../utils/logger'; // Importar logger como default
import { handleToolError } from '../../utils/errorHandler'; // Importar handleToolError como nombrada
import { allRegisteredTools } from '../index'; // Importar el array de herramientas registradas

// Definir el esquema para un solo paso del flujo de trabajo
const WorkflowStepSchema = z.object({
  tool: z.string().describe('Nombre de la herramienta a ejecutar (e.g., "word/merge")'),
  operation: z.string().describe('Nombre de la operación de la herramienta (e.g., "run")'),
  params: z.record(z.any()).optional().describe('Parámetros para la operación de la herramienta'),
});

// Definir el esquema para la definición completa del flujo de trabajo
const WorkflowDefinitionSchema = z.array(WorkflowStepSchema).describe('Lista de pasos a ejecutar en el flujo de trabajo');

// Definir el esquema de entrada para la operación 'run'
const RunOperationInputSchema = z.object({
  workflow: WorkflowDefinitionSchema.describe('Definición del flujo de trabajo a ejecutar'),
});

// Definir el esquema de salida para la operación 'run'
const RunOperationOutputSchema = z.object({
  results: z.array(z.object({
    step: WorkflowStepSchema,
    success: z.boolean(),
    output: z.any().optional(),
    error: z.string().optional(),
  })).describe('Resultados de la ejecución de cada paso'),
});

/**
 * @tool office/workflow
 * @description Permite definir y ejecutar flujos de trabajo multi-paso que involucren múltiples herramientas de Office o pasos.
 */
export class OfficeWorkflowTool implements McpResource {
  path = 'office/workflow'; // Usar 'path' en lugar de 'name'
  description = 'Allows defining and executing multi-step workflows involving multiple Office tools or steps.';
  schema = RunOperationInputSchema; // El esquema de entrada para la herramienta completa

  // El handler principal para la herramienta workflow
  handler = this.runWorkflow.bind(this);

  // No hay 'operations' en la interfaz McpResource, la lógica de 'run' está en el handler principal.
  // Si necesitáramos múltiples operaciones para esta herramienta, tendríamos que adaptar la interfaz McpResource
  // o definir sub-rutas (e.g., 'office/workflow/run', 'office/workflow/trigger').
  // Por ahora, nos enfocamos solo en la operación 'run' a través del handler principal.

  async runWorkflow(input: ToolRequestParams): Promise<ApiResponse<z.infer<typeof RunOperationOutputSchema>>> {
    try {
      // Validar la entrada usando el esquema de la herramienta
      const validatedInput = RunOperationInputSchema.parse(input);
      const { workflow } = validatedInput;
      const results: z.infer<typeof RunOperationOutputSchema>['results'] = [];

      logger.info(`Ejecutando flujo de trabajo con ${workflow.length} pasos.`);

      for (const step of workflow) {
        logger.info(`Ejecutando paso: ${step.tool}/${step.operation}`);
        let stepResult: any = { step, success: false };

        try {
          // Buscar la herramienta por su 'path' en el array de herramientas registradas
          // No necesitamos verificar 'operations' o usar type assertion a McpTool
          const tool = allRegisteredTools.find(t => t.path === step.tool);

          if (!tool) {
            throw new Error(`Herramienta no encontrada: ${step.tool}`);
          }

          // Para McpResource, no hay 'operations'. La lógica de la operación debe estar en el handler
          // o la herramienta llamada debe manejar la operación internamente basada en los parámetros.
          // Asumimos que el 'tool' en el paso se refiere al 'path' de una McpResource
          // y que los 'params' incluyen la información de la operación si la herramienta llamada la necesita.
          // Por ejemplo, un paso podría ser: { tool: 'word/text', operation: 'find', params: { text: 'buscar' } }
          // El handler de 'word/text' debería ser capaz de interpretar el parámetro 'operation'.

          // Crear un objeto de contexto básico para pasar al handler de la herramienta llamada
          const basicContext = {
            log: logger, // Usar el logger existente
            reportProgress: async (progress: { progress: number; total?: number }) => { // Aceptar total opcional y hacerla async
              logger.info(`Progreso del paso ${step.tool}: ${progress.progress}${progress.total !== undefined ? '/' + progress.total : ''}`);
              // Aquí podríamos añadir lógica para reportar progreso del workflow si fuera necesario
            },
            session: { // Objeto session placeholder
              id: 'workflow-session',
              user: 'workflow-executor',
              // Añadir otras propiedades de sesión si son relevantes
            },
            // Añadir otras propiedades del contexto si son necesarias para las herramientas llamadas
          };

          // Llamar al handler de la herramienta encontrada
          // Pasamos los parámetros del paso y el contexto básico
          // El handler de la herramienta llamada debe manejar la validación de sus propios parámetros.
          const output = await tool.handler(step.params || {}, basicContext);

          // El handler de la herramienta debe devolver un ApiResponse
          if (output.success) {
            stepResult.success = true;
            stepResult.output = output.data; // Usar 'data' de ApiResponse
            logger.info(`Paso ${step.tool} completado con éxito.`);
          } else {
            // Si el handler de la herramienta devuelve un ErrorResponse
            stepResult.success = false;
            stepResult.error = output.error?.message || 'Error desconocido en la herramienta llamada';
            logger.error(`Error en la herramienta llamada ${step.tool}: ${stepResult.error}`);
            // Decidimos detener el flujo de trabajo si un paso falla
            results.push(stepResult);
            logger.error('Flujo de trabajo detenido debido a un error en un paso.');
            // Devolver los resultados hasta el punto del error
            return { success: true, data: { results } }; // Devolver éxito para el workflow, pero con resultados parciales y error
          }

        } catch (error: any) {
          stepResult.success = false;
          stepResult.error = error.message || 'Error desconocido al ejecutar el paso';
          logger.error(`Error al ejecutar paso ${step.tool}: ${stepResult.error}`);
          // Decidimos detener el flujo de trabajo si un paso falla
          results.push(stepResult);
          logger.error('Flujo de trabajo detenido debido a un error.');
          // Usar handleToolError para formatear el error antes de devolverlo
          const errorResponse = handleToolError(error);
          // Devolver los resultados hasta el punto del error, formateando el error
          return { success: true, data: { results: results.map(r => ({ ...r, error: r.error || (r.success ? undefined : errorResponse.error?.message) })) } }; // Devolver éxito para el workflow, pero con resultados parciales y error formateado
        }

        results.push(stepResult);
      }

      logger.info('Flujo de trabajo completado.');
      return { success: true, data: { results } }; // Devolver éxito para el workflow con todos los resultados
    } catch (error: any) {
      // Capturar errores de validación de Zod o errores generales del handler del workflow
      logger.error(`Error en el handler principal de office/workflow: ${error.message || error}`);
      return handleToolError(error, 'WORKFLOW_EXECUTION_ERROR'); // Usar handleToolError para formatear el error
    }
  }
}