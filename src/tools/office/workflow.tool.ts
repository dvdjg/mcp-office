/**
 * @file Tool for defining and executing multi-step workflows involving multiple Office tools or steps.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types.js'; // Import relevant types from the project
import logger from '../../utils/logger.js'; // Import logger as default
import { handleToolError } from '../../utils/errorHandler.js'; // Import handleToolError as named
import { allRegisteredTools } from '../index.js'; // Import the array of registered tools

// Define the schema for a single workflow step
const WorkflowStepSchema = z.object({
  tool: z.string().describe('Name of the tool to execute (e.g., "word/merge")'),
  operation: z.string().describe('Name of the tool operation (e.g., "run")'),
  params: z.record(z.any()).optional().describe('Parameters for the tool operation'),
});

// Define the schema for the complete workflow definition
const WorkflowDefinitionSchema = z.array(WorkflowStepSchema).describe('List of steps to execute in the workflow');

// Define the input schema for the 'run' operation
const RunOperationInputSchema = z.object({
  workflow: WorkflowDefinitionSchema.describe('Definition of the workflow to execute'),
});

// Define the output schema for the 'run' operation
const RunOperationOutputSchema = z.object({
  results: z.array(z.object({
    step: WorkflowStepSchema,
    success: z.boolean(),
    output: z.any().optional(),
    error: z.string().optional(),
  })).describe('Results of each step execution'),
});

/**
 * @tool office/workflow
 * @description Allows defining and executing multi-step workflows involving multiple Office tools or steps.
 */
export class OfficeWorkflowTool implements McpResource {
  path = 'office/workflow'; // Use 'path' instead of 'name'
  description = 'Allows defining and executing multi-step workflows involving multiple Office tools or steps.';
  schema = RunOperationInputSchema; // The input schema for the complete tool

  // The main handler for the workflow tool
  handler = this.runWorkflow.bind(this);

  // There are no 'operations' in the McpResource interface, the 'run' logic is in the main handler.
  // If we needed multiple operations for this tool, we would have to adapt the McpResource interface
  // or define sub-paths (e.g., 'office/workflow/run', 'office/workflow/trigger').
  // For now, we focus only on the 'run' operation through the main handler.

  async runWorkflow(input: ToolRequestParams): Promise<ApiResponse<z.infer<typeof RunOperationOutputSchema>>> {
    try {
      // Validate input using the tool schema
      const validatedInput = RunOperationInputSchema.parse(input);
      const { workflow } = validatedInput;
      const results: z.infer<typeof RunOperationOutputSchema>['results'] = [];

      logger.info(`Executing workflow with ${workflow.length} steps.`);

      for (const step of workflow) {
        logger.info(`Executing step: ${step.tool}/${step.operation}`);
        let stepResult: any = { step, success: false };

        try {
          // Find the tool by its 'path' in the array of registered tools
          // We don't need to check 'operations' or use type assertion to McpTool
          const tool = allRegisteredTools.find(t => t.path === step.tool);

          if (!tool) {
            throw new Error(`Tool not found: ${step.tool}`);
          }

          // For McpResource, there are no 'operations'. The operation logic must be in the handler
          // or the called tool must handle the operation internally based on the parameters.
          // We assume that the 'tool' in the step refers to the 'path' of an McpResource
          // and that the 'params' include the operation information if the called tool needs it.
          // For example, a step could be: { tool: 'word/text', operation: 'find', params: { text: 'search' } }
          // The handler of 'word/text' should be able to interpret the 'operation' parameter.

          // Create a basic context object to pass to the called tool's handler
          const basicContext = {
            log: logger, // Use the existing logger
            reportProgress: async (progress: { progress: number; total?: number }) => { // Accept optional total and make it async
              logger.info(`Step progress ${step.tool}: ${progress.progress}${progress.total !== undefined ? '/' + progress.total : ''}`);
              // Here we could add logic to report workflow progress if needed
            },
            session: { // Placeholder session object
              id: 'workflow-session',
              user: 'workflow-executor',
              // Add other session properties if relevant
            },
            // Add other context properties if necessary for the called tools
          };

          // Call the handler of the found tool
          // We pass the step parameters and the basic context
          // The called tool's handler should handle its own parameter validation.
          const output = await tool.handler(step.params || {}, basicContext);

          // The tool handler must return an ApiResponse
          if (output.success) {
            stepResult.success = true;
            stepResult.output = output.data; // Use 'data' from ApiResponse
            logger.info(`Step ${step.tool} completed successfully.`);
          } else {
            // If the tool handler returns an ErrorResponse
            stepResult.success = false;
            stepResult.error = output.error?.message || 'Unknown error in called tool';
            logger.error(`Error in called tool ${step.tool}: ${stepResult.error}`);
            // We decide to stop the workflow if a step fails
            results.push(stepResult);
            logger.error('Workflow stopped due to an error in a step.');
            // Return results up to the point of error
            return { success: true, data: { results } }; // Return success for the workflow, but with partial results and error
          }

        } catch (error: any) {
          stepResult.success = false;
          stepResult.error = error.message || 'Unknown error executing step';
          logger.error(`Error executing step ${step.tool}: ${stepResult.error}`);
          // We decide to stop the workflow if a step fails
          results.push(stepResult);
          logger.error('Workflow stopped due to an error.');
          // Use handleToolError to format the error before returning it
          const errorResponse = handleToolError(error);
          // Return results up to the point of error, formatting the error
          return { success: true, data: { results: results.map(r => ({ ...r, error: r.error || (r.success ? undefined : errorResponse.error?.message) })) } }; // Return success for the workflow, but with partial results and formatted error
        }

        results.push(stepResult);
      }

      logger.info('Workflow completed.');
      return { success: true, data: { results } }; // Return success for the workflow with all results
    } catch (error: any) {
      // Catch Zod validation errors or general workflow handler errors
      logger.error(`Error in main office/workflow handler: ${error.message || error}`);
      return handleToolError(error, 'WORKFLOW_EXECUTION_ERROR'); // Use handleToolError to format the error
    }
  }
}