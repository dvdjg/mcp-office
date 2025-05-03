/**
 * @file Implements the 'word/batch' tool for executing multiple operations on a Word document in a single batch.
 * This tool aims to improve efficiency by opening the document only once for multiple operations.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse, ErrorResponse } from '../../types/common.types'; // Import ToolRequestParams, ApiResponse, ErrorResponse
import logger from '../../utils/logger';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { handleToolError } from '../../utils/errorHandler';

// Import handlers for other Word tools that can be executed in a batch.
// Note: This is a placeholder. The actual implementation would need a way to map
// tool names to their execution logic.
// For initial simplicity, we will simulate execution or assume a structure
// where the main logic of each tool is in an exportable function.

/** Schema for a single operation within the batch. */
const BatchOperationSchema = z.object({
  /** The name of the Word tool to execute (e.g., 'word/styles', 'word/text'). */
  tool: z.string().describe('Name of the Word tool to execute (e.g., word/styles, word/text)').min(1),
  /** Parameters for the specified tool. Can be any key-value pair. */
  params: z.record(z.any()).describe('Parameters for the specified tool'),
});

/** Input schema for the 'word/batch' tool. */
const BatchToolInputSchema = z.object({
  /** The type of batch operation to execute ('run' or 'transaction'). 'transaction' is not yet implemented. */
  operationType: z.enum(['run', 'transaction']).describe('Type of operation to execute (run or transaction)').default('run'),
  /** The path to the .docx file (relative to the current workspace directory). */
  filePath: z.string().describe('Path to the .docx file').min(1),
  /** A list of operations to execute in the batch. */
  operations: z.array(BatchOperationSchema).describe('List of operations to execute in batch').min(1),
});

/** Output schema for the 'run' operation. */
const RunOperationOutputSchema = z.object({
  /** Results of each operation in the batch. */
  results: z.array(z.object({
    /** The name of the tool that was executed. */
    tool: z.string(),
    /** Indicates if the individual operation was successful. */
    success: z.boolean(),
    /** An optional message related to the operation's outcome. */
    message: z.string().optional(),
    /** Optional output data from the operation. Can be any type. */
    output: z.any().optional(),
  })).describe('Results of each operation in the batch'),
});

/** Output schema for the 'transaction' operation (placeholder). */
const TransactionOperationOutputSchema = z.object({
  /** Indicates if the transaction was successful. */
  success: z.boolean(),
  /** A message describing the outcome of the transaction. */
  message: z.string(),
});


/**
 * McpResource definition for the 'word/batch' tool.
 * Allows executing multiple Word tool operations in a single batch.
 */
export const batchTool: McpResource = {
  path: 'word/batch',
  description: 'Executes multiple Word tool operations in a single batch.',
  schema: BatchToolInputSchema,
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => { // Added explicit return type
    let wordApp: any = null;
    let doc: any = null;
    const results: z.infer<typeof RunOperationOutputSchema>['results'] = [];

    try {
      // Validate input parameters using the Zod schema
      const input = BatchToolInputSchema.parse(params);

      wordApp = await getOfficeApplication('Word.Application');
      doc = wordApp.Documents.Open(input.filePath);

      logger.info(`[word/batch] Executing batch operation '${input.operationType}' on ${input.filePath}`);

      if (input.operationType === 'run') {
        logger.info(`[word/batch] Executing batch of ${input.operations.length} operations in 'run' mode.`);

        for (const operation of input.operations) {
          logger.info(`[word/batch] Executing operation: ${operation.tool} with parameters: ${JSON.stringify(operation.params)}`);
          let toolResult: { success: boolean; message?: string; output?: any; };
          try {
            // --- Tool Dispatch Logic ---
            // This is where the actual logic of the specified tool would be invoked.
            // Since we cannot modify other files to refactor handlers
            // to accept the 'doc' instance directly, we will simulate the call
            // to a hypothetical function that would handle dispatch and execution.
            // In a real implementation, we would need a map or registry of tools
            // and their execution functions that can operate with an open document instance.

            // Simulation of calling the individual tool logic
            toolResult = await executeSingleOperation(doc, operation);

            results.push({
              tool: operation.tool,
              success: toolResult.success,
              message: toolResult.message,
              output: toolResult.output,
            });
            logger.info(`[word/batch] Operation ${operation.tool} completed successfully: ${toolResult.success}.`);

          } catch (opError: any) {
            logger.error(`[word/batch] Error executing operation ${operation.tool}: ${opError.message}`);
            results.push({
              tool: operation.tool,
              success: false,
              message: `Error executing operation: ${opError.message}`,
            });
            // In 'run' mode, we continue even if an operation fails.
          }
        }

        logger.info('[word/batch] Batch operations in \'run\' mode completed.');

        // Return the result in the format expected by RunOperationOutputSchema
        return { success: true, data: { results } };

      } else if (input.operationType === 'transaction') {
        logger.warn('[word/batch] \'transaction\' operation is not implemented yet.');
        // Transaction implementation (more complex with COM, requires undoing operations on failure)
        // For now, we simply report that it's not implemented.
        return { success: false, error: { code: 'NOT_IMPLEMENTED', message: 'The transaction operation is not implemented yet.' } };

      } else {
        // This should not occur if the Zod schema works correctly, but it's a safeguard.
        const errorMessage = `Unsupported operation type: ${input.operationType}`;
        logger.error(`[word/batch] ${errorMessage}`);
        return { success: false, error: { code: 'INVALID_OPERATION_TYPE', message: errorMessage } };
      }


    } catch (error: any) {
      logger.error(`[word/batch] General error in word/batch tool: ${error.message}`);
      const errorResponse = handleToolError(error); // Use handleToolError to format the error
      // If there's a general error (e.g., opening the file), report it.
      // In this case, there are no partial results from individual operations to return in the general error.
      return errorResponse;
    } finally {
      // Ensure the document and Word application are closed if they were opened
      if (doc) {
        try {
          // Do not save automatically, as some operations may not want to save
          // The save logic should be part of individual operations if needed,
          // or an explicit operation in the batch.
          // doc.Save();
          doc.Close(); // Close without saving
          logger.info('[word/batch] Document closed.');
        } catch (closeError: any) {
          logger.warn(`[word/batch] Error closing document: ${closeError.message}`);
        }
        releaseObject(doc); // Release the document object
      }
      // Do not close the Word application here, as it might be reused by other tools
      // releaseObject(wordApp); // Release the application object - Be careful not to close the user's instance
    }
  },
};

/**
 * Hypothetical function to execute a single operation within the batch.
 * This function simulates calling the logic of an individual tool.
 * In a real implementation, this would dispatch to the correct handler function
 * for the specified tool, passing it the 'doc' instance and the operation parameters.
 * @param doc - The Word document COM object.
 * @param operation - The operation details, validated against `BatchOperationSchema`.
 * @returns A promise resolving to an object indicating the success, message, and optional output of the operation.
 * @throws {Error} If the simulated operation fails.
 */
async function executeSingleOperation(doc: any, operation: z.infer<typeof BatchOperationSchema>): Promise<{ success: boolean; message?: string; output?: any; }> {
  logger.info(`[word/batch] [Simulation] Attempting to execute logic for tool: ${operation.tool}`);
  logger.debug(`[word/batch] [Simulation] Parameters received: ${JSON.stringify(operation.params)}`);

  // The actual logic to map operation.tool to the corresponding tool's execution function
  // and call it with 'doc' and 'operation.params' would go here.
  // For now, we simply simulate a result.

  // Simulation: 80% probability of success
  const success = Math.random() < 0.8;
  const message = success ? `${operation.tool} executed successfully (simulated)` : `Simulated error executing ${operation.tool}`;
  const output = success ? { simulatedOutput: `Result of ${operation.tool}` } : undefined;

  // Simulate a small delay
  await new Promise(resolve => setTimeout(resolve, 50));

  if (!success) {
    // Simulate an error by throwing an exception
    throw new Error(message);
  }

  return { success, message, output };
}