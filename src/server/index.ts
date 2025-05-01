// /src/server/index.ts
// =============================================================================
/**
 * @file Main entry point for the msoffice-mcp server.
 * Initializes FastMCP and registers all defined tools.
 */
import { FastMCP, UserError, UnexpectedStateError, Context as FastMCPContext, ContentResult, TextContent, ToolParameters, audioContent, imageContent, SerializableValue } from 'fastmcp'; // Import SerializableValue from fastmcp
import { ZodError } from 'zod';
import { StandardSchemaV1 } from '@standard-schema/spec'; // Import StandardSchemaV1
import allTools from '@/tools';
import logger from '@/utils/logger';
import { handleToolError, createErrorResponse } from '@/utils/errorHandler';
import { McpResource, ToolRequestParams, ApiResponse } from '@/types/common.types'; // Removed SerializableValue import from here
import { version as packageVersion, name as packageName } from '../../package.json'; // Import name and version

// Define a type alias for the FastMCP context specific to this server (no auth for now)
type ServerContext = FastMCPContext<undefined>;
// Define Extras type based on FastMCP definition if not directly importable
type Extras = Record<string, unknown>;

// --- FastMCP Server Configuration ---
const PORT = process.env.PORT || 3000;

// Instantiate FastMCP with name and version from package.json
const mcpServer = new FastMCP({
    name: packageName,
    version: packageVersion as `${number}.${number}.${number}`, // Assert type for version
    // instructions: "Optional instructions for the AI assistant on how to use this server",
});

// --- Register Tools ---
logger.info(`Registering ${allTools.length} MCP resources...`);

allTools.forEach((tool: McpResource) => {
    try {
        // Use addTool instead of registerResource
        mcpServer.addTool({
            name: tool.path, // Use the unique path as the tool name
            description: tool.description,
            // Use tool.schema as parameters if it exists and conforms to StandardSchemaV1
            parameters: tool.schema as unknown as ToolParameters | undefined, // Cast schema, ensure it's compatible
            // annotations: { title: tool.path }, // Optional: Add annotations like title
            // Adjust return type based on removed imports if necessary, FastMCP handles ContentResult union
            execute: async (args: StandardSchemaV1.InferOutput<any>, context: ServerContext): Promise<ContentResult | string | TextContent> => {
                const startTime = Date.now();
                // Use context.log provided by FastMCP
                context.log.info(`[${tool.path}] Request received`, { params: hideSensitiveParams(args) as SerializableValue }); // Ensure logged params are serializable

                try {
                    // Validation is typically handled by FastMCP based on 'parameters' schema
                    // Authorization checks can use context.session if authentication is implemented

                    // Execute the original tool handler
                    // Pass args and potentially context.session if needed by the original handler
                    const apiResponse: ApiResponse<any> = await tool.handler(args, undefined); // Pass args, context original era undefined

                    const duration = Date.now() - startTime;

                    // Transform the ApiResponse to the format expected by FastMCP
                    if (apiResponse.success) {
                        context.log.info(`[${tool.path}] Request successful`, { durationMs: duration });
                        const data = apiResponse.data;

                        // Map data to FastMCP return types
                        if (typeof data === 'string') {
                            // Return string directly or as TextContent
                            return data;
                            // return { type: 'text', text: data };
                        } else if (data && typeof data === 'object') {
                            // Attempt to return structured data if possible, otherwise stringify
                            // This might need refinement based on specific tool outputs
                            // For now, return as JSON string within TextContent
                            try {
                                return { type: 'text', text: JSON.stringify(data, null, 2) };
                            } catch (stringifyError) {
                                // Convert error to string for logging
                                context.log.error(`[${tool.path}] Error stringifying successful response data`, { error: String(stringifyError) });
                                throw new UnexpectedStateError("Failed to serialize successful response data.");
                            }
                        } else {
                            // Handle null, undefined, or other types
                            return { type: 'text', text: 'Operation completed successfully.' }; // Default success message
                        }
                    } else {
                        // Throw a UserError for FastMCP to handle client-side errors
                        // Convert error details to string for logging and UserError
                        const errorDetailsString = apiResponse.error?.details ? String(apiResponse.error.details) : undefined;
                        context.log.warn(`[${tool.path}] Request failed`, { durationMs: duration, code: apiResponse.error?.code, message: apiResponse.error?.message, details: errorDetailsString });
                        // Pass serializable details to UserError
                        const userErrorDetails = errorDetailsString ? { details: errorDetailsString } : undefined;
                        throw new UserError(apiResponse.error?.message || 'Tool execution failed.', userErrorDetails);
                    }

                } catch (error) {
                    const duration = Date.now() - startTime;
                     // Convert error to string for logging
                    context.log.error(`[${tool.path}] Unhandled error in tool execution wrapper`, { durationMs: duration, error: String(error) });

                    // Ensure errors are thrown in FastMCP's expected format
                    if (error instanceof UserError || error instanceof UnexpectedStateError) {
                        throw error; // Re-throw FastMCP specific errors
                    }

                    // Convert other errors to UserError using the existing handler logic
                    const apiErrorResponse = handleToolError(error); // Get standardized error response
                    // Convert details to string for UserError
                    const errorDetailsString = apiErrorResponse.error?.details ? String(apiErrorResponse.error.details) : undefined;
                    const userErrorDetails = errorDetailsString ? { details: errorDetailsString } : undefined;
                    throw new UserError(apiErrorResponse.error?.message || 'An unexpected error occurred during tool execution.', userErrorDetails);
                }
            },
            // Completions are handled differently in FastMCP (e.g., via Prompt/Resource arguments), remove from here.
        });
        logger.debug(`Registered tool: ${tool.path}`);
    } catch (error) {
         logger.error(`Failed to register tool: ${tool.path}`, { error });
    }
});

logger.info("All resources registered.");

// --- Helper to hide sensitive data from logs ---
function hideSensitiveParams(params: ToolRequestParams): ToolRequestParams {
    const sensitiveKeys = ['password', 'apiKey', 'secret', 'token', 'content']; // Add keys to hide
    const loggedParams = { ...params };
    for (const key of sensitiveKeys) {
        if (loggedParams[key]) {
             // Check if content is long binary data (Buffer)
             if (key === 'content' && Buffer.isBuffer(loggedParams[key])) {
                  loggedParams[key] = `<Buffer length=${loggedParams[key].length}>`;
             } else if (key === 'content' && typeof loggedParams[key] === 'string' && loggedParams[key].length > 100) {
                 loggedParams[key] = `<String length=${loggedParams[key].length}>`;
             } else if (key !== 'content') { // Avoid hiding short content strings unless explicitly needed
                 loggedParams[key] = '********';
             }
        }
    }
    return loggedParams;
}


// --- Start Server ---
// --- Prueba COM Interop (Opcional con winax) ---
import { getOfficeApplication, releaseObject } from '@/utils/officeInterop'; // Usar alias

async function testCom() {
  try {
    logger.info("Testing COM Interop (winax): Getting Word Application...");
    const wordApp = await getOfficeApplication('Word.Application');
    // Verificar que el objeto y la propiedad existen antes de acceder
    if (wordApp && typeof wordApp.Version !== 'undefined') {
        logger.info(`Got Word Application, Version: ${wordApp.Version}`);
    } else if (wordApp) {
        logger.warn("Got Word Application object, but couldn't retrieve Version property.");
    } else {
        logger.warn("Failed to get Word Application object in test.");
    }

    // Intenta abrir un documento (opcional, requiere manejo de errores adicional)
    // try {
    //   if (wordApp) { // Solo intentar si wordApp existe
    //      const doc = wordApp.Documents.Add();
    //      logger.info('Created new document.');
    //      // Asegurarse de que Close toma los argumentos correctos si se descomenta
    //      // El primer argumento suele ser si guardar cambios (false = no guardar)
    //      doc.Close(false);
    //      logger.info('Closed test document.');
    //   }
    // } catch (docError) {
    //    logger.error('Failed to create or close test document:', docError);
    // }

    // Intentar liberar el objeto solo si se obtuvo
    if (wordApp) {
        releaseObject(wordApp);
        logger.info("Word Application release attempted.");
    }
  } catch (error) {
    // Registrar el error pero no detener el inicio del servidor
    logger.error("COM Interop test failed:", error);
  }
}
testCom(); // Llama a la función de prueba al inicio
// --- Fin Prueba COM Interop ---
// Start the server using SSE transport
mcpServer.start({
    transportType: "sse",
    sse: {
        endpoint: "/mcp", // Define the SSE endpoint path
        port: Number(PORT) // Ensure PORT is a number
    }
})
    .then(() => {
        logger.info(`🚀 msoffice-mcp server listening on port ${PORT} at endpoint /mcp`);
    })
    .catch((error: Error) => {
        logger.error('Failed to start msoffice-mcp server:', error);
        process.exit(1);
    });

// Graceful shutdown handling
process.on('SIGTERM', () => {
    logger.info('SIGTERM signal received. Closing server...');
    mcpServer.stop().then(() => { // Assuming a stop method exists
         logger.info('Server closed.');
         process.exit(0);
    });
});

process.on('SIGINT', () => {
     logger.info('SIGINT signal received. Closing server...');
    mcpServer.stop().then(() => {
         logger.info('Server closed.');
         process.exit(0);
    });
});