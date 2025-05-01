// /src/server/index.ts
// =============================================================================
/**
 * @file Main entry point for the msoffice-mcp server.
 * Initializes FastMCP and registers all defined tools.
 */
import { FastMCP } from 'fastmcp'; // Assuming FastMCP is correctly imported
import { ZodError } from 'zod';
import allTools from '@/tools';
import logger from '@/utils/logger';
import { handleToolError, createErrorResponse } from '@/utils/errorHandler';
import { McpResource, ToolContext, ToolRequestParams, ApiResponse } from '@/types/common.types';

// --- FastMCP Server Configuration ---
const PORT = process.env.PORT || 3000;

// Assuming FastMCP constructor takes an options object
// Adjust based on actual FastMCP API
const mcpServer = new FastMCP({
    port: PORT,
    // Add other FastMCP configurations: authentication, middleware, etc.
});

// --- Register Tools ---
logger.info(`Registering ${allTools.length} MCP resources...`);

allTools.forEach((tool: McpResource) => {
    try {
        // Adapt this registration logic based on FastMCP's API
        // This example assumes a `registerResource` method
        mcpServer.registerResource({
            path: tool.path,
            description: tool.description, // Pass description if FastMCP supports it
            handler: async (params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<any>> => {
                const startTime = Date.now();
                logger.info(`[${tool.path}] Request received`, { params: hideSensitiveParams(params) }); // Avoid logging sensitive data

                try {
                    // 1. Input Validation (if schema provided)
                    if (tool.schema) {
                        try {
                            params = tool.schema.parse(params); // Validate and potentially transform params
                        } catch (validationError) {
                            if (validationError instanceof ZodError) {
                                logger.warn(`[${tool.path}] Validation failed`, { errors: validationError.errors });
                                return createErrorResponse('VALIDATION_ERROR', 'Input validation failed.', validationError.flatten());
                            }
                            throw validationError; // Re-throw unexpected validation errors
                        }
                    }

                    // 2. Authorization/Permission Checks (Example)
                    // checkPermissions(tool.path, context); // Implement this function based on your auth logic

                    // 3. Execute Tool Handler
                    const result = await tool.handler(params, context);

                    const duration = Date.now() - startTime;
                    if (result.success) {
                         logger.info(`[${tool.path}] Request successful`, { durationMs: duration });
                    } else {
                         logger.warn(`[${tool.path}] Request failed`, { durationMs: duration, error: result.error });
                    }
                    return result;

                } catch (error) {
                    const duration = Date.now() - startTime;
                    logger.error(`[${tool.path}] Unhandled error in handler`, { durationMs: duration, error });
                    return handleToolError(error); // Use centralized error handler
                }
            },
            // Pass completions if FastMCP supports it
            completions: tool.completions ? async (context?: ToolContext) => {
                 try {
                     return await tool.completions(context);
                 } catch (error) {
                     logger.error(`[${tool.path}] Error generating completions`, { error });
                     return {}; // Return empty object on error
                 }
            } : undefined,
        });
        logger.debug(`Registered resource: ${tool.path}`);
    } catch (error) {
         logger.error(`Failed to register resource: ${tool.path}`, { error });
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
mcpServer.start()
    .then(() => {
        logger.info(`🚀 msoffice-mcp server listening on port ${PORT}`);
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