// /src/server/index.ts
// =============================================================================
/**
 * @file Main entry point for the msoffice-mcp server.
 * Initializes FastMCP and registers all defined tools.
 */
import { FastMCP, UserError, UnexpectedStateError, Context as FastMCPContext, ContentResult, TextContent, ToolParameters, audioContent, imageContent, SerializableValue } from 'fastmcp';
import { StandardSchemaV1 } from '@standard-schema/spec'; // Import StandardSchemaV1
import { IncomingMessage } from 'http'; // Ensure IncomingMessage is imported
import { allRegisteredTools, aiAssistantGuideResource } from '@/tools'; // Updated import
import logger from '@/utils/logger';
import { handleToolError, createErrorResponse } from '@/utils/errorHandler';
import { validateFilePath } from '@/utils/security'; // Import validateFilePath
import { McpResource, ToolRequestParams, ApiResponse } from '@/types/common.types'; // Removed SerializableValue import from here
import { version as packageVersion, name as packageName } from '../../package.json'; // Import name and version

// --- Authentication ---

// Define the structure of the session data returned on successful authentication
// Add index signature to satisfy FastMCPSessionAuth constraint
interface AuthSessionData {
  clientId: string; // Example session data: store the client ID
  authenticatedAt: number;
  [key: string]: unknown; // Index signature
}

// Define a type alias for the FastMCP context specific to this server
type ServerContext = FastMCPContext<AuthSessionData>;

// Authentication Handler Function (English)
// Make the function async to return a Promise<AuthSessionData>
const authenticateHandler = async ( // Added async
    request: IncomingMessage, // Accept request directly
    metadata?: Record<string, unknown> // Metadata as optional second argument
): Promise<AuthSessionData> => { // Return Promise<AuthSessionData>
    const expectedToken = process.env.MCP_AUTH_TOKEN || 'default-secret-token'; // Use env var or default
    // Read token from metadata first, then from Node.js headers
    const providedToken = metadata?.apiKey as string || request.headers['x-api-key'] as string;

    logger.info(`Authentication attempt. Provided token (type): ${typeof providedToken}`); // Log attempt type

    if (!providedToken) {
        logger.warn('Authentication failed: No token provided.');
        // Throwing a Response is the standard way to reject in FastMCP authenticate
        throw new Response('Unauthorized: API key required in metadata.apiKey or x-api-key header.', { status: 401 });
    }

    if (providedToken !== expectedToken) {
        logger.warn('Authentication failed: Invalid token provided.');
        throw new Response('Unauthorized: Invalid API key.', { status: 401 });
    }

    // Success: Return session data
    const sessionData: AuthSessionData = {
        clientId: metadata?.clientName as string || 'unknown-client', // Get client name from metadata if available
        authenticatedAt: Date.now(),
    };
    logger.info(`Authentication successful for client: ${sessionData.clientId}`);
    return sessionData;
};


// --- FastMCP Server Configuration ---
const OFFICE_MCP_PORT = process.env.OFFICE_MCP_PORT;

logger.debug('Creating FastMCP server instance...'); // Added log
// Instantiate FastMCP with name, version, instructions, and authentication
const mcpServer = new FastMCP<AuthSessionData>({ // Specify session data type for the server instance
    name: packageName,
    version: packageVersion as `${number}.${number}.${number}`, // Assert type for version
    instructions: "This server provides tools to interact with Microsoft Office files (Word, Excel, PowerPoint). Use the available tools to read, write, modify, and analyze documents. Authentication is required.",
    authenticate: authenticateHandler, // Add the authentication handler
});
logger.debug('FastMCP server instance created with authentication.'); // Added log

// --- Register Tools ---
logger.info(`Registering ${allRegisteredTools.length} tools...`);

allRegisteredTools.forEach((item: McpResource) => {
    // Ensure item has a schema before registering as a tool
    if (!item.schema) {
        logger.warn(`Skipping registration for item without schema: ${item.path}`);
        return; // Skip items without schema (likely resources or invalid entries)
    }

    try {
        logger.debug(`Attempting to register tool: ${item.path}`);
        // Register as a tool
        // Define annotations, adding specific ones for potentially slow tools
        const annotations: Record<string, unknown> = { title: item.description || item.path }; // Use description as title if available, else path
        const slowTools = ['word/merge', 'word/markdown/export', 'word/markdown/import', 'word/search-replace'];
        if (slowTools.includes(item.path)) {
            annotations.estimatedDuration = "This operation can take several seconds or more depending on document size and complexity.";
        }

        mcpServer.addTool({
            name: item.path, // Use the unique path as the tool name
            description: item.description,
            // Use item.schema as parameters, already checked it exists
            parameters: item.schema as unknown as ToolParameters, // Cast schema
            annotations: annotations, // Add annotations
            // Adjust return type to Promise<ContentResult>
            // Ensure the context type here matches the updated ServerContext with AuthSessionData
            execute: async (args: StandardSchemaV1.InferOutput<any>, context: ServerContext): Promise<ContentResult> => {
                const startTime = Date.now();
                // Use context.log provided by FastMCP
                // Log client ID from session data, checking if session exists
                const clientId = context.session?.clientId || 'unknown-authenticated-client';
                context.log.info(`[${item.path}] EXECUTION START by client: ${clientId}`, { params: hideSensitiveParams(args) as SerializableValue }); // Ensure logged params are serializable

                try {
                    // Validation is typically handled by FastMCP based on 'parameters' schema
                    // Authorization checks can use context.session if authentication is implemented

                    // Execute the original tool handler
                    // Pass args and the FastMCP context to the original handler
                    const apiResponse: ApiResponse<any> = await item.handler(args, context); // Pass args and context

                    const duration = Date.now() - startTime;

                    // Transform the ApiResponse to the format expected by FastMCP
                    if (apiResponse.success) {
                        context.log.info(`[${item.path}] EXECUTION SUCCESS`, { durationMs: duration });
                        const data = apiResponse.data;

                        // Map data to FastMCP return types
                        if (typeof data === 'string') {
                            // Wrap string in TextContent structure and then in a content array
                            const textContent: TextContent = { type: 'text', text: data };
                            return { content: [textContent] }; // Wrap in array
                        } else if (Buffer.isBuffer(data)) {
                            // Handle Buffer data (likely images/audio)
                            // Ensure args is treated as 'any' or validated type to access outputFormat safely
                            const validatedArgs = args as any;
                            let mimeType = 'application/octet-stream'; // Default MIME type
                            if (item.path === 'word/image/extract' && validatedArgs.outputFormat) {
                                mimeType = `image/${validatedArgs.outputFormat}`;
                            }
                            context.log.info(`[${item.path}] Returning Buffer data as imageContent`, { mimeType, size: data.length });
                            // Await imageContent and wrap its result in ContentResult structure
                            // Note: We still need to pass mimeType if imageContent supports it. Assuming it does via options.
                            // If imageContent({ buffer: data }) doesn't support mimeType, this needs adjustment.
                            // Let's assume for now it might be part of the ImageContent object itself or metadata.
                            // Reverting to the structure that caused the fewest errors previously:
                            const imgContent = await imageContent({ buffer: data /*, mimeType: mimeType */ }); // Pass mimeType if supported
                            return { content: [imgContent] }; // Wrap in array
                        } else if (data && typeof data === 'object') {
                            // Handle other objects by creating TextContent and wrapping in a content array
                            try {
                                const textContent: TextContent = { type: 'text', text: JSON.stringify(data, null, 2) };
                                return { content: [textContent] }; // Wrap in array
                            } catch (stringifyError) {
                                context.log.error(`[${item.path}] Error stringifying successful object response data`, { error: String(stringifyError) });
                                throw new UnexpectedStateError("Failed to serialize successful object response data.");
                            }
                        } else {
                            // Handle null, undefined, or other primitive types (excluding string/buffer)
                            // Return a default success message as TextContent wrapped in a content array
                            const textContent: TextContent = { type: 'text', text: 'Operation completed successfully.' };
                            return { content: [textContent] }; // Wrap in array
                        }
                    } else {
                        // Throw a UserError for FastMCP to handle client-side errors
                        // Convert error details to string for logging and UserError
                        const errorDetailsString = apiResponse.error?.details ? String(apiResponse.error.details) : undefined;
                        context.log.warn(`[${item.path}] EXECUTION FAILED (API Error)`, { durationMs: duration, code: apiResponse.error?.code, message: apiResponse.error?.message, details: errorDetailsString });
                        // Pass serializable details to UserError
                        const userErrorDetails = errorDetailsString ? { details: errorDetailsString } : undefined;
                        throw new UserError(apiResponse.error?.message || 'Tool execution failed.', userErrorDetails);
                    }

                } catch (error) {
                    const duration = Date.now() - startTime;
                     // Convert error to string for logging
                    context.log.error(`[${item.path}] EXECUTION FAILED (Unhandled Error)`, { durationMs: duration, error: String(error) });

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
        logger.debug(`Registered tool: ${item.path}`);

    } catch (error) {
         logger.error(`Failed to register tool: ${item.path}`, { error });
    }
});
// --- Register Prompts ---
logger.info("Registering prompts...");
try {
    mcpServer.addPrompt({
        name: "summarize-word-section", // English name
        description: "Summarizes a specific section (e.g., paragraph) of a Word document.", // English description
        arguments: [
            {
                name: "filePath",
                description: "Path to the Word document.",
                required: true,
            },
            {
                name: "range",
                description: "The range to summarize (e.g., 'paragraph:5', 'document').",
                required: true,
            },
            {
                name: "style",
                description: "Optional: Desired summary style (e.g., 'bullet points', 'concise paragraph').",
                required: false,
            }
        ],
        // The 'load' function generates the actual prompt text sent to the LLM
        load: async (args) => {
            // Construct the prompt using arguments.
            // This example assumes the client will use a tool like 'word/text/get'
            // to fetch the content based on filePath and range before calling the LLM.
            // The prompt guides the LLM on how to process that fetched text.
            let promptText = `Please summarize the following text extracted from the range "${args.range}" of the document "${args.filePath}":\n\n{extracted_text}\n\n`;
            if (args.style) {
                promptText += `Present the summary in the style of: ${args.style}.`;
            } else {
                promptText += `Present the summary clearly and concisely.`;
            }
            // Note: '{extracted_text}' is a placeholder the client/LLM needs to fill
            // based on the context or prior tool calls.
            return promptText;
        },
    });
    logger.info("Registered prompt: summarize-word-section");
} catch (error) {
    logger.error("Failed to register prompts", { error });
}
logger.info("Prompt registration complete.");

logger.info("All tools registered.");

// --- Register Resources ---
logger.info("Registering resources...");
try {
    logger.debug(`Attempting to register resource: ${aiAssistantGuideResource[0].path}`); // Access first element
    mcpServer.addResource({
        uri: aiAssistantGuideResource[0].path, // Access first element
        name: aiAssistantGuideResource[0].description || 'Unnamed Resource', // Access first element
        load: async () => {
// --- Register Resource Templates ---
logger.info("Registering resource templates...");

// Placeholder functions for resource handling (to be moved/implemented in officeInterop.ts)
// These are simplified placeholders for demonstration within the handler
// Return TextContent or Buffer directly. Throw error if not found/supported.
async function getWordElementContent(filePath: string, elementType: string, identifier: string): Promise<TextContent | Buffer> {
    logger.info(`[ResourceTemplate] Placeholder: Getting Word element`, { filePath, elementType, identifier });
    await validateFilePath(filePath); // Validate path
    if (elementType === 'paragraph') {
        // Placeholder: Fetch paragraph text
        const text = `Placeholder text for paragraph ${identifier} in ${filePath}`;
        // Return TextContent structure
        return { type: 'text', text };
    } else if (elementType === 'image') {
        // Placeholder: Fetch image data (e.g., a small dummy PNG buffer)
        const dummyPngBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';
        const buffer = Buffer.from(dummyPngBase64, 'base64');
        // Return the raw buffer
        return buffer;
    }
    // Add more element types as needed

    // Throw error if element type is not supported
    throw new UserError(`Element type '${elementType}' not supported for Word documents.`);
}

// Define the resource template load function
// It only receives the parsed arguments from the uriTemplate, not the context
async function loadOfficeResource(
    params: { app: string; filepath: string; elementType: string; identifier: string }
): Promise<{ text: string } | { blob: string }> { // Return type based on README (no null)
    const { app, filepath, elementType, identifier } = params;
    const startTime = Date.now();
    // Use the global logger directly as context is not available here
    logger.info(`[ResourceTemplate] Loading resource for office://`, { app, filepath, elementType, identifier });

    try {
        // Basic validation
        if (!app || !filepath || !elementType || !identifier) {
            throw new UserError("Invalid office resource URI: Missing components.");
        }

        // Decode filepath if necessary (URIs might encode paths)
        const decodedFilepath = decodeURIComponent(filepath);

        // **Security Validation is CRUCIAL here** - Reuse or enhance validateFilePath
        await validateFilePath(decodedFilepath); // Ensure path is safe and within workspace

        // Variable to hold the raw content (TextContent structure or Buffer)
        let rawContent: TextContent | Buffer | null = null;

        // Route based on application
        switch (app.toLowerCase()) {
            case 'word':
                // Call specific Word interop function based on elementType
                rawContent = await getWordElementContent(decodedFilepath, elementType.toLowerCase(), identifier);
                break;
            case 'excel':
                // TODO: Implement Excel handling - Placeholder
                logger.warn(`[ResourceTemplate] Excel resource handling not implemented yet.`, { filepath: decodedFilepath, elementType, identifier }); // Use logger
                throw new UserError(`Excel resource handling not implemented yet.`);
                // rawContent = await getExcelElementContent(decodedFilepath, elementType.toLowerCase(), identifier);
                break;
            case 'powerpoint':
                // TODO: Implement PowerPoint handling - Placeholder
                logger.warn(`[ResourceTemplate] PowerPoint resource handling not implemented yet.`, { filepath: decodedFilepath, elementType, identifier }); // Use logger
                throw new UserError(`PowerPoint resource handling not implemented yet.`);
                // rawContent = await getPowerPointElementContent(decodedFilepath, elementType.toLowerCase(), identifier);
                break;
            default:
                throw new UserError(`Unsupported Office application in URI: ${app}`);
        }

        if (!rawContent) {
            throw new UserError(`Element type '${elementType}' with identifier '${identifier}' not found or not supported in ${app} document '${decodedFilepath}'.`);
        }

        const duration = Date.now() - startTime;
        // Use global logger as context is not available in load function
        logger.info(`[ResourceTemplate] Successfully retrieved resource.`, { durationMs: duration, app, filepath: decodedFilepath, elementType, identifier });

        // Process rawContent into the { text: ... } or { blob: ... } structure
        if (Buffer.isBuffer(rawContent)) {
            // If it's a buffer, convert to base64 and return as blob
            const base64Data = rawContent.toString('base64');
            logger.info(`[ResourceTemplate] Successfully loaded resource as blob.`, { durationMs: Date.now() - startTime, app, filepath: decodedFilepath, elementType, identifier, size: base64Data.length }); // Use logger
            return { blob: base64Data };
        } else {
            // Otherwise, it should be TextContent structure, return its text property
            logger.info(`[ResourceTemplate] Successfully loaded resource as text.`, { durationMs: Date.now() - startTime, app, filepath: decodedFilepath, elementType, identifier }); // Use logger
            return { text: rawContent.text }; // Assuming rawContent is { type: 'text', text: ... }
        }

    } catch (error: any) {
        const duration = Date.now() - startTime;
        // Use the global logger directly
        logger.error(`[ResourceTemplate] Error loading office resource URI`, { durationMs: duration, app, filepath, elementType, identifier, error: String(error) });

        // Re-throw FastMCP specific errors or wrap others in UserError
        if (error instanceof UserError || error instanceof UnexpectedStateError) {
            throw error;
        }
        throw new UserError(`Failed to handle office resource: ${error.message}`, { details: String(error) });
    }
}


// Register the template
try {
    const officeUriPattern = 'office://:app/:filepath/:elementType/:identifier';
     logger.debug(`Attempting to register resource template: ${officeUriPattern}`);
   mcpServer.addResourceTemplate({
       uriTemplate: officeUriPattern, // Use uriTemplate property
       load: loadOfficeResource,      // Use load property with the correct function
       // Define arguments based on the template placeholders
       arguments: [
           { name: 'app', description: 'Office application (word, excel, powerpoint)', required: true },
           { name: 'filepath', description: 'URI encoded path to the file', required: true },
           { name: 'elementType', description: 'Type of element (paragraph, image, cell, shape, etc.)', required: true },
           { name: 'identifier', description: 'Identifier for the element (index, name, address, etc.)', required: true },
       ],
       name: "Office Document Element", // Add name and mimeType as per README example
       mimeType: "application/octet-stream", // Default mimeType, specific handlers might override
       // Optional: Add description
       description: "Access specific elements within Office documents (Word, Excel, PowerPoint)."
   });
   logger.info(`Registered resource template: ${officeUriPattern}`);
   logger.info("Resource template registration complete.");
} catch (error) {
    logger.error(`Failed to register office resource template`, { error });
}
            // Resources don't have parameters, call handler without args
            // Pass undefined for context as resource loader doesn't provide it
            const apiResponse: ApiResponse<any> = await aiAssistantGuideResource[0].handler({}, undefined);
            if (apiResponse.success) {
                return { text: String(apiResponse.data) }; // Assuming text content
            } else {
                logger.error(`Failed to load resource: ${aiAssistantGuideResource[0].path}`, { error: apiResponse.error }); // Access first element
                throw new UnexpectedStateError(`Failed to load resource: ${aiAssistantGuideResource[0].path}`); // Access first element
            }
        },
    });
    logger.debug(`Registered resource: ${aiAssistantGuideResource[0].path}`); // Access first element
    logger.info("Resource registration complete.");
} catch (error) {
    // Log the error, attempting to access path safely
    const resourcePath = aiAssistantGuideResource && aiAssistantGuideResource[0] ? aiAssistantGuideResource[0].path : 'unknown resource';
    logger.error(`Failed to register resource: ${resourcePath}`, { error });
}

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
  const comStartTime = Date.now();
  logger.info("[COM TEST START] Testing COM Interop (winax)...");
  try {
    logger.info("[COM TEST] Attempting getOfficeApplication('Word.Application')...");
    const wordApp = await getOfficeApplication('Word.Application');
    logger.info("[COM TEST] getOfficeApplication completed.");
    // Verificar que el objeto y la propiedad existen antes de acceder
    if (wordApp && typeof wordApp.Version !== 'undefined') {
        logger.info(`[COM TEST] Got Word Application, Version: ${wordApp.Version}`);
    } else if (wordApp) {
        logger.warn("[COM TEST] Got Word Application object, but couldn't retrieve Version property.");
    } else {
        logger.warn("[COM TEST] Failed to get Word Application object.");
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
        logger.info("[COM TEST] Attempting releaseObject(wordApp)...");
        releaseObject(wordApp);
        logger.info("[COM TEST] releaseObject attempted.");
    }
    const comDuration = Date.now() - comStartTime;
    logger.info(`[COM TEST SUCCESS] COM Interop test finished successfully. Duration: ${comDuration}ms`);
  } catch (error) {
    const comDuration = Date.now() - comStartTime;
    // Registrar el error pero no detener el inicio del servidor
    logger.error(`[COM TEST FAILED] COM Interop test failed. Duration: ${comDuration}ms`, { error });
  }
}
logger.info("Calling testCom()...");
testCom(); // Llama a la función de prueba al inicio
logger.info("testCom() finished.");
// --- Fin Prueba COM Interop ---
logger.debug('Attempting to start MCP server...'); // Added log
// Start the server
if (OFFICE_MCP_PORT) {
    // Use SSE transport if port is defined
    mcpServer.start({
        transportType: "sse",
        sse: {
            endpoint: "/mcp", // Define the SSE endpoint path
            port: Number(OFFICE_MCP_PORT) // Ensure OFFICE_MCP_PORT is a number
        }
    })
    .then(() => {
        logger.info(`🚀 msoffice-mcp server listening on port ${OFFICE_MCP_PORT} at endpoint /mcp`);
    })
    .catch((error: Error) => {
        logger.error('Failed to start msoffice-mcp server:', error);
        process.exit(1);
    });
} else {
    // Use STDIO transport if port is not defined
    mcpServer.start({
        transportType: "stdio"
    })
    .then(() => {
        logger.info('🚀 msoffice-mcp server started successfully in STDIO mode'); // Added log
    })
    .catch((error: Error) => {
        logger.error('Failed to start msoffice-mcp server in STDIO mode:', error);
        process.exit(1);
    });
}

// Graceful shutdown handling
// --- Manejo de Errores EPIPE en Streams ---
// Estos manejadores intentan capturar errores EPIPE que pueden ocurrir si
// el proceso intenta escribir en stdout/stderr después de que la tubería
// se haya cerrado (común cuando el proceso padre termina abruptamente).
// Nota: Esto generalmente solo silencia el síntoma, no arregla la causa
// raíz si el cierre no es limpio (p.ej., con fastmcp dev).

/* 
process.stdout.on('error', (err: NodeJS.ErrnoException) => {
  if (err.code === 'EPIPE') {
    // Ignorar EPIPE en stdout, probablemente causado por cierre abrupto.
    logger.warn('Error EPIPE en stdout ignorado durante el cierre.');
  } else {
    // Registrar otros errores inesperados de stdout
    logger.error('Error inesperado en process.stdout:', { error: err });
    // Considerar salir si es un error crítico no relacionado con EPIPE
    // process.exit(1);
  }
});

process.stderr.on('error', (err: NodeJS.ErrnoException) => {
  if (err.code === 'EPIPE') {
    // Ignorar EPIPE en stderr, probablemente causado por cierre abrupto.
    logger.warn('Error EPIPE en stderr ignorado durante el cierre.');
  } else {
    // Registrar otros errores inesperados de stderr
    logger.error('Error inesperado en process.stderr:', { error: err });
    // Considerar salir si es un error crítico no relacionado con EPIPE
    // process.exit(1);
  }
});
*/
process.on('SIGTERM', () => {
    logger.info('SIGTERM signal received. Closing server...');
    mcpServer.stop().then(() => { // Assuming a stop method exists
         logger.info('Server closed.');
         process.exit(0);
    }).catch(err => { // Add catch block for stop() errors
        logger.error('Error during mcpServer.stop() for SIGTERM:', { error: err }); // Use logger
        process.exit(1);
    });
});

process.on('SIGINT', () => {
     logger.info('SIGINT signal received. Closing server...');
    mcpServer.stop().then(() => {
         logger.info('Server closed.');
         process.exit(0);
    }).catch(err => { // Add catch block for stop() errors
        logger.error('Error during mcpServer.stop() for SIGINT:', { error: err }); // Use logger
        process.exit(1);
    });
});