/**
 * @file Main entry point for the msoffice-mcp server.
 * Initializes FastMCP, registers tools and resources, and handles server startup/shutdown.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { FastMCP, UserError, UnexpectedStateError, Context as FastMCPContext, ContentResult, TextContent, ToolParameters, audioContent, imageContent, SerializableValue, ResourceResult } from 'fastmcp';
import { StandardSchemaV1 } from '@standard-schema/spec';
import { IncomingMessage } from 'http';
import { allRegisteredTools, aiAssistantGuideResource, registerWordCodeFormatTool, registerArchiveTools } from '@/tools/index.js'; // Import registerArchiveTools
import logger from '@/utils/logger.js';
import { handleToolError, createErrorResponse } from '@/utils/errorHandler.js';
import { validateFilePath, isFsAccessAllowed } from '@/utils/security.js';
import { getWordElementContent as getOfficeElementContentInterop } from '@/utils/officeInterop.js';
import { McpResource, ToolRequestParams, ApiResponse } from '@/types/common.types.js';
// Replaced JSON import assertion with fs.readFileSync due to SyntaxError with Node.js v22.13.0
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
// 'path' module is already imported globally for this file (see line 22).

// Determine the directory of the current module to correctly resolve package.json
// Using unique variable names to avoid potential clashes if __filename/__dirname are used elsewhere.
const __filename_server_index_for_pkg_json = fileURLToPath(import.meta.url);
const __dirname_server_index_for_pkg_json = path.dirname(__filename_server_index_for_pkg_json);
const packageJsonPath = path.resolve(__dirname_server_index_for_pkg_json, '../../package.json');
const packageInfo = JSON.parse(readFileSync(packageJsonPath, 'utf-8'));
const packageVersion = packageInfo.version;
const packageName = packageInfo.name;
import * as fs from 'fs/promises'; // Import fs.promises for async file operations
import * as path from 'path'; // Import path module

// --- Authentication ---

/**
 * Defines the structure of the session data returned on successful authentication.
 * Includes an index signature to satisfy FastMCPSessionAuth constraint.
 */
interface AuthSessionData {
  /** The client ID of the authenticated user. */
  clientId: string;
  /** Timestamp of successful authentication. */
  authenticatedAt: number;
  [key: string]: unknown; // Allows for additional session data properties.
}

/**
 * Define a type alias for the FastMCP context specific to this server, including authentication session data.
 */
type ServerContext = FastMCPContext<AuthSessionData>;

/**
 * Authentication Handler Function.
 * Validates the provided API key against an expected token from environment variables.
 * @param request - The incoming HTTP request.
 * @param metadata - Optional metadata provided during the authentication attempt.
 * @returns A promise resolving to the session data if authentication is successful.
 * @throws {Response} A Response object with status 401 if authentication fails.
 */
const authenticateHandler = async (
    request: IncomingMessage,
    metadata?: Record<string, unknown>
): Promise<AuthSessionData> => {
    const expectedToken = process.env.MCP_AUTH_TOKEN || 'default-secret-token'; // Use env var or default
    // Read token from metadata first, then from Node.js headers
    const providedToken = metadata?.apiKey as string || request.headers['x-api-key'] as string;

    logger.info(`[Auth] Authentication attempt. Provided token (type): ${typeof providedToken}`);

    if (!providedToken) {
        logger.warn('[Auth] Authentication failed: No token provided.');
        // Throwing a Response is the standard way to reject in FastMCP authenticate
        throw new Response('Unauthorized: API key required in metadata.apiKey or x-api-key header.', { status: 401 });
    }

    if (providedToken !== expectedToken) {
        logger.warn('[Auth] Authentication failed: Invalid token provided.');
        throw new Response('Unauthorized: Invalid API key.', { status: 401 });
    }

    // Success: Return session data
    const sessionData: AuthSessionData = {
        clientId: metadata?.clientName as string || 'unknown-client', // Get client name from metadata if available
        authenticatedAt: Date.now(),
    };
    logger.info(`[Auth] Authentication successful for client: ${sessionData.clientId}`);
    return sessionData;
};


// --- FastMCP Server Configuration ---

/**
 * The port the MCP server will listen on for SSE connections.
 * If not defined, the server will run in STDIO mode.
 */
const OFFICE_MCP_PORT = process.env.OFFICE_MCP_PORT;

logger.debug('Creating FastMCP server instance...');
/**
 * The FastMCP server instance for the MS Office MCP.
 * Configured with server name, version, instructions, and authentication handler.
 */
const mcpServer = new FastMCP<AuthSessionData>({ // Specify session data type for the server instance
    name: packageName,
    version: packageVersion as `${number}.${number}.${number}`, // Assert type for version
    instructions: "This server provides tools to interact with Microsoft Office files (Word, Excel, PowerPoint). Use the available tools to read, write, modify, and analyze documents. Authentication is required.",
    authenticate: authenticateHandler, // Add the authentication handler
});
logger.debug('FastMCP server instance created with authentication.');

// --- Register Tools ---

/**
 * Registers all available tools with the FastMCP server instance.
 * Skips FS tools if FS access is disabled or not in STDIO mode.
 */
logger.info(`Registering ${allRegisteredTools.length} tools...`);

/**
 * Determines if the server is running in STDIO mode (local execution).
 * This is true if the OFFICE_MCP_PORT environment variable is not set.
 */
const isStdioMode = !OFFICE_MCP_PORT;

for (let i = 0; i < allRegisteredTools.length; i++) {
    const item: McpResource = allRegisteredTools[i];

    // Check if the tool is an FS tool
    const isFsTool = item.path.startsWith('fs/');

    // If it's an FS tool, check if FS access is allowed and if running in STDIO mode
    if (isFsTool) {
        if (!isFsAccessAllowed || !isStdioMode) {
            logger.info(`[Tool Registration] Skipping registration for FS tool '${item.path}' because FS access is disabled or not in STDIO mode.`);
            continue; // Skip registration if conditions are not met
        }
        logger.info(`[Tool Registration] Registering FS tool '${item.path}' as FS access is allowed and in STDIO mode.`);
    }


    // Ensure item has a schema before registering as a tool
    if (!item.schema) {
        logger.warn(`[Tool Registration] Skipping registration for item without schema at index ${i}: ${item.path}`);
        continue; // Skip items without schema (likely resources or invalid entries)
    }

    try {
        logger.debug(`[Tool Registration] Attempting to register tool at index ${i}: ${item.path}`);
        logger.debug(`[Tool Registration] Schema ${i}:: ${JSON.stringify(item.schema)}`);
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
            parameters: item.schema as ToolParameters, // Cast schema to ToolParameters
            annotations: annotations, // Add annotations
            // Adjust return type to Promise<ContentResult>
            // Ensure the context type here matches the updated ServerContext with AuthSessionData
            execute: async (args: StandardSchemaV1.InferOutput<any>, context: ServerContext): Promise<ContentResult> => {
                const startTime = Date.now();
                // Use context.log provided by FastMCP
                // Log client ID from session data, checking if session exists
                const clientId = context.session?.clientId || 'unknown-authenticated-client';
                context.log.info(`[${item.path}] EXECUTION START by client: ${clientId}`, { params: hideSensitiveParams(args) as SerializableValue }); // Ensure logged params are serializable

                // Add logging for requestSampling availability
                logger.info(`[${item.path}] requestSampling availability: ${typeof context.session?.requestSampling === 'function' ? 'available' : 'unavailable'}`);

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
        logger.debug(`[Tool Registration] Registered tool at index ${i}: ${item.path}`);

    } catch (error) {
         logger.error(`[Tool Registration] Failed to register tool at index ${i}: ${item.path}`, { error });
    }
}

// Register the new archive tools
logger.info("Registering archive tools...");
try {
    registerArchiveTools(mcpServer);
    logger.info("Archive tool registration complete.");
} catch (error) {
    logger.error("[Tool Registration] Failed to register archive tools", { error });
}


// --- Register Prompts ---

/**
 * Registers prompts with the FastMCP server instance.
 * Prompts define templates for generating LLM input based on user arguments.
 */
logger.info("Registering prompts...");
try {
    mcpServer.addPrompt({
        name: "summarize-word-section",
        description: "Summarizes a specific section (e.g., paragraph) of a Word document.",
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
        /**
         * Generates the actual prompt text sent to the LLM for summarizing a Word document section.
         * @param args - The arguments provided to the prompt.
         * @returns The generated prompt text.
         */
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
    logger.info("[Prompt Registration] Registered prompt: summarize-word-section");
} catch (error) {
    logger.error("[Prompt Registration] Failed to register prompts", { error });
}
logger.info("Prompt registration complete.");

logger.info("All tools registered.");

// --- Register Resources ---

/**
 * Registers static resources with the FastMCP server instance.
 * Resources provide access to static data or information.
 */
logger.info("Registering resources...");
try {
    logger.debug(`[Resource Registration] Attempting to register resource: ${aiAssistantGuideResource[0].path}`); // Access first element
    mcpServer.addResource({
        uri: aiAssistantGuideResource[0].path, // Access first element
        name: aiAssistantGuideResource[0].description || 'Unnamed Resource', // Access first element
        mimeType: 'text/markdown', // Specify MIME type for Markdown
        /**
         * Loads the content of the AI assistant guide resource.
         * @returns A promise resolving to an object with the text content.
         * @throws {UnexpectedStateError} If the resource fails to load.
         */
        load: async () => {
            // Resources don't have parameters, call handler without args
            // Pass undefined for context as resource loader doesn't provide it
            const apiResponse: ApiResponse<any> = await aiAssistantGuideResource[0].handler({}, undefined);
            if (apiResponse.success) {
                return { text: String(apiResponse.data) }; // Assuming text content
            } else {
                logger.error(`[Resource Registration] Failed to load resource: ${aiAssistantGuideResource[0].path}`, { error: apiResponse.error }); // Access first element
                throw new UnexpectedStateError(`Failed to load resource: ${aiAssistantGuideResource[0].path}`); // Access first element
            }
        },
    });
    logger.debug(`[Resource Registration] Registered resource: ${aiAssistantGuideResource[0].path}`); // Access first element
    logger.info("Resource registration complete.");
    logger.debug(`[Resource Registration] Attempting to register resource: memory://office_api_doc`);
    mcpServer.addResource({
        uri: 'memory://office_api_doc',
        name: 'Office API Documentation',
        mimeType: 'text/markdown',
        load: async () => {
            // Use the fs/file/read tool via FastMCP context or directly if available
            // Since this is server-side, we can use fs.promises.readFile directly
            try {
                const content = await fs.readFile(path.join(__dirname, '../docs/office_api_doc.md'), 'utf-8');
                return { text: content };
            } catch (error) {
                logger.error(`[Resource Registration] Failed to load resource: memory://office_api_doc`, { error });
                throw new UnexpectedStateError(`Failed to load resource: memory://office_api_doc`);
            }
        },
    });
    logger.debug(`[Resource Registration] Registered resource: memory://office_api_doc`);
} catch (error) {
    // Log the error, attempting to access path safely
    const resourcePath = aiAssistantGuideResource && aiAssistantGuideResource[0] ? aiAssistantGuideResource[0].path : 'unknown resource';
    logger.error(`[Resource Registration] Failed to register resource: ${resourcePath}`, { error });
}

// --- Register Resource Templates ---

/**
 * Loads content for an Office document element resource template.
 * Handles retrieving text from paragraphs or images from Word documents.
 * @param params - Parsed arguments from the URI template.
 * @returns A promise resolving to an object with either text or blob content.
 * @throws {UserError} If the URI is invalid, the application/element type is unsupported, or the element is not found.
 * @throws {UnexpectedStateError} If an unexpected content type is returned from the interop function.
 */
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
        validateFilePath(decodedFilepath); // Ensure path is safe and within workspace

        // Variable to hold the raw content (TextContent structure or Buffer)
        let rawContent: TextContent | Buffer | null = null;

        // Route based on application
        switch (app.toLowerCase()) {
            case 'word':
                // Validate element type and identifier for Word
                const elementTypeLower = elementType.toLowerCase();
                const elementIdentifierNum = parseInt(identifier, 10);

                if (isNaN(elementIdentifierNum)) {
                    throw new UserError(`Invalid numeric identifier for Word element: ${identifier}`);
                }
                if (elementTypeLower !== 'paragraph' && elementTypeLower !== 'image') {
                     throw new UserError(`Unsupported element type for Word resource: ${elementType}`);
                }

                // Call the actual interop function
                const elementContent = await getOfficeElementContentInterop(decodedFilepath, elementTypeLower, elementIdentifierNum);
                // Assign to rawContent based on type
                if (typeof elementContent === 'string') {
                    rawContent = { type: 'text', text: elementContent };
                } else { // It must be a Buffer
                    rawContent = elementContent;
                }
                break;
            case 'excel':
                // TODO: Implement Excel handling - Placeholder
                logger.warn(`[ResourceTemplate] Excel resource handling not implemented yet.`, { filepath: decodedFilepath, elementType, identifier });
                throw new UserError(`Excel resource handling not implemented yet.`);
                // rawContent = await getExcelElementContent(decodedFilepath, elementType.toLowerCase(), identifier);
                break;
            case 'powerpoint':
                // TODO: Implement PowerPoint handling - Placeholder
                logger.warn(`[ResourceTemplate] PowerPoint resource handling not implemented yet.`, { filepath: decodedFilepath, elementType, identifier });
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
            logger.info(`[ResourceTemplate] Successfully loaded resource as blob.`, { durationMs: Date.now() - startTime, app, filepath: decodedFilepath, elementType, identifier, size: base64Data.length });
            return { blob: base64Data };
        } else if (typeof rawContent === 'string') {
             // If it's a string (e.g., paragraph text), return as text
             logger.info(`[ResourceTemplate] Successfully loaded resource as text.`, { durationMs: Date.now() - startTime, app, filepath: decodedFilepath, elementType, identifier });
             return { text: rawContent };
        } else {
             // Should not happen if interop function returns string or Buffer
             throw new UnexpectedStateError('Unexpected content type returned from interop function.');
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
     logger.debug(`[Resource Template Registration] Attempting to register resource template: ${officeUriPattern}`);
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
   logger.info(`[Resource Template Registration] Registered resource template: ${officeUriPattern}`);
   logger.info("Resource template registration complete.");
} catch (error) {
    // Log the error, attempting to access path safely
    const resourcePath = aiAssistantGuideResource && aiAssistantGuideResource[0] ? aiAssistantGuideResource[0].path : 'unknown resource';
    logger.error(`[Resource Registration] Failed to register resource: ${resourcePath}`, { error });
}


// --- Helper to hide sensitive data from logs ---
/**
 * Hides sensitive parameter values from logs.
 * Replaces values for specified keys with '********' or a summary for large content.
 * @param params - The parameters object to process.
 * @returns A new object with sensitive values hidden.
 */
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


// --- COM Interop Test (Optional with winax) ---
import { getOfficeApplication, releaseObject } from '@/utils/officeInterop.js'; // Use alias

/**
 * Performs a basic test of COM Interop by attempting to get a Word application instance.
 * Logs the outcome of the test.
 */
async function testCom() {
  const comStartTime = Date.now();
  logger.info("[COM TEST START] Testing COM Interop (winax)...");
  try {
    logger.info("[COM TEST] Attempting getOfficeApplication('Word.Application')...");
    const wordApp = await getOfficeApplication('Word.Application');
    logger.info("[COM TEST] getOfficeApplication completed.");
    // Check if the object and property exist before accessing
    if (wordApp && typeof wordApp.Version !== 'undefined') {
        logger.info(`[COM TEST] Got Word Application, Version: ${wordApp.Version}`);
    } else if (wordApp) {
        logger.warn("[COM TEST] Got Word Application object, but couldn't retrieve Version property.");
    } else {
        logger.warn("[COM TEST] Failed to get Word Application object.");
    }

    // Attempt to open a document (optional, requires additional error handling)
    // try {
    //   if (wordApp) { // Only attempt if wordApp exists
    //      const doc = wordApp.Documents.Add();
    //      logger.info('Created new document.');
    //      // Ensure Close takes the correct arguments if uncommented
    //      // The first argument is usually whether to save changes (false = do not save)
    //      doc.Close(false);
    //      logger.info('Closed test document.');
    //   }
    // } catch (docError) {
    //    logger.error('Failed to create or close test document:', docError);
    // }

    // Attempt to release the object only if obtained
    if (wordApp) {
        logger.info("[COM TEST] Attempting releaseObject(wordApp)...");
        releaseObject(wordApp);
        logger.info("[COM TEST] releaseObject attempted.");
    }
    const comDuration = Date.now() - comStartTime;
    logger.info(`[COM TEST SUCCESS] COM Interop test finished successfully. Duration: ${comDuration}ms`);
  } catch (error) {
    const comDuration = Date.now() - comStartTime;
    // Log the error but do not stop server startup
    logger.error(`[COM TEST FAILED] COM Interop test failed. Duration: ${comDuration}ms`, { error });
  }
}
logger.info("Calling testCom()...");
// Call the test function on startup
testCom();
logger.info("testCom() finished.");
// --- End COM Interop Test ---

logger.debug('Attempting to start MCP server...');
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
        logger.info('🚀 msoffice-mcp server started successfully in STDIO mode');
        //mcpServer.sessions.forEach((session) => {
        //    logger.info(`[word/generate-and-insert-text] FastMCP context=${session.context} context.session=${session.context?.session}.`);
        //});
    })
    .catch((error: Error) => {
        logger.error('Failed to start msoffice-mcp server in STDIO mode:', error);
        process.exit(1);
    });
}

mcpServer.on("connect", (event) => {
    logger.info("Client connected:", event.session);
    /*
    if (typeof event.session.requestSampling === 'function') {
        logger.info("Client requestSampling");
        event.session.requestSampling({
            messages: [
              {
                role: "user",
                content: {
                  type: "text",
                  text: "Me llamo David. ¿y tú?",
                },
              },
            ],
            systemPrompt: "Eres un amable asistente.",
            includeContext: "thisServer",
            maxTokens: 100,
          }).then((response) => {
            logger.info("Request sampling response:", response);
            // Handle the response as needed
            // For example, send the response back to the client
            // event.session.sendResponse(response);
          }).catch((error) => {
            logger.error("Error during request sampling:", error);
            // Handle the error as needed
          });
    }
    */
});

mcpServer.on("disconnect", (event) => {
    logger.info("Client disconnected:", event.session);
});

// Graceful shutdown handling

/**
 * Handles SIGTERM and SIGINT signals for graceful server shutdown.
 */
process.on('SIGTERM', () => {
    logger.info('SIGTERM signal received. Closing server...');
    mcpServer.stop().then(() => { // Assuming a stop method exists
         logger.info('Server closed.');
         process.exit(0);
    }).catch(err => { // Add catch block for stop() errors
        logger.error('Error during mcpServer.stop() for SIGTERM:', { error: err });
        process.exit(1);
    });
});

process.on('SIGINT', () => {
     logger.info('SIGINT signal received. Closing server...');
    mcpServer.stop().then(() => {
         logger.info('Server closed.');
         process.exit(0);
    }).catch(err => { // Add catch block for stop() errors
        logger.error('Error during mcpServer.stop() for SIGINT:', { error: err });
        process.exit(1);
    });
});