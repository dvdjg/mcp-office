import * as Zod from 'zod';
import { Context as FastMCPContext } from 'fastmcp'; // Import FastMCP Context
/**
 * @file Defines common types and interfaces used across the MCP server.
 */

/**
 * Represents the standard structure for a successful tool response.
 */
export interface SuccessResponse<T = unknown> {
  success: true;
  data: T;
  message?: string;
}

/**
 * Represents the standard structure for an error response.
 */
export interface ErrorResponse {
  success: false;
  error: {
    code: string; // e.g., 'VALIDATION_ERROR', 'FILE_NOT_FOUND', 'OFFICE_API_ERROR'
    message: string;
    details?: unknown; // Additional error details
  };
}

/**
 * Union type for API responses.
 */
export type ApiResponse<T = unknown> = SuccessResponse<T> | ErrorResponse;

/**
 * Interface for parameters passed to FastMCP tool handlers.
 * Adjust based on FastMCP's actual request object structure.
 */
export interface ToolRequestParams {
  [key: string]: any; // Tool-specific parameters
}

// Remove the old ToolContext interface as it's replaced by FastMCPContext
// export interface ToolContext {
//   userId?: string;
//   permissions?: string[];
//   // Add other relevant context properties
// }

/**
 * Defines the structure for a FastMCP resource (tool or resource).
 * The handler now accepts the FastMCP context.
 */
export interface McpResource {
  path: string; // e.g., 'word/styles/apply' or 'memory/ai_assistant_guide/read'
  // Handler now expects FastMCPContext<undefined> (aliased as ServerContext in index.ts)
  handler: (params: ToolRequestParams, context: FastMCPContext<undefined>) => Promise<ApiResponse<any>>;
  schema?: Zod.ZodSchema<any>; // Optional Zod schema for input validation (for tools)
  // Completions might need adjustment if they also need the FastMCP context
  completions?: (context: FastMCPContext<undefined>) => Promise<Record<string, any[]>>; // For AI completions
  description?: string; // Tool/Resource description for documentation/AI
}