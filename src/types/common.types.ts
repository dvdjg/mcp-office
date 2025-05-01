import * as Zod from 'zod';
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

/**
 * Interface for the context passed to tool handlers, potentially
 * containing user info, permissions, etc.
 * Adjust based on FastMCP's actual context object structure.
 */
export interface ToolContext {
  userId?: string;
  permissions?: string[];
  // Add other relevant context properties
}

/**
 * Defines the structure for a FastMCP resource (tool).
 */
export interface McpResource {
  path: string; // e.g., 'word/styles/apply'
  handler: (params: ToolRequestParams, context?: ToolContext) => Promise<ApiResponse<any>>;
  schema?: Zod.ZodSchema<any>; // Optional Zod schema for input validation
  completions?: (context?: ToolContext) => Promise<Record<string, any[]>>; // For AI completions
  description?: string; // Tool description for documentation/AI
}