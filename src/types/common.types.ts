/**
 * @file Defines common types and interfaces used across the MS Office MCP server.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import * as Zod from 'zod';
import { Context } from 'fastmcp'; // Import only Context
export { Context as FastMCPContext }; // Export with alias

/**
 * Represents the standard structure for a successful tool response.
 * @template T The type of the data payload. Defaults to `unknown`.
 */
export interface SuccessResponse<T = unknown> {
  /** Indicates if the operation was successful. Always `true` for a success response. */
  success: true;
  /** The data payload returned by the tool. Type is `T`. */
  data: T;
  /** An optional message providing additional information about the success. */
  message?: string;
}

/**
 * Represents the standard structure for an error response.
 */
export interface ErrorResponse {
  /** Indicates if the operation was successful. Always `false` for an error response. */
  success: false;
  /** Details about the error that occurred. */
  error: {
    /** A specific error code string (e.g., 'VALIDATION_ERROR', 'FILE_NOT_FOUND', 'OFFICE_API_ERROR'). */
    code: string;
    /** A user-friendly error message. */
    message: string;
    /** Optional additional details about the error. Can be any type (`unknown`). */
    details?: unknown;
  };
}

/**
 * Union type for API responses, representing either a successful outcome or an error.
 * @template T The type of the data payload in case of success. Defaults to `unknown`.
 */
export type ApiResponse<T = unknown> = SuccessResponse<T> | ErrorResponse;

/**
 * Interface for parameters passed to FastMCP tool handlers.
 * This is a generic type allowing any string key with any value.
 * Specific tool implementations should define more precise Zod schemas for validation.
 */
export interface ToolRequestParams {
  [key: string]: any; // Tool-specific parameters. Using `any` here as parameters are tool-dependent and validated by schema.
}


/**
 * Defines the structure for a FastMCP resource (tool or resource).
 * Includes the path, handler function, optional schema, completions, and description.
 */
export interface McpResource {
  /** The unique path identifier for the resource (e.g., 'word/styles/apply' or 'memory/ai_assistant_guide/read'). */
  path: string;
  /**
   * The handler function that executes the resource's logic.
   * It receives tool-specific parameters and the optional FastMCP context.
   * @param params - Tool-specific parameters.
   * @param context - The FastMCP context, provided for tools. Optional as resources might not always have context.
   * @returns A promise resolving to an ApiResponse.
   */
  handler: (params: ToolRequestParams, context?: Context<any>) => Promise<ApiResponse<any>>; // Use Context<any> for flexibility with context data
  /** An optional Zod schema for validating the input parameters (`params`) for tools. */
  schema?: Zod.ZodSchema<any>; // Optional Zod schema for input validation (for tools). Using `any` as schema structure is tool-specific.
  /**
   * Optional function to provide completion suggestions for tool parameters.
   * @param context - The FastMCP context.
   * @returns A promise resolving to a record of completion suggestions.
   */
  completions?: (context?: Context<any>) => Promise<Record<string, any[]>>; // Use Context<any> for flexibility with context data. Using `any[]` as completion values are tool-specific.
  /** A brief description of the tool/resource for documentation or AI assistance. */
  description?: string;
}