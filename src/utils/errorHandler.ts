// =============================================================================
/**
 * @file Provides utility functions for handling and formatting errors.
 */
import { ZodError } from 'zod';
import { ErrorResponse } from '@/types/common.types';
import logger from './logger';

/**
 * Creates a standardized error response object.
 * @param code - A specific error code string.
 * @param message - A user-friendly error message.
 * @param details - Optional additional details about the error.
 * @returns An ErrorResponse object.
 */
export function createErrorResponse(code: string, message: string, details?: unknown): ErrorResponse {
  logger.error(`Error [${code}]: ${message}`, { details });
  return {
    success: false,
    error: {
      code,
      message,
      details,
    },
  };
}

/**
 * Handles errors caught in try-catch blocks, converting them to standard ErrorResponse.
 * @param error - The error object caught.
 * @param defaultCode - A default error code if the error type is unknown.
 * @returns An ErrorResponse object.
 */
export function handleToolError(error: unknown, defaultCode = 'INTERNAL_SERVER_ERROR'): ErrorResponse {
  if (error instanceof ZodError) {
    // Handle validation errors specifically
    return createErrorResponse('VALIDATION_ERROR', 'Input validation failed.', error.errors);
  } else if (error instanceof Error) {
    // Handle standard JavaScript errors
    // Check for specific error types if needed (e.g., Office API errors)
    if (error.name === 'OfficeExtension.Error') { // Example check for Office JS errors
        return createErrorResponse('OFFICE_API_ERROR', error.message, { debugInfo: (error as any).debugInfo });
    }
    return createErrorResponse(defaultCode, error.message, { stack: error.stack });
  } else {
    // Handle unknown errors
    return createErrorResponse(defaultCode, 'An unexpected error occurred.', error);
  }
}