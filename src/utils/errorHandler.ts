/**
 * @file Provides utility functions for handling and formatting errors.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { ZodError } from 'zod';
import { ErrorResponse } from '@/types/common.types';
import logger from './logger';

/**
 * Creates a standardized error response object.
 * Logs the error using the configured logger.
 * @param code - A specific error code string (e.g., 'VALIDATION_ERROR', 'INTERNAL_SERVER_ERROR').
 * @param message - A user-friendly error message.
 * @param details - Optional additional details about the error. Can be any type.
 * @returns An ErrorResponse object indicating failure.
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
 * Handles errors caught in try-catch blocks, converting them to a standard ErrorResponse format.
 * Provides specific handling for Zod validation errors and standard Error objects.
 * @param error - The error object caught. Can be of any type (`unknown`).
 * @param defaultCode - A default error code string if the error type is unknown or no specific code is determined. Defaults to 'INTERNAL_SERVER_ERROR'.
 * @returns An ErrorResponse object representing the handled error.
 */
export function handleToolError(error: unknown, defaultCode = 'INTERNAL_SERVER_ERROR'): ErrorResponse {
  if (error instanceof ZodError) {
    // Handle validation errors specifically
    return createErrorResponse('VALIDATION_ERROR', 'Input validation failed.', error.errors);
  } else if (error instanceof Error) {
    // Handle standard JavaScript errors
    // Example check for Office JS errors - the type assertion `as any` is used here because `OfficeExtension.Error` might not be universally typed.
    // If `OfficeExtension.Error` type is available, replace `any` with the correct type.
    if (error.name === 'OfficeExtension.Error') {
        return createErrorResponse('OFFICE_API_ERROR', error.message, { debugInfo: (error as any).debugInfo });
    }
    return createErrorResponse(defaultCode, error.message, { stack: error.stack });
  } else {
    // Handle unknown errors
    // The 'error' parameter is of type 'unknown', so we pass it directly as details.
    return createErrorResponse(defaultCode, 'An unexpected error occurred.', error);
  }
}