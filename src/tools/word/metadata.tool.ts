/**
 * @file Implements the 'word/metadata' tool using COM Interop for managing document properties and comments in Word documents.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse, FastMCPContext } from '../../types/common.types'; // Normalized relative path
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Normalized relative path
import { handleToolError, createErrorResponse } from '../../utils/errorHandler'; // Normalized relative path
import { validateFilePath } from '../../utils/security'; // Normalized relative path
import logger from '../../utils/logger'; // Normalized relative path

// Schema for the 'set' operation
/**
 * Schema for the 'set' operation, requiring a property name and value.
 */
const setOperationSchema = z.object({
  /** The path to the Word document. */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  operation: z.literal('set'),
  /** The name of the property to set (e.g., "Author", "Title"). */
  propertyName: z.string().min(1, 'Property name is required for set operation.').describe('The name of the property to set (e.g., "Author", "Title").'),
  /** The value to set for the property. Can be any type. */
  propertyValue: z.any().describe('The value to set for the property.'),
});

// Schema for the 'get' operation
/**
 * Schema for the 'get' operation, allowing optional property name or comment index.
 * Validation that at least one is provided is handled in the handler.
 */
const getOperationSchema = z.object({
  /** The path to the Word document. */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  operation: z.literal('get'),
  /** The name of the property to get. If omitted, gets all properties. */
  propertyName: z.string().optional().describe('The name of the property to get. If omitted, gets all properties.'),
  /** The 1-based index of the comment to get. If omitted, gets all comments. */
  commentIndex: z.number().int().positive('Comment index must be a positive integer.').optional().describe('The 1-based index of the comment to get. If omitted, gets all comments.'),
});


// Schema for the 'add' operation (comments)
/**
 * Schema for the 'add' operation (comments), requiring comment text and a range.
 */
const addOperationSchema = z.object({
  /** The path to the Word document. */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  operation: z.literal('add'),
  /** The text of the comment to add. */
  commentText: z.string().min(1, 'Comment text is required for add operation.').describe('The text of the comment to add.'),
  /** The range in the document where the comment should be added (e.g., { start: 10, end: 20 }). */
  range: z.object({
    start: z.number().int().positive('Range start must be a positive integer.').describe('The starting character index of the range.'),
    end: z.number().int().positive('Range end must be a positive integer.').describe('The ending character index of the range.'),
  }).describe('The range in the document where the comment should be added.'),
});

// Schema for the 'remove' operation (comments)
/**
 * Schema for the 'remove' operation (comments), requiring the comment index.
 */
const removeOperationSchema = z.object({
  /** The path to the Word document. */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  operation: z.literal('remove'),
  /** The 1-based index of the comment to remove. */
  commentIndex: z.number().int().positive('Comment index must be a positive integer for remove operation.').describe('The 1-based index of the comment to remove.'),
});

// Schema for the 'manage' operation (list properties/coments)
/**
 * Schema for the 'manage' operation, used for listing all properties and comments.
 */
const manageOperationSchema = z.object({
  /** The path to the Word document. */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  operation: z.literal('manage'),
  // No additional parameters required for listing
});

// Main input schema that combines all operation schemas using discriminatedUnion
/**
 * Zod schema for the input parameters of the 'word/metadata' tool,
 * using discriminated union based on the 'operation' field.
 */
const metadataInputSchema = z.discriminatedUnion('operation', [
  setOperationSchema,
  getOperationSchema,
  addOperationSchema,
  removeOperationSchema,
  manageOperationSchema,
]);

/**
 * @tool word/metadata
 * @description Manage document properties and comments in Word documents.
 * @param {ToolRequestParams} params - Input parameters validated against `metadataInputSchema`.
 * @param {FastMCPContext} [context] - The FastMCP context (optional), providing logging.
 * @returns {Promise<ApiResponse<any>>} - Result of the operation.
 * @throws {Error} If validation fails, a COM error occurs, or an operation fails.
 */
const metadataTool: McpResource = { // Changed type to McpResource
  path: 'word/metadata', // Define the path here
  description: 'Manage document properties and comments in Word documents.',
  schema: metadataInputSchema, // Use schema for input validation
  handler: async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> => { // Adjusted handler signature
    const log = context?.log ?? logger; // Use context logger or fallback

    let wordApp: any = null;
    let doc: any = null;
    let officeAppInstance: any = null; // To manage the application instance lifecycle

    try {
      const validatedParams = metadataInputSchema.parse(params); // Validate params
      const { filePath, operation } = validatedParams;

      log.info(`Executing word/metadata operation: '${operation}' on file: ${filePath}`);

      try {
        officeAppInstance = await getOfficeApplication('Word.Application');
        wordApp = officeAppInstance.app;
        wordApp.Visible = false; // Run in background
        wordApp.DisplayAlerts = 0; // wdAlertsNone = 0
        doc = wordApp.Documents.Open(filePath);
        log.debug(`Document opened successfully: ${filePath}`);

        switch (operation) {
          case 'set': {
            const { propertyName, propertyValue } = validatedParams; // Use validatedParams
            log.info(`Setting property '${propertyName}' to '${propertyValue}'`);
            // Implement set operation for properties (Builtin or Custom)
            try {
              // Try Builtin properties first
              const builtinProps = doc.BuiltinDocumentProperties;
              let propFound = false;
              for (let i = 1; i <= builtinProps.Count; i++) {
                const prop = builtinProps.Item(i);
                if (prop.Name === propertyName) {
                  prop.Value = propertyValue;
                  propFound = true;
                  log.debug(`Builtin property '${propertyName}' set.`);
                  break;
                }
              }

              if (!propFound) {
                // Try Custom properties
                const customProps = doc.CustomDocumentProperties;
                try {
                  const prop = customProps.Item(propertyName);
                  prop.Value = propertyValue;
                  propFound = true;
                  log.debug(`Custom property '${propertyName}' updated.`);
                } catch (e: any) {
                  // Property not found in Custom properties either
                   log.debug(`Custom property '${propertyName}' not found for update.`);
                }
              }

              if (propFound) {
                doc.Save();
                log.info(`Document saved after setting property.`);
                return { success: true, data: { message: `Property '${propertyName}' set successfully.` } }; // Return ApiResponse
              } else {
                // If property not found in Builtin or Custom, maybe create a Custom one?
                // For now, just report not found. Creating requires specifying type.
                log.warn(`Property '${propertyName}' not found. Cannot set.`);
                return createErrorResponse('NOT_FOUND', `Property '${propertyName}' not found. Cannot set.`); // Use createErrorResponse
              }

            } catch (error: any) {
              log.error(`Error setting property '${propertyName}': ${error.message}`, { error });
              return handleToolError(error, 'METADATA_SET_ERROR'); // Use handleToolError
            }
          }
          case 'get': {
            const { propertyName, commentIndex } = validatedParams; // Use validatedParams
            // Handle validation that at least one is provided here
            if (propertyName === undefined && commentIndex === undefined) {
                 log.warn("Neither propertyName nor commentIndex provided for 'get' operation.");
                 return createErrorResponse('VALIDATION_ERROR', "Either propertyName or commentIndex is required for the 'get' operation.");
            }

            if (propertyName) {
              // Get a specific property
              log.info(`Getting property: '${propertyName}'`);
              try {
                // Try Builtin properties
                const builtinProps = doc.BuiltinDocumentProperties;
                for (let i = 1; i <= builtinProps.Count; i++) {
                  const prop = builtinProps.Item(i);
                  if (prop.Name === propertyName) {
                    log.debug(`Found Builtin property '${propertyName}'.`);
                    return { success: true, data: { name: prop.Name, value: prop.Value } }; // Return ApiResponse
                  }
                }

                // Try Custom properties
                const customProps = doc.CustomDocumentProperties;
                try {
                  const prop = customProps.Item(propertyName);
                  log.debug(`Found Custom property '${propertyName}'.`);
                  return { success: true, data: { name: prop.Name, value: prop.Value } }; // Return ApiResponse
                } catch (e: any) {
                   log.debug(`Custom property '${propertyName}' not found.`);
                  // Property not found in Custom properties either
                }

                log.warn(`Property '${propertyName}' not found.`);
                return createErrorResponse('NOT_FOUND', `Property '${propertyName}' not found.`); // Use createErrorResponse

              } catch (error: any) {
                log.error(`Error getting property '${propertyName}': ${error.message}`, { error });
                return handleToolError(error, 'METADATA_GET_ERROR'); // Use handleToolError
              }
            } else if (commentIndex) {
              // Get a specific comment
              log.info(`Getting comment at index: ${commentIndex}`);
              try {
                const comments = doc.Comments;
                if (commentIndex > 0 && commentIndex <= comments.Count) {
                  const comment = comments.Item(commentIndex);
                  log.debug(`Found comment at index ${commentIndex}.`);
                  return { // Return ApiResponse
                    success: true,
                    data: {
                      index: commentIndex,
                      author: comment.Author,
                      initials: comment.Initials,
                      date: comment.Date,
                      text: comment.Range.Text,
                    }
                  };
                } else {
                  log.warn(`Comment index ${commentIndex} out of bounds. Document has ${comments.Count} comments.`);
                  return createErrorResponse('OUT_OF_BOUNDS', `Comment index ${commentIndex} out of bounds. Document has ${comments.Count} comments.`); // Use createErrorResponse
                }
              } catch (error: any) {
                log.error(`Error getting comment at index ${commentIndex}: ${error.message}`, { error });
                return handleToolError(error, 'METADATA_GET_ERROR'); // Use handleToolError
              }
            } else {
              // This case should not be reached due to the validation above, but included for safety
              log.error("Neither propertyName nor commentIndex provided for 'get' operation after validation.");
              return createErrorResponse('VALIDATION_ERROR', "Internal error: Validation failed for 'get' operation.");
            }
          }
          case 'add': {
            const { commentText, range } = validatedParams; // Use validatedParams
            log.info(`Adding comment with text "${commentText}" at range [${range.start}, ${range.end}]`);
            // Add a comment
            try {
              const docRange = doc.Range(range.start, range.end);
              if (!docRange) {
                log.warn(`Invalid range specified for adding comment: start=${range.start}, end=${range.end}`);
                return createErrorResponse('INVALID_RANGE', `Invalid range specified for adding comment: start=${range.start}, end=${range.end}`); // Use createErrorResponse
              }
              doc.Comments.Add(docRange, commentText);
              doc.Save();
              log.info(`Comment added and document saved.`);
              return { success: true, data: { message: 'Comment added successfully.' } }; // Return ApiResponse
            } catch (error: any) {
              log.error(`Error adding comment: ${error.message}`, { error });
              return handleToolError(error, 'METADATA_ADD_ERROR'); // Use handleToolError
            }
          }
          case 'remove': {
            const { commentIndex } = validatedParams; // Use validatedParams
            log.info(`Removing comment at index: ${commentIndex}`);
            // Remove a comment
            try {
              const comments = doc.Comments;
              if (commentIndex > 0 && commentIndex <= comments.Count) {
                comments.Item(commentIndex).Delete();
                doc.Save();
                log.info(`Comment ${commentIndex} removed and document saved.`);
                return { success: true, data: { message: `Comment ${commentIndex} removed successfully.` } }; // Return ApiResponse
              } else {
                log.warn(`Comment index ${commentIndex} out of bounds for removal. Document has ${comments.Count} comments.`);
                return createErrorResponse('OUT_OF_BOUNDS', `Comment index ${commentIndex} out of bounds. Document has ${comments.Count} comments.`); // Use createErrorResponse
              }
            } catch (error: any) {
              log.error(`Error removing comment at index ${commentIndex}: ${error.message}`, { error });
              return handleToolError(error, 'METADATA_REMOVE_ERROR'); // Use handleToolError
            }
          }
          case 'manage': {
            log.info('Listing all properties and comments.');
            // List all properties and comments
            const properties: { name: string; value: any }[] = [];
            const comments: { index: number; author: string; initials: string; date: any; text: string }[] = [];

            try {
              // List Builtin properties
              const builtinProps = doc.BuiltinDocumentProperties;
              for (let i = 1; i <= builtinProps.Count; i++) {
                const prop = builtinProps.Item(i);
                // Some properties might not have a value or might throw errors when accessed
                try {
                   properties.push({ name: prop.Name, value: prop.Value });
                } catch (e: any) {
                   log.warn(`Error accessing BuiltinDocumentProperty '${prop.Name}': ${e.message}`);
                   properties.push({ name: prop.Name, value: `[Error accessing value: ${e.message}]` });
                }
              }
              log.debug(`Listed ${properties.length} BuiltinDocumentProperties.`);
            } catch (error: any) {
               log.error(`Error listing BuiltinDocumentProperties: ${error.message}`, { error });
               properties.push({ name: 'BuiltinDocumentProperties', value: `Error: ${error.message}` });
            }

            try {
              // List Custom properties
              const customProps = doc.CustomDocumentProperties;
               for (let i = 1; i <= customProps.Count; i++) {
                const prop = customProps.Item(i);
                 try {
                   properties.push({ name: prop.Name, value: prop.Value });
                 } catch (e: any) {
                   log.warn(`Error accessing CustomDocumentProperty '${prop.Name}': ${e.message}`);
                   properties.push({ name: prop.Name, value: `[Error accessing value: ${e.message}]` });
                 }
               }
               log.debug(`Listed ${customProps.Count} CustomDocumentProperties.`);
            } catch (error: any) {
               log.error(`Error listing CustomDocumentProperties: ${error.message}`, { error });
               properties.push({ name: 'CustomDocumentProperties', value: `Error: ${error.message}` });
            }


            try {
              // List comments
              const docComments = doc.Comments;
              for (let i = 1; i <= docComments.Count; i++) {
                const comment = docComments.Item(i);
                comments.push({
                  index: i,
                  author: comment.Author,
                  initials: comment.Initials,
                  date: comment.Date,
                  text: comment.Range.Text,
                });
              }
              log.debug(`Listed ${docComments.Count} comments.`);
            } catch (error: any) {
               log.error(`Error listing Comments: ${error.message}`, { error });
               comments.push({ index: -1, author: 'Error', initials: '', date: null, text: `Error listing comments: ${error.message}` });
            }

            log.info('Metadata listing complete.');
            return { success: true, data: { properties, comments } }; // Return ApiResponse
          }
          default:
            // This case should not be reached due to discriminatedUnion, but included for safety
            log.error(`Unknown operation reached in switch: ${operation}`);
            return createErrorResponse('METADATA_ERROR', `Unknown operation: ${operation}`); // Use createErrorResponse
        }

      } catch (error: any) {
        // Handle errors that occur after validation but before or during COM interop
        log.error(`Error during COM interop for operation '${operation}': ${error.message}`, { error });
        return handleToolError(error, 'METADATA_COM_ERROR'); // Use handleToolError
      } finally {
        if (doc) {
          try {
            // Close document without saving unless it was saved in a specific operation
            // A more robust approach would track if save was called. For now, assume save was handled in operation.
            doc.Close(0); // wdDoNotSaveChanges = 0
            log.debug(`Document closed: ${filePath}`);
          } catch (closeError: any) {
            log.error(`Error closing document: ${closeError.message}`, { error: closeError });
          }
          releaseObject(doc);
        }
        if (officeAppInstance) {
          officeAppInstance.release();
          log.debug("Office application instance released.");
        }
      }
    } catch (error: any) {
        // Handle validation errors or other errors before COM interop
        log.error(`Input validation or unexpected error before COM interop: ${error.message}`, { error });
        return handleToolError(error, 'METADATA_VALIDATION_ERROR'); // Use handleToolError
    }
  },
};

export default metadataTool;