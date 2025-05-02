import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse, FastMCPContext } from '@/types/common.types'; // Import necessary types
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { handleToolError } from '@/utils/errorHandler'; // Import handleToolError

// Esquema base para todas las operaciones
const baseMetadataSchema = z.object({
  filePath: z.string().describe('The path to the Word document.'),
  operation: z.enum(['set', 'get', 'add', 'remove', 'manage']).describe('The operation to perform.'),
});

// Esquema para la operación 'set'
const setOperationSchema = baseMetadataSchema.extend({
  operation: z.literal('set'),
  propertyName: z.string().describe('The name of the property to set (e.g., "Author", "Title").'),
  propertyValue: z.any().describe('The value to set for the property.'),
});

// Esquema para la operación 'get'
const getOperationSchema = baseMetadataSchema.extend({
  operation: z.literal('get'),
  propertyName: z.string().optional().describe('The name of the property to get. If omitted, gets all properties.'),
  commentIndex: z.number().int().positive().optional().describe('The 1-based index of the comment to get. If omitted, gets all comments.'),
});

// Esquema para la operación 'add' (comentarios)
const addOperationSchema = baseMetadataSchema.extend({
  operation: z.literal('add'),
  commentText: z.string().describe('The text of the comment to add.'),
  range: z.object({
    start: z.number().int().positive().describe('The starting character index of the range.'),
    end: z.number().int().positive().describe('The ending character index of the range.'),
  }).describe('The range in the document where the comment should be added.'),
});

// Esquema para la operación 'remove' (comentarios)
const removeOperationSchema = baseMetadataSchema.extend({
  operation: z.literal('remove'),
  commentIndex: z.number().int().positive().describe('The 1-based index of the comment to remove.'),
});

// Esquema para la operación 'manage' (listar propiedades/comentarios)
const manageOperationSchema = baseMetadataSchema.extend({
  operation: z.literal('manage'),
  // No se requieren parámetros adicionales para listar
});

// Esquema de entrada principal que combina todos los esquemas de operación
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
 * @param filePath - The path to the Word document.
 * @param operation - The operation to perform ('set', 'get', 'add', 'remove', 'manage').
 * @param propertyName - (For 'set' and 'get') The name of the property to manage.
 * @param propertyValue - (For 'set') The value to set for the property.
 * @param commentText - (For 'add') The text of the comment to add.
 * @param range - (For 'add') The range in the document where the comment should be added (e.g., { start: 10, end: 20 }).
 * @param commentIndex - (For 'get' and 'remove') The 1-based index of the comment to manage.
 */
const metadataTool: McpResource = { // Changed type to McpResource
  path: 'word/metadata', // Define the path here
  description: 'Manage document properties and comments in Word documents.',
  schema: metadataInputSchema, // Use schema for input validation
  handler: async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> => { // Adjusted handler signature
    try {
      const validatedParams = metadataInputSchema.parse(params); // Validate params
      const { filePath, operation } = validatedParams;

      let wordApp: any = null;
      let doc: any = null;

      try {
        wordApp = await getOfficeApplication('Word.Application');
        doc = wordApp.Documents.Open(filePath);

        switch (operation) {
          case 'set': {
            const { propertyName, propertyValue } = validatedParams as z.infer<typeof setOperationSchema>; // Use validatedParams
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
                } catch (e) {
                  // Property not found in Custom properties either
                }
              }

              if (propFound) {
                doc.Save();
                return { success: true, data: { message: `Property '${propertyName}' set successfully.` } }; // Return ApiResponse
              } else {
                // If property not found in Builtin or Custom, maybe create a Custom one?
                // For now, just report not found. Creating requires specifying type.
                return handleToolError(new Error(`Property '${propertyName}' not found. Cannot set.`), 'METADATA_SET_ERROR'); // Use handleToolError
              }

            } catch (error: any) {
              return handleToolError(error, 'METADATA_SET_ERROR'); // Use handleToolError
            }
          }
          case 'get': {
            const { propertyName, commentIndex } = validatedParams as z.infer<typeof getOperationSchema>; // Use validatedParams
            if (propertyName) {
              // Get a specific property
              try {
                // Try Builtin properties
                const builtinProps = doc.BuiltinDocumentProperties;
                for (let i = 1; i <= builtinProps.Count; i++) {
                  const prop = builtinProps.Item(i);
                  if (prop.Name === propertyName) {
                    return { success: true, data: { name: prop.Name, value: prop.Value } }; // Return ApiResponse
                  }
                }

                // Try Custom properties
                const customProps = doc.CustomDocumentProperties;
                try {
                  const prop = customProps.Item(propertyName);
                  return { success: true, data: { name: prop.Name, value: prop.Value } }; // Return ApiResponse
                } catch (e) {
                  // Property not found in Custom properties either
                }

                return handleToolError(new Error(`Property '${propertyName}' not found.`), 'METADATA_GET_ERROR'); // Use handleToolError

              } catch (error: any) {
                return handleToolError(error, 'METADATA_GET_ERROR'); // Use handleToolError
              }
            } else if (commentIndex) {
              // Get a specific comment
              try {
                const comments = doc.Comments;
                if (commentIndex > 0 && commentIndex <= comments.Count) {
                  const comment = comments.Item(commentIndex);
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
                  return handleToolError(new Error(`Comment index ${commentIndex} out of bounds. Document has ${comments.Count} comments.`), 'METADATA_GET_ERROR'); // Use handleToolError
                }
              } catch (error: any) {
                return handleToolError(error, 'METADATA_GET_ERROR'); // Use handleToolError
              }
            } else {
              // Get all properties and comments (manage-like behavior)
              // This falls through to the 'manage' case logic below
            }
          }
          case 'add': {
            const { commentText, range } = validatedParams as z.infer<typeof addOperationSchema>; // Use validatedParams
            // Add a comment
            try {
              const docRange = doc.Range(range.start, range.end);
              if (!docRange) {
                return handleToolError(new Error(`Invalid range specified for adding comment: start=${range.start}, end=${range.end}`), 'METADATA_ADD_ERROR'); // Use handleToolError
              }
              doc.Comments.Add(docRange, commentText);
              doc.Save();
              return { success: true, data: { message: 'Comment added successfully.' } }; // Return ApiResponse
            } catch (error: any) {
              return handleToolError(error, 'METADATA_ADD_ERROR'); // Use handleToolError
            }
          }
          case 'remove': {
            const { commentIndex } = validatedParams as z.infer<typeof removeOperationSchema>; // Use validatedParams
            // Remove a comment
            try {
              const comments = doc.Comments;
              if (commentIndex > 0 && commentIndex <= comments.Count) {
                comments.Item(commentIndex).Delete();
                doc.Save();
                return { success: true, data: { message: `Comment ${commentIndex} removed successfully.` } }; // Return ApiResponse
              } else {
                return handleToolError(new Error(`Comment index ${commentIndex} out of bounds. Document has ${comments.Count} comments.`), 'METADATA_REMOVE_ERROR'); // Use handleToolError
              }
            } catch (error: any) {
              return handleToolError(error, 'METADATA_REMOVE_ERROR'); // Use handleToolError
            }
          }
          case 'manage': {
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
                } catch (e) {
                   properties.push({ name: prop.Name, value: '[Error accessing value]' });
                }
              }
            } catch (error: any) {
               console.error(`Error listing BuiltinDocumentProperties: ${error.message}`);
               properties.push({ name: 'BuiltinDocumentProperties', value: `Error: ${error.message}` });
            }

            try {
              // List Custom properties
              const customProps = doc.CustomDocumentProperties;
               for (let i = 1; i <= customProps.Count; i++) {
                const prop = customProps.Item(i);
                 try {
                   properties.push({ name: prop.Name, value: prop.Value });
                 } catch (e) {
                   properties.push({ name: prop.Name, value: '[Error accessing value]' });
                 }
              }
            } catch (error: any) {
               console.error(`Error listing CustomDocumentProperties: ${error.message}`);
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
            } catch (error: any) {
               console.error(`Error listing Comments: ${error.message}`);
               comments.push({ index: -1, author: 'Error', initials: '', date: null, text: `Error listing comments: ${error.message}` });
            }


            return { success: true, data: { properties, comments } }; // Return ApiResponse
          }
          default:
            // This case should not be reached due to discriminatedUnion, but included for safety
            return handleToolError(new Error(`Unknown operation: ${operation}`), 'METADATA_ERROR'); // Use handleToolError
        }

      } catch (error: any) {
        return handleToolError(error, 'METADATA_ERROR'); // Use handleToolError
      } finally {
        if (doc) {
          try {
            // Close document without saving unless it was saved in a specific operation
            // A more robust approach would track if save was called. For now, assume save was handled in operation.
            doc.Close(0); // wdDoNotSaveChanges = 0
          } catch (closeError: any) {
            console.error(`Error closing document: ${closeError.message}`);
          }
          releaseObject(doc);
        }
        if (wordApp) {
          // Decide whether to quit the application. Quitting might close a user's open instance.
          // A safer approach might be to only quit if we know we started the instance.
          // For now, let's not quit the application automatically.
          // wordApp.Quit();
          releaseObject(wordApp); // Release the reference
        }
      }
    } catch (error: any) {
        // Handle validation errors or other errors before COM interop
        return handleToolError(error, 'METADATA_VALIDATION_ERROR'); // Use handleToolError
    }
  },
};

export default metadataTool;