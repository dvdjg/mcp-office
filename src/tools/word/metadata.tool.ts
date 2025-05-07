/**
 * @file Implements the 'word/metadata' tool using COM Interop for managing document properties and comments in Word documents.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import fs from 'fs-extra';
import { Document, Packer, IPropertiesOptions } from 'docx';
// Note: Mammoth might be used for comment extraction if docx doesn't suffice for reading them.
import * as mammoth from 'mammoth';
import { McpResource, ToolRequestParams, ApiResponse, FastMCPContext } from '../../types/common.types.js';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js';
import { handleToolError, createErrorResponse } from '../../utils/errorHandler.js';
import { validateFilePath } from '../../utils/security.js';
import logger from '../../utils/logger.js';

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
  useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
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
  useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
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
  useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
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
  useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
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
  useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."),
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
      const validatedParams = metadataInputSchema.parse(params);
      const { filePath, operation, useComInterop } = validatedParams as any; // Cast to any to access useComInterop
      const safeFilePath = validateFilePath(filePath);

      log.info(`Executing word/metadata operation: '${operation}' on file: ${safeFilePath}, useComInterop: ${useComInterop}`);

      if (useComInterop) {
        log.info(`[word/metadata] Using COM Interop path for operation: ${operation}`);
        officeAppInstance = await getOfficeApplication('Word.Application');
        wordApp = officeAppInstance; // getOfficeApplication returns the app directly
        // wordApp.Visible = false; // Default behavior or handled by getOfficeApplication
        // wordApp.DisplayAlerts = 0; // Consider if this is globally desired

        try {
            doc = wordApp.Documents.Open(safeFilePath); // filePath is already validated as safeFilePath
            log.debug(`COM: Document opened successfully: ${safeFilePath}`);

            switch (operation) {
              case 'set': {
                const { propertyName, propertyValue } = validatedParams as z.infer<typeof setOperationSchema>;
                log.info(`COM: Setting property '${propertyName}' to '${propertyValue}'`);
                // ... (existing COM 'set' logic - lines 146-183)
                // For brevity, this detailed logic is not duplicated here but assumed to be the same.
                // Ensure it uses `doc.Save()` and returns ApiResponse.
                // --- Placeholder for COM set logic ---
                const builtinProps = doc.BuiltinDocumentProperties;
                let propFound = false;
                for (let i = 1; i <= builtinProps.Count; i++) {
                    const prop = builtinProps.Item(i);
                    if (prop.Name === propertyName) {
                        prop.Value = propertyValue;
                        propFound = true;
                        releaseObject(prop); break;
                    }
                    releaseObject(prop);
                }
                if (!propFound) {
                    const customProps = doc.CustomDocumentProperties;
                    try {
                        const prop = customProps.Item(propertyName);
                        prop.Value = propertyValue;
                        propFound = true;
                        releaseObject(prop);
                    } catch (e) { /* not found */ }
                    releaseObject(customProps);
                }
                releaseObject(builtinProps);
                if (propFound) {
                    doc.Save();
                    return { success: true, data: { message: `COM: Property '${propertyName}' set.` } };
                } else {
                    return createErrorResponse('NOT_FOUND', `COM: Property '${propertyName}' not found.`);
                }
                // --- End Placeholder ---
              }
              case 'get': {
                const { propertyName, commentIndex } = validatedParams as z.infer<typeof getOperationSchema>;
                if (propertyName === undefined && commentIndex === undefined) {
                    return createErrorResponse('VALIDATION_ERROR', "COM: Either propertyName or commentIndex required for 'get'.");
                }
                // ... (existing COM 'get' logic for properties and comments - lines 198-260)
                // --- Placeholder for COM get logic ---
                if (propertyName) {
                    const builtinProps = doc.BuiltinDocumentProperties;
                    for (let i = 1; i <= builtinProps.Count; i++) {
                        const prop = builtinProps.Item(i);
                        if (prop.Name === propertyName) {
                            const val = prop.Value; releaseObject(prop); releaseObject(builtinProps);
                            return { success: true, data: { name: propertyName, value: val } };
                        }
                        releaseObject(prop);
                    }
                    releaseObject(builtinProps);
                    const customProps = doc.CustomDocumentProperties;
                     try {
                        const prop = customProps.Item(propertyName);
                        const val = prop.Value; releaseObject(prop); releaseObject(customProps);
                        return { success: true, data: { name: propertyName, value: val } };
                    } catch (e) { /* not found */ }
                    releaseObject(customProps);
                    return createErrorResponse('NOT_FOUND', `COM: Property '${propertyName}' not found.`);
                } else if (commentIndex) {
                    const comments = doc.Comments;
                    if (commentIndex > 0 && commentIndex <= comments.Count) {
                        const comment = comments.Item(commentIndex);
                        const data = { index: commentIndex, author: comment.Author, initials: comment.Initials, date: comment.Date, text: comment.Range.Text };
                        releaseObject(comment); releaseObject(comments);
                        return { success: true, data };
                    } else {
                        releaseObject(comments);
                        return createErrorResponse('OUT_OF_BOUNDS', `COM: Comment index ${commentIndex} out of bounds.`);
                    }
                }
                // --- End Placeholder ---
                return createErrorResponse('INTERNAL_ERROR', "COM: 'get' operation failed unexpectedly."); // Should be caught by specific logic
              }
              case 'add': {
                const { commentText, range } = validatedParams as z.infer<typeof addOperationSchema>;
                // ... (existing COM 'add' comment logic - lines 266-279)
                // --- Placeholder for COM add comment logic ---
                const docRange = doc.Range(range.start, range.end);
                if (!docRange) return createErrorResponse('INVALID_RANGE', `COM: Invalid range for comment.`);
                doc.Comments.Add(docRange, commentText);
                releaseObject(docRange);
                doc.Save();
                return { success: true, data: { message: 'COM: Comment added.' } };
                // --- End Placeholder ---
              }
              case 'remove': {
                const { commentIndex } = validatedParams as z.infer<typeof removeOperationSchema>;
                // ... (existing COM 'remove' comment logic - lines 285-299)
                // --- Placeholder for COM remove comment logic ---
                const comments = doc.Comments;
                if (commentIndex > 0 && commentIndex <= comments.Count) {
                    comments.Item(commentIndex).Delete();
                    releaseObject(comments);
                    doc.Save();
                    return { success: true, data: { message: `COM: Comment ${commentIndex} removed.` } };
                } else {
                    releaseObject(comments);
                    return createErrorResponse('OUT_OF_BOUNDS', `COM: Comment index ${commentIndex} out of bounds for removal.`);
                }
                // --- End Placeholder ---
              }
              case 'manage': {
                // ... (existing COM 'manage' logic for listing all - lines 304-365)
                // --- Placeholder for COM manage logic ---
                const properties: any[] = []; const commentsData: any[] = []; // Simplified
                const builtinProps = doc.BuiltinDocumentProperties;
                for (let i = 1; i <= builtinProps.Count; i++) { const p = builtinProps.Item(i); properties.push({name: p.Name, value: p.Value }); releaseObject(p); }
                releaseObject(builtinProps);
                const customProps = doc.CustomDocumentProperties;
                for (let i = 1; i <= customProps.Count; i++) { const p = customProps.Item(i); properties.push({name: p.Name, value: p.Value }); releaseObject(p); }
                releaseObject(customProps);
                const docComments = doc.Comments;
                for (let i = 1; i <= docComments.Count; i++) { const c = docComments.Item(i); commentsData.push({index: i, author: c.Author, text: c.Range.Text }); releaseObject(c); }
                releaseObject(docComments);
                return { success: true, data: { properties, comments: commentsData } };
                // --- End Placeholder ---
              }
              default:
                log.error(`COM: Unknown operation: ${operation}`);
                return createErrorResponse('METADATA_ERROR_COM', `COM: Unknown operation: ${operation}`);
            }
        } catch (error: any) {
            log.error(`[word/metadata] COM Error during operation '${operation}': ${error.message}`, { error });
            return handleToolError(error, 'METADATA_COM_OPERATION_ERROR');
        } finally {
            if (doc) {
                try { doc.Close(false); } catch (e: any) { log.warn(`COM: Error closing document: ${e.message}`); }
                releaseObject(doc);
            }
            if (wordApp) releaseObject(wordApp); // Changed from officeAppInstance.release()
            log.debug("COM: Office objects released.");
        }
      } else {
        // Library path
        log.info(`[word/metadata] Using Library path for operation: ${operation}`);
        const fileBuffer = await fs.readFile(safeFilePath);

        switch (operation) {
            case 'set': {
                const { propertyName, propertyValue } = validatedParams as z.infer<typeof setOperationSchema>;
                log.warn(`Library path: Setting metadata property '${propertyName}' is complex for existing files. This feature is primarily for new file creation or COM interop.`);
                // For docx.js, properties are typically set on Document creation.
                // Modifying existing properties involves unzipping, editing XML (core.xml, app.xml), and rezipping.
                // Packer.patch might be an option for specific, known XML paths.
                // For now, this will be marked as not fully implemented for existing files.
                if (await fs.pathExists(safeFilePath)) {
                     return createErrorResponse('NOT_IMPLEMENTED_LIB_MODIFY', `Library path: Setting properties on existing files is not fully implemented. Use COM Interop or create a new file with properties.`);
                }
                // If creating a new file
                const docCreationProps: any = {}; // Build as a plain object
                const lowerPropName = propertyName.toLowerCase();

                if (lowerPropName === 'title') docCreationProps.title = propertyValue;
                else if (lowerPropName === 'subject') docCreationProps.subject = propertyValue;
                else if (lowerPropName === 'creator' || lowerPropName === 'author') docCreationProps.creator = propertyValue;
                else if (lowerPropName === 'keywords') docCreationProps.keywords = propertyValue;
                else if (lowerPropName === 'description') docCreationProps.description = propertyValue;
                else if (lowerPropName === 'lastmodifiedby') docCreationProps.lastModifiedBy = propertyValue;
                else if (lowerPropName === 'revision') docCreationProps.revision = propertyValue.toString();
                // Add other settable properties from IPropertiesOptions as needed
                else {
                    return createErrorResponse('NOT_SUPPORTED_LIB', `Library path: Property '${propertyName}' is not a standard core property supported for setting via new document creation.`);
                }
                // Ensure sections is always present for a valid document
                docCreationProps.sections = [{ children: [] }];

                const newDoc = new Document(docCreationProps as IPropertiesOptions);
                const buffer = await Packer.toBuffer(newDoc);
                await fs.writeFile(safeFilePath, buffer);
                return { success: true, data: { message: `Library: New document created with property '${propertyName}' set.` } };
            }
            case 'get': {
                const { propertyName, commentIndex } = validatedParams as z.infer<typeof getOperationSchema>;
                 if (propertyName === undefined && commentIndex === undefined) {
                    return createErrorResponse('VALIDATION_ERROR', "Library: Either propertyName or commentIndex required for 'get'.");
                }
                if (propertyName) {
                    log.warn("Library path: Reading specific document properties from an existing file is not directly supported by 'docx' in a simple way. This feature is COM-dependent for reliability.");
                    // Attempting to read docx properties would require unzipping and parsing XML (e.g. docProps/core.xml)
                    // For now, return not implemented for property get.
                    return createErrorResponse('NOT_IMPLEMENTED_LIB_PROP_GET', `Library path: Getting specific document property '${propertyName}' is not implemented. Use COM Interop.`);
                } else if (commentIndex) {
                    // Comment extraction with Mammoth
                    const { value: html } = await mammoth.convertToHtml({ buffer: fileBuffer });
                    // Basic HTML parsing for comments (very fragile, specific to Mammoth's HTML output for comments)
                    // Mammoth typically converts Word comments to HTML footnotes or similar structures.
                    // A more robust parser would be needed for general HTML.
                    // This regex is highly specific and might break with Mammoth updates.
                    const commentRegex = /<a href="#footnote-ref-(\d+)" id="footnote-(\d+)"><sup>\[\d+\]<\/sup><\/a>.*?<aside id="footnote-ref-\d+"[^>]*>.*?<p>(.*?)<\/p>/gs;
                    const comments = [];
                    let match;
                    let internalCommentIdx = 0;
                    while((match = commentRegex.exec(html)) !== null) {
                        internalCommentIdx++;
                        // We use internalCommentIdx because Mammoth's footnote IDs might not be sequential or start from 1 as user expects.
                        comments.push({ index: internalCommentIdx, text: match[3].trim().replace(/<[^>]*>/g, "") }); // Strip any inner HTML tags from comment text
                    }

                    if (commentIndex > 0 && commentIndex <= comments.length) {
                        const foundComment = comments.find(c => c.index === commentIndex);
                        if (foundComment) {
                           return { success: true, data: { index: commentIndex, text: foundComment.text } };
                        } else {
                           // This case should ideally not be hit if logic is correct, but as a safeguard:
                           return createErrorResponse('NOT_FOUND_LIB', `Library: Comment with effective index ${commentIndex} not found after parsing.`);
                        }
                    } else {
                         return createErrorResponse('OUT_OF_BOUNDS_LIB', `Library: Comment index ${commentIndex} out of bounds. Found ${comments.length} comments (via HTML parsing).`);
                    }
                }
                return createErrorResponse('INTERNAL_ERROR_LIB', "Library: 'get' operation failed unexpectedly.");
            }
            case 'add':
            case 'remove':
                log.warn(`Library path: Comment 'add'/'remove' operations are complex and not implemented. Use COM Interop.`);
                return createErrorResponse('NOT_IMPLEMENTED_LIB_COMMENT_MODIFY', `Library path: Comment 'add'/'remove' operations are not implemented. Use COM Interop.`);
            case 'manage': {
                log.warn("Library path: Reading document properties from an existing file is not directly supported by 'docx' in a simple way. Property listing will be empty or limited. Use COM Interop for full property details.");
                const { value: html } = await mammoth.convertToHtml({ buffer: fileBuffer });
                const commentRegex = /<a href="#footnote-ref-(\d+)" id="footnote-(\d+)"><sup>\[\d+\]<\/sup><\/a>.*?<aside id="footnote-ref-\d+"[^>]*>.*?<p>(.*?)<\/p>/gs; // Adjusted regex
                const commentsText: any[] = [];
                let match;
                let internalCommentIdx = 0;
                while((match = commentRegex.exec(html)) !== null) {
                    internalCommentIdx++;
                    commentsText.push({ index: internalCommentIdx, text: match[3].trim().replace(/<[^>]*>/g, "") });
                }
                // Properties part is removed as Packer.parseProperties was incorrect.
                return { success: true, data: { properties: "Property listing via library path is limited/not implemented.", comments: commentsText } };
            }
            default:
                log.error(`Library: Unknown operation: ${operation}`);
                return createErrorResponse('METADATA_ERROR_LIB', `Library: Unknown operation: ${operation}`);
        }
      }
    } catch (error: any) {
        log.error(`[word/metadata] Error: ${error.message}`, { error });
        return handleToolError(error, 'METADATA_ERROR_GENERAL');
    }
  },
};

export default metadataTool;