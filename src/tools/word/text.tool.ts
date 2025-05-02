import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types'; // Import FastMCPContext, remove ToolContext
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Removed getRangeFromSpecifier import
import { validateFilePath } from '../../utils/security';
import logger from '../../utils/logger';

// --- Schemas ---

const textOpBaseSchema = z.object({
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.", // Added validation refine
  }),
});

const getSchema = textOpBaseSchema.extend({
  range: z.string().min(1, 'Range specifier is required (e.g., "paragraph:N", "document", "selection").'),
});

const insertSchema = textOpBaseSchema.extend({
  text: z.string(),
  position: z.string().min(1, 'Position specifier is required (e.g., "start", "end", "paragraph:N:start", "paragraph:N:end", "selection").'),
});

const modifySchema = textOpBaseSchema.extend({
  range: z.string().min(1, 'Range specifier is required.'),
  newText: z.string(),
});

const deleteSchema = textOpBaseSchema.extend({
  range: z.string().min(1, 'Range specifier is required.'),
});

// --- Helper Functions ---

// Helper to create standard error responses
const createErrorResponse = (message: string, code = 'TOOL_EXECUTION_ERROR', details?: unknown): ApiResponse<never> => ({
    success: false,
    error: { code, message, details },
});

/**
 * Gets a COM Range object based on a string specifier.
 * Adapted from word/styles/applyStyle logic.
 * @param doc The Word Document COM object.
 * @param rangeSpecifier String like 'selection', 'document', 'paragraph:N'.
 * @param wordApp The Word Application COM object (needed for 'selection').
 * @returns The COM Range object.
 * @throws Error if the specifier is invalid or the range cannot be obtained.
 */
function getRangeFromSpecifier(doc: any, rangeSpecifier: string, wordApp: any): any {
    const rangeStringLower = rangeSpecifier.toLowerCase();
    let selectedRange: any = null;

    logger.debug(`Attempting to get range for specifier: ${rangeSpecifier}`);

    if (rangeStringLower === 'selection') {
        if (!wordApp || !wordApp.Selection || !wordApp.Selection.Range) {
             logger.warn("wordApp.Selection.Range was null or undefined. This might happen if there's no active selection or the app is headless/not focused.");
             throw new Error("Could not get range from selection. Ensure the document is active and has a selection, or use a different range specifier.");
        }
        selectedRange = wordApp.Selection.Range;
    } else if (rangeStringLower === 'document') {
        if (!doc || !doc.Content) {
             throw new Error("Could not get document content range.");
        }
        selectedRange = doc.Content;
    } else if (rangeStringLower.startsWith('paragraph:')) {
        const indexStr = rangeSpecifier.split(':')[1];
        const index = parseInt(indexStr, 10);
        if (isNaN(index) || index <= 0) {
            throw new Error(`Invalid paragraph index format: '${indexStr}'. Use 'paragraph:N' where N is a positive integer.`);
        }
        if (!doc || !doc.Paragraphs || typeof doc.Paragraphs.Count === 'undefined') {
             throw new Error("Could not access document paragraphs collection.");
        }
        const paraCount = doc.Paragraphs.Count;
        if (index > paraCount) {
            throw new Error(`Paragraph index ${index} is out of bounds. Document has ${paraCount} paragraphs.`);
        }
        selectedRange = doc.Paragraphs(index).Range; // COM indices are 1-based
    } else {
        throw new Error(`Unsupported range format: '${rangeSpecifier}'. Supported formats: 'selection', 'document', 'paragraph:N'.`);
    }

    if (!selectedRange) {
        // This case should ideally be caught by earlier checks, but as a safeguard:
        throw new Error(`Failed to obtain a valid range object for specifier: ${rangeSpecifier}`);
    }
    logger.debug(`Successfully obtained range for specifier: ${rangeSpecifier}`);
    return selectedRange; // Remember: Caller must release this object!
}


// --- Handlers ---

async function getText(requestParams: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<string>> { // Use FastMCPContext<undefined>
  let wordApp: any = null;
  let doc: any = null;
  let selectedRange: any = null;
  let officeAppInstance: any = null;

  try {
    const params = getSchema.parse(requestParams);
    logger.info(`Executing word/text/get for file: ${params.filePath}, range: ${params.range}`);

    // File path validated by Zod refine
    officeAppInstance = await getOfficeApplication('Word.Application');
    wordApp = officeAppInstance.app;
    doc = officeAppInstance.openDocument(params.filePath); // Open read-only by default
    if (!doc) {
      return createErrorResponse(`Failed to open document: ${params.filePath}`, 'FILE_OPEN_FAILED');
    }

    selectedRange = getRangeFromSpecifier(doc, params.range, wordApp); // Use local helper

    const textContent = selectedRange.Text || '';

    logger.info(`Successfully retrieved text from range "${params.range}" in ${params.filePath}`);
    return { success: true, data: textContent };

  } catch (error: any) {
    logger.error(`Error in word/text/get: ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
     if (error instanceof z.ZodError) {
        return createErrorResponse('Input validation failed', 'VALIDATION_ERROR', error.errors);
    }
    // Catch errors from getRangeFromSpecifier too
    return createErrorResponse(`Failed to get text: ${error.message}`, 'GET_TEXT_FAILED', error);
  } finally {
    releaseObject(selectedRange); // Release range obtained from helper
    if (doc) {
      try { doc.Close(false); } catch (e) { logger.warn('Error closing document (read-only)', e); }
      releaseObject(doc);
    }
    if (officeAppInstance) {
      officeAppInstance.release();
    }
  }
}

async function insertText(requestParams: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> { // Use FastMCPContext<undefined>
    let wordApp: any = null;
    let doc: any = null;
    let insertionRange: any = null;
    let officeAppInstance: any = null;
    let paraRange: any = null; // Specific range for paragraph logic

    try {
        const params = insertSchema.parse(requestParams);
        logger.info(`Executing word/text/insert for file: ${params.filePath}, position: ${params.position}`);

        // File path validated by Zod refine
        officeAppInstance = await getOfficeApplication('Word.Application');
        wordApp = officeAppInstance.app;
        doc = officeAppInstance.openDocument(params.filePath, false, false); // Open read/write
        if (!doc) {
             return createErrorResponse(`Failed to open document: ${params.filePath}`, 'FILE_OPEN_FAILED');
        }

        const positionLower = params.position.toLowerCase();

        if (positionLower === 'start') {
            insertionRange = doc.Range(0, 0);
        } else if (positionLower === 'end') {
            const endPos = doc.Content.End;
            insertionRange = doc.Range(endPos, endPos);
        } else if (positionLower === 'selection') {
             if (!wordApp.Selection) return createErrorResponse("Cannot insert at selection: No selection found.", 'NO_SELECTION');
             insertionRange = wordApp.Selection.Range;
             if (wordApp.Selection.Type !== 2 /* wdSelectionIP = 2 */) {
                  insertionRange.Collapse(1); // wdCollapseStart = 1
             }
        } else if (positionLower.startsWith('paragraph:')) {
            const parts = positionLower.split(':');
            if (parts.length < 2 || parts.length > 3) return createErrorResponse(`Invalid paragraph position format: ${params.position}`, 'INVALID_POSITION');

            const indexStr = parts[1];
            const paraIndex = parseInt(indexStr, 10);
             if (isNaN(paraIndex) || paraIndex <= 0) {
                return createErrorResponse(`Invalid paragraph index format: '${indexStr}'. Use 'paragraph:N' where N is a positive integer.`, 'INVALID_PARAM');
            }
             if (!doc || !doc.Paragraphs || typeof doc.Paragraphs.Count === 'undefined') {
                 return createErrorResponse("Could not access document paragraphs collection.", 'COM_ERROR');
            }
            const paraCount = doc.Paragraphs.Count;
            if (paraIndex > paraCount) {
                return createErrorResponse(`Paragraph index ${paraIndex} is out of bounds. Document has ${paraCount} paragraphs.`, 'INVALID_PARAM');
            }

            paraRange = doc.Paragraphs(paraIndex).Range; // Get the paragraph range

            if (parts.length === 3) { // Position specified (start/end)
                if (parts[2] === 'start') {
                    insertionRange = doc.Range(paraRange.Start, paraRange.Start);
                } else if (parts[2] === 'end') {
                    const endPos = paraRange.End > paraRange.Start ? paraRange.End - 1 : paraRange.Start; // Adjust for para mark
                    insertionRange = doc.Range(endPos, endPos);
                } else {
                    releaseObject(paraRange); // Release paraRange before throwing
                    return createErrorResponse(`Invalid paragraph position specifier: ${parts[2]}`, 'INVALID_POSITION');
                }
            } else { // Default to start of paragraph
                 insertionRange = doc.Range(paraRange.Start, paraRange.Start);
            }
            // paraRange is released in the finally block now
        } else {
            return createErrorResponse(`Unsupported position specifier: ${params.position}`, 'INVALID_POSITION');
        }

        if (!insertionRange) {
             // Should be caught earlier, but safeguard
             return createErrorResponse(`Could not determine insertion range for position: ${params.position}`, 'RANGE_ERROR');
        }

        insertionRange.Text = params.text;

        doc.Save();
        logger.info(`Successfully inserted text at position "${params.position}" and saved ${params.filePath}`);
        return { success: true, data: {} };

    } catch (error: any) {
        logger.error(`Error in word/text/insert: ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
         if (error instanceof z.ZodError) {
            return createErrorResponse('Input validation failed', 'VALIDATION_ERROR', error.errors);
        }
        return createErrorResponse(`Failed to insert text: ${error.message}`, 'INSERT_TEXT_FAILED', error);
    } finally {
        releaseObject(insertionRange);
        releaseObject(paraRange); // Release the paragraph range if it was obtained
        if (doc) {
            try { doc.Close(false); } catch (e) { logger.warn('Error closing document after insert', e); }
            releaseObject(doc);
        }
        if (officeAppInstance) {
            officeAppInstance.release();
        }
    }
}

async function modifyText(requestParams: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> { // Use FastMCPContext<undefined>
    let wordApp: any = null;
    let doc: any = null;
    let selectedRange: any = null;
    let officeAppInstance: any = null;

    try {
        const params = modifySchema.parse(requestParams);
        logger.info(`Executing word/text/modify for file: ${params.filePath}, range: ${params.range}`);

        // File path validated by Zod refine
        officeAppInstance = await getOfficeApplication('Word.Application');
        wordApp = officeAppInstance.app;
        doc = officeAppInstance.openDocument(params.filePath, false, false); // Open read/write
        if (!doc) {
            return createErrorResponse(`Failed to open document: ${params.filePath}`, 'FILE_OPEN_FAILED');
        }

        selectedRange = getRangeFromSpecifier(doc, params.range, wordApp); // Use local helper

        selectedRange.Text = params.newText;

        doc.Save();
        logger.info(`Successfully modified text in range "${params.range}" and saved ${params.filePath}`);
        return { success: true, data: {} };

    } catch (error: any) {
        logger.error(`Error in word/text/modify: ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
         if (error instanceof z.ZodError) {
            return createErrorResponse('Input validation failed', 'VALIDATION_ERROR', error.errors);
        }
        return createErrorResponse(`Failed to modify text: ${error.message}`, 'MODIFY_TEXT_FAILED', error);
    } finally {
        releaseObject(selectedRange); // Release range obtained from helper
        if (doc) {
             try { doc.Close(false); } catch (e) { logger.warn('Error closing document after modify', e); }
            releaseObject(doc);
        }
        if (officeAppInstance) {
            officeAppInstance.release();
        }
    }
}

async function deleteText(requestParams: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> { // Use FastMCPContext<undefined>
    let wordApp: any = null;
    let doc: any = null;
    let selectedRange: any = null;
    let officeAppInstance: any = null;

    try {
        const params = deleteSchema.parse(requestParams);
        logger.info(`Executing word/text/delete for file: ${params.filePath}, range: ${params.range}`);

        // File path validated by Zod refine
        officeAppInstance = await getOfficeApplication('Word.Application');
        wordApp = officeAppInstance.app;
        doc = officeAppInstance.openDocument(params.filePath, false, false); // Open read/write
        if (!doc) {
            return createErrorResponse(`Failed to open document: ${params.filePath}`, 'FILE_OPEN_FAILED');
        }

        selectedRange = getRangeFromSpecifier(doc, params.range, wordApp); // Use local helper

        selectedRange.Delete(); // Use Delete method

        doc.Save();
        logger.info(`Successfully deleted text in range "${params.range}" and saved ${params.filePath}`);
        return { success: true, data: {} };

    } catch (error: any) {
        logger.error(`Error in word/text/delete: ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
         if (error instanceof z.ZodError) {
            return createErrorResponse('Input validation failed', 'VALIDATION_ERROR', error.errors);
        }
        return createErrorResponse(`Failed to delete text: ${error.message}`, 'DELETE_TEXT_FAILED', error);
    } finally {
        releaseObject(selectedRange); // Release range obtained from helper
        if (doc) {
             try { doc.Close(false); } catch (e) { logger.warn('Error closing document after delete', e); }
            releaseObject(doc);
        }
        if (officeAppInstance) {
            officeAppInstance.release();
        }
    }
}


// --- Resource Definitions ---

export const wordTextTool: McpResource[] = [
  {
    path: 'word/text/get', // Use 'path'
    description: 'Gets text content from a specified range within a Word document (e.g., specific paragraph, whole document, current selection).',
    // icon: '📄', // Icon not part of McpResource definition
    schema: getSchema, // Use 'schema'
    handler: getText,
  },
  {
    path: 'word/text/insert', // Use 'path'
    description: 'Inserts text at a specified position within a Word document (e.g., start, end, start/end of a paragraph, current selection). Saves the document after insertion.',
    // icon: '➕📄',
    schema: insertSchema, // Use 'schema'
    handler: insertText,
  },
  {
    path: 'word/text/modify', // Use 'path'
    description: 'Replaces the text content of a specified range within a Word document with new text. Saves the document after modification.',
    // icon: '✏️📄',
    schema: modifySchema, // Use 'schema'
    handler: modifyText,
  },
  {
    path: 'word/text/delete', // Use 'path'
    description: 'Deletes the text content of a specified range within a Word document. Saves the document after deletion.',
    // icon: '🗑️📄',
    schema: deleteSchema, // Use 'schema'
    handler: deleteText,
  },
];