import { z } from 'zod';
import * as path from 'path';
import * as fs from 'fs-extra'; // Added fs import
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { resolveNaturalLanguageRange } from '../../utils/wordRangeResolver';
import { applyMarkdownFormattingToWord } from '../../utils/markdownToOffice';
import { validateFilePath } from '../../utils/security';
import logger from '../../utils/logger';
import { saveResource } from '../dynamic/resources.tool';

// --- Schemas --- (Keep existing schemas)
const textOpBaseSchema = z.object({
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
});

const getSchema = textOpBaseSchema.extend({
  range: z.string().min(1, 'Range specifier is required (e.g., "paragraph:N", "document", "selection", or natural language like "the third paragraph" or "the section \'Introduction\'").'),
});

const insertSchema = textOpBaseSchema.extend({
  text: z.string(),
  position: z.string().min(1, 'Position specifier is required (e.g., "start", "end", "paragraph:N:start", "paragraph:N:end", "selection", or natural language like "after the heading \'Introduction\'").'),
  format: z.enum(['plaintext', 'markdown']).default('markdown').optional().describe("Format of the text to insert ('plaintext' or 'markdown'). Defaults to 'markdown'."),
});

const modifySchema = textOpBaseSchema.extend({
  range: z.string().min(1, 'Range specifier is required.'),
  newText: z.string(),
});

const deleteSchema = textOpBaseSchema.extend({
  range: z.string().min(1, 'Range specifier is required.'),
});


// --- Helper Functions --- (Keep existing helpers)
const createErrorResponse = (message: string, code = 'TOOL_EXECUTION_ERROR', details?: unknown): ApiResponse<never> => ({
    success: false,
    error: { code, message, details },
});

export function getRangeFromSpecifier(doc: any, rangeSpecifier: string, wordApp: any): any {
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

// getText function remains unchanged
export async function getText(requestParams: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<string>> {
  let wordApp: any = null;
  let doc: any = null;
  let selectedRange: any = null;
  let officeAppInstance: any = null;

  try {
    const params = getSchema.parse(requestParams);
    logger.info(`Executing word/text/get for file: ${params.filePath}, range: ${params.range}`);

    // File path validated by Zod refine
    wordApp = await getOfficeApplication('Word.Application');
    const absoluteFilePath = path.resolve(params.filePath);

    // Check if file exists before opening read-only
    if (!await fs.pathExists(absoluteFilePath)) {
        return createErrorResponse(`File not found: ${params.filePath}`, 'FILE_NOT_FOUND');
    }

    doc = wordApp.Documents.Open(absoluteFilePath, false, true, false, "", "", false, "", "", 0, false); // Open read-only, not visible
    if (!doc) {
      return createErrorResponse(`Failed to open document: ${params.filePath}`, 'FILE_OPEN_FAILED');
    }

    // Attempt to resolve natural language range first
    selectedRange = resolveNaturalLanguageRange(doc, params.range, wordApp);

    if (!selectedRange) {
        // If natural language resolution failed, try specific formats
        selectedRange = getRangeFromSpecifier(doc, params.range, wordApp); // Use local helper
    }

    if (!selectedRange) {
        // If neither resolution method worked
        return createErrorResponse(`Could not determine range for extraction based on specifier: ${params.range}`, 'RANGE_ERROR');
    }

    const textContent = selectedRange.Text || '';

    logger.info(`Successfully retrieved text from range "${params.range}" in ${params.filePath}`);
    return { success: true, data: textContent };

  } catch (error: any) {
    logger.error(`Error in word/text/get: ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
     if (error instanceof z.ZodError) {
        return createErrorResponse('Input validation failed', 'VALIDATION_ERROR', error.errors);
    }
    // Catch errors from getRangeFromSpecifier and resolveNaturalLanguageRange too
    return createErrorResponse(`Failed to get text: ${error.message}`, 'GET_TEXT_FAILED', error);
  } finally {
    releaseObject(selectedRange); // Release range obtained from helper
    if (doc) {
      try { doc.Close(false); } catch (e) { logger.warn('Error closing document (read-only)', e); }
      releaseObject(doc);
    }
    if (wordApp) { // Changed from officeAppInstance
      releaseObject(wordApp); // Release the application object
    }
  }
}

// Modified insertText function
export async function insertText(requestParams: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> {
    let wordApp: any = null;
    let doc: any = null;
    let insertionRange: any = null;
    let paraRange: any = null; // Specific range for paragraph logic
    let fileCreated = false; // Flag to track if the file was created

    try {
        const params = insertSchema.parse(requestParams);
        logger.info(`Executing word/text/insert for file: ${params.filePath}, position: ${params.position}`);

        // File path validated by Zod refine
        wordApp = await getOfficeApplication('Word.Application');
        const absoluteFilePath = path.resolve(params.filePath);

        // --- Create if not exists logic ---
        try {
            if (await fs.pathExists(absoluteFilePath)) {
                logger.info(`Opening existing document: ${absoluteFilePath}`);
                doc = wordApp.Documents.Open(absoluteFilePath, false, false); // Open read/write
            } else {
                logger.info(`File not found. Creating new document at: ${absoluteFilePath}`);
                doc = wordApp.Documents.Add(); // Create new document
                // Save the new document immediately to the target path
                const wdFormatDocumentDefault = 16; // .docx format
                doc.SaveAs2(absoluteFilePath, wdFormatDocumentDefault);
                fileCreated = true;
                logger.info(`Successfully created and saved new document: ${absoluteFilePath}`);
            }
        } catch (fileError: any) {
             logger.error(`Error opening or creating document '${absoluteFilePath}': ${fileError.message}`, { error: fileError });
             return createErrorResponse(`Failed to open or create document: ${fileError.message}`, 'FILE_OPERATION_FAILED', fileError);
        }
        // --- End create if not exists logic ---

        if (!doc) {
             // This check might be redundant if the try/catch above handles errors, but good as a safeguard
             return createErrorResponse(`Failed to obtain document object for: ${params.filePath}`, 'FILE_OPEN_FAILED');
        }

        // Attempt to resolve natural language position first
        insertionRange = resolveNaturalLanguageRange(doc, params.position, wordApp);

        if (!insertionRange) {
            // If natural language resolution failed, try specific formats
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
        }


        if (!insertionRange) {
             // Should be caught earlier, but safeguard
             return createErrorResponse(`Could not determine insertion range for position: ${params.position}`, 'RANGE_ERROR');
        }

        // Insert and format text based on the specified format
        if (params.format === 'markdown') {
            logger.debug('Inserting and formatting text as Markdown.');
            await applyMarkdownFormattingToWord(insertionRange, params.text, wordApp, doc); // Pass doc object
        } else { // Default to plaintext
            logger.debug('Inserting text as plaintext.');
            insertionRange.Text = params.text;
        }


        // Save only if the file wasn't just created (SaveAs already saved it)
        if (!fileCreated) {
            doc.Save();
            logger.info(`Successfully inserted text at position "${params.position}" with format "${params.format}" and saved ${params.filePath}`);
        } else {
             logger.info(`Successfully inserted text at position "${params.position}" with format "${params.format}" into newly created file ${params.filePath}`);
        }


        // Guardar el documento modificado como un recurso dinámico
        try {
            // Assuming officeAppInstance.readDocumentContent is a helper that uses the COM object
            // Need to replace this with direct COM calls or a helper that takes wordApp/doc
            // For now, commenting out or adapting based on available info
            // const updatedContent = await officeAppInstance.readDocumentContent(params.filePath);
            // await saveResource('word/text', path.basename(params.filePath), updatedContent);
            // logger.info(`Saved ${params.filePath} as a dynamic resource.`);
             logger.warn("Dynamic resource saving commented out due to reliance on officeAppInstance.readDocumentContent");
        } catch (resourceSaveError: any) {
            logger.error(`Failed to save ${params.filePath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continuar la ejecución aunque falle el guardado del recurso
        }

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
            // Close without saving changes if we just created it (already saved by SaveAs)
            // Otherwise, close normally (changes should have been saved by doc.Save())
            try { doc.Close(false); } catch (e) { logger.warn('Error closing document after insert', e); }
            releaseObject(doc);
        }
        if (wordApp) {
            releaseObject(wordApp); // Release the application object
        }
    }
}

// modifyText function remains unchanged
export async function modifyText(requestParams: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> {
    let wordApp: any = null;
    let doc: any = null;
    let selectedRange: any = null;
    let officeAppInstance: any = null;

    try {
        const params = modifySchema.parse(requestParams);
        logger.info(`Executing word/text/modify for file: ${params.filePath}, range: ${params.range}`);

        // File path validated by Zod refine
        wordApp = await getOfficeApplication('Word.Application');
        const absoluteFilePath = path.resolve(params.filePath);

        // Check if file exists before opening
        if (!await fs.pathExists(absoluteFilePath)) {
            return createErrorResponse(`File not found: ${params.filePath}`, 'FILE_NOT_FOUND');
        }

        doc = wordApp.Documents.Open(absoluteFilePath, false, false); // Open read/write
        if (!doc) {
            return createErrorResponse(`Failed to open document: ${params.filePath}`, 'FILE_OPEN_FAILED');
        }

        selectedRange = getRangeFromSpecifier(doc, params.range, wordApp); // Use local helper

        selectedRange.Text = params.newText;

        doc.Save();
        logger.info(`Successfully modified text in range "${params.range}" and saved ${params.filePath}`);

        // Guardar el documento modificado como un recurso dinámico
        try {
            // Assuming officeAppInstance.readDocumentContent is a helper that uses the COM object
            // Need to replace this with direct COM calls or a helper that takes wordApp/doc
            // For now, commenting out or adapting based on available info
            // const updatedContent = await officeAppInstance.readDocumentContent(params.filePath);
            // await saveResource('word/text', path.basename(params.filePath), updatedContent);
            // logger.info(`Saved ${params.filePath} as a dynamic resource.`);
             logger.warn("Dynamic resource saving commented out due to reliance on officeAppInstance.readDocumentContent");
        } catch (resourceSaveError: any) {
            logger.error(`Failed to save ${params.filePath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continuar la ejecución aunque falle el guardado del recurso
        }

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
        if (wordApp) {
            releaseObject(wordApp); // Release the application object
        }
    }
}

// deleteText function remains unchanged
export async function deleteText(requestParams: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> {
    let wordApp: any = null;
    let doc: any = null;
    let selectedRange: any = null;
    let officeAppInstance: any = null;

    try {
        const params = deleteSchema.parse(requestParams);
        logger.info(`Executing word/text/delete for file: ${params.filePath}, range: ${params.range}`);

        // File path validated by Zod refine
        wordApp = await getOfficeApplication('Word.Application');
        const absoluteFilePath = path.resolve(params.filePath);

        // Check if file exists before opening
        if (!await fs.pathExists(absoluteFilePath)) {
            return createErrorResponse(`File not found: ${params.filePath}`, 'FILE_NOT_FOUND');
        }

        doc = wordApp.Documents.Open(absoluteFilePath, false, false); // Open read/write
        if (!doc) {
            return createErrorResponse(`Failed to open document: ${params.filePath}`, 'FILE_OPEN_FAILED');
        }

        selectedRange = getRangeFromSpecifier(doc, params.range, wordApp); // Use local helper

        selectedRange.Delete(); // Use Delete method

        doc.Save();
        logger.info(`Successfully deleted text in range "${params.range}" and saved ${params.filePath}`);

        // Guardar el documento modificado como un recurso dinámico
        try {
            // Assuming officeAppInstance.readDocumentContent is a helper that uses the COM object
            // Need to replace this with direct COM calls or a helper that takes wordApp/doc
            // For now, commenting out or adapting based on available info
            // const updatedContent = await officeAppInstance.readDocumentContent(params.filePath);
            // await saveResource('word/text', path.basename(params.filePath), updatedContent);
            // logger.info(`Saved ${params.filePath} as a dynamic resource.`);
             logger.warn("Dynamic resource saving commented out due to reliance on officeAppInstance.readDocumentContent");
        } catch (resourceSaveError: any) {
            logger.error(`Failed to save ${params.filePath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continuar la ejecución aunque falle el guardado del recurso
        }

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
        if (wordApp) {
            releaseObject(wordApp); // Release the application object
        }
    }
}


// --- Resource Definitions --- (Keep existing definitions)
export const wordTextTool: McpResource[] = [
  {
    path: 'word/text/get',
    description: 'Gets text content from a specified range within a Word document (e.g., specific paragraph, whole document, current selection).',
    // icon: '📄', // Icon not part of McpResource definition
    schema: getSchema,
    handler: getText,
  },
  {
    path: 'word/text/insert',
    description: 'Inserts text at a specified position within a Word document (e.g., start, end, start/end of a paragraph, current selection). Creates the file if it does not exist. Saves the document after insertion.', // Updated description
    // icon: '➕📄',
    schema: insertSchema,
    handler: insertText,
  },
  {
    path: 'word/text/modify',
    description: 'Replaces the text content of a specified range within a Word document with new text. Saves the document after modification.',
    // icon: '✏️📄',
    schema: modifySchema,
    handler: modifyText,
  },
  {
    path: 'word/text/delete',
    description: 'Deletes the text content of a specified range within a Word document. Saves the document after deletion.',
    // icon: '🗑️📄',
    schema: deleteSchema,
    handler: deleteText,
  },
];