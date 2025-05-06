import { z } from 'zod';
import * as path from 'path';
import * as fs from 'fs-extra';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import mammoth from 'mammoth';
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
  useComInterop: z.boolean().optional().default(false).describe("Specifies the processing path. `true` uses COM Interop (for local, interactive scenarios). `false` uses platform-independent libraries like docx/mammoth (for server-side, batch processing). Defaults to `false`."),
});

const getSchema = textOpBaseSchema.extend({
  range: z.string().min(1, 'Range specifier is required (e.g., "paragraph:N", "document", "selection", or natural language like "the third paragraph" or "the section \'Introduction\'").'),
});

const insertSchema = textOpBaseSchema.extend({
  text: z.string(),
  position: z.string().min(1, 'Position specifier is required (e.g., "start", "end", "paragraph:N:start", "paragraph:N:end", "selection", or natural language like "after the heading \'Introduction\'").'),
  format: z.enum(['plaintext', 'markdown']).default('markdown').optional().describe("Format of the text to insert ('plaintext' or 'markdown'). Defaults to 'markdown'."),
});

export const modifySchema = textOpBaseSchema.extend({
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
  const params = getSchema.parse(requestParams);
  logger.info(`Executing word/text/get for file: ${params.filePath}, range: ${params.range}, useComInterop: ${params.useComInterop}`);
  const absoluteFilePath = path.resolve(params.filePath);

  if (!await fs.pathExists(absoluteFilePath)) {
    return createErrorResponse(`File not found: ${params.filePath}`, 'FILE_NOT_FOUND');
  }

  if (params.useComInterop) {
    let wordApp: any = null;
    let doc: any = null;
    let selectedRange: any = null;
    try {
      wordApp = await getOfficeApplication('Word.Application');
      doc = wordApp.Documents.Open(absoluteFilePath, false, true, false, "", "", false, "", "", 0, false); // Open read-only, not visible
      if (!doc) {
        return createErrorResponse(`Failed to open document via COM: ${params.filePath}`, 'FILE_OPEN_FAILED_COM');
      }

      selectedRange = resolveNaturalLanguageRange(doc, params.range, wordApp);
      if (!selectedRange) {
        selectedRange = getRangeFromSpecifier(doc, params.range, wordApp);
      }

      if (!selectedRange) {
        return createErrorResponse(`Could not determine range for extraction via COM based on specifier: ${params.range}`, 'RANGE_ERROR_COM');
      }
      const textContent = selectedRange.Text || '';
      logger.info(`Successfully retrieved text via COM from range "${params.range}" in ${params.filePath}`);
      return { success: true, data: textContent };
    } catch (error: any) {
      logger.error(`Error in word/text/get (COM path): ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
      if (error instanceof z.ZodError) {
        return createErrorResponse('Input validation failed (COM path)', 'VALIDATION_ERROR', error.errors);
      }
      return createErrorResponse(`Failed to get text (COM path): ${error.message}`, 'GET_TEXT_FAILED_COM', error);
    } finally {
      releaseObject(selectedRange);
      if (doc) {
        try { doc.Close(false); } catch (e) { logger.warn('Error closing document (COM path, read-only)', e); }
        releaseObject(doc);
      }
      if (wordApp) {
        releaseObject(wordApp);
      }
    }
  } else {
    // Library path (Mammoth)
    try {
      const rangeLower = params.range.toLowerCase();
      if (rangeLower === 'document') {
        const { value } = await mammoth.extractRawText({ path: absoluteFilePath });
        logger.info(`Successfully retrieved text using mammoth from ${params.filePath}`);
        return { success: true, data: value };
      } else if (rangeLower === 'selection' || rangeLower.startsWith('paragraph:') || resolveNaturalLanguageRange(null, params.range, null)) {
         // Crude check for natural language, as resolveNaturalLanguageRange would be used by COM path
        return createErrorResponse(
            `Library path for getText currently only supports 'document' range. '${params.range}' requires COM Interop. Set useComInterop to true.`,
            'RANGE_REQUIRES_COM_LIB'
        );
      } else {
        return createErrorResponse(
            `Library path for getText currently only supports 'document' range. '${params.range}' is not supported.`,
            'RANGE_NOT_SUPPORTED_LIB'
        );
      }
    } catch (mammothError: any) {
      logger.error(`Error in word/text/get with mammoth: ${mammothError.message}`, { error: mammothError });
      return createErrorResponse(`Failed to get text with mammoth: ${mammothError.message}`, 'MAMMOTH_GET_TEXT_FAILED', mammothError);
    }
  }
}

// Modified insertText function
export async function insertText(requestParams: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> {
    const params = insertSchema.parse(requestParams);
    logger.info(`Executing word/text/insert for file: ${params.filePath}, position: ${params.position}, useComInterop: ${params.useComInterop}`);
    const absoluteFilePath = path.resolve(params.filePath);

    if (params.useComInterop) {
        let wordApp: any = null;
        let doc: any = null;
        let insertionRange: any = null;
        let paraRange: any = null;
        let fileCreated = false;
        try {
            wordApp = await getOfficeApplication('Word.Application');

            try {
                if (await fs.pathExists(absoluteFilePath)) {
                    logger.info(`Opening existing document (COM): ${absoluteFilePath}`);
                    doc = wordApp.Documents.Open(absoluteFilePath, false, false);
                } else {
                    logger.info(`File not found. Creating new document (COM) at: ${absoluteFilePath}`);
                    doc = wordApp.Documents.Add();
                    const wdFormatDocumentDefault = 16; // .docx format
                    doc.SaveAs2(absoluteFilePath, wdFormatDocumentDefault);
                    fileCreated = true;
                    logger.info(`Successfully created and saved new document (COM): ${absoluteFilePath}`);
                }
            } catch (fileError: any) {
                logger.error(`Error opening or creating document (COM) '${absoluteFilePath}': ${fileError.message}`, { error: fileError });
                return createErrorResponse(`Failed to open or create document (COM): ${fileError.message}`, 'FILE_OPERATION_FAILED_COM', fileError);
            }

            if (!doc) {
                return createErrorResponse(`Failed to obtain document object (COM) for: ${params.filePath}`, 'FILE_OPEN_FAILED_COM');
            }

            insertionRange = resolveNaturalLanguageRange(doc, params.position, wordApp);
            if (!insertionRange) {
                const positionLower = params.position.toLowerCase();
                if (positionLower === 'start') {
                    insertionRange = doc.Range(0, 0);
                } else if (positionLower === 'end') {
                    const endPos = doc.Content.End;
                    insertionRange = doc.Range(endPos, endPos);
                } else if (positionLower === 'selection') {
                    if (!wordApp.Selection) return createErrorResponse("Cannot insert at selection (COM): No selection found.", 'NO_SELECTION_COM');
                    insertionRange = wordApp.Selection.Range;
                    if (wordApp.Selection.Type !== 2 /* wdSelectionIP */) {
                        insertionRange.Collapse(1); // wdCollapseStart
                    }
                } else if (positionLower.startsWith('paragraph:')) {
                    const parts = positionLower.split(':');
                    if (parts.length < 2 || parts.length > 3) return createErrorResponse(`Invalid paragraph position format (COM): ${params.position}`, 'INVALID_POSITION_COM');
                    const indexStr = parts[1];
                    const paraIndex = parseInt(indexStr, 10);
                    if (isNaN(paraIndex) || paraIndex <= 0) return createErrorResponse(`Invalid paragraph index format (COM): '${indexStr}'.`, 'INVALID_PARAM_COM');
                    if (!doc.Paragraphs || typeof doc.Paragraphs.Count === 'undefined') return createErrorResponse("Could not access document paragraphs collection (COM).", 'COM_ERROR');
                    const paraCount = doc.Paragraphs.Count;
                    if (paraIndex > paraCount) return createErrorResponse(`Paragraph index ${paraIndex} out of bounds (COM). Document has ${paraCount} paragraphs.`, 'INVALID_PARAM_COM');
                    paraRange = doc.Paragraphs(paraIndex).Range;
                    if (parts.length === 3) {
                        if (parts[2] === 'start') insertionRange = doc.Range(paraRange.Start, paraRange.Start);
                        else if (parts[2] === 'end') {
                            const endPos = paraRange.End > paraRange.Start ? paraRange.End - 1 : paraRange.Start;
                            insertionRange = doc.Range(endPos, endPos);
                        } else {
                            releaseObject(paraRange);
                            return createErrorResponse(`Invalid paragraph position specifier (COM): ${parts[2]}`, 'INVALID_POSITION_COM');
                        }
                    } else insertionRange = doc.Range(paraRange.Start, paraRange.Start);
                } else return createErrorResponse(`Unsupported position specifier (COM): ${params.position}`, 'INVALID_POSITION_COM');
            }

            if (!insertionRange) return createErrorResponse(`Could not determine insertion range (COM) for position: ${params.position}`, 'RANGE_ERROR_COM');

            if (params.format === 'markdown') {
                logger.debug('Inserting and formatting text as Markdown (COM).');
                await applyMarkdownFormattingToWord(insertionRange, params.text, wordApp, doc);
            } else {
                logger.debug('Inserting text as plaintext (COM).');
                insertionRange.Text = params.text;
            }

            if (!fileCreated) {
                doc.Save();
                logger.info(`Successfully inserted text (COM) at position "${params.position}" with format "${params.format}" and saved ${params.filePath}`);
            } else {
                logger.info(`Successfully inserted text (COM) at position "${params.position}" with format "${params.format}" into newly created file ${params.filePath}`);
            }

            // Save the modified document as a dynamic resource
            try {
                logger.warn("Dynamic resource saving (COM path) commented out due to reliance on officeAppInstance.readDocumentContent");
            } catch (resourceSaveError: any) {
                logger.error(`Failed to save ${params.filePath} as a dynamic resource (COM): ${resourceSaveError.message}`);
                // Continue execution even if saving the resource fails
            }
            return { success: true, data: {} };
        } catch (error: any) {
            logger.error(`Error in word/text/insert (COM path): ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
            if (error instanceof z.ZodError) return createErrorResponse('Input validation failed (COM path)', 'VALIDATION_ERROR', error.errors);
            return createErrorResponse(`Failed to insert text (COM path): ${error.message}`, 'INSERT_TEXT_FAILED_COM', error);
        } finally {
            releaseObject(insertionRange);
            releaseObject(paraRange);
            if (doc) {
                try { doc.Close(false); } catch (e) { logger.warn('Error closing document after insert (COM)', e); }
                releaseObject(doc);
            }
            if (wordApp) releaseObject(wordApp);
        }
    } else {
        // Library path (docx)
        logger.warn(`Library path for word/text/insert is not fully implemented yet for file: ${params.filePath}`);
        // TODO: Implement docx logic for insertText
        // This will involve:
        // 1. Reading the document if it exists (or creating a new one).
        //    - `docx` doesn't "open" files like COM. You'd read it into memory.
        //    - `fs.readFile` then parse with a `docx` parser if modifying, or `new Document()` if new.
        // 2. Determining insertion point (start, end, paragraph:N). This is complex with `docx`.
        // 3. If 'markdown', using a new `markdownToDocx` utility (to be created).
        // 4. If 'plaintext', creating `Paragraph` and `TextRun` objects.
        // 5. Using `Packer.toBuffer(doc)` and `fs.writeFile` to save.
        return createErrorResponse(
            'Library path for insertText is not yet implemented. Use COM Interop (useComInterop: true) for this functionality.',
            'NOT_IMPLEMENTED_LIB'
        );
    }
}

export async function modifyText(requestParams: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> {
    const params = modifySchema.parse(requestParams);
    logger.info(`Executing word/text/modify for file: ${params.filePath}, range: ${params.range}, useComInterop: ${params.useComInterop}`);
    const absoluteFilePath = path.resolve(params.filePath);

    if (!await fs.pathExists(absoluteFilePath)) {
        return createErrorResponse(`File not found: ${params.filePath}`, 'FILE_NOT_FOUND');
    }

    if (params.useComInterop) {
        let wordApp: any = null;
        let doc: any = null;
        let selectedRange: any = null;
        try {
            wordApp = await getOfficeApplication('Word.Application');
            doc = wordApp.Documents.Open(absoluteFilePath, false, false); // Open read/write
            if (!doc) {
                return createErrorResponse(`Failed to open document (COM): ${params.filePath}`, 'FILE_OPEN_FAILED_COM');
            }
            selectedRange = getRangeFromSpecifier(doc, params.range, wordApp);
            selectedRange.Text = params.newText;
            doc.Save();
            logger.info(`Successfully modified text (COM) in range "${params.range}" and saved ${params.filePath}`);

            // Save the modified document as a dynamic resource
            try {
                logger.warn("Dynamic resource saving (COM path) commented out due to reliance on officeAppInstance.readDocumentContent");
            } catch (resourceSaveError: any) {
                logger.error(`Failed to save ${params.filePath} as a dynamic resource (COM): ${resourceSaveError.message}`);
                // Continue execution even if saving the resource fails
            }
            return { success: true, data: {} };
        } catch (error: any) {
            logger.error(`Error in word/text/modify (COM path): ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
            if (error instanceof z.ZodError) {
                return createErrorResponse('Input validation failed (COM path)', 'VALIDATION_ERROR', error.errors);
            }
            return createErrorResponse(`Failed to modify text (COM path): ${error.message}`, 'MODIFY_TEXT_FAILED_COM', error);
        } finally {
            releaseObject(selectedRange);
            if (doc) {
                try { doc.Close(false); } catch (e) { logger.warn('Error closing document after modify (COM)', e); }
                releaseObject(doc);
            }
            if (wordApp) {
                releaseObject(wordApp);
            }
        }
    } else {
        // Library path (docx)
        logger.warn(`Library path for word/text/modify is not fully implemented yet for file: ${params.filePath}`);
        // TODO: Implement docx logic for modifyText
        // This is complex. It might involve:
        // 1. Reading the document using a docx parser.
        // 2. Traversing the document structure to find the text matching the range (if possible to map ranges).
        // 3. Replacing the content.
        // 4. Re-packing and saving the document.
        // Simple string replacement on full text extracted by Mammoth would lose formatting.
        return createErrorResponse(
            'Library path for modifyText is not yet implemented. Use COM Interop (useComInterop: true) for this functionality.',
            'NOT_IMPLEMENTED_LIB'
        );
    }
}

export async function deleteText(requestParams: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> {
    const params = deleteSchema.parse(requestParams);
    logger.info(`Executing word/text/delete for file: ${params.filePath}, range: ${params.range}, useComInterop: ${params.useComInterop}`);
    const absoluteFilePath = path.resolve(params.filePath);

    if (!await fs.pathExists(absoluteFilePath)) {
        return createErrorResponse(`File not found: ${params.filePath}`, 'FILE_NOT_FOUND');
    }

    if (params.useComInterop) {
        let wordApp: any = null;
        let doc: any = null;
        let selectedRange: any = null;
        try {
            wordApp = await getOfficeApplication('Word.Application');
            doc = wordApp.Documents.Open(absoluteFilePath, false, false); // Open read/write
            if (!doc) {
                return createErrorResponse(`Failed to open document (COM): ${params.filePath}`, 'FILE_OPEN_FAILED_COM');
            }
            selectedRange = getRangeFromSpecifier(doc, params.range, wordApp);
            selectedRange.Delete();
            doc.Save();
            logger.info(`Successfully deleted text (COM) in range "${params.range}" and saved ${params.filePath}`);

            // Save the modified document as a dynamic resource
            try {
                logger.warn("Dynamic resource saving (COM path) commented out due to reliance on officeAppInstance.readDocumentContent");
            } catch (resourceSaveError: any) {
                logger.error(`Failed to save ${params.filePath} as a dynamic resource (COM): ${resourceSaveError.message}`);
                // Continue execution even if saving the resource fails
            }
            return { success: true, data: {} };
        } catch (error: any) {
            logger.error(`Error in word/text/delete (COM path): ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
            if (error instanceof z.ZodError) {
                return createErrorResponse('Input validation failed (COM path)', 'VALIDATION_ERROR', error.errors);
            }
            return createErrorResponse(`Failed to delete text (COM path): ${error.message}`, 'DELETE_TEXT_FAILED_COM', error);
        } finally {
            releaseObject(selectedRange);
            if (doc) {
                try { doc.Close(false); } catch (e) { logger.warn('Error closing document after delete (COM)', e); }
                releaseObject(doc);
            }
            if (wordApp) {
                releaseObject(wordApp);
            }
        }
    } else {
        // Library path (docx)
        logger.warn(`Library path for word/text/delete is not fully implemented yet for file: ${params.filePath}`);
        // TODO: Implement docx logic for deleteText
        // Similar complexity to modifyText.
        return createErrorResponse(
            'Library path for deleteText is not yet implemented. Use COM Interop (useComInterop: true) for this functionality.',
            'NOT_IMPLEMENTED_LIB'
        );
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