/**
 * @file Implements the 'word/reformat' tool using COM Interop for reformatting Word documents.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { FastMCPContext, ApiResponse, ToolRequestParams, McpResource } from '../../types/common.types.js'; // Normalized relative path
import { getOfficeApplication, openWordDocument, releaseObject } from '../../utils/officeInterop.js'; // Normalized relative path
import logger from '../../utils/logger.js'; // Normalized relative path
import { handleToolError, createErrorResponse } from '../../utils/errorHandler.js'; // Normalized relative path
import { validateFilePath } from '../../utils/security.js'; // Normalized relative path

/**
 * Zod schema for the input parameters of the 'word/reformat' tool.
 */
const ReformatParamsSchema = z.object({
  /** Path to the document. */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  /** Operation to perform: 'analyze', 'apply', or 'standardize'. */
  operation: z.enum(['analyze', 'apply', 'standardize']).describe('Operation to perform: \'analyze\', \'apply\', or \'standardize\'.'),
  /** Name of the style set to apply or analyze against. */
  styleSet: z.string().optional().describe('Name of the style set to apply or analyze against.'),
  /** Scope of the reformatting: 'document', 'section', or 'range'. */
  scope: z.enum(['document', 'section', 'range']).optional().default('document').describe('Scope of the reformatting: \'document\', \'section\', or \'range\'.'),
});

/**
 * Infers the type for the validated reformat parameters.
 */
type ReformatParams = z.infer<typeof ReformatParamsSchema>;

/**
 * Handler function for the 'word/reformat' tool.
 * @param params - The parameters for the tool, validated against `ReformatParamsSchema`.
 * @param context - The FastMCP context (optional), providing logging.
 * @returns A promise resolving to an ApiResponse containing the result of the operation.
 * @throws {Error} If validation fails, a COM error occurs, or an operation fails.
 */
const handler = async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> => {
  const log = context?.log ?? logger; // Use context logger or fallback

  let validatedParams: ReformatParams;
  try {
    validatedParams = ReformatParamsSchema.parse(params);
  } catch (error: any) {
    if (error instanceof z.ZodError) {
      log.warn(`Input validation failed at handler entry for word/reformat: ${error.message}`, { errors: JSON.stringify(error.errors), params });
      return createErrorResponse('VALIDATION_ERROR', `Input validation failed: ${error.errors.map(e => `${e.path.join('.')} - ${e.message}`).join(', ')}`);
    }
    log.error(`Unexpected error parsing params in word/reformat handler: ${error.message}`, { error, params });
    return createErrorResponse('INTERNAL_ERROR', 'Failed to parse tool parameters.');
  }

  const { filePath, operation, styleSet, scope } = validatedParams;

  log.info(`Executing word/reformat operation: ${operation} on ${filePath} with scope ${scope}`);

  let wordApp: any = null;
  let doc: any = null;
  let officeAppInstance: any = null; // To manage the application instance lifecycle

  try {
    officeAppInstance = await getOfficeApplication('Word.Application');
    wordApp = officeAppInstance.app;
    wordApp.Visible = false; // Run in background
    wordApp.DisplayAlerts = 0; // wdAlertsNone = 0
    doc = await openWordDocument(wordApp, filePath);
    log.debug(`Document opened successfully: ${filePath}`);

    let result;
    switch (operation) {
      case 'analyze':
        result = await analyzeDocument(doc, styleSet, scope, log); // Pass log
        break;
      case 'apply':
        result = await applyFormatting(doc, styleSet, scope, log); // Pass log
        break;
      case 'standardize':
        result = await standardizeDocument(doc, styleSet, scope, log); // Pass log
        break;
      default:
        // This case should not be reached due to Zod enum validation, but included for safety
        log.error(`Unsupported operation reached in switch: ${operation}`);
        return createErrorResponse('UNSUPPORTED_OPERATION', `Unsupported operation: ${operation}`);
    }

    await doc.Save();
    log.info(`Document saved: ${filePath}`);
    await doc.Close();
    log.info(`Document closed.`);

    return { success: true, data: result }; // Return ApiResponse structure

  } catch (error: any) {
    log.error(`Error in word/reformat: ${error.message}`, { error });
    // Attempt to close document and release objects in case of error
    if (doc) {
      try {
        doc.Close(0); // wdDoNotSaveChanges = 0
        log.debug(`Document closed in error handler: ${filePath}`);
      } catch (closeError: any) {
        log.warn(`Error closing document in error handler: ${closeError.message}`, { error: closeError });
      }
      releaseObject(doc);
    }
    // Quitting the app in the error handler might be too aggressive
    // if (wordApp) {
    //   try {
    //     await wordApp.Quit();
    //   } catch (quitError: any) {
    //     logger.error(`Error quitting Word application in error handler: ${quitError.message}`);
    //   }
    // }
    return handleToolError(error, 'WORD_REFORMAT_ERROR'); // Use handleToolError
  } finally {
      // Ensure objects are released
      if (doc) releaseObject(doc); // Redundant if closed in catch, but safe
      if (officeAppInstance) {
          officeAppInstance.release(); // Release the application instance
          log.debug("Office application instance released.");
      }
      if (wordApp) releaseObject(wordApp); // Release the app object reference
  }
};

/**
 * Analyzes the formatting of a Word document or a specific scope.
 * @param doc - The Word document COM object.
 * @param styleSet - Optional name of the style set to analyze against.
 * @param scope - The scope of the analysis ('document', 'section', or 'range').
 * @param log - The logger instance.
 * @returns A promise resolving to an object containing analysis results.
 */
async function analyzeDocument(doc: any, styleSet?: string, scope?: 'document' | 'section' | 'range', log: any = logger): Promise<any> {
  log.info(`Analyzing document: ${doc.Name} with scope ${scope}`);
  const analysisResult: { paragraphs: any[], tables: any[], inconsistencies: string[] } = {
    paragraphs: [],
    tables: [],
    inconsistencies: [],
  };

  try {
    // Analyze Paragraphs (assuming document scope for now)
    if (scope === 'document' || scope === 'section') { // Apply to document or sections within scope
        const paragraphs = doc.Paragraphs;
        for (let i = 1; i <= paragraphs.Count; i++) {
          const paragraph = paragraphs.Item(i);
          try {
            const styleName = paragraph.Style ? paragraph.Style.NameLocal : 'No Style';
            analysisResult.paragraphs.push({
              index: i,
              style: styleName,
              // Add more properties as needed, e.g., Font, ParagraphFormat
            });

            // Basic inconsistency check (example: check if style is not 'Normal' and no styleSet is provided)
            if (!styleSet && styleName !== 'Normal') {
               analysisResult.inconsistencies.push(`Paragraph ${i} uses style "${styleName}" which is not 'Normal'.`);
            }
            // TODO: Implement more sophisticated analysis based on styleSet if provided
          } catch (paraError: any) {
            log.warn(`Error analyzing paragraph ${i}: ${paraError.message}`, { error: paraError });
            analysisResult.inconsistencies.push(`Error analyzing paragraph ${i}: ${paraError.message}`);
          } finally {
            releaseObject(paragraph);
          }
        }
        releaseObject(paragraphs); // Release paragraphs collection
    }


    // Analyze Tables (assuming document scope for now)
    if (scope === 'document' || scope === 'section') { // Apply to document or sections within scope
        const tables = doc.Tables;
        for (let i = 1; i <= tables.Count; i++) {
          const table = tables.Item(i);
          try {
            const styleName = table.Style ? table.Style.NameLocal : 'No Style';
             analysisResult.tables.push({
              index: i,
              style: styleName,
              rows: table.Rows.Count,
              columns: table.Columns.Count,
              // Add more properties as needed
            });

            // Basic inconsistency check (example: check if table style is not 'No Style' and no styleSet is provided)
            if (!styleSet && styleName !== 'No Style') {
               analysisResult.inconsistencies.push(`Table ${i} uses style "${styleName}" which is not 'No Style'.`);
            }
            // TODO: Implement more sophisticated analysis for tables based on styleSet if provided
          } catch (tableError: any) {
            log.warn(`Error analyzing table ${i}: ${tableError.message}`, { error: tableError });
            analysisResult.inconsistencies.push(`Error analyzing table ${i}: ${tableError.message}`);
          } finally {
            releaseObject(table);
          }
        }
        releaseObject(tables); // Release tables collection
    }


    // TODO: Implement analysis for other document elements (sections, shapes, etc.) based on scope
    // TODO: Implement range-specific analysis

  } catch (error: any) {
    log.error(`Error during document analysis: ${error.message}`, { error });
    throw error; // Re-throw to be caught by the main handler's catch block
  }

  return analysisResult;
}

/**
 * Applies formatting to a Word document or a specific scope.
 * @param doc - The Word document COM object.
 * @param styleSet - Optional name of the style set to apply.
 * @param scope - The scope of the formatting application ('document', 'section', or 'range').
 * @param log - The logger instance.
 * @returns A promise resolving to an object containing details of applied changes.
 */
async function applyFormatting(doc: any, styleSet?: string, scope?: 'document' | 'section' | 'range', log: any = logger): Promise<any> {
  log.info(`Applying formatting to document: ${doc.Name} with scope ${scope}`);
  const appliedChanges: string[] = [];

  try {
    // Apply formatting to Paragraphs (example: apply 'Normal' style)
    if (scope === 'document' || scope === 'section') { // Apply to document or sections within scope
        const paragraphs = doc.Paragraphs;
        for (let i = 1; i <= paragraphs.Count; i++) {
          const paragraph = paragraphs.Item(i);
          try {
            // TODO: Implement logic to apply styles/formatting based on styleSet or other rules
            // For now, a simple example: apply 'Normal' style if styleSet is not provided
            if (!styleSet) {
               paragraph.Style = 'Normal';
               appliedChanges.push(`Applied 'Normal' style to paragraph ${i}.`);
            } else {
              // Example: Try to apply a style from the styleSet
              try {
                paragraph.Style = styleSet; // Assuming styleSet is a valid style name
                appliedChanges.push(`Attempted to apply style "${styleSet}" to paragraph ${i}.`);
              } catch (styleError: any) {
                log.warn(`Could not apply style "${styleSet}" to paragraph ${i}: ${styleError.message}`, { error: styleError });
                appliedChanges.push(`Failed to apply style "${styleSet}" to paragraph ${i}. Error: ${styleError.message}`);
              }
            }
            // TODO: Apply other formatting properties (Font, ParagraphFormat, etc.)
          } catch (paraError: any) {
            log.warn(`Error applying formatting to paragraph ${i}: ${paraError.message}`, { error: paraError });
            appliedChanges.push(`Error applying formatting to paragraph ${i}: ${paraError.message}`);
          } finally {
            releaseObject(paragraph);
          }
        }
        releaseObject(paragraphs); // Release paragraphs collection
    }


    // TODO: Implement formatting application for Tables and other elements based on scope
    // TODO: Implement range-specific formatting application

  } catch (error: any) {
    log.error(`Error during formatting application: ${error.message}`, { error });
    throw error; // Re-throw to be caught by the main handler's catch block
  }

  return { status: 'Formatting application attempted.', changes: appliedChanges };
}

/**
 * Standardizes the formatting of a Word document or a specific scope.
 * @param doc - The Word document COM object.
 * @param styleSet - Optional name of the style set to standardize against.
 * @param scope - The scope of the standardization ('document', 'section', or 'range').
 * @param log - The logger instance.
 * @returns A promise resolving to an object containing details of standardization results.
 */
async function standardizeDocument(doc: any, styleSet?: string, scope?: 'document' | 'section' | 'range', log: any = logger): Promise<any> {
  log.info(`Standardizing document: ${doc.Name} with scope ${scope}`);
  // Standardization can be a combination of analysis and application,
  // or applying a predefined set of rules.
  // For this implementation, we'll apply a default style or the specified styleSet.
  // This is similar to applyFormatting but might imply a more rigid set of rules.

  const standardizationResult: { status: string, changes: string[], analysis?: any } = {
    status: 'Document standardization attempted.',
    changes: [],
  };

  try {
    // Optionally, perform an analysis first
    // standardizationResult.analysis = await analyzeDocument(doc, styleSet, scope, log); // Pass log

    // Apply standardization rules (example: apply a default style or the styleSet)
    const targetStyle = styleSet || 'Normal'; // Use 'Normal' as default if no styleSet is provided

    // Apply style to Paragraphs
    if (scope === 'document' || scope === 'section') { // Apply to document or sections within scope
        const paragraphs = doc.Paragraphs;
        for (let i = 1; i <= paragraphs.Count; i++) {
          const paragraph = paragraphs.Item(i);
          try {
            if (paragraph.Style.NameLocal !== targetStyle) {
              paragraph.Style = targetStyle;
              standardizationResult.changes.push(`Applied style "${targetStyle}" to paragraph ${i}.`);
            }
            // TODO: Apply other standardization rules (Font, ParagraphFormat, etc.)
          } catch (paraError: any) {
            log.warn(`Error standardizing paragraph ${i}: ${paraError.message}`, { error: paraError });
            standardizationResult.changes.push(`Error standardizing paragraph ${i}: ${paraError.message}`);
          } finally {
            releaseObject(paragraph);
          }
        }
        releaseObject(paragraphs); // Release paragraphs collection
    }


    // TODO: Implement standardization for Tables and other elements based on scope
    // TODO: Implement range-specific standardization

  } catch (error: any) {
    log.error(`Error during document standardization: ${error.message}`, { error });
    throw error; // Re-throw to be caught by the main handler's catch block
  }

  return standardizationResult;
}

/**
 * McpResource definition for the 'word/reformat' tool.
 * Reformat documents.
 */
const reformatTool: McpResource = {
    path: 'word/reformat',
    description: 'Reformat documents.',
    schema: ReformatParamsSchema,
    handler: handler,
};

export default reformatTool;