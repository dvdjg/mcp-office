import { z } from 'zod';
import { FastMCPContext, ApiResponse, ToolRequestParams, McpResource } from '@/types/common.types'; // Import ApiResponse, ToolRequestParams, McpResource
import { getOfficeApplication, openWordDocument, releaseObject } from '../../utils/officeInterop';
import logger from '../../utils/logger';
import { handleToolError } from '@/utils/errorHandler'; // Import handleToolError

/**
 * @tool word/reformat
 * @description Reformat documents.
 * @operations analyze, apply, standardize.
 * @param filePath Path to the document.
 * @param operation Operation to perform: 'analyze', 'apply', or 'standardize'.
 * @param styleSet Name of the style set to apply or analyze against.
 * @param scope Scope of the reformatting: 'document', 'section', or 'range'.
 */

const ReformatParamsSchema = z.object({
  filePath: z.string(),
  operation: z.enum(['analyze', 'apply', 'standardize']),
  styleSet: z.string().optional(),
  scope: z.enum(['document', 'section', 'range']).optional().default('document'),
});

type ReformatParams = z.infer<typeof ReformatParamsSchema>;

const handler = async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> => {
  const { filePath, operation, styleSet, scope } = ReformatParamsSchema.parse(params);

  logger.info(`Executing word/reformat operation: ${operation} on ${filePath} with scope ${scope}`);

  let wordApp: any = null;
  let doc: any = null;
  try {
    wordApp = await getOfficeApplication('Word.Application');
    doc = await openWordDocument(wordApp, filePath);

    let result;
    switch (operation) {
      case 'analyze':
        result = await analyzeDocument(doc, styleSet, scope);
        break;
      case 'apply':
        result = await applyFormatting(doc, styleSet, scope);
        break;
      case 'standardize':
        result = await standardizeDocument(doc, styleSet, scope);
        break;
      default:
        throw new Error(`Unsupported operation: ${operation}`);
    }

    await doc.Save();
    await doc.Close();
    // Decide whether to quit Word application.
    // For simplicity, we might leave Word running if it was already open.
    // A more robust solution would track if we opened the instance.
    // await wordApp.Quit(); // Commenting out for now to avoid closing user's instance

    return { success: true, data: result }; // Changed to return ApiResponse structure

  } catch (error: any) {
    logger.error(`Error in word/reformat: ${error.message}`);
    // Attempt to close document and release objects in case of error
    if (doc) {
      try {
        doc.Close(0); // wdDoNotSaveChanges = 0
      } catch (closeError: any) {
        logger.warn(`Error closing document in error handler: ${closeError.message}`);
      }
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
      if (doc) releaseObject(doc);
      // if (wordApp) releaseObject(wordApp); // Be cautious about releasing the app object
  }
};

async function analyzeDocument(doc: any, styleSet?: string, scope?: 'document' | 'section' | 'range'): Promise<any> {
  logger.info(`Analyzing document: ${doc.Name}`);
  const analysisResult: { paragraphs: any[], tables: any[], inconsistencies: string[] } = {
    paragraphs: [],
    tables: [],
    inconsistencies: [],
  };

  try {
    // Analyze Paragraphs
    for (let i = 1; i <= doc.Paragraphs.Count; i++) {
      const paragraph = doc.Paragraphs.Item(i);
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
        logger.warn(`Error analyzing paragraph ${i}: ${paraError.message}`);
        analysisResult.inconsistencies.push(`Error analyzing paragraph ${i}: ${paraError.message}`);
      } finally {
        releaseObject(paragraph);
      }
    }

    // Analyze Tables
    for (let i = 1; i <= doc.Tables.Count; i++) {
      const table = doc.Tables.Item(i);
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
        logger.warn(`Error analyzing table ${i}: ${tableError.message}`);
        analysisResult.inconsistencies.push(`Error analyzing table ${i}: ${tableError.message}`);
      } finally {
        releaseObject(table);
      }
    }

    // TODO: Implement analysis for other document elements (sections, shapes, etc.) based on scope

  } catch (error: any) {
    logger.error(`Error during document analysis: ${error.message}`);
    throw error; // Re-throw to be caught by the main handler's catch block
  }

  return analysisResult;
}

async function applyFormatting(doc: any, styleSet?: string, scope?: 'document' | 'section' | 'range'): Promise<any> {
  logger.info(`Applying formatting to document: ${doc.Name}`);
  const appliedChanges: string[] = [];

  try {
    // Apply formatting to Paragraphs (example: apply 'Normal' style)
    for (let i = 1; i <= doc.Paragraphs.Count; i++) {
      const paragraph = doc.Paragraphs.Item(i);
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
            logger.warn(`Could not apply style "${styleSet}" to paragraph ${i}: ${styleError.message}`);
            appliedChanges.push(`Failed to apply style "${styleSet}" to paragraph ${i}. Error: ${styleError.message}`);
          }
        }
        // TODO: Apply other formatting properties (Font, ParagraphFormat, etc.)
      } catch (paraError: any) {
        logger.warn(`Error applying formatting to paragraph ${i}: ${paraError.message}`);
        appliedChanges.push(`Error applying formatting to paragraph ${i}: ${paraError.message}`);
      } finally {
        releaseObject(paragraph);
      }
    }

    // TODO: Implement formatting application for Tables and other elements based on scope

  } catch (error: any) {
    logger.error(`Error during formatting application: ${error.message}`);
    throw error; // Re-throw to be caught by the main handler's catch block
  }

  return { status: 'Formatting application attempted.', changes: appliedChanges };
}

async function standardizeDocument(doc: any, styleSet?: string, scope?: 'document' | 'section' | 'range'): Promise<any> {
  logger.info(`Standardizing document: ${doc.Name}`);
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
    // standardizationResult.analysis = await analyzeDocument(doc, styleSet, scope);

    // Apply standardization rules (example: apply a default style or the styleSet)
    const targetStyle = styleSet || 'Normal'; // Use 'Normal' as default if no styleSet is provided

    // Apply style to Paragraphs
    for (let i = 1; i <= doc.Paragraphs.Count; i++) {
      const paragraph = doc.Paragraphs.Item(i);
      try {
        if (paragraph.Style.NameLocal !== targetStyle) {
          paragraph.Style = targetStyle;
          standardizationResult.changes.push(`Applied style "${targetStyle}" to paragraph ${i}.`);
        }
        // TODO: Apply other standardization rules (Font, ParagraphFormat, etc.)
      } catch (paraError: any) {
        logger.warn(`Error standardizing paragraph ${i}: ${paraError.message}`);
        standardizationResult.changes.push(`Error standardizing paragraph ${i}: ${paraError.message}`);
      } finally {
        releaseObject(paragraph);
      }
    }

    // TODO: Implement standardization for Tables and other elements based on scope

  } catch (error: any) {
    logger.error(`Error during document standardization: ${error.message}`);
    throw error; // Re-throw to be caught by the main handler's catch block
  }

  return standardizationResult;
}

const reformatTool: McpResource = {
    path: 'word/reformat',
    description: 'Reformat documents.',
    schema: ReformatParamsSchema,
    handler: handler,
};

export default reformatTool;