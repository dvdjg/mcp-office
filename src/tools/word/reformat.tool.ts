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
  // Implementación de analyze
  // Recorrer elementos del documento (párrafos, tablas, etc.)
  // Comparar estilos/formato con styleSet (si se proporciona)
  // Devolver un informe de inconsistencias
  logger.info(`Analyzing document: ${doc.Name}`);
  // Placeholder implementation
  return { analysis: 'Document analysis not yet fully implemented.' };
}

async function applyFormatting(doc: any, styleSet?: string, scope?: 'document' | 'section' | 'range'): Promise<any> {
  // Implementación de apply
  // Aplicar estilos o reglas de formato basadas en styleSet
  // Usar COM Interop para modificar propiedades de formato
  logger.info(`Applying formatting to document: ${doc.Name}`);
  // Placeholder implementation
  return { status: 'Formatting application not yet fully implemented.' };
}

async function standardizeDocument(doc: any, styleSet?: string, scope?: 'document' | 'section' | 'range'): Promise<any> {
  // Implementación de standardize
  // Combinar analyze y apply, o aplicar un conjunto de reglas de estandarización predefinidas
  logger.info(`Standardizing document: ${doc.Name}`);
  // Placeholder implementation
  return { status: 'Document standardization not yet fully implemented.' };
}

const reformatTool: McpResource = {
    path: 'word/reformat',
    description: 'Reformat documents.',
    schema: ReformatParamsSchema,
    handler: handler,
};

export default reformatTool;