/**
 * @file Tool for adapting a Word document to a PowerPoint presentation.
 * This tool orchestrates getting an active Word document (if not provided),
 * converting it to PowerPoint using wordToPowerpointTool, and potentially
 * applying AI enhancements in the future.
 * @author Roo
 * @date 2025-05-15
 */
import { z } from 'zod';
import path from 'path';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext as Context } from '../../types/common.types.js';
import logger from '../../utils/logger.js';
import getActiveOfficeDocumentsTool from '../os/getActiveOfficeDocuments.tool.js';
import wordToPowerpointTool from './wordToPowerpoint.tool.js';

// Define the input schema for the adaptWordToPowerpoint tool
const AdaptWordToPowerpointInputSchema = z.object({
  inputWordPath: z.string().optional().describe('Optional path to the source Word file. If not provided, the tool will attempt to use the active Word document.'),
  outputPowerpointPath: z.string().optional().describe('Optional path where the output PowerPoint presentation will be saved. If not provided, it defaults to the same name as the input Word document with a .pptx extension in the same directory.'),
  headingLevelForNewSlide: z.number().int().min(1).max(9).optional().default(1).describe('Word heading level that will start a new slide (e.g., 1 for Heading 1). Passed to the underlying wordToPowerpointTool.'),
  enableAiEnhancement: z.boolean().optional().default(false).describe('Future placeholder: Whether to enable AI-driven enhancements for faithfulness and presentation quality.'),
});

type AdaptWordToPowerpointInput = z.infer<typeof AdaptWordToPowerpointInputSchema>;

// Define an interface for the structure of an active Office document (mirrors the one in getActiveOfficeDocuments.tool.ts)
interface ActiveOfficeDocument {
  filePath: string;
  applicationType: 'Word' | 'Excel' | 'PowerPoint';
}

interface AdaptWordToPowerpointOutput {
  message: string;
  powerpointFilePath: string;
  details?: any;
}

const adaptWordToPowerpointTool: McpResource = {
  path: 'office/adaptWordToPowerpoint',
  description: 'Adapts a Word document to a PowerPoint presentation, optionally using the active Word document.',
  schema: AdaptWordToPowerpointInputSchema,
  handler: async (params: ToolRequestParams, context?: Context<any>): Promise<ApiResponse<AdaptWordToPowerpointOutput>> => {
    const parsedParams = AdaptWordToPowerpointInputSchema.safeParse(params);

    if (!parsedParams.success) {
      return {
        success: false,
        error: {
          code: 'VALIDATION_ERROR',
          message: 'Invalid input parameters for adaptWordToPowerpoint',
          details: parsedParams.error.errors,
        },
      };
    }

    let { inputWordPath, outputPowerpointPath, headingLevelForNewSlide, enableAiEnhancement } = parsedParams.data;
    let finalWordPath = inputWordPath;

    try {
      // 1. Identify Active Word Document if inputWordPath is not provided
      if (!finalWordPath) {
        logger.info('Input Word path not provided, attempting to find active Word document.');
        const activeDocsResponse = await getActiveOfficeDocumentsTool.handler({});
        if (!activeDocsResponse.success) {
          // Now we know activeDocsResponse is an ErrorResponse
          logger.error('Failed to get active office documents.', activeDocsResponse.error);
          return { success: false, error: { code: 'ACTIVE_DOC_ERROR', message: 'Failed to retrieve active Office documents.', details: activeDocsResponse.error } };
        }
        // If success is true, activeDocsResponse.data is guaranteed to exist.
        if (!activeDocsResponse.data) {
            logger.error('Failed to get active office documents data, though success was true.');
            return { success: false, error: { code: 'ACTIVE_DOC_DATA_ERROR', message: 'Failed to retrieve active Office documents data.'} };
        }

        const wordDocuments = activeDocsResponse.data.documents.filter((doc: ActiveOfficeDocument) => doc.applicationType === 'Word');

        if (wordDocuments.length === 0) {
          logger.warn('No active Word document found and no inputWordPath provided.');
          return { success: false, error: { code: 'NO_WORD_DOC_ACTIVE', message: 'No active Word document found. Please provide an inputWordPath or open a Word document.' } };
        }

        if (wordDocuments.length > 1) {
          logger.warn(`Multiple Word documents active: ${wordDocuments.map((d: ActiveOfficeDocument) => d.filePath).join(', ')}. Using the first one: ${wordDocuments[0].filePath}`);
          // Potentially ask user or add a parameter to select which one if multiple are open. For now, using the first.
        }
        finalWordPath = wordDocuments[0].filePath;
        logger.info(`Using active Word document: ${finalWordPath}`);
      }

      if (!finalWordPath) { // Should be caught by previous checks, but as a safeguard
        return { success: false, error: { code: 'NO_INPUT_PATH', message: 'Could not determine the Word document path to process.'}};
      }

      // 2. Determine Output PowerPoint Path
      let finalOutputPowerpointPath = outputPowerpointPath;
      if (!finalOutputPowerpointPath) {
        const dirName = path.dirname(finalWordPath);
        const baseName = path.basename(finalWordPath, path.extname(finalWordPath));
        finalOutputPowerpointPath = path.join(dirName, `${baseName}.pptx`);
        logger.info(`Output PowerPoint path not provided, defaulting to: ${finalOutputPowerpointPath}`);
      }

      // 3. Perform Base Word to PowerPoint Conversion
      logger.info(`Starting Word to PowerPoint conversion for: ${finalWordPath} to ${finalOutputPowerpointPath}`);
      const conversionParams: ToolRequestParams = {
        wordFilePath: finalWordPath,
        powerpointFilePath: finalOutputPowerpointPath,
        operation: 'transfer', // Key operation for conversion
        headingLevelForNewSlide: headingLevelForNewSlide,
      };

      // Directly call the handler of the wordToPowerpointTool
      const conversionResult = await wordToPowerpointTool.handler(conversionParams, context);

      if (!conversionResult.success) {
        logger.error('Word to PowerPoint conversion failed.', conversionResult.error);
        return { success: false, error: { code: 'CONVERSION_FAILED', message: 'Core Word to PowerPoint conversion failed.', details: conversionResult.error } };
      }

      logger.info(`Word to PowerPoint conversion successful. Output: ${finalOutputPowerpointPath}`);

      // 4. Placeholder for AI Enhancement
      if (enableAiEnhancement) {
        logger.info('AI enhancement is enabled (placeholder). Future implementation needed.');
        // Here you would call aiSuggestTool or similar, potentially passing finalOutputPowerpointPath
        // For now, just log it.
      }

      // 5. Output/Return Value
      return {
        success: true,
        data: {
          message: `Successfully adapted Word document '${path.basename(finalWordPath)}' to PowerPoint presentation '${path.basename(finalOutputPowerpointPath)}'.`,
          powerpointFilePath: finalOutputPowerpointPath,
          details: conversionResult.data,
        },
      };

    } catch (error: any) {
      logger.error(`Error in adaptWordToPowerpointTool: ${error.message}`, error);
      return {
        success: false,
        error: {
          code: 'ADAPT_TOOL_ERROR',
          message: `An unexpected error occurred in adaptWordToPowerpointTool: ${error.message}`,
          details: error.stack,
        },
      };
    }
  },
};

export default adaptWordToPowerpointTool;