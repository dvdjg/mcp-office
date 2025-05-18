import { z } from 'zod';
import { McpResource, FastMCPContext, ApiResponse, SuccessResponse, ToolRequestParams } from '../../types/common.types.js';
import getActiveOfficeDocumentsTool, { ActiveOfficeDocument } from '../os/getActiveOfficeDocuments.tool.js'; // Import ActiveOfficeDocument
import { wordTextTool } from './text.tool.js'; // getText is part of wordTextTool array
import { wordGenerateAndInsertTextTool } from './generateAndInsertText.tool.js';
import logger from '../../utils/logger.js';
import { createErrorResponse, handleToolError } from '../../utils/errorHandler.js';

const ConcludeStoryInDocumentInputSchema = z.object({
  documentNameOrPath: z.string().optional().default("Historias de luis.docx").describe(
    "The name or full path of the Word document. Defaults to 'Historias de luis.docx'."
  ),
  contextParagraphsCount: z.number().int().positive().optional().default(5).describe(
    "The number of preceding paragraphs to use as context for story conclusion. Defaults to 5."
  ),
  useComInterop: z.boolean().optional().default(false).describe(
    "Specifies the processing path. `true` uses COM Interop. `false` uses platform-independent libraries. Defaults to `false`."
  )
});

type ConcludeStoryInDocumentInput = z.infer<typeof ConcludeStoryInDocumentInputSchema>;

const concludeStoryInDocumentTool: McpResource = {
  path: 'word/concludeStoryInDocument', // Changed name to path
  description: 'Identifies a Word document, extracts preceding story context, generates a conclusion using AI, and inserts it at the end of the document.',
  schema: ConcludeStoryInDocumentInputSchema,
  handler: async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<string>> => { // Changed execute to handler, params type, context type and return type
    const validatedParams = params as ConcludeStoryInDocumentInput; // Assuming params are validated by FastMCP before handler is called
    logger.info(`Executing concludeStoryInDocumentTool with params: ${JSON.stringify(validatedParams)}`);
    const { documentNameOrPath, contextParagraphsCount, useComInterop } = validatedParams;

    try {
      let targetPath = documentNameOrPath;

      // Step 2: Identify Target Word Document
      if (!targetPath.endsWith('.docx') && !targetPath.includes('/') && !targetPath.includes('\\\\')) {
        // Assuming it's a name, try to find it in active documents
        const activeDocsResponse = await getActiveOfficeDocumentsTool.handler({}, context);
        if (!activeDocsResponse.success) {
          logger.error('Failed to get active office documents', activeDocsResponse.error);
          return createErrorResponse('ACTIVE_DOCS_ERROR', 'Failed to retrieve active Office documents.', activeDocsResponse.error);
        }
        const activeWordDocs = activeDocsResponse.data.documents.filter((doc: ActiveOfficeDocument) => doc.applicationType === 'Word');
        const foundDoc = activeWordDocs.find((doc: ActiveOfficeDocument) => doc.resolvedPath && doc.resolvedPath.endsWith(targetPath)); // Simple name match

        if (foundDoc && foundDoc.resolvedPath) {
          targetPath = foundDoc.resolvedPath;
          logger.info(`Found active document: ${targetPath}`);
        } else {
          // For this task, if a name is given and it's not active, an error is returned.
          // Future enhancement: attempt to find it in a default location or prompt user.
          logger.warn(`Document '${documentNameOrPath}' not found among active Word documents and no full path provided.`);
          return createErrorResponse('DOCUMENT_NOT_FOUND', `Document '${documentNameOrPath}' not found among active Word documents. Please open the document or provide a full path.`);
        }
      }
      logger.info(`Target document path: ${targetPath}`);

      // Step 3: Extract Preceding Story Context
      const getTextHandler = wordTextTool.find(tool => tool.path === 'word/text/get')?.handler;
      if (!getTextHandler) {
        return createErrorResponse('TOOL_NOT_FOUND', 'getTextTool handler not found.');
      }
      const textContentResponse = await getTextHandler({ filePath: targetPath, range: 'document', useComInterop }, context);

      if (!textContentResponse.success) {
        logger.error('Failed to get text content from document', textContentResponse.error);
        return createErrorResponse('GET_TEXT_FAILED', `Failed to extract text from '${targetPath}'.`, textContentResponse.error);
      }
      const fullTextContent = textContentResponse.data as string; // Cast based on getText return type
      const paragraphs = fullTextContent.split(/\r?\n\s*\r?\n/); // Split by one or more empty lines
      const storyContext = paragraphs.slice(-contextParagraphsCount).join('\n\n');
      logger.info(`Extracted context (first 100 chars): "${storyContext.substring(0, 100)}..."`);

      // Step 4: Generate and Insert Conclusion
      const generationPrompt = `Based on the following story context, please write a compelling conclusion for the story:\n\nContext:\n${storyContext}\n\nConclusion:`;
      const generateAndInsertHandler = wordGenerateAndInsertTextTool.handler;

      const generationResult = await generateAndInsertHandler({
        filePath: targetPath,
        position: 'end', // Insert at the end of the document
        prompt: generationPrompt,
        // maxTokens: undefined, // Let the tool use its default or be configured elsewhere if needed
        // useComInterop is not a direct param of generateAndInsertTextTool, it's handled internally by it.
      }, context);

      if (!generationResult.success) {
        logger.error('Failed to generate and insert text', generationResult.error);
        return createErrorResponse('GENERATION_INSERTION_FAILED', `Failed to generate and insert conclusion into '${targetPath}'.`, generationResult.error);
      }
      logger.warn("Note: generateAndInsertTextTool's useComInterop is handled internally by that tool, not passed from here.");


      // Step 5: Output/Return Value
      const successMessage = `Successfully concluded the story in document: ${targetPath}`;
      logger.info(successMessage);
      return { success: true, data: successMessage } as SuccessResponse<string>;

    } catch (error) {
      logger.error(`Error in concludeStoryInDocumentTool: ${error instanceof Error ? error.message : 'Unknown error'}`, error);
      return handleToolError(error, 'CONCLUDE_STORY_FAILED');
    }
  },
};

export { concludeStoryInDocumentTool, ConcludeStoryInDocumentInputSchema }; // Path is used for registration, name is for user display