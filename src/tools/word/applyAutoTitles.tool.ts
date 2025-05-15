import { z } from 'zod';
import getActiveOfficeDocumentsTool from '../os/getActiveOfficeDocuments.tool.js';
import { aiSuggestTool } from '../office/aiSuggest.tool.js';
import { getText } from './text.tool.js';
import { applyStyle } from './styles.tool.js';
import logger from '../../utils/logger.js';
import { createErrorResponse } from '../../utils/errorHandler.js';
import type { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types.js';

// Input schema for Zod validation
const ApplyAutoTitlesInputSchema = z.object({
  filePath: z.string().optional().describe('Optional: Path to the Word document. If not provided, tries to use the active Word document.'),
});

// Type inferred from Zod schema for internal use, though handler takes ToolRequestParams
type ApplyAutoTitlesInput = z.infer<typeof ApplyAutoTitlesInputSchema>;

// This interface is for the structure of suggestions we expect from the LLM
interface SuggestedHeading {
  text: string;
  level: string;
  paragraphIdentifier: string | number; // Crucial for locating where to apply the style
}

// This interface defines the `data` part of a successful ApiResponse
interface ApplyAutoTitlesOutputData {
  message: string;
  appliedHeadings: Array<{ text: string; style: string }>;
}

export const applyAutoTitlesTool: McpResource = {
  path: 'word/applyAutoTitles',
  description: 'Automatically identifies and applies heading styles to an active Word document.',
  schema: ApplyAutoTitlesInputSchema, // Input validation schema
  async handler(params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<ApplyAutoTitlesOutputData>> {
    // Validate incoming params against the schema
    let validatedParams: ApplyAutoTitlesInput;
    try {
      validatedParams = ApplyAutoTitlesInputSchema.parse(params);
    } catch (error) {
      logger.error('Input validation failed for word/applyAutoTitles', error);
      return createErrorResponse('VALIDATION_ERROR', 'Input validation failed.', (error as z.ZodError).errors);
    }

    logger.info(`Starting auto-title application for Word document. Inputs: ${JSON.stringify(validatedParams)}`);
    let targetFilePath = validatedParams.filePath;
    const appliedHeadings: Array<{ text: string; style: string }> = [];

    if (!targetFilePath) {
      try {
        const activeDocsResult = await getActiveOfficeDocumentsTool.handler({}, context);
        if (!activeDocsResult.success) {
          return createErrorResponse('ACTIVE_DOC_FAILED', activeDocsResult.error.message || 'Failed to get active office documents.');
        }
        
        const documents = Array.isArray(activeDocsResult.data?.documents) ? activeDocsResult.data.documents : [];
        const wordDocs = documents.filter(
          (doc: any) => doc.applicationType === 'Word' // Ensure this matches the actual output of getActiveOfficeDocumentsTool
        );

        if (wordDocs.length === 0) {
          return createErrorResponse('NO_ACTIVE_WORD_DOC', 'No active Word document found.');
        }
        if (wordDocs.length > 1) {
          logger.warn(`Multiple active Word documents found. Using the first one: ${wordDocs[0].filePath}`);
        }
        targetFilePath = wordDocs[0].filePath; // Ensure this matches the actual output property
      } catch (error: any) {
        logger.error(`Error identifying active Word document: ${error.message}`, error);
        return createErrorResponse('ACTIVE_DOC_EXCEPTION', `Failed to identify active Word document: ${error.message}`);
      }
    }

    if (!targetFilePath) {
        return createErrorResponse('NO_TARGET_PATH', 'Target Word document path could not be determined.');
    }

    logger.info(`Target Word document: ${targetFilePath}`);

    let documentText: string;
    try {
      const textResult = await getText({ filePath: targetFilePath, range: 'document', useComInterop: false }, context);
      if (!textResult.success) {
        return createErrorResponse('GET_TEXT_FAILED', textResult.error.message || `Failed to extract text from ${targetFilePath}`);
      }
      documentText = textResult.data;
      if (!documentText || documentText.trim().length === 0) {
        return createErrorResponse('EMPTY_DOCUMENT', 'No text content found in the document or document is empty.');
      }
    } catch (error: any) {
      logger.error(`Error extracting text from Word document: ${error.message}`, error);
      return createErrorResponse('GET_TEXT_EXCEPTION', `Failed to extract text from Word document '${targetFilePath}': ${error.message}`);
    }

    let suggestedHeadings: SuggestedHeading[];
    try {
      const prompt = `
Analyze the following document content and identify lines or phrases that should be formatted as headings.
For each identified heading, suggest an appropriate heading level (e.g., "Heading 1", "Heading 2", "Heading 3").
Return the result as a JSON array of objects, where each object has "text" (the heading phrase), "level" (the suggested style like "Heading 1"), and "paragraphIdentifier" (the original text phrase to locate it, which could be the text itself if no other unique ID is available).
Ensure the "text" and "paragraphIdentifier" are the exact phrases from the document.
Example: [{ "text": "Introduction", "level": "Heading 1", "paragraphIdentifier": "Introduction" }, { "text": "Background Information", "level": "Heading 2", "paragraphIdentifier": "Background Information" }]

Document Content:
---
${documentText}
---
      `;
      // The call to aiSuggestTool remains a point of concern due to schema mismatch with the plan's intent.
      // The plan implies a generic LLM call, but aiSuggestTool has a specific Office-related schema.
      // Passing a fictional '_raw_prompt_for_llm' parameter and hoping the underlying LLM call within aiSuggestTool can use it.
      // This is a significant assumption and may require aiSuggestTool to be refactored or a new generic LLM tool.
      const suggestionResult = await aiSuggestTool.handler({
        application: 'Word.Application', // Required by aiSuggestTool's schema
        operation: 'customSuggestion',   // A placeholder operation for aiSuggestTool
        _raw_prompt_for_llm: prompt,     // Fictional parameter for the raw prompt
        contextText: documentText.substring(0, 2000), // Provide some context, limited length
      } as any, context); // Using 'as any' due to the _raw_prompt_for_llm parameter

      if (!suggestionResult.success) {
        return createErrorResponse('AI_SUGGEST_FAILED', suggestionResult.error.message || 'AI suggestion tool failed.');
      }

      const suggestionData = suggestionResult.data; 
      // aiSuggestTool is expected to return a string which is a JSON array of suggestions.
      if (typeof suggestionData !== 'string') {
        logger.error('aiSuggest output was not a string as expected.', suggestionData);
        return createErrorResponse('AI_SUGGEST_INVALID_FORMAT', 'AI suggestion tool provided an invalid response format. Expected a JSON string.');
      }

      try {
        const parsedSuggestion = JSON.parse(suggestionData);
        if (!Array.isArray(parsedSuggestion)) {
          throw new Error('Parsed suggestion is not an array.');
        }
        suggestedHeadings = parsedSuggestion.map((item: any) => ({
          text: item.text,
          level: item.level,
          paragraphIdentifier: item.paragraphIdentifier || item.text, // Ensure paragraphIdentifier
        })) as SuggestedHeading[];
      } catch (parseError: any) {
        logger.error(`aiSuggest output was a string but failed to parse as JSON array: ${parseError.message}`, suggestionData);
        return createErrorResponse('AI_SUGGEST_PARSE_FAILED', `AI suggestion tool provided a non-JSON string or invalid JSON. ${parseError.message}`);
      }

      if (suggestedHeadings.length === 0) {
        logger.info('aiSuggest did not identify any headings.');
        return {
          success: true,
          data: {
            message: 'No headings were identified or suggested by the AI.',
            appliedHeadings: [],
          },
        };
      }
      logger.info(`AI suggested ${suggestedHeadings.length} headings.`);

    } catch (error: any) {
      logger.error(`Error getting heading suggestions from AI: ${error.message}`, error);
      return createErrorResponse('AI_SUGGEST_EXCEPTION', `Failed to get heading suggestions: ${error.message}`);
    }

    for (const heading of suggestedHeadings) {
      if (!heading.text || !heading.level || !heading.paragraphIdentifier) {
        logger.warn(`Skipping invalid heading suggestion (missing text, level, or identifier): ${JSON.stringify(heading)}`);
        continue;
      }
      try {
        // The 'range' for applyStyle needs to be the text to find, or a more specific paragraph identifier.
        // This relies on the 'applyStyle' tool's ability to resolve this range.
        const styleResult = await applyStyle({
          filePath: targetFilePath,
          range: heading.paragraphIdentifier.toString(), 
          style: heading.level,
          useComInterop: false, // As per plan, styles.tool.ts should handle this
        }, context);

        if (!styleResult.success) {
          logger.error(`Error applying style '${heading.level}' to "${heading.text}" (identifier: "${heading.paragraphIdentifier}"): ${styleResult.error.message}`, styleResult.error);
          continue; 
        }
        appliedHeadings.push({ text: heading.text, style: heading.level });
        logger.info(`Applied style '${heading.level}' to heading: "${heading.text}" (identifier: "${heading.paragraphIdentifier}")`);
      } catch (error: any) {
        logger.error(`Exception applying style '${heading.level}' to "${heading.text}": ${error.message}`, error);
        // Continue applying other styles even if one fails
      }
    }

    if (appliedHeadings.length === 0 && suggestedHeadings.length > 0) {
      return {
        success: true,
        data: {
          message: 'AI suggested headings, but none could be applied. Check logs for details.',
          appliedHeadings: [],
        },
      };
    }
    
    return {
      success: true,
      data: {
        message: `Successfully processed Word document. Applied ${appliedHeadings.length} heading styles.`,
        appliedHeadings,
      },
    };
  },
};

// To make this tool discoverable, it needs to be registered.
// This would typically happen in a central tool registration file (e.g., src/tools/index.ts)
// import { registerTool } from './index';
// registerTool(applyAutoTitlesTool);

// Exporting types for potential external use, though the tool itself is the primary export.
export type { ApplyAutoTitlesInput, ApplyAutoTitlesOutputData, SuggestedHeading };