/**
 * @file Tool for providing AI-powered suggestions based on the current context of an Office application.
 * Allows suggesting formats, search terms, or chart types.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse, FastMCPContext } from '@/types/common.types'; // Import McpResource, ToolRequestParams, ApiResponse and FastMCPContext
import { getOfficeApplication, OfficeAppName } from '../../utils/officeInterop'; // Import OfficeAppName
import { TextContent, ContentResult } from 'fastmcp'; // Import TextContent and ContentResult

// Define the input schema for the office/ai-suggest tool
const AiSuggestInputSchema = z.object({
  application: z.enum(['Word.Application', 'Excel.Application', 'PowerPoint.Application']), // Use full names
  operation: z.enum(['format', 'search', 'chart']),
  contextRange: z.string().optional(), // Could be an Excel range, a paragraph identifier, etc.
  filePath: z.string().optional(), // File path if relevant
});

type AiSuggestInput = z.infer<typeof AiSuggestInputSchema>;

/**
 * @tool office/ai-suggest
 * @description Provides AI-powered suggestions based on the current context of an Office application.
 * Allows suggesting formats, search terms, or chart types.
 * @input AiSuggestInputSchema
 * @output z.string() // The AI-generated suggestion
 */
export const aiSuggestTool: McpResource = { // Use McpResource
  path: 'office/ai-suggest', // Define the path
  description: 'Provides AI-powered suggestions based on the current context of an Office application.',
  schema: AiSuggestInputSchema, // Use schema instead of inputSchema
  handler: async (params: ToolRequestParams, context?: FastMCPContext<any>) => { // Use ToolRequestParams and FastMCPContext<any>
    // Access the sampling function through the session context
    const requestSampling = context?.session?.requestSampling;

    try {
      // Validate input parameters
      const input = AiSuggestInputSchema.parse(params);

      if (!requestSampling) {
          throw new Error('The requestSampling function is not available in the session context.');
      }

      const { application, operation, contextRange, filePath } = input;

      // Get the active Office application
      const officeApp = await getOfficeApplication(application as OfficeAppName); // Use await and OfficeAppName
      if (!officeApp) {
        throw new Error(`Office application not found or not supported: ${application}`);
      }

      let contextText = ''; // Rename to avoid conflict with the context parameter
      try {
        // Attempt to get the selected text through the COM object
        // The way to access the selection varies slightly between applications
        if (application === 'Word.Application') {
            contextText = officeApp.Selection.Text;
        } else if (application === 'Excel.Application') {
            contextText = officeApp.Selection.Text; // Or Value, depending on the data type
        } else if (application === 'PowerPoint.Application') {
             // PowerPoint selection is more complex, might need to check ActiveWindow.Selection
             // For simplicity, we might skip selection context for now or get slide text
             contextText = 'Could not get selection context in PowerPoint.';
        }

      } catch (selectErr) {
        // If selection fails, attempt to get context from the active document
        try {
           const activeDoc = officeApp.ActiveDocument || officeApp.ActivePresentation || officeApp.ActiveWorkbook; // Property varies by application
           if (activeDoc) {
               if (application === 'Word.Application' && activeDoc.Content) {
                   contextText = activeDoc.Content.Text;
                   // Limit context size to avoid overly long prompts
                   if (contextText.length > 1000) {
                       contextText = contextText.substring(0, 1000) + '...';
                   }
               } else if (application === 'Excel.Application' && activeDoc.ActiveSheet) {
                   // We could try to get data from the active sheet or a specific range if contextRange is defined
                   // For now, we just indicate that specific context could not be obtained
                   contextText = `Active document: ${activeDoc.Name}. Could not get detailed context.`;
               } else if (application === 'PowerPoint.Application' && activeDoc.Slides) {
                   // We could try to get text from the current slide
                   contextText = `Active presentation: ${activeDoc.Name}. Could not get detailed context.`;
               } else {
                   contextText = `Could not get detailed context from the active document in ${application}.`;
               }
           } else {
               contextText = `Could not get context from the active document in ${application}.`;
           }
        } catch (docErr) {
            contextText = `Error attempting to get context from document or selection in ${application}.`;
            console.error("Error getting context:", docErr);
        }
      }


      let prompt = '';
      switch (operation) {
        case 'format':
          prompt = `Based on the following text or context from a ${application} document, suggest appropriate formatting options (styles, bold, italics, alignment, etc.). Context: "${contextText}"`;
          break;
        case 'search':
          prompt = `Based on the following text or context from a ${application} document, suggest relevant search terms or possible locations within the document to find related information. Context: "${contextText}"`;
          break;
        case 'chart':
          if (application !== 'Excel.Application') { // Use full name
              throw new Error(`The 'chart' operation is only supported for Excel. Current application: ${application}`);
          }
          prompt = `Based on the following context from an Excel worksheet, suggest appropriate chart types and possible data ranges to visualize. Context: "${contextText}"`;
          break;
        default:
          throw new Error(`Unsupported operation: ${operation}`);
      }

      // Generate suggestion using FastMCP sampling
      const samplingResult = await requestSampling({
          prompt: prompt, // Use the generated prompt
          maxTokens: 500, // Limit suggestion length (adjust as needed)
          // Other sampling parameters can be added here if relevant
      });

      // Process the sampling result
      let suggestion = '';
      if (samplingResult && samplingResult.content && samplingResult.content.length > 0) {
          // Assume the first text content block is the suggestion
          // Use type assertion to access 'type' and 'text' due to type error
          const textContent = samplingResult.content.find((c: any) => c.type === 'text') as TextContent | undefined;
          if (textContent) {
              suggestion = textContent.text;
          }
      }

      if (!suggestion) {
          // If no text was obtained from sampling, throw an error or return a default message
          throw new Error('FastMCP sampling did not generate any text suggestion.');
      }

      // Return a success response with the suggestion
      return { success: true, data: suggestion };

    } catch (error: any) {
      console.error(`Error in office/ai-suggest tool: ${error.message}`);
      // Return an error response
      return { success: false, error: { code: 'AI_SUGGEST_ERROR', message: `Error processing AI suggestion: ${error.message}` } };
    }
  },
};