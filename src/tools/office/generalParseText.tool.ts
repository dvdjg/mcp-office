import { z } from 'zod';
import officeparser from 'officeparser';
import fs from 'fs-extra';
import { McpResource, ToolRequestParams, ApiResponse, FastMCPContext } from '@/types/common.types';
import { createErrorResponse, handleToolError } from '@/utils/errorHandler';
import logger from '@/utils/logger'; // Import logger for more detailed logging if needed

const generalParseTextToolInputSchema = z.object({
  filePath: z.string().min(1, 'File path cannot be empty.'),
});

// This is the schema for the 'data' part of the successful ApiResponse
const generalParseTextToolSuccessDataSchema = z.object({
  extractedText: z.string(),
  message: z.string(),
  detectedFileType: z.string().optional(),
});

// Infer the type for the success data
type GeneralParseTextSuccessData = z.infer<typeof generalParseTextToolSuccessDataSchema>;

/**
 * @tool generalParseTextHandler
 * @description Handler function for the generalParseText tool.
 * Extracts text content from various Office document types using the officeparser library.
 *
 * @param {ToolRequestParams} params - The input parameters, validated against generalParseTextToolInputSchema.
 * @param {FastMCPContext} [context] - The FastMCP context (optional).
 * @returns {Promise<ApiResponse<GeneralParseTextSuccessData>>}
 *          An ApiResponse object containing the extracted text, a status message, and optionally the detected file type,
 *          or an error response.
 */
const generalParseTextHandler = async (
  params: ToolRequestParams,
  _context?: FastMCPContext<any> // context is available but not used in this specific tool
): Promise<ApiResponse<GeneralParseTextSuccessData>> => {
  let validatedInput: { filePath: string };
  try {
    validatedInput = generalParseTextToolInputSchema.parse(params);
  } catch (error) {
    return handleToolError(error, 'VALIDATION_ERROR');
  }

  const { filePath } = validatedInput;

  try {
    const fileExists = await fs.pathExists(filePath);
    if (!fileExists) {
      logger.warn(`File not found at path: ${filePath} for generalParseText tool.`);
      return createErrorResponse('FILE_NOT_FOUND', `Error: File not found at path: ${filePath}`);
    }

    const result = await officeparser.parseOfficeAsync(filePath);
    let extractedText = '';
    let detectedFileType: string | undefined = undefined;

    if (typeof result === 'string') {
      extractedText = result;
    } else if (typeof result === 'object' && result !== null && 'content' in result) {
      extractedText = String((result as any).content || ''); // Ensure content is string
      if ('meta' in result && typeof (result as any).meta === 'object' && (result as any).meta !== null && 'type' in (result as any).meta) {
        detectedFileType = String((result as any).meta.type);
      }
    } else if (result !== null && result !== undefined) {
      // If result is not string or expected object, try to stringify it, could be an unexpected structure
      logger.warn(`Unexpected result type from officeparser for ${filePath}: ${typeof result}`);
      extractedText = JSON.stringify(result);
    } else {
        logger.warn(`Null or undefined result from officeparser for ${filePath}`);
        extractedText = ''; // Default to empty string if result is null/undefined
    }
    
    logger.info(`Text extracted successfully from ${filePath}. Length: ${extractedText.length}`);
    return {
      success: true,
      data: {
        extractedText,
        message: 'Text extracted successfully.',
        detectedFileType,
      },
    };
  } catch (error: any) {
    logger.error(`Error parsing file ${filePath} with officeparser:`, error);
    // Use handleToolError for consistent error formatting
    // Provide a more specific error code if possible, e.g., 'OFFICEPARSER_ERROR'
    return handleToolError(error, 'OFFICEPARSER_ERROR');
  }
};

export const generalParseTextTool: McpResource = {
  path: 'office/generalParseText', // Define a unique path for the tool
  handler: generalParseTextHandler,
  schema: generalParseTextToolInputSchema, // Zod schema for input validation
  description:
    'Extracts text content from various Office document types (.docx, .xlsx, .pptx, .odt, .ods, .odp, etc.) using officeparser. For server-side quick text extraction. Relies on officeparser for actual parsing and supported types.',
};

// Export input and output types for external use if needed, though ApiResponse<GeneralParseTextSuccessData> is the main output type.
export type GeneralParseTextInput = z.infer<typeof generalParseTextToolInputSchema>;
export type GeneralParseTextOutput = ApiResponse<GeneralParseTextSuccessData>;