/**
 * @file Defines the static resource for the AI assistant usage guide.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { promises as fs } from 'fs';
import path from 'path';
import { McpResource, ApiResponse, ToolRequestParams } from '../../types/common.types';
import { handleToolError } from '../../utils/errorHandler';
import logger from '../../utils/logger'; // Changed to default import
import { Context as FastMCPContext } from 'fastmcp'; // Import FastMCP Context

const guidePath = path.resolve(__dirname, '../../docs/ai_assistant_guide.md');

/**
 * Handler function to read the AI assistant usage guide.
 * @param params - Tool-specific parameters (not used for this resource). Defaults to an empty object.
 * @param context - The FastMCP context, optional for resources.
 * @returns A promise resolving to an ApiResponse containing the guide content or an error.
 */
const readAiAssistantGuide = async (
    params: ToolRequestParams = {},
    context?: FastMCPContext<undefined> // Add optional context parameter
): Promise<ApiResponse<string>> => {
  try {
    // Use context.log if available, otherwise use the global logger
    const log = context?.log ?? logger;
    log.info('[Resource] Reading AI assistant guide...', { params }); // Log params if needed
    const content = await fs.readFile(guidePath, 'utf-8');
    log.info('[Resource] AI assistant guide read successfully.');
    return { success: true, data: content };
  } catch (error) {
    return handleToolError(error, 'Error reading AI assistant guide');
  }
};

/**
 * Definition of the AI assistant guide resource.
 */
export const aiAssistantGuideReadResource: McpResource = {
  /** The unique path identifier for the resource. */
  path: 'office-mcp://ai_assistant_guide', // Full path for the operation
  /** The handler function to execute when the resource is accessed. */
  handler: readAiAssistantGuide,
  // No input schema needed for a simple read operation
  // schema: Zod.object({}), // Example if input validation was needed
  /** A brief description of the resource. */
  description: 'Reads the static AI assistant usage guide from memory/ai_assistant_guide.md.',
  // completions and outputSchema are not part of the current McpResource interface definition
};

// Exporting as an array containing the single resource for easy spreading
export default [aiAssistantGuideReadResource];