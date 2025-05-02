import { promises as fs } from 'fs';
import path from 'path';
import { McpResource, ApiResponse, ToolRequestParams } from '../../types/common.types'; // Added ToolRequestParams
import { handleToolError } from '../../utils/errorHandler';
import logger from '../../utils/logger'; // Changed to default import
import { Context as FastMCPContext } from 'fastmcp'; // Import FastMCP Context

const guidePath = path.resolve(__dirname, '../../docs/ai_assistant_guide.md');

// Handler now accepts params and an optional context (FastMCPContext<undefined>)
// Context is optional here because the resource loader in server/index.ts doesn't provide it.
const readAiAssistantGuide = async (
    params: ToolRequestParams = {},
    context?: FastMCPContext<undefined> // Add optional context parameter
): Promise<ApiResponse<string>> => {
  try {
    // Use context.log if available, otherwise use the global logger
    const log = context?.log ?? logger;
    log.info('Reading AI assistant guide...', { params }); // Log params if needed
    const content = await fs.readFile(guidePath, 'utf-8');
    logger.info('AI assistant guide read successfully.');
    return { success: true, data: content };
  } catch (error) {
    return handleToolError(error, 'Error reading AI assistant guide');
  }
};

// Define the resource according to the McpResource interface
export const aiAssistantGuideReadResource: McpResource = {
  path: 'memory/ai_assistant_guide/read', // Full path for the operation
  handler: readAiAssistantGuide,
  // No input schema needed for a simple read operation
  // schema: Zod.object({}), // Example if input validation was needed
  description: 'Reads the static AI assistant usage guide from memory/ai_assistant_guide.md.',
  // completions and outputSchema are not part of the current McpResource interface definition
};

// Exporting as an array containing the single resource for easy spreading
export default [aiAssistantGuideReadResource];