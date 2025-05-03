/**
 * @file Aggregates all tool and resource definitions for registration with the MCP server.
 * Exports lists of tools and resources conforming to the McpResource interface.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { McpResource } from '@/types/common.types';
import { z } from 'zod';
// Import error handling functions used in placeholders and adapter
import { createErrorResponse, handleToolError } from '@/utils/errorHandler';

// Import tool definitions (assuming default or named exports matching the variable names)
import { fsDirectoryTool } from './fs/directory.tool'; // Exports McpResource[]
import { fsFileTool } from './fs/file.tool'; // Exports McpResource[]
import { wordStylesTool } from './word/styles.tool'; // Exports McpResource[]
import { wordMarkdownTool } from './word/markdown.tool'; // Exports McpResource[]
import { wordMergeTool } from './word/merge.tool'; // Exports McpResource (single object)
import { wordTemplateTool } from './word/template.tool'; // Exports McpResource[]
import { wordTextTool } from './word/text.tool'; // Exports McpResource[]
import { wordSearchReplaceTool } from './word/searchReplace.tool'; // Exports McpResource (single object)
import { wordTablesTool } from './word/tables.tool'; // Exports McpResource (single object)
import { wordChartsTool } from './word/charts.tool'; // Exports McpResource (single object)
import { wordImageExtractTool, wordImageInsertTool } from './word/image.tool'; // Exports individual McpResource objects
import { wordGenerateAndInsertTextTool } from './word/generateAndInsertText.tool'; // Exports McpResource (single object)
import { wordPageTool } from './word/page.tool'; // Exports McpResource (single object)
import { wordHeadersFootersTool } from './word/headersFooters.tool'; // Exports McpResource (single object)
import { embeddedObjectsTool } from './word/embeddedObjects.tool'; // Import the embeddedObjectsTool directly
import { batchTool } from './word/batch.tool'; // Import the new batch tool
import mermaidImportTool from './word/mermaidImport.tool'; // Import the mermaid import tool
import mermaidExportTool from './word/mermaidExport.tool'; // Import the mermaid export tool
import reformatTool from './word/reformat.tool'; // Import the new reformat tool
import { registerWordCodeFormatTool } from './word/codeFormat.tool'; // Import the new code format tool registration function

import analyzeToolDefinition from './word/analyze.tool'; // Import the new analyze tool definition

// Import Excel tools
import { excelWorksheetsTool } from './excel/worksheets.tool';
import excelRangeTool from './excel/range.tool'; // Import the new excel/range tool
import excelTablesTool from './excel/tables.tool'; // Import the new excel/tables tool
import excelChartsTool from './excel/charts.tool'; // Import the new excel/charts tool
import { excelDataAnalysisTool } from './excel/dataAnalysis.tool'; // Import the new excel/data-analysis tool

// Import PowerPoint tools
import slidesTool from './powerpoint/slides.tool'; // Import the new powerpoint/slides tool
import powerpointShapesTool from './powerpoint/shapes.tool'; // Import the new powerpoint/shapes tool
import { powerpointPropertiesTool } from './powerpoint/properties.tool'; // Import the new powerpoint/properties tool
import animationsTool from './powerpoint/animations.tool'; // Import the new powerpoint/animations tool
import officeTransferTool from './office/transfer.tool'; // Import the new office/transfer tool
import { OfficeWorkflowTool } from './office/workflow.tool'; // Import the new office/workflow tool
import { pdfExportTool } from './office/pdfExport.tool'; // Import the new office/pdf/export tool
import pdfParseTool from './office/pdfParse.tool'; // Import the new office/pdf/parse tool
import { officeCombineTool } from './office/combine.tool'; // Import the new office/combine tool
// Import dynamic tools
import { dynamicResourcesTool } from './dynamic/resources.tool';

// Import static resources
import aiAssistantGuideResource from '../resources/static/ai_assistant_guide.resource';

/**
 * Adapts the analyze tool definition to the McpResource interface.
 * Provides a placeholder handler as the execution is delegated to the server wrapper.
 */
const analyzeMcpResource: McpResource = {
    path: analyzeToolDefinition.name, // Use name as path
    handler: async (params, context) => {
        try {
             // Placeholder handler, execution delegated to server.ts wrapper
             return { success: true, data: "Handler called, but execution delegated to server.ts wrapper." };
        } catch (error) {
             return handleToolError(error, 'ANALYZE_ERROR'); // Use a specific error code
        }
    },
    schema: analyzeToolDefinition.inputSchema as z.ZodSchema<any>, // Use inputSchema as schema
    description: analyzeToolDefinition.description,
};


/**
 * A flattened array containing all implemented tool definitions conforming to the McpResource interface.
 * This list is used by the server to register the tools.
 */
const allImplementedTools: McpResource[] = [
    ...fsDirectoryTool, // Array
    ...fsFileTool,      // Array
    ...wordStylesTool,  // Array
    ...wordMarkdownTool,// Array
    ...wordTemplateTool,// Array
    ...wordTextTool,    // Array
    wordImageExtractTool, // Single object
    wordImageInsertTool,  // Single object
    // Single objects
    wordMergeTool,
    wordSearchReplaceTool,
    wordTablesTool,
    wordChartsTool,
    wordGenerateAndInsertTextTool,
    wordPageTool,
    wordHeadersFootersTool,
    embeddedObjectsTool,
    batchTool,
    mermaidImportTool,
    mermaidExportTool,
    reformatTool,
    analyzeMcpResource,
    excelWorksheetsTool,
    excelRangeTool,
    excelTablesTool,
    excelChartsTool,
    excelDataAnalysisTool,
    slidesTool,
    powerpointShapesTool,
    powerpointPropertiesTool,
    animationsTool,
    officeTransferTool,
    new OfficeWorkflowTool(),
    pdfExportTool,
    pdfParseTool,
    officeCombineTool,
    dynamicResourcesTool,
     // Add other imported tools here
 ].flat(); // Flatten the array to ensure it only contains McpResource objects


// Placeholder tools for unimplemented features (ensure they conform to McpResource)
const placeholderTools: McpResource[] = [
    // Add placeholders for any tools that are defined but not yet fully implemented
    // Example:
    // {
    //     path: 'office/ai-suggest',
    //     handler: async () => createErrorResponse('NOT_IMPLEMENTED', 'Tool office/ai-suggest not implemented.'),
    //     description: 'Provides AI-powered suggestions within Office documents (Not Implemented).',
    //     schema: z.object({}), // Define a basic schema even for placeholders if it's a tool
    // },
 ];


/**
 * The final list of all tools and resources to be registered with the FastMCP server.
 * Combines implemented tools and any defined placeholders.
 */
const allRegisteredTools = [...allImplementedTools, ...placeholderTools];

// Export the list used by server.ts and the resource
export { allRegisteredTools, aiAssistantGuideResource, registerWordCodeFormatTool };