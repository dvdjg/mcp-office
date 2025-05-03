// /src/tools/index.ts
// =============================================================================
/**
 * @file Aggregates all tool definitions for registration with the MCP server.
 * Exports tools conforming to the McpResource interface.
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
import { wordImageTools } from './word/image.tool'; // Exports McpResource[]
import { wordGenerateAndInsertTextTool } from './word/generateAndInsertText.tool'; // Exports McpResource (single object)
import { wordPageTool } from './word/page.tool'; // Exports McpResource (single object)
import { wordHeadersFootersTool } from './word/headersFooters.tool'; // Exports McpResource (single object)
import { embeddedObjectsTool as embeddedObjectsToolDefinition } from './word/embeddedObjects.tool'; // Import the FastMCP definition
import { batchTool } from './word/batch.tool'; // Import the new batch tool
import mermaidImportTool from './word/mermaidImport.tool'; // Import the mermaid import tool
import mermaidExportTool from './word/mermaidExport.tool'; // Import the mermaid export tool
import reformatTool from './word/reformat.tool'; // Import the new reformat tool
import { registerWordCodeFormatTool } from './word/codeFormat.tool'; // Import the new code format tool registration function

import analyzeToolDefinition from './word/analyze.tool'; // Import the new analyze tool definition

// Import Excel tools
import { excelWorksheetsTool } from './excel/worksheets.tool';
import excelRangeTool from './excel/range.tool'; // Importar la nueva herramienta excel/range
import excelTablesTool from './excel/tables.tool'; // Importar la nueva herramienta excel/tables
import excelChartsTool from './excel/charts.tool'; // Importar la nueva herramienta excel/charts
import { excelDataAnalysisTool } from './excel/dataAnalysis.tool'; // Importar la nueva herramienta excel/data-analysis

// Import PowerPoint tools
import slidesTool from './powerpoint/slides.tool'; // Importar la nueva herramienta powerpoint/slides
import powerpointShapesTool from './powerpoint/shapes.tool'; // Importar la nueva herramienta powerpoint/shapes
import { powerpointPropertiesTool } from './powerpoint/properties.tool'; // Importar la nueva herramienta powerpoint/properties
import animationsTool from './powerpoint/animations.tool'; // Importar la nueva herramienta powerpoint/animations
import officeTransferTool from './office/transfer.tool'; // Importar la nueva herramienta office/transfer
import { OfficeWorkflowTool } from './office/workflow.tool'; // Importar la nueva herramienta office/workflow
import { pdfExportTool } from './office/pdfExport.tool'; // Importar la nueva herramienta office/pdf/export
import pdfParseTool from './office/pdfParse.tool'; // Importar la nueva herramienta office/pdf/parse
import { officeCombineTool } from './office/combine.tool'; // Importar la nueva herramienta office/combine
// Import dynamic tools
import { dynamicResourcesTool } from './dynamic/resources.tool';

// Import static resources
import aiAssistantGuideResource from '../resources/static/ai_assistant_guide.resource';

// Adapt the FastMCP tool definition to the McpResource interface
const embeddedObjectsMcpResource: McpResource = {
    path: embeddedObjectsToolDefinition.name, // Use name as path
    handler: async (params, context) => {
        try {
             // Placeholder handler, execution delegated to server.ts wrapper
             return { success: true, data: "Handler called, but execution delegated to server.ts wrapper." };
        } catch (error) {
             return handleToolError(error, 'EMBEDDED_OBJECTS_ERROR');
        }
    },
    schema: embeddedObjectsToolDefinition.parameters as z.ZodSchema<any>, // Use parameters as schema
    description: embeddedObjectsToolDefinition.description,
};

// Adapt the analyze tool definition to the McpResource interface
const analyzeMcpResource: McpResource = {
    path: analyzeToolDefinition.name, // Use name as path
    handler: async (params, context) => {
        try {
             // Placeholder handler, execution delegated to server.ts wrapper
             return { success: true, data: "Handler called, but execution delegated to server.ts wrapper." };
        } catch (error) {
             return handleToolError(error, 'ANALYZE_ERROR'); // Usar un código de error específico
        }
    },
    schema: analyzeToolDefinition.inputSchema as z.ZodSchema<any>, // Use inputSchema as schema
    description: analyzeToolDefinition.description,
};


// Combine all tools into a potentially nested array first
const nestedToolsList = [
    fsDirectoryTool, // Array
    fsFileTool,      // Array
    wordStylesTool,  // Array
    wordMarkdownTool,// Array
    wordTemplateTool,// Array
    wordTextTool,    // Array
    wordImageTools,  // Array
    // Single objects
    wordSearchReplaceTool,
    wordTablesTool,
    wordChartsTool,
    wordGenerateAndInsertTextTool,
    wordPageTool,
    wordHeadersFootersTool,
    embeddedObjectsMcpResource,
    batchTool, // Añadida la nueva herramienta batchTool
    mermaidImportTool, // Añadida la herramienta de importación de Mermaid
    mermaidExportTool, // Añadida la herramienta de exportación de Mermaid
    reformatTool, // Añadida la herramienta reformatTool
    analyzeMcpResource, // Añadida la herramienta analyzeTool adaptada
   excelWorksheetsTool, // Añadida la herramienta excel/worksheets
   excelRangeTool, // Añadida la nueva herramienta excel/range
   excelTablesTool, // Añadida la nueva herramienta excel/tables
   excelChartsTool, // Añadida la nueva herramienta excel/charts
   excelDataAnalysisTool, // Añadida la nueva herramienta excel/data-analysis
   slidesTool, // Añadida la nueva herramienta powerpoint/slides
   powerpointShapesTool, // Añadida la nueva herramienta powerpoint/shapes
   powerpointPropertiesTool, // Añadida la nueva herramienta powerpoint/properties
   animationsTool, // Añadida la nueva herramienta powerpoint/animations
   officeTransferTool, // Añadida la nueva herramienta office/transfer
   new OfficeWorkflowTool(), // Añadida la nueva herramienta office/workflow
   pdfExportTool, // Añadida la nueva herramienta office/pdf/export
   pdfParseTool, // Añadida la nueva herramienta office/pdf/parse
   officeCombineTool, // Añadida la nueva herramienta office/combine
   dynamicResourcesTool, // Añadida la herramienta dynamic/resources
     // Add other imported tools here
 ];

// Flatten the array to ensure it only contains McpResource objects
const allImplementedTools: McpResource[] = nestedToolsList.flat();

// Placeholder tools for unimplemented features (ensure they conform to McpResource)
const placeholderTools: McpResource[] = [
    // { path: 'word/embedded-objects', ... }, // REMOVED
    // { path: 'word/metadata', handler: async () => createErrorResponse('NOT_IMPLEMENTED', 'Tool word/metadata not implemented.'), description: 'Manage document properties and comments (Not Implemented)', schema: z.object({}) }, // REMOVED - Implemented
    // { path: 'word/batch', handler: async () => createErrorResponse('NOT_IMPLEMENTED', 'Tool word/batch not implemented.'), description: 'Execute multiple operations (Not Implemented)', schema: z.object({}) }, // Eliminado el placeholder
    // { path: 'word/mermaid/import', handler: async () => createErrorResponse('NOT_IMPLEMENTED', 'Tool word/mermaid/import not implemented.'), description: 'Import and render Mermaid diagrams (Not Implemented)', schema: z.object({}) }, // Eliminado el placeholder
    // { path: 'word/mermaid/export', handler: async () => createErrorResponse('NOT_IMPLEMENTED', 'Tool word/mermaid/export not implemented.'), description: 'Export Mermaid diagrams (Not Implemented)', schema: z.object({}) }, // Eliminado el placeholder
    // { path: 'word/reformat', handler: async () => createErrorResponse('NOT_IMPLEMENTED', 'Tool word/reformat not implemented.'), description: 'Reformat a document professionally (Not Implemented)', schema: z.object({}) }, // Eliminado el placeholder
    // { path: 'word/code-format', handler: async () => createErrorResponse('NOT_IMPLEMENTED', 'Tool word/code-format not implemented.'), description: 'Format code and metadata with syntax highlighting (Not Implemented)', schema: z.object({}) }, // Eliminado el placeholder
    // Excel (Placeholders)
    // { path: 'excel/worksheets', handler: async () => createErrorResponse('NOT_IMPLEMENTED', 'Tool excel/worksheets not implemented.'), description: 'Manage Excel worksheets (Not Implemented - Requires Office JS/Scripts)', schema: z.object({}) }, // Eliminado el placeholder
    // { path: 'excel/range', handler: async () => createErrorResponse('NOT_IMPLEMENTED', 'Tool excel/range not implemented.'), description: 'Manipulate cell ranges (Not Implemented - Requires Office JS/Scripts)', schema: z.object({}) }, // Eliminado el placeholder
   // PowerPoint (Placeholders) - REMOVED as implemented
 ]; // Eliminado el placeholder para office/ai-suggest


// Combine implemented and placeholder tools
const allRegisteredTools = [...allImplementedTools, ...placeholderTools];

// Export the list used by server.ts and the resource
export { allRegisteredTools, aiAssistantGuideResource, registerWordCodeFormatTool };