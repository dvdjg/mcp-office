/**
 * @file Aggregates all tool and resource definitions for registration with the MCP server.
 * Exports lists of tools and resources conforming to the McpResource interface.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { McpResource } from '@/types/common.types.js';
import { z } from 'zod';
// Import error handling functions used in placeholders and adapter
import { createErrorResponse, handleToolError } from '@/utils/errorHandler.js';

// Import tool definitions (assuming default or named exports matching the variable names)
import { fsDirectoryTool } from './fs/directory.tool.js'; // Exports McpResource[]
import { fsFileTool } from './fs/file.tool.js'; // Exports McpResource[]
import {
  fileContentTools,
  InsertTextLinesArgs,
  DeleteTextLinesArgs,
  ReplaceTextLinesArgs,
  ExtractTextFromRangeArgs,
  SearchReplaceInTextRangeArgs,
  SortTextLinesArgs,
  DeduplicateConsecutiveLinesArgs,
  TrimLineWhitespaceArgs,
  ReadBinaryAsHexArgs,
  WriteHexAsBinaryArgs,
  ReadBase64FileArgs,
  WriteToBase64FileArgs
} from './fs/fileContent.tool.js';
import { directoryOperationsTools } from './fs/directoryOperations.tool.js';
import {
  structuredDataTools,
  structuredDataSchemas,
} from './fs/structuredData.tool.js';
import { registerArchiveTools } from './fs/archive.tool.js'; // Import the new archive tool registration function
import { wordStylesTool } from './word/styles.tool.js'; // Exports McpResource[]
import { wordMarkdownTool } from './word/markdown.tool.js'; // Exports McpResource[]
import { wordMergeTool } from './word/merge.tool.js'; // Exports McpResource (single object)
import { wordTemplateTool } from './word/template.tool.js'; // Exports McpResource[]
import { wordTextTool } from './word/text.tool.js'; // Exports McpResource[]
import { wordSearchReplaceTool } from './word/searchReplace.tool.js'; // Exports McpResource (single object)
import { wordTablesTool } from './word/tables.tool.js'; // Exports McpResource (single object)
import { wordChartsTool } from './word/charts.tool.js'; // Exports McpResource (single object)
import { wordImageExtractTool, wordImageInsertTool } from './word/image.tool.js'; // Exports individual McpResource objects
import { wordGenerateAndInsertTextTool } from './word/generateAndInsertText.tool.js'; // Exports McpResource (single object)
import { wordPageTool } from './word/page.tool.js'; // Exports McpResource (single object)
import { wordHeadersFootersTool } from './word/headersFooters.tool.js'; // Exports McpResource (single object)
import { embeddedObjectsTool } from './word/embeddedObjects.tool.js'; // Import the embeddedObjectsTool directly
import { batchTool } from './word/batch.tool.js'; // Import the new batch tool
import mermaidImportTool from './word/mermaidImport.tool.js'; // Import the mermaid import tool
import mermaidExportTool from './word/mermaidExport.tool.js'; // Import the mermaid export tool
import reformatTool from './word/reformat.tool.js'; // Import the new reformat tool
import { registerWordCodeFormatTool } from './word/codeFormat.tool.js'; // Import the new code format tool registration function

import analyzeToolDefinition from './word/analyze.tool.js'; // Import the new analyze tool definition

// Import Excel tools
import { excelWorksheetsTool } from './excel/worksheets.tool.js';
import excelRangeTool from './excel/range.tool.js'; // Import the new excel/range tool
import excelTablesTool from './excel/tables.tool.js'; // Import the new excel/tables tool
import excelChartsTool from './excel/charts.tool.js'; // Import the new excel/charts tool
import { excelDataAnalysisTool } from './excel/dataAnalysis.tool.js'; // Import the new excel/data-analysis tool

// Import PowerPoint tools
import slidesTool from './powerpoint/slides.tool.js'; // Import the new powerpoint/slides tool
import powerpointShapesTool from './powerpoint/shapes.tool.js'; // Import the new powerpoint/shapes tool
import { powerpointPropertiesTool } from './powerpoint/properties.tool.js'; // Import the new powerpoint/properties tool
import animationsTool from './powerpoint/animations.tool.js'; // Import the new powerpoint/animations tool
import officeTransferTool from './office/transfer.tool.js'; // Import the new office/transfer tool
import { OfficeWorkflowTool } from './office/workflow.tool.js'; // Import the new office/workflow tool
import { pdfExportTool } from './office/pdfExport.tool.js'; // Import the new office/pdf/export tool
import pdfParseTool from './office/pdfParse.tool.js'; // Import the new office/pdf/parse tool
import { officeCombineTool } from './office/combine.tool.js'; // Import the new office/combine tool
// Import dynamic tools
import { dynamicResourcesTool } from './dynamic/resources.tool.js'; // Assuming single object
// Import the new image analysis/generation tool
import { officeImageAnalysisGenerationTool } from './office/imageAnalysisGeneration.tool.js'; // Single object
import { generalParseTextTool } from './office/generalParseText.tool.js'; // Import the new general parse text tool

// Import static resources
import aiAssistantGuideResource from '../resources/static/ai_assistant_guide.resource.js';

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

// Schemas for fileContentTools
const bufferEncodingSchema = z.enum(['ascii', 'utf8', 'utf-8', 'utf16le', 'ucs2', 'ucs-2', 'base64', 'latin1', 'binary', 'hex']).optional();

const InsertTextLinesArgsSchema = z.object({
  path: z.string(),
  lineNumber: z.number().int(),
  lines: z.array(z.string()),
  encoding: bufferEncodingSchema,
});

const DeleteTextLinesArgsSchema = z.object({
  path: z.string(),
  startLine: z.number().int(),
  endLine: z.number().int(),
  encoding: bufferEncodingSchema,
});

const ReplaceTextLinesArgsSchema = z.object({
  path: z.string(),
  startLine: z.number().int(),
  endLine: z.number().int(),
  newLines: z.array(z.string()),
  encoding: bufferEncodingSchema,
});

const ExtractTextFromRangeArgsSchema = z.object({
  path: z.string(),
  startLine: z.number().int(),
  startChar: z.number().int(),
  endLine: z.number().int(),
  endChar: z.number().int(),
  encoding: bufferEncodingSchema,
});

const SearchReplaceInTextRangeArgsSchema = z.object({
  path: z.string(),
  searchTerm: z.string(),
  replacement: z.string(),
  startLine: z.number().int().optional(),
  endLine: z.number().int().optional(),
  isRegex: z.boolean().optional(),
  replaceAll: z.boolean().optional(),
  encoding: bufferEncodingSchema,
});

const SortTextLinesArgsSchema = z.object({
  path: z.string(),
  options: z.object({
    reverse: z.boolean().optional(),
    caseSensitive: z.boolean().optional(),
    locale: z.string().optional(),
  }).optional(),
  encoding: bufferEncodingSchema,
});

const DeduplicateConsecutiveLinesArgsSchema = z.object({
  path: z.string(),
  caseSensitive: z.boolean().optional(),
  encoding: bufferEncodingSchema,
});

const TrimLineWhitespaceArgsSchema = z.object({
  path: z.string(),
  options: z.object({
    leading: z.boolean().optional(),
    trailing: z.boolean().optional(),
  }).optional(),
  encoding: bufferEncodingSchema,
});

const ReadBinaryAsHexArgsSchema = z.object({
  path: z.string(),
});

const WriteHexAsBinaryArgsSchema = z.object({
  path: z.string(),
  hexString: z.string(),
});

const ReadBase64FileArgsSchema = z.object({
  path: z.string(),
  outputEncoding: bufferEncodingSchema.or(z.literal('binary')).optional(),
});

const WriteToBase64FileArgsSchema = z.object({
  path: z.string(),
  data: z.union([z.string(), z.instanceof(Buffer)]), // Note: Zod cannot directly validate Buffer type from client JSON. Server-side check might be needed.
  inputEncoding: bufferEncodingSchema,
});

// McpResource definitions for fileContentTools
const fileContentMcpResources: McpResource[] = [
  {
    path: 'fs/insert_text_lines',
    handler: fileContentTools.insert_text_lines as any, // Cast to any due to complex arg types for generic handler
    schema: InsertTextLinesArgsSchema,
    description: 'Inserts an array of strings as new lines into a text file at a specified line number, respecting file encoding.',
  },
  {
    path: 'fs/delete_text_lines',
    handler: fileContentTools.delete_text_lines as any,
    schema: DeleteTextLinesArgsSchema,
    description: 'Deletes a specified range of lines from a text file, respecting file encoding.',
  },
  {
    path: 'fs/replace_text_lines',
    handler: fileContentTools.replace_text_lines as any,
    schema: ReplaceTextLinesArgsSchema,
    description: 'Replaces a range of lines in a text file with new lines, respecting file encoding.',
  },
  {
    path: 'fs/extract_text_from_range',
    handler: fileContentTools.extract_text_from_range as any,
    schema: ExtractTextFromRangeArgsSchema,
    description: 'Extracts a text segment from a file defined by character positions, respecting file encoding.',
  },
  {
    path: 'fs/search_replace_in_text_range',
    handler: fileContentTools.search_replace_in_text_range as any,
    schema: SearchReplaceInTextRangeArgsSchema,
    description: 'Searches (string/regex) and replaces text within a line range, respecting file encoding.',
  },
  {
    path: 'fs/sort_text_lines',
    handler: fileContentTools.sort_text_lines as any,
    schema: SortTextLinesArgsSchema,
    description: 'Sorts the lines of a text file.',
  },
  {
    path: 'fs/deduplicate_consecutive_lines',
    handler: fileContentTools.deduplicate_consecutive_lines as any,
    schema: DeduplicateConsecutiveLinesArgsSchema,
    description: 'Removes consecutive duplicate lines from a text file.',
  },
  {
    path: 'fs/trim_line_whitespace',
    handler: fileContentTools.trim_line_whitespace as any,
    schema: TrimLineWhitespaceArgsSchema,
    description: 'Trims leading and/or trailing whitespace from each line in a text file.',
  },
  {
    path: 'fs/read_binary_as_hex',
    handler: fileContentTools.read_binary_as_hex as any,
    schema: ReadBinaryAsHexArgsSchema,
    description: 'Reads a binary file and returns its content as a hexadecimal string.',
  },
  {
    path: 'fs/write_hex_as_binary',
    handler: fileContentTools.write_hex_as_binary as any,
    schema: WriteHexAsBinaryArgsSchema,
    description: 'Writes a hexadecimal string to a file as binary data.',
  },
  {
    path: 'fs/read_base64_file',
    handler: fileContentTools.read_base64_file as any,
    schema: ReadBase64FileArgsSchema,
    description: 'Reads a Base64 encoded file and returns its decoded content.',
  },
  {
    path: 'fs/write_to_base64_file',
    handler: fileContentTools.write_to_base64_file as any,
    schema: WriteToBase64FileArgsSchema,
    description: 'Encodes string data or binary data to Base64 and writes it to a file.',
   },
  ];
  
  // McpResource definitions for directoryOperationsTools
  const directoryOperationsMcpResources: McpResource[] = Object.values(directoryOperationsTools).map(tool => ({
    path: `fs/${tool.name}`,
    handler: tool.execute as any, // Cast to any due to generic execute signature
    schema: tool.schema as z.ZodSchema<any>,
    description: tool.description,
  }));
  
  // McpResource definitions for structuredDataTools
  const structuredDataMcpResources: McpResource[] = [
    {
      path: 'fs/read_csv_data',
      handler: structuredDataTools.read_csv_data as any,
      schema: structuredDataSchemas.readCsvDataInputSchema,
      description: 'Reads data from a CSV file, allowing selection of specific columns and row ranges.',
    },
    {
      path: 'fs/write_csv_data',
      handler: structuredDataTools.write_csv_data as any,
      schema: structuredDataSchemas.writeCsvDataInputSchema,
      description: 'Writes an array of objects or arrays of data to a CSV file.',
    },
    {
      path: 'fs/read_json_path',
      handler: structuredDataTools.read_json_path as any,
      schema: structuredDataSchemas.readJsonPathInputSchema,
      description: 'Reads a specific value from a JSON file using a JSONPath-like expression or a simple dot-notation path.',
    },
    {
      path: 'fs/write_json_path',
      handler: structuredDataTools.write_json_path as any,
      schema: structuredDataSchemas.writeJsonPathInputSchema,
      description: 'Writes or updates a value at a specific path within a JSON file.',
    },
    {
      path: 'fs/extract_from_markup',
      handler: structuredDataTools.extract_from_markup as any,
      schema: structuredDataSchemas.extractFromMarkupInputSchema,
      description: 'Extracts data (text, attribute, or HTML) from an HTML or XML file using a CSS selector.',
    },
    {
      path: 'fs/update_markup_content',
      handler: structuredDataTools.update_markup_content as any,
      schema: structuredDataSchemas.updateMarkupContentInputSchema,
      description: 'Updates the content (text or HTML) or an attribute of elements matching a CSS selector in an HTML/XML file.',
    },
  ];
  
  /**
   * A flattened array containing all implemented tool definitions conforming to the McpResource interface.
 * This list is used by the server to register the tools.
 */
// Combine all tools, spreading arrays and adding single objects
const allImplementedToolsIntermediate = [
    ...fsDirectoryTool, // Spread array
    ...fsFileTool,      // Spread array
    ...fileContentMcpResources,
    ...directoryOperationsMcpResources,
    ...structuredDataMcpResources, // Add new structured data tools
    // Archive tools are registered via registerArchiveTools function call in server.ts
    ...wordStylesTool,  // Spread array
    ...wordMarkdownTool,// Spread array
    ...wordTemplateTool,// Spread array
    ...wordTextTool,    // Spread array
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
    ...excelWorksheetsTool, // Spread array
    excelRangeTool,
    excelTablesTool,
    excelChartsTool,
    ...excelDataAnalysisTool, // Spread array
    slidesTool,
    powerpointShapesTool,
    ...powerpointPropertiesTool, // Spread array
    animationsTool,
    officeTransferTool,
    new OfficeWorkflowTool(), // Assuming this instance conforms to McpResource
    pdfExportTool,
    pdfParseTool,
    officeCombineTool, // Assuming single object
    dynamicResourcesTool, // Assuming single object
    officeImageAnalysisGenerationTool, // Add the new tool here
    generalParseTextTool, // Add the new general parse text tool
     // Add other imported tools here
 ];

 // Flatten the array manually for compatibility
 const allImplementedTools: McpResource[] = Array.prototype.concat.apply([], allImplementedToolsIntermediate);


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

// Export the list used by server.ts and the resources/tool registration functions
export { allRegisteredTools, aiAssistantGuideResource, registerWordCodeFormatTool, registerArchiveTools };