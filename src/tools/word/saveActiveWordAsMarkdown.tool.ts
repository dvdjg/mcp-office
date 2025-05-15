import { z } from 'zod';
import path from 'path';
import fs from 'fs-extra';
import getActiveOfficeDocumentsTool from '../os/getActiveOfficeDocuments.tool.js';
import { wordMarkdownExportTool, exportSchema as wordMarkdownExportInputSchema } from './markdown.tool.js';
// Define BaseTool, ToolContext, ToolResponse, SuccessResponse, ErrorResponse, createErrorResponse, createSuccessResponse locally
// as they are not found in common.types.js based on the error
interface SuccessResponse<T = unknown> {
  success: true;
  data: T;
  message?: string;
}

interface ErrorResponse {
  success: false;
  error: {
    code: string;
    message: string;
    details?: unknown;
  };
}

type ToolResponse<T = unknown> = SuccessResponse<T> | ErrorResponse;

interface ToolContext {
  log: { // Made non-optional
    info: (message: string, data?: any) => void; // Adjusted signature
    warn: (message: string, data?: any) => void; // Adjusted signature
    error: (message: string, data?: any) => void; // Adjusted signature
    debug: (message: string, data?: any) => void; // Adjusted signature
  };
  reportProgress: (progress: { message: string; value: number; total?: number; type?: 'info' | 'error' | 'warning' }) => Promise<void>; // Made non-optional
  session?: Map<string, any>;
  // Add other context properties if needed
}

interface BaseTool<T_Input, T_Output> {
  path: string;
  description: string;
  inputSchema: z.ZodSchema<T_Input>;
  outputSchema: z.ZodSchema<T_Output>;
  handler: (params: T_Input, context: ToolContext) => Promise<ToolResponse<T_Output>>;
}

function createSuccessResponse<T>(data: T, message?: string): SuccessResponse<T> {
  return { success: true, data, message };
}

function createErrorResponse(code: string, message: string, details?: unknown): ErrorResponse {
  return { success: false, error: { code, message, details } };
}
import logger from '../../utils/logger.js'; // Assuming a logger utility

// Define the input schema for this orchestration tool
// It includes options for specifying input/output and passthrough options for markdown export
export const saveActiveWordAsMarkdownInputSchema = z.object({
  inputFilePath: z.string().optional().describe('Optional: Full path to a specific active Word document to convert. If not provided, the tool will attempt to use the sole active Word document. If multiple are active, this field becomes mandatory.'),
  outputFilePath: z.string().optional()
    .describe('Optional: Full path where the Markdown file should be saved. If not provided, defaults to the same directory and basename as the input Word document, with a .md extension.')
    .refine(value => {
        if (!value) return true; // Optional, so valid if not provided
        // Basic check: disallow overly simple directory paths like '/' or '.'
        // More robust validation (like preventing path traversal) should be handled
        // by the underlying file system operations or specific validation utilities.
        const dir = path.dirname(value);
        return dir !== '.' && dir !== '/' && dir !== '\\' && dir.includes(path.sep);
      }, {
        message: "Invalid or potentially unsafe output file path or directory. Ensure it's a valid relative or absolute path to a file, not just a root or current directory.",
      }),
  // Passthrough options from wordMarkdownExportTool schema, excluding filePath and output
  imageDir: wordMarkdownExportInputSchema.shape.imageDir.optional(),
  imagePrefix: wordMarkdownExportInputSchema.shape.imagePrefix.optional(),
  tableFormat: wordMarkdownExportInputSchema.shape.tableFormat.optional(),
  zipOutput: wordMarkdownExportInputSchema.shape.zipOutput.optional(),
  zipFileName: wordMarkdownExportInputSchema.shape.zipFileName.optional(),
  comments: wordMarkdownExportInputSchema.shape.comments.optional(),
  useComInterop: wordMarkdownExportInputSchema.shape.useComInterop.optional(),
});

type SaveActiveWordAsMarkdownInput = z.infer<typeof saveActiveWordAsMarkdownInputSchema>;
type SaveActiveWordAsMarkdownOutput = {
  outputPath: string;
  messages?: string[];
};

interface ActiveOfficeDoc {
  filePath: string;
  applicationType: 'Word' | 'Excel' | 'PowerPoint' | string;
}

const saveActiveWordAsMarkdownTool: BaseTool<SaveActiveWordAsMarkdownInput, SaveActiveWordAsMarkdownOutput> = {
  path: 'word/saveActiveWordAsMarkdown',
  description: 'Saves an active Microsoft Word document as a Markdown file. Can optionally specify input and output paths.',
  inputSchema: saveActiveWordAsMarkdownInputSchema,
  outputSchema: z.object({
    outputPath: z.string().describe('The full path to the saved Markdown file.'),
    messages: z.array(z.string()).optional().describe('Optional messages, e.g., warnings or info.'),
  }),

  handler: async (params: SaveActiveWordAsMarkdownInput, context: ToolContext): Promise<ToolResponse<SaveActiveWordAsMarkdownOutput>> => {
    // Ensure toolLogger adheres to the expected signature.
    // The default 'logger' from 'utils/logger.js' should be compatible or adaptable.
    // If context.log is directly passed, it must match ToolContext.log.
    // For simplicity, we assume context.log will be provided and correctly typed by the calling environment,
    // or the default logger is compatible.
    const toolLogger = context.log; // Directly use context.log, assuming it's correctly provided.
    toolLogger.info(`Starting saveActiveWordAsMarkdownTool with params:`, { params }); // Pass params as data
    await context.reportProgress?.({ message: 'Starting to save active Word document as Markdown...', value: 0 });

    let targetWordDocPath: string | undefined = undefined;

    try {
      await context.reportProgress?.({ message: 'Fetching active Office documents...', value: 10 });
      const activeDocsResponse = await getActiveOfficeDocumentsTool.handler({}, context as any); // Cast context for now

      if (!activeDocsResponse.success) {
        toolLogger.error('Failed to get active documents.', { error: activeDocsResponse.error });
        return createErrorResponse(
          activeDocsResponse.error?.code || 'GET_ACTIVE_DOCS_FAILED',
          `Failed to retrieve active Office documents: ${activeDocsResponse.error?.message || 'Unknown error'}`,
          activeDocsResponse.error?.details
        );
      }

      const activeWordDocs: ActiveOfficeDoc[] = activeDocsResponse.data?.documents?.filter(
        (doc: ActiveOfficeDoc) => doc.applicationType === 'Word'
      ) || [];
      toolLogger.debug(`Found ${activeWordDocs.length} active Word documents.`);

      if (params.inputFilePath) {
        const specifiedDoc = activeWordDocs.find(
          (doc: ActiveOfficeDoc) => doc.filePath === params.inputFilePath
        );
        if (specifiedDoc) {
          targetWordDocPath = specifiedDoc.filePath;
          toolLogger.info(`Using specified active Word document: ${targetWordDocPath}`);
        } else {
          return createErrorResponse('SPECIFIED_DOC_NOT_ACTIVE', `Specified input file path '${params.inputFilePath}' is not an active Word document or is not a Word document.`);
        }
      } else {
        if (activeWordDocs.length === 0) {
          return createErrorResponse('NO_ACTIVE_WORD_DOC', 'No active Word document found.');
        }
        if (activeWordDocs.length === 1) {
          targetWordDocPath = activeWordDocs[0].filePath;
          toolLogger.info(`Using the only active Word document: ${targetWordDocPath}`);
        } else {
          return createErrorResponse('MULTIPLE_ACTIVE_WORD_DOCS', `Multiple Word documents are active. Please specify 'inputFilePath'. Active documents: ${activeWordDocs.map((d: ActiveOfficeDoc) => d.filePath).join(', ')}`);
        }
      }
    } catch (error: any) {
      toolLogger.error('Error while getting active documents:', { error });
      return createErrorResponse('GET_ACTIVE_DOCS_ERROR', `An unexpected error occurred while fetching active documents: ${error.message}`, { stack: error.stack });
    }

    if (!targetWordDocPath) {
      // This case should ideally be caught by earlier logic, but as a safeguard:
      return createErrorResponse('INTERNAL_ERROR', 'Target Word document path could not be determined.');
    }

    let determinedOutputMdPath: string;
    if (!params.outputFilePath) {
      const docDir = path.dirname(targetWordDocPath as string); // targetWordDocPath is string here
      const docBaseName = path.basename(targetWordDocPath as string, path.extname(targetWordDocPath as string));
      determinedOutputMdPath = path.join(docDir, `${docBaseName}.md`);
      toolLogger.info(`Output file path not provided. Defaulting to: ${determinedOutputMdPath}`);
    } else {
      determinedOutputMdPath = params.outputFilePath;
      toolLogger.info(`Using provided output file path: ${determinedOutputMdPath}`);
    }

    // Ensure output directory exists
    try {
      const outputDir = path.dirname(determinedOutputMdPath);
      if (!await fs.pathExists(outputDir)) {
        await fs.ensureDir(outputDir);
        toolLogger.info(`Created output directory: ${outputDir}`);
      }
    } catch (error: any) {
      toolLogger.error(`Failed to ensure output directory for ${determinedOutputMdPath}:`, error);
      return createErrorResponse('OUTPUT_DIR_CREATION_FAILED', `Failed to ensure output directory exists for '${determinedOutputMdPath}': ${error.message}`);
    }

    const exportParams = {
      filePath: targetWordDocPath,
      output: determinedOutputMdPath,
      imageDir: params.imageDir,
      imagePrefix: params.imagePrefix,
      tableFormat: params.tableFormat,
      zipOutput: params.zipOutput,
      zipFileName: params.zipFileName,
      comments: params.comments,
      useComInterop: params.useComInterop,
    };

    try {
      await context.reportProgress({ message: `Exporting '${path.basename(targetWordDocPath as string)}' to Markdown...`, value: 50 }); // Removed optional chaining
      toolLogger.info(`Calling wordMarkdownExportTool with params:`, { exportParams }); // Pass params as data
      const exportResult = await wordMarkdownExportTool.handler(exportParams, context as any); // Cast context for now

      if (exportResult.success) {
        toolLogger.info(`Successfully exported Word document to Markdown: ${exportResult.data?.outputPath}`);
        await context.reportProgress?.({ message: 'Markdown export successful!', value: 100 });
        return createSuccessResponse({ outputPath: exportResult.data.outputPath }, `Successfully saved Word document as Markdown: ${exportResult.data.outputPath}`);
      } else {
        toolLogger.error('wordMarkdownExportTool failed:', exportResult.error);
        return createErrorResponse(
          exportResult.error?.code || 'MARKDOWN_EXPORT_FAILED',
          `Failed to export Word document to Markdown. ${exportResult.error?.code || ''}: ${exportResult.error?.message || 'Unknown export error'}`,
          exportResult.error?.details
        );
      }
    } catch (error: any) {
      toolLogger.error('Unhandled error during markdown export:', error);
      await context.reportProgress?.({ message: 'Markdown export failed due to an unhandled error.', value: 100, type: 'error' });
      return createErrorResponse('MARKDOWN_EXPORT_UNHANDLED_ERROR', `An unexpected error occurred during Markdown export: ${error.message}`, { stack: error.stack });
    }
  },
};

export default saveActiveWordAsMarkdownTool;