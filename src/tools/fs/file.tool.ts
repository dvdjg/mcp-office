/**
 * @file Implements the 'fs/file' tool for file operations.
 */
import fs from 'fs-extra';
import path from 'path';
import { z } from 'zod';
import { McpResource, ApiResponse, ToolContext, ToolRequestParams } from '@/types/common.types';
import { createErrorResponse, handleToolError } from '@/utils/errorHandler';
import { validateFilePath } from '@/utils/security';
import logger from '@/utils/logger';

// --- Schemas ---
const fileOpSchema = z.object({
    path: z.string().min(1),
});

const readSchema = fileOpSchema.extend({
    format: z.enum(['text', 'binary']).default('text'),
    encoding: z.string().optional().default('utf8'), // Used only for text format
});

const writeSchema = fileOpSchema.extend({
    content: z.union([z.string(), z.instanceof(Buffer)]), // Accept string or Buffer
    encoding: z.string().optional().default('utf8'), // Used only if content is string
    append: z.boolean().optional().default(false),
});

const renameSchema = z.object({
    oldPath: z.string().min(1),
    newPath: z.string().min(1),
});


// --- Handlers ---
async function readFile(params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<string | Buffer>> {
    try {
        const validatedParams = readSchema.parse(params);
        const safePath = validateFilePath(validatedParams.path);

        if (!await fs.pathExists(safePath) || !(await fs.stat(safePath)).isFile()) {
            return createErrorResponse('NOT_FOUND', `File not found: ${validatedParams.path}`);
        }

        let data: string | Buffer;
        if (validatedParams.format === 'binary') {
            data = await fs.readFile(safePath);
        } else {
            data = await fs.readFile(safePath, { encoding: validatedParams.encoding as BufferEncoding });
        }

        logger.info(`Read file: ${safePath}`, { format: validatedParams.format });
        return { success: true, data: data };
    } catch (error) {
        return handleToolError(error, 'FS_READ_ERROR');
    }
}

async function writeFile(params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<{ path: string }>> {
    try {
        const validatedParams = writeSchema.parse(params);
        // Validate the *directory* where the file will be written
        const parentDir = path.dirname(validatedParams.path);
        validateFilePath(parentDir);

        const fileToWrite = path.resolve(validatedParams.path);

        const writeOptions: fs.WriteFileOptions = {};
        if (typeof validatedParams.content === 'string') {
            writeOptions.encoding = validatedParams.encoding as BufferEncoding;
        }
        if (validatedParams.append) {
            writeOptions.flag = 'a'; // Append mode
        }

        await fs.writeFile(fileToWrite, validatedParams.content, writeOptions);

        logger.info(`Wrote file: ${fileToWrite}`, { append: validatedParams.append, encoding: validatedParams.encoding });
        return { success: true, data: { path: fileToWrite } };
    } catch (error) {
        return handleToolError(error, 'FS_WRITE_ERROR');
    }
}

async function deleteFile(params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<{}>> {
    try {
        const validatedParams = fileOpSchema.parse(params);
        const safePath = validateFilePath(validatedParams.path);

        if (!await fs.pathExists(safePath) || !(await fs.stat(safePath)).isFile()) {
            return createErrorResponse('NOT_FOUND', `File not found: ${validatedParams.path}`);
        }

        await fs.remove(safePath); // fs-extra remove works for files too
        logger.info(`Deleted file: ${safePath}`);
        return { success: true, data: {} };
    } catch (error) {
        return handleToolError(error, 'FS_DELETE_ERROR');
    }
}

async function renameFile(params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<{ newPath: string }>> {
     try {
        const validatedParams = renameSchema.parse(params);
        const safeOldPath = validateFilePath(validatedParams.oldPath);
        // Validate the *directory* of the new path
        const newParentDir = path.dirname(validatedParams.newPath);
        validateFilePath(newParentDir);
        const safeNewPath = path.resolve(validatedParams.newPath);


        if (!await fs.pathExists(safeOldPath) || !(await fs.stat(safeOldPath)).isFile()) {
            return createErrorResponse('NOT_FOUND', `Source file not found: ${validatedParams.oldPath}`);
        }
        if (await fs.pathExists(safeNewPath)) {
             return createErrorResponse('ALREADY_EXISTS', `Target file already exists: ${validatedParams.newPath}`);
        }

        await fs.rename(safeOldPath, safeNewPath);
        logger.info(`Renamed file: ${safeOldPath} -> ${safeNewPath}`);
        return { success: true, data: { newPath: safeNewPath } };
    } catch (error) {
        return handleToolError(error, 'FS_RENAME_ERROR');
    }
}

// --- Resource Definition ---
export const fsFileTool: McpResource[] = [
    {
        path: 'fs/file/read',
        handler: readFile,
        schema: readSchema,
        description: 'Reads content from a file as text or binary.',
        completions: async () => ({ format: ['text', 'binary'], encoding: ['utf8', 'base64'] }),
    },
    {
        path: 'fs/file/write',
        handler: writeFile,
        schema: writeSchema,
        description: 'Writes text or binary content to a file, optionally appending.',
    },
     {
        path: 'fs/file/delete',
        handler: deleteFile,
        schema: fileOpSchema,
        description: 'Deletes a file.',
    },
     {
        path: 'fs/file/rename',
        handler: renameFile,
        schema: renameSchema,
        description: 'Renames or moves a file.',
    },
];