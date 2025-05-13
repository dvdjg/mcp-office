/**
 * @file Implements the 'fs/file' tool for file operations (read, write, delete, rename).
 * Provides functionality to interact with files within allowed paths.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import fs from 'fs-extra';
import path from 'path';
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '@/types/common.types.js'; // Import FastMCPContext, remove ToolContext
import { createErrorResponse, handleToolError } from '@/utils/errorHandler.js';
import { validateFilePath } from '@/utils/security.js';
import logger from '@/utils/logger.js';

// --- Schemas ---

/** Base schema for file operations requiring a path. */
const fileOpSchema = z.object({
    /** The path of the file (relative to the current workspace directory). */
    path: z.string().min(1),
});

/** Schema for the 'fs/file/read' tool parameters. */
const readSchema = fileOpSchema.extend({
    /** The format to read the file content as ('text' or 'binary'). Defaults to 'text'. */
    format: z.enum(['text', 'binary']).default('text'),
    /** The encoding to use when reading the file as text. Defaults to 'utf8'. Used only for text format. */
    encoding: z.string().optional().default('utf8'), // Used only for text format
});

/** Schema for the 'fs/file/write' tool parameters. */
const writeSchema = fileOpSchema.extend({
    /** The content to write to the file. Can be a string or a Buffer. */
    content: z.union([z.string(), z.instanceof(Buffer)]), // Accept string or Buffer
    /** The encoding to use when writing string content. Defaults to 'utf8'. Used only if content is string. */
    encoding: z.string().optional().default('utf8'), // Used only if content is string
    /** Whether to append the content to the file instead of overwriting. Defaults to false. */
    append: z.boolean().optional().default(false),
});

/** Schema for the 'fs/file/rename' tool parameters. */
const renameSchema = z.object({
    /** The current path of the file to rename or move (relative to the current workspace directory). */
    oldPath: z.string().min(1),
    /** The new path for the file (relative to the current workspace directory). */
    newPath: z.string().min(1),
});


// --- Handlers ---

/**
 * Reads content from a file as text or binary.
 * @param params - The parameters for the read operation, validated against `readSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse containing the file content as a string or Buffer.
 */
async function readFile(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<string | Buffer>> { // Use FastMCPContext<undefined>
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

        logger.info(`[fs/file/read] Read file: ${safePath}`, { format: validatedParams.format });
        return { success: true, data: data };
    } catch (error) {
        return handleToolError(error, 'FS_READ_ERROR');
    }
}

/**
 * Writes text or binary content to a file, optionally appending.
 * @param params - The parameters for the write operation, validated against `writeSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse containing the path of the written file.
 */
async function writeFile(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{ path: string }>> { // Use FastMCPContext<undefined>
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

        logger.info(`[fs/file/write] Wrote file: ${fileToWrite}`, { append: validatedParams.append, encoding: validatedParams.encoding });
        return { success: true, data: { path: fileToWrite } };
    } catch (error) {
        return handleToolError(error, 'FS_WRITE_ERROR');
    }
}

/**
 * Deletes a file at the specified path.
 * @param params - The parameters for the delete operation, validated against `fileOpSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse containing the path of the deleted file.
 */
async function deleteFile(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{ path: string }>> { // Use FastMCPContext<undefined>
    try {
        const validatedParams = fileOpSchema.parse(params);
        const safePath = validateFilePath(validatedParams.path);

        if (!await fs.pathExists(safePath)) { // Check if path exists (can be file or dir for fs.remove)
            return createErrorResponse('NOT_FOUND', `Path not found: ${validatedParams.path}`);
        }
        // fs.remove can handle both files and directories, so specific isFile check might be too restrictive if we want to extend this.
        // For now, the tool is fs/file/delete, so we expect a file.
        if (!(await fs.stat(safePath)).isFile()) {
            return createErrorResponse('INVALID_TYPE', `Path is not a file: ${validatedParams.path}`);
        }

        await fs.remove(safePath); // fs-extra remove works for files
        logger.info(`[fs/file/delete] Deleted file: ${safePath}`);
        return { success: true, data: { path: safePath } };
    } catch (error) {
        return handleToolError(error, 'FS_DELETE_ERROR');
    }
}

/**
 * Renames or moves a file or directory from one path to another.
 * @param params - The parameters for the rename operation, validated against `renameSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse containing the old and new paths.
 */
async function renameFile(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{ oldPath: string, newPath: string }>> { // Use FastMCPContext<undefined>
    const validatedParams = renameSchema.parse(params); // Parse params early for access in catch
    let safeOldPath: string = ''; // Declare here for wider scope
    let safeNewPath: string = ''; // Declare here for wider scope
    try {
        safeOldPath = validateFilePath(validatedParams.oldPath);
        // Validate the *directory* of the new path
        const newParentDir = path.dirname(validatedParams.newPath);
        validateFilePath(newParentDir); // Ensures the target directory is within allowed bounds
        safeNewPath = path.resolve(validatedParams.newPath); // Get absolute path for consistency

        if (!await fs.pathExists(safeOldPath)) {
            return createErrorResponse('NOT_FOUND', `Source path not found: ${validatedParams.oldPath}`);
        }
        // Note: fs.stat will throw if path does not exist, so pathExists check is important.
        // We don't need to check if it's a file or directory here, fs.rename will handle it or error appropriately.

        if (await fs.pathExists(safeNewPath)) {
             return createErrorResponse('ALREADY_EXISTS', `Target path already exists: ${validatedParams.newPath}`);
        }

        await fs.rename(safeOldPath, safeNewPath); // fs.rename works for files and directories
        logger.info(`[fs/file/rename] Renamed path: ${safeOldPath} -> ${safeNewPath}`);
        return { success: true, data: { oldPath: safeOldPath, newPath: safeNewPath } };
    } catch (error) {
        // Catch specific errors from fs.rename if needed, e.g., trying to move directory to different device (EXDEV)
        // For EXDEV, fs.move from fs-extra would be a solution, but it's a more complex operation (copy then delete).
        // For now, we rely on standard fs.rename behavior.
        if (error && typeof error === 'object' && 'code' in error && error.code === 'EXDEV') {
            // Now safeOldPath and safeNewPath are accessible here
            logger.error(`[fs/file/rename] Failed to rename across devices: ${safeOldPath || validatedParams.oldPath} -> ${safeNewPath || validatedParams.newPath}`, error);
            const errorMessage = (error as { message?: string }).message || 'Unknown EXDEV error';
            return createErrorResponse('FS_RENAME_ERROR', `Cannot move path across different devices/partitions: ${errorMessage}`);
        }
        return handleToolError(error, 'FS_RENAME_ERROR');
    }
}

// --- Resource Definition ---

/**
 * Array of McpResource definitions for file system file operations.
 */
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
        description: 'Deletes a file. Returns the path of the deleted file.',
    },
     {
        path: 'fs/file/rename',
        handler: renameFile,
        schema: renameSchema,
        description: 'Renames or moves a file or directory. Returns the old and new paths.',
    },
];