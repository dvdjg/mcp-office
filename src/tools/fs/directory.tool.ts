/**
 * @file Implements the 'fs/directory' tool for directory operations (list, create, delete).
 * Provides functionality to interact with the file system directories within allowed paths.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import fs from 'fs-extra';
import path from 'path';
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '@/types/common.types'; // Import FastMCPContext from common types
import { createErrorResponse, handleToolError } from '@/utils/errorHandler';
import { validateFilePath } from '@/utils/security';
import logger from '@/utils/logger';

// --- Schemas for Input Validation ---

/** Schema for the 'fs/directory/list' tool parameters. */
const listSchema = z.object({
    /** The path of the directory to list contents for (relative to the current workspace directory). */
    path: z.string().min(1),
    /** Optional glob pattern to filter files (e.g., '*.ts' for TypeScript files). If not provided, it will list all files (*). */
    filter: z.string().optional(), // e.g., "*.docx", "image/*" (simple glob pattern)
    /** Whether to list files recursively. Use true for recursive listing, false or omit for top-level only. */
    recursive: z.boolean().optional().default(false),
});

/** Schema for the 'fs/directory/create' tool parameters. */
const createSchema = z.object({
    /** The path of the directory to create (relative to the current workspace directory). */
    path: z.string().min(1),
});

/** Schema for the 'fs/directory/delete' tool parameters. */
const deleteSchema = z.object({
    /** The path of the directory to delete (relative to the current workspace directory). */
    path: z.string().min(1),
    /** Whether to delete the directory recursively. Required for non-empty directories. Defaults to false. */
    recursive: z.boolean().optional().default(false), // Safety: require explicit recursive delete
});

// --- Tool Handlers ---

/**
 * Lists files and subdirectories within a given path.
 * @param params - The parameters for the list operation, validated against `listSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse containing a list of directory entries (strings).
 */
async function listDirectory(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<string[]>> { // Changed generic type to undefined
    try {
        const validatedParams = listSchema.parse(params);
        // Validate the base path first (directory itself must be allowed)
        const safeBasePath = validateFilePath(validatedParams.path);

        if (!await fs.pathExists(safeBasePath) || !(await fs.stat(safeBasePath)).isDirectory()) {
            return createErrorResponse('NOT_FOUND', `Directory not found: ${validatedParams.path}`);
        }

        const entries = await fs.readdir(safeBasePath);
        let results: string[] = [];

        for (const entry of entries) {
            const fullPath = path.join(safeBasePath, entry);
            const stat = await fs.stat(fullPath);

            if (stat.isDirectory() && validatedParams.recursive) {
                // Recursive listing is not fully implemented in this example, only top level is listed.
                logger.warn(`[fs/directory/list] Recursive listing for ${fullPath} not fully implemented.`);
                results.push(`${entry}/`); // Indicate directory
            } else if (!stat.isDirectory()) {
                // Apply filter if provided (simple wildcard matching)
                if (!validatedParams.filter || entry.match(new RegExp(validatedParams.filter.replace('*','.*')))) {
                   results.push(entry);
                }
            } else {
                 results.push(`${entry}/`); // Indicate directory
            }
        }

        logger.info(`[fs/directory/list] Listed directory: ${safeBasePath}`, { count: results.length, params: validatedParams });
        return { success: true, data: results };
    } catch (error) {
        return handleToolError(error, 'FS_LIST_ERROR');
    }
}

/**
 * Creates a new directory at the specified path.
 * @param params - The parameters for the create operation, validated against `createSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse containing the path of the created directory.
 */
async function createDirectory(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{ path: string }>> { // Changed generic type to undefined
    try {
        const validatedParams = createSchema.parse(params);
        // Validate the *parent* directory of the path to be created
        const parentDir = path.dirname(validatedParams.path);
        validateFilePath(parentDir); // Ensure parent is allowed

        const dirToCreate = path.resolve(validatedParams.path); // Use resolved path

        if (await fs.pathExists(dirToCreate)) {
            return createErrorResponse('ALREADY_EXISTS', `Directory already exists: ${validatedParams.path}`);
        }

        await fs.mkdir(dirToCreate, { recursive: true }); // Use recursive true for intermediate dirs
        logger.info(`[fs/directory/create] Created directory: ${dirToCreate}`);
        return { success: true, data: { path: dirToCreate } };
    } catch (error) {
        return handleToolError(error, 'FS_CREATE_ERROR');
    }
}

/**
 * Deletes a directory at the specified path.
 * Requires the `recursive` flag to be true for non-empty directories.
 * @param params - The parameters for the delete operation, validated against `deleteSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an empty ApiResponse indicating success.
 */
async function deleteDirectory(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> { // Changed generic type to undefined
     try {
        const validatedParams = deleteSchema.parse(params);
        const safePath = validateFilePath(validatedParams.path); // Validate the path to delete

        if (!await fs.pathExists(safePath) || !(await fs.stat(safePath)).isDirectory()) {
            return createErrorResponse('NOT_FOUND', `Directory not found: ${validatedParams.path}`);
        }

        // Safety check: Require explicit recursive flag for non-empty directories
        const entries = await fs.readdir(safePath);
        if (entries.length > 0 && !validatedParams.recursive) {
             return createErrorResponse('DIRECTORY_NOT_EMPTY', `Directory is not empty. Use recursive=true to delete: ${validatedParams.path}`);
        }

        await fs.remove(safePath); // fs-extra remove handles recursive deletion
        logger.info(`[fs/directory/delete] Deleted directory: ${safePath}`, { recursive: validatedParams.recursive });
        return { success: true, data: {} };
     } catch (error) {
        return handleToolError(error, 'FS_DELETE_ERROR');
     }
}

// --- Resource Definition ---

/**
 * Array of McpResource definitions for file system directory operations.
 */
export const fsDirectoryTool: McpResource[] = [
    {
        path: 'fs/directory/list',
        handler: listDirectory,
        schema: listSchema,
        description: 'Lists files and subdirectories in a specified path.',
        completions: async () => ({ path: ['/documents/', '/templates/'] }), // Example completions
    },
    {
        path: 'fs/directory/create',
        handler: createDirectory,
        schema: createSchema,
        description: 'Creates a new directory at the specified path.',
    },
    {
        path: 'fs/directory/delete',
        handler: deleteDirectory,
        schema: deleteSchema,
        description: 'Deletes a directory. Requires recursive=true for non-empty directories.',
    }
];