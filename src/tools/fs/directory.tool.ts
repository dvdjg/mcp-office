// =============================================================================
/**
 * @file Implements the 'fs/directory' tool for directory operations.
 */
import fs from 'fs-extra';
import path from 'path';
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '@/types/common.types'; // Import FastMCPContext from common types
import { createErrorResponse, handleToolError } from '@/utils/errorHandler';
import { validateFilePath } from '@/utils/security';
import logger from '@/utils/logger';
// Removed incorrect import: import { FastMCPContext } from 'fastmcp';

// --- Schemas for Input Validation ---
const listSchema = z.object({
    path: z.string().min(1),
    filter: z.string().optional(), // e.g., "*.docx", "image/*" (simple glob pattern)
    recursive: z.boolean().optional().default(false),
});

const createSchema = z.object({
    path: z.string().min(1),
});

const deleteSchema = z.object({
    path: z.string().min(1),
    recursive: z.boolean().optional().default(false), // Safety: require explicit recursive delete
});

// --- Tool Handlers ---

/**
 * Lists files and subdirectories within a given path.
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
                // Recursively list if requested (implement recursive logic carefully)
                // For simplicity, this example only lists the top level
                logger.warn(`Recursive listing for ${fullPath} not fully implemented in this example.`);
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

        logger.info(`Listed directory: ${safeBasePath}`, { count: results.length, params: validatedParams });
        return { success: true, data: results };
    } catch (error) {
        return handleToolError(error, 'FS_LIST_ERROR');
    }
}

/**
 * Creates a new directory.
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
        logger.info(`Created directory: ${dirToCreate}`);
        return { success: true, data: { path: dirToCreate } };
    } catch (error) {
        return handleToolError(error, 'FS_CREATE_ERROR');
    }
}

/**
 * Deletes a directory.
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
        logger.info(`Deleted directory: ${safePath}`, { recursive: validatedParams.recursive });
        return { success: true, data: {} };
    } catch (error) {
        return handleToolError(error, 'FS_DELETE_ERROR');
    }
}

// --- Resource Definition ---
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