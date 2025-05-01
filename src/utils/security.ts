// =============================================================================
/**
 * @file Provides security-related utility functions.
 */
import path from 'path';
import fs from 'fs-extra'; // Use fs-extra for path existence checks
import { createErrorResponse } from './errorHandler';

// Define allowed base paths for file system operations
// IMPORTANT: Configure this securely based on your deployment environment!
const ALLOWED_BASE_PATHS = [
    path.resolve(process.env.ALLOWED_FS_PATH_1 || './data/user_files'), // Example user data path
    path.resolve(process.env.ALLOWED_FS_PATH_2 || './data/templates'), // Example template path
];

/**
 * Validates if a given file path is within the allowed base directories.
 * Prevents directory traversal attacks.
 * @param filePath - The absolute or relative file path to validate.
 * @param basePathOverride - Optional override for the allowed base paths (for specific tools).
 * @returns The resolved, absolute path if valid.
 * @throws An error if the path is invalid or outside allowed directories.
 */
export function validateFilePath(filePath: string, basePathOverride?: string[]): string {
    const resolvedPath = path.resolve(filePath);
    const allowedPaths = basePathOverride || ALLOWED_BASE_PATHS;

    const isAllowed = allowedPaths.some(basePath => {
        const relative = path.relative(basePath, resolvedPath);
        // Check if the path starts within the base path and doesn't climb up ('..')
        return relative && !relative.startsWith('..') && !path.isAbsolute(relative);
    });

    if (!isAllowed) {
        throw new Error(`Access denied: Path '${filePath}' is outside allowed directories.`);
    }

    // Optional: Check if the path actually exists if required by the operation
    // if (!fs.pathExistsSync(resolvedPath)) {
    //     throw new Error(`File or directory not found: ${resolvedPath}`);
    // }

    return resolvedPath;
}

// Add other security functions as needed (e.g., permission checks based on ToolContext)