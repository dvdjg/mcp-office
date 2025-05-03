// =============================================================================
/**
 * @file Provides security-related utility functions.
 */
import os from 'os'; // Importado os
import path from 'path';
import fs from 'fs-extra'; // Use fs-extra for path existence checks
import { createErrorResponse } from './errorHandler';
import logger from './logger'; // Import logger

// Define allowed base paths for file system operations
// IMPORTANT: Configure this securely based on your deployment environment!
// Define allowed base paths for file system operations based on ALLOWED_FS_PATHS environment variable.
// IMPORTANT: Configure this securely based on your deployment environment!
const allowedFsPathsEnv = process.env.ALLOWED_FS_PATHS;

// Function to resolve '~' to the home directory
function resolveHome(filePath: string): string {
    if (filePath.startsWith('~')) {
        const homeDir = os.homedir();
        if (!homeDir) {
            logger.warn(`Could not determine home directory for path: ${filePath}.`);
            // Return original path if home directory cannot be determined
            return filePath;
        }
        return path.join(homeDir, filePath.substring(1));
    }
    return filePath;
}

// Determine the allowed base paths based on the environment variable
const ALLOWED_BASE_PATHS: string[] = [];
let isFsAccessAllowed = true;

if (allowedFsPathsEnv === 'none') {
    isFsAccessAllowed = false;
    logger.info('ALLOWED_FS_PATHS is set to "none". File system access is disabled.');
} else if (allowedFsPathsEnv) {
    // Split the environment variable by semicolon and resolve each path
    // Split the environment variable by semicolon or colon and resolve each path
    const rawAllowedPaths = allowedFsPathsEnv.split(/[:;]/);
    ALLOWED_BASE_PATHS.push(...rawAllowedPaths.map(p => path.resolve(resolveHome(p))));
    logger.info(`Allowed file system base paths: ${ALLOWED_BASE_PATHS.join(';')}`); // Still log with ; for consistency, but accept : or ; as input
} else {
    // If ALLOWED_FS_PATHS is not set, allow all paths.
    // In a real-world scenario, you might want to default to a restricted set.
    // For this implementation, we interpret undefined as allowing all paths.
    logger.warn('ALLOWED_FS_PATHS environment variable is not set. Allowing all file system paths.');
    // An empty ALLOWED_BASE_PATHS array will be handled in validateFilePath
}


/**
 * Validates if a given file path is within the allowed base directories,
 * or if file system access is globally allowed.
 * Prevents directory traversal attacks.
 * @param filePath - The absolute or relative file path to validate.
 * @returns The resolved, absolute path if valid.
 * @throws An error if the path is invalid or outside allowed directories when restrictions are in place.
 */
export function validateFilePath(filePath: string): string {
    // If file system access is globally disabled, throw an error immediately.
    if (!isFsAccessAllowed) {
        throw new Error(`File system access is disabled. Path '${filePath}' cannot be accessed.`);
    }

    // Resolve the input path to get a canonical absolute path
    const resolvedInputPath = path.resolve(resolveHome(filePath));

    // If ALLOWED_BASE_PATHS is empty, it means no restrictions are in place (ALLOWED_FS_PATHS was not set or empty string),
    // so any path is considered allowed.
    if (ALLOWED_BASE_PATHS.length === 0) {
        return resolvedInputPath; // Return the resolved path as it's allowed
    }

    // If ALLOWED_BASE_PATHS is not empty, validate against the list.
    const isAllowed = ALLOWED_BASE_PATHS.some(basePath => {
        // Check if the resolved input path starts with the allowed base path.
        // This is more robust than path.relative for checking containment.
        // Ensure basePath ends with a separator to avoid partial matches
        // e.g.: /allowed/path vs /allowed/path-other
        const basePathWithSeparator = basePath.endsWith(path.sep) ? basePath : basePath + path.sep;
        return resolvedInputPath.startsWith(basePathWithSeparator) || resolvedInputPath === basePath;
    });

    if (!isAllowed) {
        // Log detailed info for debugging (optional, consider information sensitivity)
        // console.error(`Access Denied: Path '${resolvedInputPath}' not within allowed bases: ${ALLOWED_BASE_PATHS.join(', ')}`);
        throw new Error(`Access denied: Path '${filePath}' is outside allowed directories.`);
    }

    // The existence check is optional and depends on the use case.
    // If the tool needs the file to exist, uncomment this.
    // if (!fs.pathExistsSync(resolvedInputPath)) {
    //     throw new Error(`File or directory not found: ${resolvedInputPath}`);
    // }

    return resolvedInputPath; // Return the resolved absolute path
}

// Export the flag to check if FS access is allowed
export { isFsAccessAllowed };