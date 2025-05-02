// =============================================================================
/**
 * @file Provides security-related utility functions.
 */
import os from 'os'; // Importado os
import path from 'path';
import fs from 'fs-extra'; // Use fs-extra for path existence checks
import { createErrorResponse } from './errorHandler';

// Define allowed base paths for file system operations
// IMPORTANT: Configure this securely based on your deployment environment!
const rawAllowedPaths = [ // Renombrado y añadido nueva ruta
    process.env.ALLOWED_FS_PATH_1 || './data/user_files', // Example user data path
    process.env.ALLOWED_FS_PATH_2 || './data/templates', // Example template path
    'HOME/Downloads'
];

// Mapeo para resolver HOME/ y luego path.resolve
const ALLOWED_BASE_PATHS = rawAllowedPaths.map(p => {
    let resolvedPath = p;
    if (p.startsWith('HOME/')) {
        const homeDir = os.homedir();
        if (!homeDir) {
            // Manejar caso donde homedir no se puede determinar (poco probable pero seguro)
            console.warn(`Could not determine home directory for path: ${p}. Skipping.`);
            // Podrías lanzar un error o devolver un valor que indique fallo
            // Por ahora, devolvemos la ruta original sin resolver para evitar errores inesperados
            // pero esto significa que la ruta HOME/ no funcionará.
            return p; // O podrías retornar null y filtrar después, o lanzar error.
        }
        resolvedPath = path.join(homeDir, p.substring(5));
    }
    // path.resolve maneja rutas absolutas (como las de HOME/ resueltas)
    // y rutas relativas al directorio de trabajo actual (como las de ./data/)
    return path.resolve(resolvedPath);
}).filter(p => p !== null); // Filtrar posibles nulos si decides usar esa estrategia arriba

/**
 * Validates if a given file path is within the allowed base directories.
 * Prevents directory traversal attacks.
 * @param filePath - The absolute or relative file path to validate.
 * @param basePathOverride - Optional override for the allowed base paths (for specific tools).
 * @returns The resolved, absolute path if valid.
 * @throws An error if the path is invalid or outside allowed directories.
 */
export function validateFilePath(filePath: string, basePathOverride?: string[]): string {
    // Primero, resuelve la ruta de entrada para obtener una ruta absoluta canónica
    const resolvedInputPath = path.resolve(filePath);

    // Usa las rutas base permitidas (ya resueltas a absolutas) o un override
    const allowedPaths = basePathOverride
        ? basePathOverride.map(p => path.resolve(p)) // Asegura que los overrides también sean absolutos
        : ALLOWED_BASE_PATHS;

    const isAllowed = allowedPaths.some(basePath => {
        // Comprueba si la ruta resuelta comienza con la ruta base permitida.
        // Esto es más robusto que path.relative para verificar la contención.
        // Asegúrate de que basePath termine con un separador para evitar coincidencias parciales
        // ej: /allowed/path vs /allowed/path-other
        const basePathWithSeparator = basePath.endsWith(path.sep) ? basePath : basePath + path.sep;
        return resolvedInputPath.startsWith(basePathWithSeparator) || resolvedInputPath === basePath;
    });

    if (!isAllowed) {
        // Log detallado para depuración (opcional, considera la sensibilidad de la información)
        // console.error(`Access Denied: Path '${resolvedInputPath}' not within allowed bases: ${allowedPaths.join(', ')}`);
        throw new Error(`Access denied: Path '${filePath}' is outside allowed directories.`);
    }

    // La comprobación de existencia es opcional y depende del caso de uso.
    // Si la herramienta necesita que el archivo exista, descomenta esto.
    // if (!fs.pathExistsSync(resolvedInputPath)) {
    //     throw new Error(`File or directory not found: ${resolvedInputPath}`);
    // }

    return resolvedInputPath; // Devuelve la ruta absoluta resuelta
}

// Add other security functions as needed (e.g., permission checks based on ToolContext)