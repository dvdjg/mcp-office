import { z } from 'zod';
import * as fs from 'fs-extra';
import path from 'path';
import NodeCache from 'node-cache';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '@/types/common.types';
import logger from '@/utils/logger'; // Importar logger

// Directorio base para almacenar recursos dinámicos
const DYNAMIC_STORAGE_DIR = path.join(__dirname, '..', '..', '..', 'dynamic_storage');

// Opciones de configuración para el sistema de recursos
const RESOURCE_SYSTEM_ENABLED = process.env.RESOURCE_SYSTEM_ENABLED !== 'false'; // Habilitado por defecto
const RESOURCE_LIMIT_PER_TYPE = process.env.RESOURCE_LIMIT_PER_TYPE ? parseInt(process.env.RESOURCE_LIMIT_PER_TYPE, 10) : undefined; // Límite por tipo de recurso
const RESOURCE_LIMIT_TOTAL = process.env.RESOURCE_LIMIT_TOTAL ? parseInt(process.env.RESOURCE_LIMIT_TOTAL, 10) : undefined; // Límite total

// Estructura de directorios virtual: <tool_name>/<YYYY-MM-DD>/<filename>
const getOrganizedResourcePath = (toolName: string, filename: string): string => {
  const today = new Date();
  const datePath = today.toISOString().split('T')[0]; // YYYY-MM-DD
  const resourceRelativePath = path.join(toolName, datePath, filename);
  return path.join(DYNAMIC_STORAGE_DIR, resourceRelativePath);
};

// Inicializar caché para metadatos (opcional, para mejorar rendimiento)
const metadataCache = new NodeCache({ stdTTL: 600, checkperiod: 120 }); // Cache por 10 minutos

/**
 * @tool dynamic/resources
 * @description Gestiona recursos dinámicos generados por el usuario (documentos, imágenes, etc.).
 * Permite listar, leer, escribir, eliminar, obtener metadatos y buscar recursos.
 * Los recursos se almacenan en el directorio `/dynamic_storage`.
 */

// Esquema de entrada para las operaciones del recurso
// Esquema combinado para todas las operaciones (usando z.object)
const DynamicResourcesInputSchema = z.object({
    operation: z.enum(['list', 'read', 'write', 'delete', 'metadata', 'search']).describe('The operation to perform (list, read, write, delete, metadata, or search).'),
    // Incluir todos los campos posibles de las operaciones
    resourceId: z.string().optional().describe("(read, delete, metadata) ID único del recurso."), // Hacer opcional aquí, validar en handler
    filePath: z.string().optional().describe("(read, write, delete, metadata) Ruta relativa dentro de dynamic_storage. Requerido para read, write, delete, metadata."), // Hacer opcional aquí, validar en handler
    content: z.string().optional().describe("(write) Contenido del recurso. Requerido para write."), // Hacer opcional aquí, validar en handler
    type: z.string().optional().describe("(list, search) Filtrar por tipo de recurso (ej: 'document', 'image')."),
    query: z.string().optional().describe("(list, search) Consulta de búsqueda (para la operación 'search'). Requerido para search."), // Hacer opcional aquí, validar en handler
}).refine(data => {
    // Validaciones condicionales basadas en la operación
    if (data.operation === 'read' || data.operation === 'delete' || data.operation === 'metadata') {
        return data.resourceId !== undefined || data.filePath !== undefined; // Requiere resourceId o filePath
    } else if (data.operation === 'write') {
        return data.filePath !== undefined && data.content !== undefined; // Requiere filePath y content
    } else if (data.operation === 'search') {
        return data.query !== undefined; // Requiere query
    }
    // 'list' no requiere campos adicionales
    return true; // Pasa la validación si la operación no requiere campos específicos o si los tiene
}, {
    message: "Missing required fields for the specified operation.",
    path: [], // Apply error to the whole object
});


// Inferir el tipo combinado para usar en el handler
type DynamicResourcesInput = z.infer<typeof DynamicResourcesInputSchema>;


// Función para obtener la ruta completa de un recurso a partir de su ruta relativa dentro de dynamic_storage
const getFullPathFromRelative = (relativePath: string): string => {
  const fullPath = path.join(DYNAMIC_STORAGE_DIR, relativePath);
  // Asegurarse de que la ruta esté dentro del directorio de almacenamiento dinámico
  // Esto es una validación de seguridad crucial
  if (!fullPath.startsWith(DYNAMIC_STORAGE_DIR)) {
    throw new Error('Acceso denegado: la ruta del recurso está fuera del directorio de almacenamiento permitido.');
  }
  return fullPath;
};

// Función para guardar un recurso con la estructura organizada
const saveResource = async (toolName: string, filename: string, content: string | Buffer): Promise<string> => {
  if (!RESOURCE_SYSTEM_ENABLED) {
    logger.info('[dynamic/resources] Resource system is disabled. Skipping save.');
    return ''; // Devolver cadena vacía o similar si está deshabilitado
  }

  const fullPath = getOrganizedResourcePath(toolName, filename);
  await fs.ensureDir(path.dirname(fullPath)); // Asegurar que el directorio padre exista
  await fs.writeFile(fullPath, content); // Usar writeFile que maneja strings y Buffers

  // Implementar lógica de límites aquí si es necesario
  // Esto podría ser complejo (listar todos los archivos, ordenar por fecha, eliminar los más antiguos)
  // Por ahora, solo guardamos. La lógica de límites se puede añadir después.

  // Devolver la ruta relativa dentro de dynamic_storage
  return path.relative(DYNAMIC_STORAGE_DIR, fullPath);
};

// Función para eliminar recursos antiguos según los límites configurados
const enforceResourceLimits = async () => {
  if (!RESOURCE_SYSTEM_ENABLED || (!RESOURCE_LIMIT_PER_TYPE && !RESOURCE_LIMIT_TOTAL)) {
    return; // No hacer nada si el sistema está deshabilitado o no hay límites configurados
  }

  logger.info('[dynamic/resources] Enforcing resource limits...');

  // Implementación de límites (simplificada por ahora)
  // Esto requeriría listar todos los archivos, agruparlos por tipo/herramienta,
  // ordenar por fecha de modificación y eliminar los excedentes.
  // Dada la complejidad, esta parte se deja como un TODO para una implementación futura más robusta.
  logger.warn('[dynamic/resources] Resource limit enforcement is not fully implemented yet.');
};

// Implementación del manejador del recurso
const handler = async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> => {
  const log = context?.log || logger; // Usar el logger del contexto si está disponible, sino el global

  let validatedRequest: DynamicResourcesInput;
  try {
    // Validar los parámetros de entrada usando el nuevo esquema
    validatedRequest = DynamicResourcesInputSchema.parse(params);
  } catch (error: any) {
     // Si la validación inicial falla, devolver un error de validación
     if (error instanceof z.ZodError) {
         // Serializar error.errors para que sea serializable
         const errorDetails = JSON.stringify(error.errors, null, 2);
         log.warn(`[dynamic/resources] Input validation failed: ${error.message}`, { errors: errorDetails, params });
         return {
             success: false,
             error: {
                 code: 'VALIDATION_ERROR',
                 message: 'Error de validación de entrada.',
                 details: errorDetails,
             },
         };
     }
     // Otro error inesperado durante el parseo inicial
     log.error(`[dynamic/resources] Unexpected error parsing params: ${error.message}`, { error: String(error), params });
     return {
         success: false,
         error: {
             code: 'INTERNAL_ERROR',
             message: 'Failed to parse tool parameters.',
             details: String(error),
         },
     };
  }

  const { operation, ...args } = validatedRequest; // Extraer operation y el resto como args
  const { resourceId, filePath, content, type, query } = args; // Extraer campos específicos

  try {
    // Asegurarse de que el directorio de almacenamiento exista
    await fs.ensureDir(DYNAMIC_STORAGE_DIR);

    switch (operation) {
      case 'list': {
        // Listar recursos de forma recursiva para reflejar la estructura organizada
        const listFilesRecursively = async (dir: string, relativeDir = ''): Promise<any[]> => {
          const entries = await fs.readdir(dir, { withFileTypes: true });
          let resources: any[] = [];

          for (const entry of entries) {
            const fullPath = path.join(dir, entry.name);
            const relativePath = path.join(relativeDir, entry.name);
            const stats = await fs.stat(fullPath);

            if (entry.isDirectory()) {
              // Si es un directorio, listar recursivamente
              resources.push({
                id: relativePath,
                name: entry.name,
                type: 'directory',
                filePath: relativePath,
                createdAt: stats.birthtime,
                updatedAt: stats.mtime,
                size: stats.size, // Tamaño del directorio puede no ser preciso
              });
              resources = resources.concat(await listFilesRecursively(fullPath, relativePath));
            } else {
              // Si es un archivo
              resources.push({
                id: relativePath,
                name: entry.name,
                type: 'file',
                filePath: relativePath,
                createdAt: stats.birthtime,
                updatedAt: stats.mtime,
                size: stats.size,
              });
            }
          }
          return resources;
        };

        const allResources = await listFilesRecursively(DYNAMIC_STORAGE_DIR);
        // Filtrar por tipo si se especifica
        const filteredResources = type ? allResources.filter(res => res.type === type) : allResources;

        return { success: true, data: filteredResources };
      }

      case 'read': {
        // Validar que se proporcionó filePath o resourceId (ya hecho en refine)
        const targetPath = filePath ? getFullPathFromRelative(filePath) : (resourceId ? getFullPathFromRelative(resourceId) : undefined);
        if (!targetPath) throw new Error("filePath or resourceId is required for 'read' operation.");

        if (!await fs.pathExists(targetPath)) {
          return { success: false, error: { code: 'NOT_FOUND', message: `Recurso no encontrado: ${filePath || resourceId}` } };
        }
        const content = await fs.readFile(targetPath, 'utf-8');
        return { success: true, data: { content } };
      }

      case 'write': {
        // Esta operación ahora se usará internamente por otras herramientas a través de saveResource
        // No debería ser llamada directamente por el cliente con esta lógica.
        // Podríamos mantenerla para flexibilidad, pero la lógica de organización y límites
        // debería pasar por saveResource.
        // Por ahora, la dejamos como estaba, pero con la validación de ruta completa.
         if (!filePath || content === undefined) throw new Error("filePath and content are required for 'write' operation.");
         const fullPath = getFullPathFromRelative(filePath);
         await fs.ensureDir(path.dirname(fullPath));
         await fs.writeFile(fullPath, content, 'utf-8');
         metadataCache.del(filePath);
         // Después de escribir, aplicar límites
         await enforceResourceLimits();
         return { success: true, data: { filePath: filePath } };
      }

      case 'delete': {
        // Validar que se proporcionó filePath o resourceId (ya hecho en refine)
        const targetPath = filePath ? getFullPathFromRelative(filePath) : (resourceId ? getFullPathFromRelative(resourceId) : undefined);
        if (!targetPath) throw new Error("filePath or resourceId is required for 'delete' operation.");

         if (!await fs.pathExists(targetPath)) {
          return { success: false, error: { code: 'NOT_FOUND', message: `Recurso no encontrado: ${filePath || resourceId}` } };
        }
        await fs.remove(targetPath);
        // Invalidar caché para este recurso
        const cacheKey = filePath || resourceId;
        if (cacheKey) metadataCache.del(cacheKey); // Invalidar usando filePath o resourceId
        // Después de eliminar, aplicar límites (esto podría ser redundante si la eliminación fue manual)
        await enforceResourceLimits();
        return { success: true, data: { filePath: filePath || resourceId } };
      }

      case 'metadata': {
        // Validar que se proporcionó filePath o resourceId (ya hecho en refine)
        const targetPath = filePath ? getFullPathFromRelative(filePath) : (resourceId ? getFullPathFromRelative(resourceId) : undefined);
        if (!targetPath) throw new Error("filePath or resourceId is required for 'metadata' operation.");

        const cacheKey = filePath || resourceId; // Usar filePath o resourceId como clave de caché
        if (cacheKey) { // Solo intentar obtener de caché si la clave está definida
            const cachedMetadata = metadataCache.get(cacheKey);
            if (cachedMetadata) {
              return { success: true, data: cachedMetadata };
            }
        }

        if (!await fs.pathExists(targetPath)) {
          return { success: false, error: { code: 'NOT_FOUND', message: `Recurso no encontrado: ${filePath || resourceId}` } };
        }
        const stats = await fs.stat(targetPath);
        const metadata = {
          id: filePath || resourceId, // Usar la ruta relativa o resourceId como ID
          name: path.basename(targetPath),
          type: stats.isDirectory() ? 'directory' : 'file',
          createdAt: stats.birthtime,
          updatedAt: stats.mtime,
          size: stats.size,
          filePath: filePath || resourceId, // Usar la ruta relativa o resourceId
          // Se podrían añadir más metadatos aquí si es necesario (ej: hash, permisos, etc.)
        };
        if (cacheKey) metadataCache.set(cacheKey, metadata); // Solo establecer en caché si la clave está definida
        return { success: true, data: metadata };
      }

      case 'search': {
        // Validar que query no sea undefined (ya hecho en refine)
        if (!query) throw new Error("query is required for 'search' operation.");

        // Implementación básica de búsqueda: buscar archivos que contengan la consulta en su nombre
        // Una implementación más avanzada podría indexar contenido o usar herramientas de búsqueda dedicadas
        // Para buscar en la estructura organizada, necesitamos listar recursivamente primero
        const listFilesRecursively = async (dir: string, relativeDir = ''): Promise<string[]> => {
            const entries = await fs.readdir(dir, { withFileTypes: true });
            let filePaths: string[] = [];

            for (const entry of entries) {
                const fullPath = path.join(dir, entry.name);
                const relativePath = path.join(relativeDir, entry.name);

                if (entry.isDirectory()) {
                    filePaths = filePaths.concat(await listFilesRecursively(fullPath, relativePath));
                } else {
                    filePaths.push(relativePath);
                }
            }
            return filePaths;
        };

        const allRelativePaths = await listFilesRecursively(DYNAMIC_STORAGE_DIR);
        const matchingRelativePaths = allRelativePaths.filter(relativePath => relativePath.includes(query!));

        const resources = await Promise.all(matchingRelativePaths.map(async (relativePath) => {
           const fullPath = getFullPathFromRelative(relativePath);
           const stats = await fs.stat(fullPath);
           return {
             id: relativePath,
             name: path.basename(relativePath),
             type: 'file', // Asumimos que la búsqueda es solo en archivos por ahora
             createdAt: stats.birthtime,
             updatedAt: stats.mtime,
             size: stats.size,
             filePath: relativePath,
           };
         }));
        // Opcional: filtrar por tipo si se especifica (aunque la búsqueda actual es solo en nombres de archivo)
        const filteredResources = type ? resources.filter(res => res.type === type) : resources;


        return { success: true, data: filteredResources };
      }

      default:
        // Esto no debería ocurrir gracias a discriminatedUnion, pero es un fallback seguro
        // const exhaustiveCheck: never = operation; // No longer needed with combined schema
        throw new Error(`Operación no soportada: ${operation}`); // Acceder a operation directamente
    }
  } catch (error: any) {
    const message = error instanceof Error ? error.message : String(error);
    log.error(`[dynamic/resources] Error en la herramienta: ${message}`, { error: String(error), params }); // Serializar error
    return { success: false, error: { code: 'TOOL_EXECUTION_ERROR', message: `Error en la herramienta dynamic/resources: ${message}`, details: String(error) } }; // Devolver ErrorResponse
  }
};

// Exportar la función saveResource para que otras herramientas puedan usarla
export { saveResource };

// Definición de la herramienta McpResource
export const dynamicResourcesTool: McpResource[] = [{
  path: 'dynamic/resources',
  description: 'Manages dynamic resources generated by the user (documents, images, etc.) stored in /dynamic_storage. Allows listing, reading, writing, deleting, getting metadata, and searching for resources.', // Translated description
  handler: handler,
  schema: DynamicResourcesInputSchema, // Usar el nuevo esquema z.object
}];