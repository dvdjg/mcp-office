import { z } from 'zod';
import * as fs from 'fs-extra';
import path from 'path';
import NodeCache from 'node-cache';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '@/types/common.types';
import logger from '@/utils/logger'; // Importar logger

// Directorio base para almacenar recursos dinámicos
const DYNAMIC_STORAGE_DIR = path.join(__dirname, '..', '..', '..', 'dynamic_storage');

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


// Función para obtener la ruta completa de un recurso
const getResourcePath = (filePath: string): string => {
  const fullPath = path.join(DYNAMIC_STORAGE_DIR, filePath);
  // Asegurarse de que la ruta esté dentro del directorio de almacenamiento dinámico
  // Esto es una validación de seguridad crucial
  if (!fullPath.startsWith(DYNAMIC_STORAGE_DIR)) {
    throw new Error('Acceso denegado: la ruta del recurso está fuera del directorio de almacenamiento permitido.');
  }
  return fullPath;
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
        const files = await fs.readdir(DYNAMIC_STORAGE_DIR);
        // Implementación básica de listado, se podría añadir filtrado por tipo o búsqueda aquí
        // Para una búsqueda completa, se usaría la operación 'search'
        const resources = await Promise.all(files.map(async (file) => {
          const filePath = path.join(DYNAMIC_STORAGE_DIR, file);
          const stats = await fs.stat(filePath);
          return {
            id: file, // Usar el nombre del archivo como ID simple
            name: file,
            type: stats.isDirectory() ? 'directory' : 'file', // Tipo básico
            createdAt: stats.birthtime,
            updatedAt: stats.mtime,
            size: stats.size,
            filePath: file, // Ruta relativa
          };
        }));
        const filteredResources = type ? resources.filter(res => res.type === type) : resources;
        return { success: true, data: filteredResources };
      }

      case 'read': {
        // Validar que se proporcionó filePath o resourceId (ya hecho en refine)
        const targetPath = filePath ? getResourcePath(filePath) : (resourceId ? getResourcePath(resourceId) : undefined);
        if (!targetPath) throw new Error("filePath or resourceId is required for 'read' operation.");

        if (!await fs.pathExists(targetPath)) {
          return { success: false, error: { code: 'NOT_FOUND', message: `Recurso no encontrado: ${filePath || resourceId}` } };
        }
        const content = await fs.readFile(targetPath, 'utf-8');
        return { success: true, data: { content } };
      }

      case 'write': {
        // Validar que filePath y content no sean undefined (ya hecho en refine)
        if (!filePath || content === undefined) throw new Error("filePath and content are required for 'write' operation.");

        const fullPath = getResourcePath(filePath);
        await fs.ensureDir(path.dirname(fullPath)); // Asegurar que el directorio padre exista
        await fs.writeFile(fullPath, content, 'utf-8');
        // Invalidar caché para este recurso si existía
        metadataCache.del(filePath);
        return { success: true, data: { filePath: filePath } };
      }

      case 'delete': {
        // Validar que se proporcionó filePath o resourceId (ya hecho en refine)
        const targetPath = filePath ? getResourcePath(filePath) : (resourceId ? getResourcePath(resourceId) : undefined);
        if (!targetPath) throw new Error("filePath or resourceId is required for 'delete' operation.");

         if (!await fs.pathExists(targetPath)) {
          return { success: false, error: { code: 'NOT_FOUND', message: `Recurso no encontrado: ${filePath || resourceId}` } };
        }
        await fs.remove(targetPath);
        // Invalidar caché para este recurso
        const cacheKey = filePath || resourceId;
        if (cacheKey) metadataCache.del(cacheKey); // Invalidar usando filePath o resourceId
        return { success: true, data: { filePath: filePath || resourceId } };
      }

      case 'metadata': {
        // Validar que se proporcionó filePath o resourceId (ya hecho en refine)
        const targetPath = filePath ? getResourcePath(filePath) : (resourceId ? getResourcePath(resourceId) : undefined);
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
        const files = await fs.readdir(DYNAMIC_STORAGE_DIR);
        const matchingFiles = files.filter(file => file.includes(query!));

        const resources = await Promise.all(matchingFiles.map(async (file) => {
           const filePath = path.join(DYNAMIC_STORAGE_DIR, file);
           const stats = await fs.stat(filePath);
           return {
             id: file,
             name: file,
             type: stats.isDirectory() ? 'directory' : 'file',
             createdAt: stats.birthtime,
             updatedAt: stats.mtime,
             size: stats.size,
             filePath: file,
           };
        }));
        // Opcional: filtrar por tipo si se especifica
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

// Definición de la herramienta McpResource
export const dynamicResourcesTool: McpResource[] = [{
  path: 'dynamic/resources',
  description: 'Manages dynamic resources generated by the user (documents, images, etc.) stored in /dynamic_storage. Allows listing, reading, writing, deleting, getting metadata, and searching for resources.', // Translated description
  handler: handler,
  schema: DynamicResourcesInputSchema, // Usar el nuevo esquema z.object
}];