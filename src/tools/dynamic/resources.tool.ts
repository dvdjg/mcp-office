import { z } from 'zod';
import * as fs from 'fs-extra';
import path from 'path';
import NodeCache from 'node-cache';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '@/types/common.types';

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
const DynamicResourceInputSchema = z.discriminatedUnion('operation', [
  z.object({
    operation: z.literal('list'),
    type: z.string().optional(), // Filtrar por tipo de recurso (ej: 'document', 'image')
    query: z.string().optional(), // Consulta de búsqueda (para la operación 'search')
  }),
  z.object({
    operation: z.literal('read'),
    resourceId: z.string(), // ID único del recurso
    filePath: z.string().optional(), // Ruta relativa dentro de dynamic_storage si resourceId no es suficiente
  }),
  z.object({
    operation: z.literal('write'),
    resourceId: z.string().optional(), // ID único del recurso (opcional para nuevos recursos)
    filePath: z.string(), // Ruta relativa dentro de dynamic_storage
    content: z.string(), // Contenido del recurso
    type: z.string().optional(), // Tipo de recurso
  }),
  z.object({
    operation: z.literal('delete'),
    resourceId: z.string().optional(), // ID único del recurso
    filePath: z.string(), // Ruta relativa dentro de dynamic_storage
  }),
  z.object({
    operation: z.literal('metadata'),
    resourceId: z.string().optional(), // ID único del recurso
    filePath: z.string(), // Ruta relativa dentro de dynamic_storage
  }),
  z.object({
    operation: z.literal('search'),
    query: z.string(), // Consulta de búsqueda
    type: z.string().optional(), // Filtrar por tipo de recurso
  }),
]);

type DynamicResourceInput = z.infer<typeof DynamicResourceInputSchema>;

// Función para obtener la ruta completa de un recurso
const getResourcePath = (filePath: string): string => {
  const fullPath = path.join(DYNAMIC_STORAGE_DIR, filePath);
  // Asegurarse de que la ruta esté dentro del directorio de almacenamiento dinámico
  if (!fullPath.startsWith(DYNAMIC_STORAGE_DIR)) {
    throw new Error('Acceso denegado: la ruta del recurso está fuera del directorio de almacenamiento permitido.');
  }
  return fullPath;
};

// Implementación del manejador del recurso
const handler = async (input: DynamicResourceInput, context: any): Promise<ApiResponse<any>> => {
  try {
    // Asegurarse de que el directorio de almacenamiento exista
    await fs.ensureDir(DYNAMIC_STORAGE_DIR);

    switch (input.operation) {
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
        return { success: true, data: resources };
      }

      case 'read': {
        if (!input.filePath) {
           throw new Error('filePath es requerido para la operación read.');
        }
        const fullPath = getResourcePath(input.filePath);
        if (!await fs.pathExists(fullPath)) {
          throw new Error(`Recurso no encontrado: ${input.filePath}`);
        }
        const content = await fs.readFile(fullPath, 'utf-8');
        return { success: true, data: { content } };
      }

      case 'write': {
        if (!input.filePath || input.content === undefined) {
           throw new Error('filePath y content son requeridos para la operación write.');
        }
        const fullPath = getResourcePath(input.filePath);
        await fs.ensureDir(path.dirname(fullPath)); // Asegurar que el directorio padre exista
        await fs.writeFile(fullPath, input.content, 'utf-8');
        // Invalidar caché para este recurso si existía
        metadataCache.del(input.filePath);
        return { success: true, data: { filePath: input.filePath } };
      }

      case 'delete': {
        if (!input.filePath) {
           throw new Error('filePath es requerido para la operación delete.');
        }
        const fullPath = getResourcePath(input.filePath);
         if (!await fs.pathExists(fullPath)) {
          throw new Error(`Recurso no encontrado: ${input.filePath}`);
        }
        await fs.remove(fullPath);
        // Invalidar caché para este recurso
        metadataCache.del(input.filePath);
        return { success: true, data: { filePath: input.filePath } };
      }

      case 'metadata': {
        if (!input.filePath) {
           throw new Error('filePath es requerido para la operación metadata.');
        }
        const fullPath = getResourcePath(input.filePath);
        const cachedMetadata = metadataCache.get(input.filePath);
        if (cachedMetadata) {
          return { success: true, data: cachedMetadata };
        }

        if (!await fs.pathExists(fullPath)) {
          throw new Error(`Recurso no encontrado: ${input.filePath}`);
        }
        const stats = await fs.stat(fullPath);
        const metadata = {
          id: input.filePath, // Usar la ruta relativa como ID
          name: path.basename(input.filePath),
          type: stats.isDirectory() ? 'directory' : 'file',
          createdAt: stats.birthtime,
          updatedAt: stats.mtime,
          size: stats.size,
          filePath: input.filePath,
          // Se podrían añadir más metadatos aquí si es necesario (ej: hash, permisos, etc.)
        };
        metadataCache.set(input.filePath, metadata);
        return { success: true, data: metadata };
      }

      case 'search': {
        if (!input.query) {
           throw new Error('query es requerido para la operación search.');
        }
        // Implementación básica de búsqueda: buscar archivos que contengan la consulta en su nombre
        // Una implementación más avanzada podría indexar contenido o usar herramientas de búsqueda dedicadas
        const files = await fs.readdir(DYNAMIC_STORAGE_DIR);
        const matchingFiles = files.filter(file => file.includes(input.query!));

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
        const filteredResources = input.type ? resources.filter(res => res.type === input.type) : resources;

        return { success: true, data: filteredResources };
      }

      default:
        // Esto no debería ocurrir gracias a discriminatedUnion, pero es un fallback seguro
        throw new Error(`Operación no soportada: ${(input as any).operation}`);
    }
  } catch (error: any) {
    console.error(`Error en el recurso dynamic/resources: ${error.message}`);
    throw new Error(`Error en el recurso dynamic/resources: ${error.message}`);
  }
};

// Definición de la herramienta McpResource
export const dynamicResourcesTool: McpResource[] = [{
  path: 'dynamic/resources',
  description: 'Manages dynamic resources generated by the user (documents, images, etc.) stored in /dynamic_storage. Allows listing, reading, writing, deleting, getting metadata, and searching for resources.', // Translated description
  handler: async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> => {
    try {
      // Validar los parámetros de entrada usando el esquema existente
      const input = DynamicResourceInputSchema.parse(params) as DynamicResourceInput;

      // Asegurarse de que el directorio de almacenamiento exista
      await fs.ensureDir(DYNAMIC_STORAGE_DIR);

      const op = input.operation;
      // Llamar al manejador con los parámetros validados
      switch (op) {
        case 'list': {
          const files = await fs.readdir(DYNAMIC_STORAGE_DIR);
          const resources = await Promise.all(files.map(async (file) => {
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
          const filteredResources = input.type ? resources.filter(res => res.type === input.type) : resources;
          return { success: true, data: filteredResources };
        }

        case 'read': {
          if (!input.filePath) {
             return { success: false, error: { code: 'MISSING_PARAMETER', message: 'filePath es requerido para la operación read.' } };
          }
          const fullPath = getResourcePath(input.filePath);
          if (!await fs.pathExists(fullPath)) {
            return { success: false, error: { code: 'NOT_FOUND', message: `Recurso no encontrado: ${input.filePath}` } };
          }
          const content = await fs.readFile(fullPath, 'utf-8');
          return { success: true, data: { content } };
        }

        case 'write': {
          if (!input.filePath || input.content === undefined) {
             return { success: false, error: { code: 'MISSING_PARAMETER', message: 'filePath y content son requeridos para la operación write.' } };
          }
          const fullPath = getResourcePath(input.filePath);
          await fs.ensureDir(path.dirname(fullPath));
          await fs.writeFile(fullPath, input.content, 'utf-8');
          metadataCache.del(input.filePath);
          return { success: true, data: { filePath: input.filePath } };
        }

        case 'delete': {
          if (!input.filePath) {
             return { success: false, error: { code: 'MISSING_PARAMETER', message: 'filePath es requerido para la operación delete.' } };
          }
          const fullPath = getResourcePath(input.filePath);
           if (!await fs.pathExists(fullPath)) {
            return { success: false, error: { code: 'NOT_FOUND', message: `Recurso no encontrado: ${input.filePath}` } };
          }
          await fs.remove(fullPath);
          metadataCache.del(input.filePath);
          return { success: true, data: { filePath: input.filePath } };
        }

        case 'metadata': {
          if (!input.filePath) {
             return { success: false, error: { code: 'MISSING_PARAMETER', message: 'filePath es requerido para la operación metadata.' } };
          }
          const fullPath = getResourcePath(input.filePath);
          const cachedMetadata = metadataCache.get(input.filePath);
          if (cachedMetadata) {
            return { success: true, data: cachedMetadata };
          }

          if (!await fs.pathExists(fullPath)) {
            return { success: false, error: { code: 'NOT_FOUND', message: `Recurso no encontrado: ${input.filePath}` } };
          }
          const stats = await fs.stat(fullPath);
          const metadata = {
            id: input.filePath,
            name: path.basename(input.filePath),
            type: stats.isDirectory() ? 'directory' : 'file',
            createdAt: stats.birthtime,
            updatedAt: stats.mtime,
            size: stats.size,
            filePath: input.filePath,
          };
          metadataCache.set(input.filePath, metadata);
          return { success: true, data: metadata };
        }

        case 'search': {
          if (!input.query) {
             return { success: false, error: { code: 'MISSING_PARAMETER', message: 'query es requerido para la operación search.' } };
          }
          const files = await fs.readdir(DYNAMIC_STORAGE_DIR);
          const matchingFiles = files.filter(file => file.includes(input.query!));

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
          const filteredResources = input.type ? resources.filter(res => res.type === input.type) : resources;

          return { success: true, data: filteredResources };
        }

        default:
          return { success: false, error: { code: 'UNSUPPORTED_OPERATION', message: `Operación no soportada: ${op}` } };
      }
    } catch (error: any) {
      console.error(`Error en la herramienta dynamic/resources: ${error.message}`);
      return { success: false, error: { code: 'TOOL_EXECUTION_ERROR', message: `Error en la herramienta dynamic/resources: ${error.message}`, details: error.stack } };
    }
  },
  schema: DynamicResourceInputSchema,
}];