// Quitar importaciones de tipos específicos de FastMCP si no se usan explícitamente
// import { IFastMcpToolDefinition, IFastMcpToolHandler, IFastMcpToolSchema } from 'fastmcp';
import { z } from 'zod';
// import { WordApplication } from 'winax'; // No se importa directamente, se usa 'any' o se infiere
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { validateFilePath } from '../../utils/security'; // Corregido: Nombre de función
import * as fs from 'fs/promises';
import * as path from 'path';
import logger from '../../utils/logger'; // Corregido: Importación por defecto

// --- Esquema de Entrada (Zod) ---
const EmbeddedObjectBaseSchema = z.object({
  filePath: z.string().min(1, 'El path del archivo es requerido.'),
});

const InsertSchema = EmbeddedObjectBaseSchema.extend({
  operation: z.literal('insert'),
  objectPath: z.string().min(1, 'El path del objeto a insertar es requerido.'),
  // Opcional: position, linkToFile, displayAsIcon, iconFileName, iconLabel, etc.
  // range: z.string().optional().describe("Rango donde insertar (ej: 'paragraph:N', 'selection', 'end'). Default: 'end'"),
});

const ModifySchema = EmbeddedObjectBaseSchema.extend({
  operation: z.literal('modify'),
  objectIndex: z.number().int().positive('El índice del objeto debe ser un entero positivo.'),
  newObjectPath: z.string().min(1, 'El path del nuevo objeto es requerido.'),
  // Opcional: linkToFile, displayAsIcon, etc.
});

const DeleteSchema = EmbeddedObjectBaseSchema.extend({
  operation: z.literal('delete'),
  objectIndex: z.number().int().positive('El índice del objeto debe ser un entero positivo.'),
});

const ExtractAllSchema = EmbeddedObjectBaseSchema.extend({
  operation: z.literal('extractAll'),
  outputDirectory: z.string().min(1, 'El directorio de salida es requerido.'),
});

// Opcional: GetPropertiesSchema
// const GetPropertiesSchema = EmbeddedObjectBaseSchema.extend({
//   operation: z.literal('getProperties'),
//   objectIndex: z.number().int().positive('El índice del objeto debe ser un entero positivo.'),
// });

const EmbeddedObjectsInputSchema = z.discriminatedUnion('operation', [
  InsertSchema,
  ModifySchema,
  DeleteSchema,
  ExtractAllSchema,
  // GetPropertiesSchema, // Descomentar si se implementa
]);

type EmbeddedObjectsInput = z.infer<typeof EmbeddedObjectsInputSchema>;

// --- Esquema de Salida (Zod) ---
// No es estrictamente necesario definirlo aquí si no se valida explícitamente el retorno,
// pero puede ser útil para documentación o tipos internos.
const EmbeddedObjectsOutputSchema = z.object({
  success: z.boolean(),
  message: z.string(),
  details: z.any().optional(), // Para devolver información extra, como paths de archivos extraídos
});

// --- Handler (ahora llamado 'execute') ---
// El handler recibe input y el contexto (que incluye log, reportProgress, session)
const embeddedObjectsExecute = async (input: EmbeddedObjectsInput, context: { log: any, reportProgress: (progress: any) => void, session: any }) => {
  const log = context.log; // Usar el logger del contexto
  log.info(`Iniciando operación '${input.operation}' en archivo: ${input.filePath}`);
  // Usar validateFilePath que devuelve la ruta absoluta validada o lanza error
  const absoluteFilePath = validateFilePath(input.filePath);
  log.debug(`[EmbeddedObjects] Validated absolute path: ${absoluteFilePath}`);

  let wordApp: any = null; // Usar 'any' para el objeto COM de la aplicación
  let doc: any = null; // Usar 'any' para el objeto COM del documento

  try {
    wordApp = await getOfficeApplication('Word.Application');
    // Abrir el documento. Considerar ReadOnly para 'extractAll' y 'getProperties'
    const openReadOnly = input.operation === 'extractAll'; // || input.operation === 'getProperties';
    log.debug(`[EmbeddedObjects] Opening document: ${absoluteFilePath} (ReadOnly: ${openReadOnly})`);
      // Parámetros Open: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, ... Visible
      // Abrir no visible para operaciones de fondo
      doc = wordApp.Documents.Open(absoluteFilePath, false, openReadOnly, false, undefined, undefined, undefined, undefined, undefined, undefined, false); // Visible = false al final

    switch (input.operation) {
      case 'insert':
        // Lógica para insertar objeto OLE usando doc.InlineShapes.AddOLEObject(...) o similar
        log.warn('Operación "insert" aún no implementada.');
        // Ejemplo (requiere investigación API COM):
        // const { objectPath } = input; // Asegurarse que objectPath está definido en InsertSchema
        // const absoluteObjectPath = validateFilePath(objectPath); // Validar también el path del objeto
        // await fs.access(absoluteObjectPath); // Verificar existencia
        // const range = doc.Content; // O determinar el rango según input.range
        // range.Collapse(0); // wdCollapseEnd = 0 // Ir al final por defecto
        // doc.InlineShapes.AddOLEObject({ ClassType: undefined, FileName: absoluteObjectPath, LinkToFile: false, DisplayAsIcon: false }, range);
        return { success: false, message: 'Operación "insert" no implementada.' };

      case 'modify':
        // Lógica para modificar objeto OLE
        log.warn('Operación "modify" aún no implementada.');
        // Ejemplo (requiere investigación API COM):
        // const { objectIndex, newObjectPath } = input; // Asegurarse que están definidos en ModifySchema
        // const absoluteNewObjectPath = validateFilePath(newObjectPath); // Validar path
        // await fs.access(absoluteNewObjectPath);
        // if (objectIndex <= 0 || objectIndex > doc.InlineShapes.Count) { // Corregido: <= 0
        //   throw new Error(`Índice de objeto ${objectIndex} fuera de rango (1-${doc.InlineShapes.Count}).`);
        // }
        // const shape = doc.InlineShapes.Item(objectIndex);
        // if (shape.Type === 7 /* wdInlineShapeOLEObject */ || shape.OLEFormat) {
        //    // Lógica para reemplazar: ¿eliminar y añadir? ¿o hay método directo?
        //    // shape.OLEFormat.DoVerb(VerbIndex:=wdOLEVerbPrimary) // Activar?
        //    // shape.Delete();
        //    // doc.InlineShapes.AddOLEObject(...)
        // } else {
        //    throw new Error(`El objeto en el índice ${objectIndex} no es un objeto OLE.`);
        // }
        return { success: false, message: 'Operación "modify" no implementada.' };

      case 'delete':
        // Lógica para eliminar objeto OLE
        log.warn('Operación "delete" aún no implementada.');
        // Ejemplo (requiere investigación API COM):
        // const { objectIndex } = input; // Asegurarse que está definido en DeleteSchema
        // if (objectIndex <= 0 || objectIndex > doc.InlineShapes.Count) { // Corregido: <= 0
        //   throw new Error(`Índice de objeto ${objectIndex} fuera de rango (1-${doc.InlineShapes.Count}).`);
        // }
        // const shapeToDelete = doc.InlineShapes.Item(objectIndex); // Índice 1-based
        // // Verificar si es OLE antes de borrar? shapeToDelete.Type === 7 /* wdInlineShapeOLEObject */
        // shapeToDelete.Delete();
        // doc.Save();
        return { success: false, message: 'Operación "delete" no implementada.' };

      case 'extractAll':
        // Lógica para extraer todos los objetos OLE
        log.info('Iniciando operación "extractAll".');
        const { outputDirectory } = input; // Asegurarse que está definido en ExtractAllSchema
        // Validar el directorio de salida también
        const absoluteOutputDir = validateFilePath(outputDirectory);
        log.debug(`[EmbeddedObjects] Validated output directory: ${absoluteOutputDir}`);

        // Asegurarse de que el directorio de salida exista (validateFilePath no lo hace por defecto)
        try {
            await fs.access(absoluteOutputDir);
            const stats = await fs.stat(absoluteOutputDir);
            if (!stats.isDirectory()) {
                 throw new Error(`La ruta de salida no es un directorio: ${absoluteOutputDir}`);
            }
            log.debug(`[EmbeddedObjects] Output directory exists: ${absoluteOutputDir}`);
        } catch (error: any) {
             if (error.code === 'ENOENT') {
                 log.error(`El directorio de salida no existe: ${absoluteOutputDir}`);
                 throw new Error(`El directorio de salida no existe: ${absoluteOutputDir}`);
             } else {
                 log.error(`Error al acceder al directorio de salida ${absoluteOutputDir}: ${error.message}`);
                 throw new Error(`Error al acceder al directorio de salida: ${error.message}`);
             }
        }

        const extractedFiles: string[] = [];
        let oleObjectCount = 0;

        // Iterar sobre InlineShapes y Shapes (algunos objetos pueden estar en la capa de dibujo)
        // La API COM exacta para guardar/extraer necesita investigación profunda.
        // Podría ser algo como shape.OLEFormat.Object.SaveAs(...) o shape.OLEFormat.Activate() y luego interactuar.

        // Ejemplo conceptual (requiere validación API COM):
        /*
        for (let i = 1; i <= doc.InlineShapes.Count; i++) {
            const shape = doc.InlineShapes.Item(i);
            // wdInlineShapeOLEObject = 7, wdInlineShapeLinkedOLEObject = 8? Check constants.
            if (shape.Type === 7 || shape.Type === 8 || shape.OLEFormat) {
                oleObjectCount++;
                try {
                    // Intento 1: Usar SaveAs si existe en el objeto OLE directamente
                    // const oleObject = shape.OLEFormat.Object; // Esto puede variar mucho
                    // if (oleObject && typeof oleObject.SaveAs === 'function') { // SaveAs puede no existir o requerir formato específico
                    //     // Generar nombre de archivo único y seguro
                    //     const progId = shape.OLEFormat?.ProgID?.replace(/[^a-zA-Z0-9.-]/g, '_') || 'UnknownObject';
                    //     const filename = `embedded_inline_${i}_${progId}_${Date.now()}.ole`; // Usar extensión .ole o específica si se conoce
                    //     const outputPath = path.join(absoluteOutputDir, filename); // Usar ruta absoluta validada
                    //     // oleObject.SaveAs(outputPath); // La llamada exacta puede variar
                    //     // Necesita manejo de errores específico para SaveAs
                    //     extractedFiles.push(outputPath);
                    //     log.info(`Intentando extraer objeto InlineShape ${i} a ${outputPath}`);
                    // } else {
                    //     // Intento 2: Activar y copiar/pegar o guardar desde la aplicación OLE (MUY complejo)
                    //     log.warn(`No se pudo extraer el objeto InlineShape ${i} directamente. Método alternativo no implementado.`);
                    // }

                    // Intento 3: Extraer representación de icono/imagen (si aplica)
                    // if (shape.Picture) { ... shape.Picture.SaveAs(...) ... }

                    log.warn(`Extracción real para InlineShape ${i} no implementada.`);


                } catch (extractError: any) {
                    log.error(`Error extrayendo InlineShape ${i}: ${extractError.message}`);
                }
            }
        }
        // Repetir bucle similar para doc.Shapes si es necesario (objetos flotantes)
        for (let i = 1; i <= doc.Shapes.Count; i++) {
             const shape = doc.Shapes.Item(i);
             // msoEmbeddedOLEObject = 7, msoLinkedOLEObject = 10? Check constants.
             if (shape.Type === 7 || shape.Type === 10 || shape.OLEFormat) {
                 oleObjectCount++;
                 // Lógica de extracción similar a InlineShapes
                 log.warn(`Extracción real para Shape ${i} (flotante) no implementada.`);
             }
        }
        */

        if (oleObjectCount === 0) {
          return { success: true, message: 'No se encontraron objetos OLE incrustados en el documento.' };
        } else {
          // Actualizar mensaje cuando la extracción funcione
          return { success: false, message: `Se encontraron ${oleObjectCount} objetos OLE (inline o flotantes), pero la extracción aún no está implementada. Archivos extraídos: ${extractedFiles.length}`, details: { extractedPaths: extractedFiles } };
        }

      // case 'getProperties':
      //   // Lógica para obtener propiedades
      //   log.warn('Operación "getProperties" aún no implementada.');
      //   return { success: false, message: 'Operación "getProperties" no implementada.' };

      default:
         // El error "Property 'operation' does not exist on type 'never'" indica que TS
         // ha verificado que todos los casos de la unión discriminada están cubiertos.
         // Por lo tanto, este caso 'default' es teóricamente inalcanzable.
         // Lanzar un error genérico sin acceder a 'input'.
         const unreachableCase: never = input; // Mantenemos esto para la verificación de tipos
         log.error(`Caso inalcanzable en switch detectado: ${JSON.stringify(unreachableCase)}`);
         throw new Error(`Operación desconocida o no manejada.`);
    }

  } catch (error: any) {
    const message = error instanceof Error ? error.message : String(error);
    // Usar context.log si está disponible
    const logFn = context?.log?.error || logger.error; // Fallback a logger global si context no está
    // Acceder a input.operation y input.filePath aquí es seguro porque están fuera del switch/default
    // Comprobar si input existe antes de acceder a sus propiedades en caso de error muy temprano
    const operation = (input as any)?.operation || 'desconocida';
    const filePathLog = (input as any)?.filePath || 'desconocido';
    logFn(`Error en la operación '${operation}' en archivo '${filePathLog}': ${message}`, { error });
    // Devolver un mensaje de error más informativo
    return { success: false, message: `Error durante la operación '${operation}': ${message}` };
  } finally {
    // --- Bloque Finally Mejorado ---
    const logFnDebug = context?.log?.debug || logger.debug;
    const logFnWarn = context?.log?.warn || logger.warn;
    const logFnInfo = context?.log?.info || logger.info;

    if (doc) {
      try {
        // Cerrar sin guardar cambios, especialmente si fue solo lectura o hubo error
        // Si hubo éxito en insert/modify/delete, ya se debería haber guardado.
        // wdDoNotSaveChanges = 0
        doc.Close(0);
        logFnDebug(`[EmbeddedObjects] Documento cerrado: ${absoluteFilePath}`);
      } catch (closeError: any) {
        logFnWarn(`[EmbeddedObjects] Error al cerrar el documento ${absoluteFilePath}: ${closeError.message}`);
      }
      releaseObject(doc); // Liberar objeto del documento
      doc = null; // Ayuda a GC y evita doble liberación
    }
    if (wordApp) {
      try {
        // Intentar cerrar Word solo si no quedan otros documentos abiertos
        // Esto es arriesgado si el usuario tiene otros documentos abiertos.
        // Una opción más segura es simplemente liberar el objeto y dejar que Word se cierre solo eventualmente.
        // if (wordApp.Documents.Count === 0) {
        //    wordApp.Quit();
        //    logFnDebug('[EmbeddedObjects] Word application Quit() called.');
        // } else {
        //    logFnDebug('[EmbeddedObjects] Word application has other documents open, not quitting.');
        // }
      } catch (quitError: any) {
         logFnWarn(`[EmbeddedObjects] Error al intentar cerrar Word: ${quitError.message}`);
      }
      releaseObject(wordApp); // Liberar objeto de la aplicación
      wordApp = null; // Ayuda a GC
    }
    // Acceder a input.operation y input.filePath aquí es seguro
    // Comprobar si input existe antes de acceder a sus propiedades en caso de error muy temprano
    const operationFinal = (input as any)?.operation || 'desconocida';
    const filePathFinal = (input as any)?.filePath || 'desconocido';
    logFnInfo(`Finalizada operación '${operationFinal}' en archivo: ${filePathFinal}`);
  }
};

// --- Definición de la Herramienta ---
/**
 * @tool word/embedded-objects
 * @description Gestiona objetos OLE incrustados en documentos Word (.docx) usando COM Interop.
 * Permite insertar ('insert'), modificar ('modify'), eliminar ('delete') y extraer todos ('extractAll') los objetos incrustados.
 * La operación 'extractAll' intenta guardar todos los objetos OLE detectados (tanto inline como flotantes) en un directorio especificado.
 * Requiere investigación adicional de la API COM para la implementación completa, especialmente para 'extractAll' y 'modify'.
 * Utiliza winax para la interacción COM. Asegúrate de que Word esté instalado y accesible.
 *
 * @input_schema
 * {
 *   "type": "object",
 *   "properties": {
 *     "operation": { "enum": ["insert", "modify", "delete", "extractAll"] },
 *     "filePath": { "type": "string", "description": "Path al archivo .docx." },
 *     // Propiedades específicas por operación:
 *     "objectPath": { "type": "string", "description": "(insert) Path al archivo del objeto a insertar." },
 *     "objectIndex": { "type": "integer", "description": "(modify, delete) Índice (1-based) del objeto (InlineShapes primero, luego Shapes)." },
 *     "newObjectPath": { "type": "string", "description": "(modify) Path al nuevo archivo del objeto." },
 *     "outputDirectory": { "type": "string", "description": "(extractAll) Directorio donde guardar los objetos extraídos." }
 *     // ... otras propiedades opcionales ...
 *   },
 *   "required": ["operation", "filePath"], // Requeridos base
 *   // Añadir lógica para requeridos condicionales si es posible o validar en el handler
 *   "allOf": [
 *      {
 *          "if": { "properties": { "operation": { "const": "insert" } } },
 *          "then": { "required": ["objectPath"] }
 *      },
 *      {
 *          "if": { "properties": { "operation": { "const": "modify" } } },
 *          "then": { "required": ["objectIndex", "newObjectPath"] }
 *      },
 *      {
 *          "if": { "properties": { "operation": { "const": "delete" } } },
 *          "then": { "required": ["objectIndex"] }
 *      },
 *      {
 *          "if": { "properties": { "operation": { "const": "extractAll" } } },
 *          "then": { "required": ["outputDirectory"] }
 *      }
 *   ]
 * }
 *
 * @output_schema
 * {
 *   "type": "object",
 *   "properties": {
 *     "success": { "type": "boolean" },
 *     "message": { "type": "string" },
 *     "details": { "type": "object", "optional": true, "description": "Información adicional (ej: paths extraídos)." }
 *   },
 *   "required": ["success", "message"]
 * }
 *
 * @example_usage
 * // Extraer todos los objetos
 * {
 *   "tool_name": "word/embedded-objects",
 *   "arguments": {
 *     "operation": "extractAll",
 *     "filePath": "documentos/informe_con_objetos.docx",
 *     "outputDirectory": "output/objetos_extraidos"
 *   }
 * }
 * // Eliminar el segundo objeto (considerando InlineShapes y Shapes)
 * {
 *   "tool_name": "word/embedded-objects",
 *   "arguments": {
 *     "operation": "delete",
 *     "filePath": "documentos/informe_con_objetos.docx",
 *     "objectIndex": 2
 *   }
 * }
 */
// Eliminar anotación de tipo explícita, dejar que FastMCP la infiera al usar server.addTool
export const embeddedObjectsTool = {
  name: 'word/embedded-objects',
  description: 'Gestiona objetos OLE incrustados en documentos Word (insert, modify, delete, extractAll).',
  // FastMCP espera 'parameters' y 'execute', no 'schema' y 'handler' directamente en la definición del objeto.
  // El esquema Zod se pasa a 'parameters'.
  parameters: EmbeddedObjectsInputSchema,
  // El esquema de salida no se define aquí, se infiere del retorno de 'execute'.
  execute: embeddedObjectsExecute, // Renombrar 'handler' a 'execute'
  // Añadir anotaciones opcionales si se desea
  annotations: {
    title: "Word Embedded Objects Manager",
    readOnlyHint: false, // Puede modificar (insert, modify, delete)
    // destructiveHint: true, // Podría ser destructivo (delete)
    // openWorldHint: false, // No interactúa con el mundo exterior directamente (solo sistema de archivos local)
  }
};

// Exportar para index.ts
export default embeddedObjectsTool;