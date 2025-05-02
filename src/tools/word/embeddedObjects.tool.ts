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
      case 'insert': {
        log.info('Iniciando operación "insert".');
        const { objectPath } = input;
        const absoluteObjectPath = validateFilePath(objectPath);
        await fs.access(absoluteObjectPath); // Verificar existencia del archivo a insertar

        // Insertar al final del documento por defecto
        const range = doc.Content;
        range.Collapse(0); // wdCollapseEnd = 0

        // doc.InlineShapes.AddOLEObject(ClassType, FileName, LinkToFile, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Range)
        // ClassType: Opcional. Especifica la clase del objeto OLE. Si se omite, se determina por FileName.
        // FileName: Opcional. El archivo a insertar.
        // LinkToFile: Opcional. True para vincular, False para incrustar.
        // DisplayAsIcon: Opcional. True para mostrar como icono.
        // Range: Opcional. El rango donde insertar.
        const inlineShape = doc.InlineShapes.AddOLEObject(
          undefined, // ClassType
          absoluteObjectPath, // FileName
          false, // LinkToFile (incrustar)
          false, // DisplayAsIcon
          undefined, // IconFileName
          undefined, // IconIndex
          undefined, // IconLabel
          range // Range
        );

        doc.Save(); // Guardar cambios
        log.info(`Objeto insertado desde ${objectPath}.`);
        // El ID puede no ser el índice 1-based global, pero es un identificador útil.
        // Podríamos intentar encontrar el índice después de la inserción si fuera necesario.
        return { success: true, message: `Objeto insertado desde ${objectPath}.`, details: { insertedObjectId: inlineShape.Range.InlineShape.ID } };
      }

      case 'modify': {
        log.info('Iniciando operación "modify".');
        const { objectIndex, newObjectPath } = input;
        const absoluteNewObjectPath = validateFilePath(newObjectPath);
        await fs.access(absoluteNewObjectPath); // Verificar existencia del nuevo archivo

        // La modificación directa de objetos OLE incrustados vía COM es compleja.
        // Un enfoque común es eliminar el objeto existente e insertar el nuevo.
        // Intentaremos mantener la posición si es un InlineShape.

        let shapeToModify: any = null;
        let isInline = false;

        // Buscar en InlineShapes primero
        if (objectIndex > 0 && objectIndex <= doc.InlineShapes.Count) {
          shapeToModify = doc.InlineShapes.Item(objectIndex);
          isInline = true;
          log.debug(`[EmbeddedObjects] Found InlineShape at index ${objectIndex}.`);
        } else {
          // Si no está en InlineShapes, buscar en Shapes (objetos flotantes)
          // El índice para Shapes es relativo a la colección Shapes, no global.
          // Necesitamos ajustar el índice.
          const shapesIndex = objectIndex - doc.InlineShapes.Count;
          if (shapesIndex > 0 && shapesIndex <= doc.Shapes.Count) {
            shapeToModify = doc.Shapes.Item(shapesIndex);
            isInline = false;
            log.debug(`[EmbeddedObjects] Found Shape at index ${objectIndex} (relative index ${shapesIndex}).`);
          } else {
            throw new Error(`Índice de objeto ${objectIndex} fuera de rango. Total InlineShapes: ${doc.InlineShapes.Count}, Total Shapes: ${doc.Shapes.Count}.`);
          }
        }

        // Verificar si el objeto es OLE antes de intentar modificar/eliminar
        // wdInlineShapeOLEObject = 7, wdInlineShapeLinkedOLEObject = 8
        // msoEmbeddedOLEObject = 7, msoLinkedOLEObject = 10
        const isOLE = shapeToModify.Type === 7 || shapeToModify.Type === 8 || shapeToModify.Type === 10 || shapeToModify.OLEFormat;

        if (!isOLE) {
           throw new Error(`El objeto en el índice ${objectIndex} no es un objeto OLE incrustado o vinculado.`);
        }

        let originalRange: any = null;
        let originalLeft: number | undefined;
        let originalTop: number | undefined;
        let originalAnchor: any = null;

        if (isInline) {
           originalRange = shapeToModify.Range; // Capturar el rango antes de eliminar
        } else {
           // Para Shapes flotantes, capturar la posición y el ancla.
           originalLeft = shapeToModify.Left;
           originalTop = shapeToModify.Top;
           originalAnchor = shapeToModify.Anchor;
           log.warn(`Modificación de Shape flotante (índice ${objectIndex}) intentará mantener la posición, pero puede variar.`);
        }

        shapeToModify.Delete(); // Eliminar el objeto existente
        log.debug(`[EmbeddedObjects] Deleted object at index ${objectIndex}.`);

        let newShape: any = null;
        if (isInline && originalRange) {
           // Intentar insertar el nuevo objeto en el rango original
           originalRange.Collapse(0); // wdCollapseEnd = 0
           newShape = doc.InlineShapes.AddOLEObject(
              undefined, // ClassType
              absoluteNewObjectPath, // FileName
              false, // LinkToFile (incrustar)
              false, // DisplayAsIcon
              undefined, // IconFileName
              undefined, // IconIndex
              undefined, // IconLabel
              originalRange // Range
           );
           log.info(`Objeto en índice ${objectIndex} modificado (reemplazado) con ${newObjectPath} en la posición original.`);
        } else if (!isInline && originalAnchor) {
            // Intentar reinsertar Shape flotante con la misma posición y ancla
            // AddOLEObject en Shapes colección es diferente: AddOLEObject(ClassType, FileName, LinkToFile, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Left, Top, Width, Height, Anchor)
            newShape = doc.Shapes.AddOLEObject(
                undefined, // ClassType
                absoluteNewObjectPath, // FileName
                false, // LinkToFile (incrustar)
                false, // DisplayAsIcon
                undefined, // IconFileName
                undefined, // IconIndex
                undefined, // IconLabel
                originalLeft, // Left
                originalTop, // Top
                undefined, // Width (auto)
                undefined, // Height (auto)
                originalAnchor // Anchor
            );
            log.info(`Objeto en índice ${objectIndex} modificado (reemplazado) con ${newObjectPath}, reinsertado como Shape flotante.`);

        } else {
           // Insertar al final si no se pudo mantener la posición (InlineShape sin rango o Shape sin ancla)
           const endRange = doc.Content;
           endRange.Collapse(0); // wdCollapseEnd = 0
           newShape = doc.InlineShapes.AddOLEObject(
              undefined, // ClassType
              absoluteNewObjectPath, // FileName
              false, // LinkToFile (incrustar)
              false, // DisplayAsIcon
              undefined, // IconFileName
              undefined, // IconIndex
              undefined, // IconLabel
              endRange // Range
           );
           log.info(`Objeto en índice ${objectIndex} modificado (reemplazado) con ${newObjectPath} al final del documento.`);
        }

        doc.Save(); // Guardar cambios
        return { success: true, message: `Objeto en índice ${objectIndex} modificado (reemplazado) con ${newObjectPath}.` };
      }

      case 'delete': {
        log.info('Iniciando operación "delete".');
        const { objectIndex } = input;

        let shapeToDelete: any = null;
        let isInline = false;

        // Buscar en InlineShapes primero
        if (objectIndex > 0 && objectIndex <= doc.InlineShapes.Count) {
          shapeToDelete = doc.InlineShapes.Item(objectIndex);
          isInline = true;
          log.debug(`[EmbeddedObjects] Found InlineShape at index ${objectIndex} for deletion.`);
        } else {
          // Si no está en InlineShapes, buscar en Shapes (objetos flotantes)
          const shapesIndex = objectIndex - doc.InlineShapes.Count;
          if (shapesIndex > 0 && shapesIndex <= doc.Shapes.Count) {
            shapeToDelete = doc.Shapes.Item(shapesIndex);
            isInline = false;
            log.debug(`[EmbeddedObjects] Found Shape at index ${objectIndex} (relative index ${shapesIndex}) for deletion.`);
          } else {
            throw new Error(`Índice de objeto ${objectIndex} fuera de rango. Total InlineShapes: ${doc.InlineShapes.Count}, Total Shapes: ${doc.Shapes.Count}.`);
          }
        }

        // Verificar si el objeto es OLE antes de eliminar
        const isOLE = shapeToDelete.Type === 7 || shapeToDelete.Type === 8 || shapeToDelete.Type === 10 || shapeToDelete.OLEFormat;

        if (!isOLE) {
           throw new Error(`El objeto en el índice ${objectIndex} no es un objeto OLE incrustado o vinculado y no puede ser eliminado por esta herramienta.`);
        }

        shapeToDelete.Delete(); // Eliminar el objeto
        doc.Save(); // Guardar cambios
        log.info(`Objeto en índice ${objectIndex} eliminado.`);
        return { success: true, message: `Objeto en índice ${objectIndex} eliminado.` };
      }

      case 'extractAll': {
        log.info('Iniciando operación "extractAll".');
        const { outputDirectory } = input;
        const absoluteOutputDir = validateFilePath(outputDirectory);
        log.debug(`[EmbeddedObjects] Validated output directory: ${absoluteOutputDir}`);

        // Asegurarse de que el directorio de salida exista, crearlo si no
        try {
            await fs.access(absoluteOutputDir);
            const stats = await fs.stat(absoluteOutputDir);
            if (!stats.isDirectory()) {
                 throw new Error(`La ruta de salida no es un directorio: ${absoluteOutputDir}`);
            }
            log.debug(`[EmbeddedObjects] Output directory exists: ${absoluteOutputDir}`);
        } catch (error: any) {
             if (error.code === 'ENOENT') {
                 log.info(`El directorio de salida no existe, intentando crearlo: ${absoluteOutputDir}`);
                 await fs.mkdir(absoluteOutputDir, { recursive: true });
                 log.info(`Directorio de salida creado: ${absoluteOutputDir}`);
             } else {
                 log.error(`Error al acceder al directorio de salida ${absoluteOutputDir}: ${error.message}`);
                 throw new Error(`Error al acceder al directorio de salida: ${error.message}`);
             }
        }

        const extractedFiles: string[] = [];
        let oleObjectCount = 0;

        // Extraer de InlineShapes
        for (let i = 1; i <= doc.InlineShapes.Count; i++) {
            const shape = doc.InlineShapes.Item(i);
            // wdInlineShapeOLEObject = 7, wdInlineShapeLinkedOLEObject = 8
            if (shape.Type === 7 || shape.Type === 8 || shape.OLEFormat) {
                oleObjectCount++;
                try {
                    const oleFormat = shape.OLEFormat;
                    if (oleFormat && oleFormat.Object && typeof oleFormat.Object.SaveAs === 'function') {
                        // Generar nombre de archivo único y seguro
                        const progId = oleFormat?.ProgID?.replace(/[^a-zA-Z0-9.-]/g, '_') || 'UnknownObject';
                        let baseFilename = `embedded_inline_${i}_${progId}`;
                        let filename = `${baseFilename}.ole`; // Extensión por defecto
                        let outputPath = path.join(absoluteOutputDir, filename);
                        let counter = 1;

                        // Manejar colisiones de nombres
                        while (await fs.access(outputPath).then(() => true).catch(() => false)) {
                            filename = `${baseFilename}_${counter++}.ole`;
                            outputPath = path.join(absoluteOutputDir, filename);
                        }

                        // Intentar guardar el objeto OLE
                        // La API SaveAs puede requerir un formato específico o no estar disponible para todos los tipos.
                        // Si falla, intentaremos un método alternativo si es posible.
                        try {
                           oleFormat.Object.SaveAs(outputPath);
                           extractedFiles.push(outputPath);
                           log.info(`Extraído InlineShape ${i} a ${outputPath}`);
                        } catch (saveError: any) {
                           log.warn(`Error al usar SaveAs en InlineShape ${i} (${progId}): ${saveError.message}. Intentando método alternativo (si aplica).`);
                           // Método alternativo: Si es una imagen OLE, intentar guardar la imagen
                           if (shape.Type === 7 && shape.OLEFormat?.ProgID?.toLowerCase().includes('package')) {
                              // Los objetos "Package" a menudo son archivos incrustados.
                              // No hay un método SaveAs directo en el objeto OLE.
                              // La extracción de Packages es compleja y a menudo requiere activar el objeto.
                              log.warn(`Extracción de objeto Package (InlineShape ${i}) no implementada directamente.`);
                           } else if (shape.Picture) {
                              // Si tiene una representación de imagen, intentar guardarla
                              try {
                                 filename = `${baseFilename}.png`; // O determinar extensión
                                 outputPath = path.join(absoluteOutputDir, filename);
                                 counter = 1;
                                 while (await fs.access(outputPath).then(() => true).catch(() => false)) {
                                     filename = `${baseFilename}_${counter++}.png`;
                                     outputPath = path.join(absoluteOutputDir, filename);
                                 }
                                 shape.Picture.SaveAs(outputPath); // wdFormatPNG = 13 (o usar constante)
                                 extractedFiles.push(outputPath);
                                 log.info(`Extraída imagen de InlineShape ${i} a ${outputPath}`);
                              } catch (pictureSaveError: any) {
                                 log.error(`Error al extraer imagen de InlineShape ${i}: ${pictureSaveError.message}`);
                              }
                           } else {
                              log.error(`No se pudo extraer InlineShape ${i} (${progId}) usando SaveAs ni método alternativo.`);
                           }
                        }

                    } else {
                        log.warn(`InlineShape ${i} no parece tener un objeto OLE con método SaveAs o no es un tipo OLE manejable directamente.`);
                    }
                } catch (extractError: any) {
                    log.error(`Error general extrayendo InlineShape ${i}: ${extractError.message}`);
                } finally {
                    // Liberar el objeto shape si es necesario, aunque winax a menudo maneja esto
                    releaseObject(shape);
                }
            }
        }

        // Extraer de Shapes (objetos flotantes)
        for (let i = 1; i <= doc.Shapes.Count; i++) {
             const shape = doc.Shapes.Item(i);
             // msoEmbeddedOLEObject = 7, msoLinkedOLEObject = 10
             if (shape.Type === 7 || shape.Type === 10 || shape.OLEFormat) {
                 oleObjectCount++;
                 try {
                     const oleFormat = shape.OLEFormat;
                     if (oleFormat && oleFormat.Object && typeof oleFormat.Object.SaveAs === 'function') {
                         // Generar nombre de archivo único y seguro
                         const progId = oleFormat?.ProgID?.replace(/[^a-zA-Z0-9.-]/g, '_') || 'UnknownObject';
                         let baseFilename = `embedded_shape_${i}_${progId}`;
                         let filename = `${baseFilename}.ole`; // Extensión por defecto
                         let outputPath = path.join(absoluteOutputDir, filename);
                         let counter = 1;

                         // Manejar colisiones de nombres
                         while (await fs.access(outputPath).then(() => true).catch(() => false)) {
                             filename = `${baseFilename}_${counter++}.ole`;
                             outputPath = path.join(absoluteOutputDir, filename);
                         }

                         // Intentar guardar el objeto OLE
                         try {
                            oleFormat.Object.SaveAs(outputPath);
                            extractedFiles.push(outputPath);
                            log.info(`Extraído Shape ${i} a ${outputPath}`);
                         } catch (saveError: any) {
                            log.warn(`Error al usar SaveAs en Shape ${i} (${progId}): ${saveError.message}. Intentando método alternativo (si aplica).`);
                            // Método alternativo: Si es una imagen OLE, intentar guardar la imagen
                            if (shape.Type === 7 && shape.OLEFormat?.ProgID?.toLowerCase().includes('package')) {
                               log.warn(`Extracción de objeto Package (Shape ${i}) no implementada directamente.`);
                            } else if (shape.Picture) {
                               try {
                                  filename = `${baseFilename}.png`; // O determinar extensión
                                  outputPath = path.join(absoluteOutputDir, filename);
                                  counter = 1;
                                  while (await fs.access(outputPath).then(() => true).catch(() => false)) {
                                      filename = `${baseFilename}_${counter++}.png`;
                                      outputPath = path.join(absoluteOutputDir, filename);
                                  }
                                  shape.Picture.SaveAs(outputPath); // wdFormatPNG = 13 (o usar constante)
                                  extractedFiles.push(outputPath);
                                  log.info(`Extraída imagen de Shape ${i} a ${outputPath}`);
                               } catch (pictureSaveError: any) {
                                  log.error(`Error al extraer imagen de Shape ${i}: ${pictureSaveError.message}`);
                               }
                            } else {
                               log.error(`No se pudo extraer Shape ${i} (${progId}) usando SaveAs ni método alternativo.`);
                            }
                         }

                     } else {
                         log.warn(`Shape ${i} no parece tener un objeto OLE con método SaveAs o no es un tipo OLE manejable directamente.`);
                     }
                 } catch (extractError: any) {
                     log.error(`Error general extrayendo Shape ${i}: ${extractError.message}`);
                 } finally {
                     // Liberar el objeto shape si es necesario
                     releaseObject(shape);
                 }
             }
        }


        if (oleObjectCount === 0) {
          return { success: true, message: 'No se encontraron objetos OLE incrustados o vinculados en el documento.' };
        } else {
          return { success: true, message: `Se encontraron ${oleObjectCount} objetos OLE (inline o flotantes). Archivos extraídos: ${extractedFiles.length}.`, details: { extractedPaths: extractedFiles } };
        }
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