/**
 * @file Implements the 'word/embedded-objects' tool using COM Interop (winax) for managing embedded OLE objects in Word documents.
 * Provides functionality to insert, modify, delete, and extract embedded objects.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types'; // Import necessary types
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { validateFilePath } from '../../utils/security'; // Corrected: Function name
import * as fs from 'fs/promises';
import * as path from 'path';
import logger from '../../utils/logger'; // Corrected: Default import

// --- Input Schema (Zod) ---

/** Base schema for embedded object operations requiring a file path. */
const EmbeddedObjectBaseSchema = z.object({
  /** The path to the Word file (relative to the current workspace directory). */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
});

/** Combined schema for all embedded object operations. */
const WordEmbeddedObjectsInputSchema = EmbeddedObjectBaseSchema.extend({
    /** The operation to perform ('insert', 'modify', 'delete', or 'extractAll'). */
    operation: z.enum(['insert', 'modify', 'delete', 'extractAll']).describe('The operation to perform (insert, modify, delete, or extractAll).'),
    /** The path to the file of the object to insert. Required for 'insert'. */
    objectPath: z.string().optional().describe("(insert) Path to the object file to insert. Required for 'insert'."), // Make optional here, validate in handler
    /** The 1-based index of the object (InlineShapes first, then Shapes). Required for 'modify' and 'delete'. */
    objectIndex: z.number().int().positive('Object index must be a positive integer. Required for modify, delete.').optional().describe("(modify, delete) 1-based index of the object (InlineShapes first, then Shapes)."), // Make optional here, validate in handler
    /** The path to the new object file. Required for 'modify'. */
    newObjectPath: z.string().optional().describe("(modify) Path to the new object file. Required for 'modify'."), // Make optional here, validate in handler
    /** The directory where extracted objects will be saved. Required for 'extractAll'. */
    outputDirectory: z.string().optional().describe("(extractAll) Directory where extracted objects will be saved. Required for 'extractAll'."), // Make optional here, validate in handler
    // Optional: position, linkToFile, displayAsIcon, iconFileName, iconIndex, iconLabel, etc.
    // range: z.string().optional().describe("Range where to insert (e.g.: 'paragraph:N', 'selection', 'end'). Default: 'end'"),
}).refine(data => {
    // Specific validations per operation within the refinement
    if (data.operation === 'insert') {
        return data.objectPath !== undefined; // Requires objectPath
    } else if (data.operation === 'modify') {
        return data.objectIndex !== undefined && data.newObjectPath !== undefined; // Requires objectIndex and newObjectPath
    } else if (data.operation === 'delete') {
        return data.objectIndex !== undefined; // Requires objectIndex
    } else if (data.operation === 'extractAll') {
        return data.outputDirectory !== undefined; // Requires outputDirectory
    }
    return true; // Passes validation if the operation doesn't require specific fields or if it has them
}, {
    message: "Invalid input for the specified operation. Check required fields (objectPath, objectIndex, newObjectPath, outputDirectory).",
    path: [], // Apply error to the whole object
});


/** Infer the combined type for use in the handler. */
type WordEmbeddedObjectsInput = z.infer<typeof WordEmbeddedObjectsInputSchema>;


// --- Output Schema (Zod) ---
// Not strictly necessary to define here if not explicitly validating the return,
// but can be useful for documentation or internal types.
/** Schema for the output of embedded object operations. */
const EmbeddedObjectsOutputSchema = z.object({
  /** Indicates if the operation was successful. */
  success: z.boolean(),
  /** A message describing the outcome of the operation. */
  message: z.string(),
  /** Optional additional details, such as paths of extracted files. Can be any type. */
  details: z.any().optional(), // To return extra information, like paths of extracted files
});

// --- Handler (now called 'execute') ---
/**
 * Handler function for the 'word/embedded-objects' tool.
 * Executes the specified operation on embedded OLE objects in a Word document.
 * @param params - The parameters for the tool, validated against `WordEmbeddedObjectsInputSchema`.
 * @param context - The FastMCP context (optional).
 * @returns A promise resolving to an ApiResponse indicating the outcome of the operation.
 * @throws {Error} If validation fails, the document cannot be opened, the object is not found/not OLE, or an operation fails.
 */
const embeddedObjectsExecute = async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> => {
  const log = context?.log || logger; // Use the context logger if available, otherwise the global one
  log.info(`[EmbeddedObjects] Starting operation on file: ${params.filePath}`); // Use params directly for initial filePath

  let validatedRequest: WordEmbeddedObjectsInput;
  try {
    // Validar y parsear los params genéricos usando el nuevo WordEmbeddedObjectsInputSchema
    validatedRequest = WordEmbeddedObjectsInputSchema.parse(params);
  } catch (error: any) {
     // Si la validación inicial falla, devolver un error de validación
     if (error instanceof z.ZodError) {
         // Serializar error.errors para que sea serializable
         const errorDetails = JSON.stringify(error.errors, null, 2);
         log.warn(`[EmbeddedObjects] Input validation failed at handler entry for word/embedded-objects: ${error.message}`, { errors: errorDetails, params });
         return {
             success: false,
             error: {
                 code: 'VALIDATION_ERROR',
                 message: `Input validation failed: ${error.errors.map(e => `${e.path.join('.')} - ${e.message}`).join(', ')}`,
                 details: errorDetails,
             },
         };
     }
     // Otro error inesperado durante el parseo inicial
     log.error(`[EmbeddedObjects] Unexpected error parsing params in word/embedded-objects handler: ${error.message}`, { error: String(error), params });
     return {
         success: false,
         error: {
             code: 'INTERNAL_ERROR',
             message: 'Failed to parse tool parameters.',
             details: String(error),
         },
     };
  }

  const { operation, filePath, objectPath, objectIndex, newObjectPath, outputDirectory } = validatedRequest; // Extract operation and specific fields directly

  // Use validateFilePath which returns the validated absolute path or throws an error
  const absoluteFilePath = validateFilePath(filePath);
  log.debug(`[EmbeddedObjects] Validated absolute path: ${absoluteFilePath}`);

  let wordApp: any = null; // Use 'any' for the application COM object
  let doc: any = null; // Use 'any' for the document COM object

  try {
    wordApp = await getOfficeApplication('Word.Application');
    // Open the document. Consider ReadOnly for 'extractAll' and 'getProperties'
    const openReadOnly = operation === 'extractAll'; // || operation === 'getProperties';
    log.debug(`[EmbeddedObjects] Opening document: ${absoluteFilePath} (ReadOnly: ${openReadOnly})`);
      // Open parameters: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, ... Visible
      // Open non-visible for background operations
      doc = wordApp.Documents.Open(absoluteFilePath, false, openReadOnly, false, undefined, undefined, undefined, undefined, undefined, undefined, false); // Visible = false at the end

    switch (operation) {
      case 'insert': {
        log.info('[EmbeddedObjects] Starting "insert" operation.');
        // Validar que objectPath is not undefined (already done in refine, but good practice)
        if (objectPath === undefined) throw new Error("objectPath is required for 'insert' operation.");
        const absoluteObjectPath = validateFilePath(objectPath);
        await fs.access(absoluteObjectPath); // Verify existence of the file to insert

        // Insert at the end of the document by default
        const range = doc.Content;
        range.Collapse(0); // wdCollapseEnd = 0

        // doc.InlineShapes.AddOLEObject(ClassType, FileName, LinkToFile, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Range)
        // ClassType: Optional. Specifies the class of the OLE object. If omitted, it is determined by FileName.
        // FileName: Optional. The file to insert.
        // LinkToFile: Optional. True to link, False to embed.
        // DisplayAsIcon: Optional. True to display as an icon.
        // Range: Optional. The range where to insert.
        const inlineShape = doc.InlineShapes.AddOLEObject(
          undefined, // ClassType
          absoluteObjectPath, // FileName
          false, // LinkToFile (embed)
          false, // DisplayAsIcon
          undefined, // IconFileName
          undefined, // IconIndex
          undefined, // IconLabel
          range // Range
        );

        doc.Save(); // Save changes
        log.info(`[EmbeddedObjects] Object inserted from ${objectPath}.`);
        // The ID may not be the global 1-based index, but it's a useful identifier.
        // We could try to find the index after insertion if needed.
        return { success: true, message: `Object inserted from ${objectPath}.`, data: { insertedObjectId: inlineShape.Range.InlineShape.ID } }; // Move to data
      }

      case 'modify': {
        log.info('[EmbeddedObjects] Starting "modify" operation.');
        // Validar que objectIndex and newObjectPath are not undefined
        if (objectIndex === undefined) throw new Error("objectIndex is required for 'modify' operation.");
        if (newObjectPath === undefined) throw new Error("newObjectPath is required for 'modify' operation.");

        const absoluteNewObjectPath = validateFilePath(newObjectPath);
        await fs.access(absoluteNewObjectPath); // Verify existence of the new file

        // Direct modification of embedded OLE objects via COM is complex.
        // A common approach is to delete the existing object and insert the new one.
        // We will try to maintain the position if it's an InlineShape.

        let shapeToModify: any = null;
        let isInline = false;

        // Search in InlineShapes first
        if (objectIndex > 0 && objectIndex <= doc.InlineShapes.Count) {
          shapeToModify = doc.InlineShapes.Item(objectIndex);
          isInline = true;
          log.debug(`[EmbeddedObjects] Found InlineShape at index ${objectIndex}.`);
        } else {
          // If not in InlineShapes, search in Shapes (floating objects)
          // The index for Shapes is relative to the Shapes collection, not global.
          // We need to adjust the index.
          const shapesIndex = objectIndex - doc.InlineShapes.Count;
          if (shapesIndex > 0 && shapesIndex <= doc.Shapes.Count) {
            shapeToModify = doc.Shapes.Item(shapesIndex);
            isInline = false;
            log.debug(`[EmbeddedObjects] Found Shape at index ${objectIndex} (relative index ${shapesIndex}).`);
          } else {
            throw new Error(`Object index ${objectIndex} out of range. Total InlineShapes: ${doc.InlineShapes.Count}, Total Shapes: ${doc.Shapes.Count}.`);
          }
        }

        // Verify if the object is OLE before attempting to modify/delete
        // wdInlineShapeOLEObject = 7, wdInlineShapeLinkedOLEObject = 8
        // msoEmbeddedOLEObject = 7, msoLinkedOLEObject = 10
        const isOLE = shapeToModify.Type === 7 || shapeToModify.Type === 8 || shapeToModify.Type === 10 || shapeToModify.OLEFormat;

        if (!isOLE) {
           throw new Error(`The object at index ${objectIndex} is not an embedded or linked OLE object.`);
        }

        let originalRange: any = null;
        let originalLeft: number | undefined;
        let originalTop: number | undefined;
        let originalAnchor: any = null;

        if (isInline) {
           originalRange = shapeToModify.Range; // Capture the range before deleting
        } else {
           // For floating Shapes, capture the position and anchor.
           originalLeft = shapeToModify.Left;
           originalTop = shapeToModify.Top;
           originalAnchor = shapeToModify.Anchor;
           log.warn(`[EmbeddedObjects] Modification of floating Shape (index ${objectIndex}) will attempt to maintain position, but it may vary.`);
        }

        shapeToModify.Delete(); // Delete the existing object
        log.debug(`[EmbeddedObjects] Deleted object at index ${objectIndex}.`);

        let newShape: any = null;
        if (isInline && originalRange) {
           // Attempt to insert the new object in the original range
           originalRange.Collapse(0); // wdCollapseEnd = 0
           newShape = doc.InlineShapes.AddOLEObject(
              undefined, // ClassType
              absoluteNewObjectPath, // FileName
              false, // LinkToFile (embed)
              false, // DisplayAsIcon
              undefined, // IconFileName
              undefined, // IconIndex
              undefined, // IconLabel
              originalRange // Range
            );
           log.info(`[EmbeddedObjects] Object at index ${objectIndex} modified (reemplazado) con ${newObjectPath} en la posición original.`);
        } else if (!isInline && originalAnchor) {
            // Attempt to reinsert floating Shape with the same position and anchor
            // AddOLEObject in Shapes collection is different: AddOLEObject(ClassType, FileName, LinkToFile, DisplayAsIcon, IconFileName, IconIndex, IconLabel, Left, Top, Width, Height, Anchor)
            newShape = doc.Shapes.AddOLEObject(
                undefined, // ClassType
                absoluteNewObjectPath, // FileName
                false, // LinkToFile (embed)
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
            log.info(`[EmbeddedObjects] Object at index ${objectIndex} modified (reemplazado) con ${newObjectPath}, reinsertado como Shape flotante.`);

        } else {
           // Insert at the end if position could not be maintained (InlineShape without range or Shape without anchor)
           const endRange = doc.Content;
           endRange.Collapse(0); // wdCollapseEnd = 0
           newShape = doc.InlineShapes.AddOLEObject(
              undefined, // ClassType
              absoluteNewObjectPath, // FileName
              false, // LinkToFile (embed)
              false, // DisplayAsIcon
              undefined, // IconFileName
              undefined, // IconIndex
              undefined, // IconLabel
              endRange // Range
           );
           log.info(`Objeto en índice ${objectIndex} modificado (reemplazado) con ${newObjectPath} al final del documento.`);
        }

        doc.Save(); // Save changes
        return { success: true, message: `Object at index ${objectIndex} modified (reemplazado) con ${newObjectPath}.`, data: null }; // Add data: null
      }

      case 'delete': {
        log.info('[EmbeddedObjects] Starting "delete" operation.');
        const { objectIndex } = validatedRequest; // Use validatedRequest
        // Validar que objectIndex is not undefined
        if (objectIndex === undefined) throw new Error("objectIndex is required for 'delete' operation.");


        let shapeToDelete: any = null;
        let isInline = false;

        // Search in InlineShapes first
        if (objectIndex > 0 && objectIndex <= doc.InlineShapes.Count) {
          shapeToDelete = doc.InlineShapes.Item(objectIndex);
          isInline = true;
          log.debug(`[EmbeddedObjects] Found InlineShape at index ${objectIndex} for deletion.`);
        } else {
          // If not in InlineShapes, search in Shapes (floating objects)
          const shapesIndex = objectIndex - doc.InlineShapes.Count;
          if (shapesIndex > 0 && shapesIndex <= doc.Shapes.Count) {
            shapeToDelete = doc.Shapes.Item(shapesIndex);
            isInline = false;
            log.debug(`[EmbeddedObjects] Found Shape at index ${objectIndex} (relative index ${shapesIndex}) for deletion.`);
          } else {
            throw new Error(`Object index ${objectIndex} out of range. Total InlineShapes: ${doc.InlineShapes.Count}, Total Shapes: ${doc.Shapes.Count}.`);
          }
        }

        // Verify if the object is OLE before deleting
        const isOLE = shapeToDelete.Type === 7 || shapeToDelete.Type === 8 || shapeToDelete.Type === 10 || shapeToDelete.OLEFormat;

        if (!isOLE) {
           throw new Error(`The object at index ${objectIndex} is not an embedded or linked OLE object and cannot be deleted by this tool.`);
        }

        shapeToDelete.Delete(); // Delete the object
        doc.Save(); // Save changes
        log.info(`[EmbeddedObjects] Object at index ${objectIndex} deleted.`);
        return { success: true, message: `Object at index ${objectIndex} deleted.`, data: null }; // Add data: null
      }

      case 'extractAll': {
        log.info('[EmbeddedObjects] Starting "extractAll" operation.');
        const { outputDirectory } = validatedRequest; // Use validatedRequest
        // Validar que outputDirectory is not undefined
        if (outputDirectory === undefined) throw new Error("outputDirectory is required for 'extractAll' operation.");

        const absoluteOutputDir = validateFilePath(outputDirectory);
        log.debug(`[EmbeddedObjects] Validated output directory: ${absoluteOutputDir}`);

        // Ensure the output directory exists, create it if not
        try {
            await fs.access(absoluteOutputDir);
            const stats = await fs.stat(absoluteOutputDir);
            if (!stats.isDirectory()) {
                 throw new Error(`The output path is not a directory: ${absoluteOutputDir}`);
            }
            log.debug(`[EmbeddedObjects] Output directory exists: ${absoluteOutputDir}`);
        } catch (error: any) {
             if (error.code === 'ENOENT') {
                 log.info(`[EmbeddedObjects] Output directory does not exist, attempting to create: ${absoluteOutputDir}`);
                 await fs.mkdir(absoluteOutputDir, { recursive: true });
                 log.info(`[EmbeddedObjects] Output directory created: ${absoluteOutputDir}`);
             } else {
                 log.error(`[EmbeddedObjects] Error accessing output directory ${absoluteOutputDir}: ${error.message}`);
                 throw new Error(`Error accessing output directory: ${error.message}`);
             }
        }

        const extractedFiles: string[] = [];
        let oleObjectCount = 0;

        // Extract from InlineShapes
        for (let i = 1; i <= doc.InlineShapes.Count; i++) {
            const shape = doc.InlineShapes.Item(i);
            // wdInlineShapeOLEObject = 7, wdInlineShapeLinkedOLEObject = 8
            if (shape.Type === 7 || shape.Type === 8 || shape.OLEFormat) {
                oleObjectCount++;
                try {
                    const oleFormat = shape.OLEFormat;
                    if (oleFormat && oleFormat.Object && typeof oleFormat.Object.SaveAs === 'function') {
                        // Generate a unique and safe filename
                        const progId = oleFormat?.ProgID?.replace(/[^a-zA-Z0-9.-]/g, '_') || 'UnknownObject';
                        let baseFilename = `embedded_inline_${i}_${progId}`;
                        let filename = `${baseFilename}.ole`; // Default extension
                        let outputPath = path.join(absoluteOutputDir, filename);
                        let counter = 1;

                        // Handle name collisions
                        while (await fs.access(outputPath).then(() => true).catch(() => false)) {
                            filename = `${baseFilename}_${counter++}.ole`;
                            outputPath = path.join(absoluteOutputDir, filename);
                        }

                        // Attempt to save the OLE object
                        // The SaveAs API may require a specific format or not be available for all types.
                        // If it fails, we will try an alternative method if possible.
                        try {
                           oleFormat.Object.SaveAs(outputPath);
                           extractedFiles.push(outputPath);
                           log.info(`[EmbeddedObjects] Extracted InlineShape ${i} to ${outputPath}`);
                        } catch (saveError: any) {
                           log.warn(`[EmbeddedObjects] Error using SaveAs on InlineShape ${i} (${progId}): ${saveError.message}. Attempting alternative method (if applicable).`);
                           // Alternative method: If it's an OLE image, try to save the picture
                           if (shape.Type === 7 && shape.OLEFormat?.ProgID?.toLowerCase().includes('package')) {
                              // "Package" objects are often embedded files.
                              // There is no direct SaveAs method on the OLE object.
                              // Extraction of Packages is complex and often requires activating the object.
                              log.warn(`[EmbeddedObjects] Extraction of Package object (InlineShape ${i}) not implemented directly.`);
                           } else if (shape.Picture) {
                              // If it has a picture representation, try to save it
                              try {
                                 filename = `${baseFilename}.png`; // Or determine extension
                                 outputPath = path.join(absoluteOutputDir, filename);
                                 counter = 1;
                                 while (await fs.access(outputPath).then(() => true).catch(() => false)) {
                                     filename = `${baseFilename}_${counter++}.png`;
                                     outputPath = path.join(absoluteOutputDir, filename);
                                 }
                                 shape.Picture.SaveAs(outputPath); // wdFormatPNG = 13 (or use constant)
                                 extractedFiles.push(outputPath);
                                 log.info(`[EmbeddedObjects] Extracted image from InlineShape ${i} to ${outputPath}`);
                              } catch (pictureSaveError: any) {
                                 log.error(`[EmbeddedObjects] Error extracting image from InlineShape ${i}: ${pictureSaveError.message}`);
                              }
                           } else {
                              log.error(`[EmbeddedObjects] Could not extract InlineShape ${i} (${progId}) using SaveAs or alternative method.`);
                           }
                        }

                    } else {
                        log.warn(`[EmbeddedObjects] InlineShape ${i} does not seem to have an OLE object with SaveAs method or is not a directly manageable OLE type.`);
                    }
                } catch (extractError: any) {
                    log.error(`[EmbeddedObjects] General error extracting InlineShape ${i}: ${extractError.message}`);
                } finally {
                    // Release the shape object if necessary, although winax often handles this
                    releaseObject(shape);
                }
            }
        }

        // Extract from Shapes (floating objects)
        for (let i = 1; i <= doc.Shapes.Count; i++) {
             const shape = doc.Shapes.Item(i);
             // msoEmbeddedOLEObject = 7, msoLinkedOLEObject = 10
             if (shape.Type === 7 || shape.Type === 10 || shape.OLEFormat) {
                 oleObjectCount++;
                 try {
                     const oleFormat = shape.OLEFormat;
                     if (oleFormat && oleFormat.Object && typeof oleFormat.Object.SaveAs === 'function') {
                         // Generate a unique and safe filename
                         const progId = oleFormat?.ProgID?.replace(/[^a-zA-Z0-9.-]/g, '_') || 'UnknownObject';
                         let baseFilename = `embedded_shape_${i}_${progId}`;
                         let filename = `${baseFilename}.ole`; // Default extension
                         let outputPath = path.join(absoluteOutputDir, filename);
                         let counter = 1;

                         // Handle name collisions
                         while (await fs.access(outputPath).then(() => true).catch(() => false)) {
                             filename = `${baseFilename}_${counter++}.ole`;
                             outputPath = path.join(absoluteOutputDir, filename);
                         }

                         // Attempt to save the OLE object
                         try {
                            oleFormat.Object.SaveAs(outputPath);
                            extractedFiles.push(outputPath);
                            log.info(`[EmbeddedObjects] Extracted Shape ${i} to ${outputPath}`);
                         } catch (saveError: any) {
                            log.warn(`[EmbeddedObjects] Error using SaveAs on Shape ${i} (${progId}): ${saveError.message}. Attempting alternative method (if applicable).`);
                            // Alternative method: If it's an OLE image, try to save the picture
                            if (shape.Type === 7 && shape.OLEFormat?.ProgID?.toLowerCase().includes('package')) {
                               log.warn(`[EmbeddedObjects] Extraction of Package object (Shape ${i}) not implemented directly.`);
                            } else if (shape.Picture) {
                                try {
                                   filename = `${baseFilename}.png`; // Or determine extension
                                   outputPath = path.join(absoluteOutputDir, filename);
                                   counter = 1;
                                   while (await fs.access(outputPath).then(() => true).catch(() => false)) {
                                       filename = `${baseFilename}_${counter++}.png`;
                                       outputPath = path.join(absoluteOutputDir, filename);
                                   }
                                   shape.Picture.SaveAs(outputPath); // wdFormatPNG = 13 (or use constant)
                                   extractedFiles.push(outputPath);
                                   log.info(`[EmbeddedObjects] Extracted image from Shape ${i} to ${outputPath}`);
                                } catch (pictureSaveError: any) {
                                   log.error(`[EmbeddedObjects] Error extracting image from Shape ${i}: ${pictureSaveError.message}`);
                                }
                             } else {
                                log.error(`[EmbeddedObjects] Could not extract Shape ${i} (${progId}) using SaveAs or alternative method.`);
                             }
                          }

                      } else {
                          log.warn(`[EmbeddedObjects] Shape ${i} does not seem to have an OLE object with SaveAs method or is not a directly manageable OLE type.`);
                      }
                  } catch (extractError: any) {
                      log.error(`[EmbeddedObjects] General error extracting Shape ${i}: ${extractError.message}`);
                  } finally {
                      // Release the shape object if necessary
                      releaseObject(shape);
                  }
              }
         }


         if (oleObjectCount === 0) {
           return { success: true, message: 'No embedded or linked OLE objects found in the document.', data: null }; // Add data: null
         } else {
           return { success: true, message: `Found ${oleObjectCount} OLE objects (inline or floating). Extracted files: ${extractedFiles.length}.`, data: { extractedPaths: extractedFiles } }; // Move to data
         }
       }

       // case 'getProperties':
       //   // Logic to get properties
       //   log.warn('Operation "getProperties" not implemented yet.');
       //   return { success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Operation "getProperties" not implemented.' } }; // Return ErrorResponse

       default:
          // The error "Property 'operation' does not exist on type 'never'" indicates that TS
          // has verified that all cases of the discriminated union are covered.
          // Therefore, this 'default' case is theoretically unreachable.
          // Throw a generic error without accessing 'input'.
          // const unreachableCase: never = validatedRequest; // We keep this for type checking
          // log.error(`Unreachable case in switch detected: ${JSON.stringify(unreachableCase)}`);
          throw new Error(`Unknown or unhandled operation.`);
     }
   } catch (error: any) {
     const message = error instanceof Error ? error.message : String(error);
     // Use context.log if available
     const logFn = context?.log?.error || logger.error; // Fallback to global logger if context is not available
     // Access operation and filePath here is safe because they are outside the switch/default
     // Check if validatedRequest exists before accessing its properties in case of a very early error
     const operation = (validatedRequest as any)?.operation || 'unknown'; // Use validatedRequest
     const filePathLog = (validatedRequest as any)?.filePath || 'unknown'; // Use validatedRequest
     logFn(`[EmbeddedObjects] Error in operation '${operation}' on file '${filePathLog}': ${message}`, { error: String(error) }); // Serialize error
     // Return a more informative error message
     return { success: false, error: { code: 'EMBEDDED_OBJECTS_ERROR', message: `Error during operation '${operation}': ${message}`, details: String(error) } }; // Return ErrorResponse
   } finally {
     // --- Improved Finally Block ---
     const logFnDebug = context?.log?.debug || logger.debug;
     const logFnWarn = context?.log?.warn || logger.warn;
     const logFnInfo = context?.log?.info || logger.info;

     if (doc) {
       try {
         // Close without saving changes, especially if it was read-only or there was an error
         // If insert/modify/delete was successful, it should have already been saved.
         // wdDoNotSaveChanges = 0
         doc.Close(0);
         logFnDebug(`[EmbeddedObjects] Document closed: ${absoluteFilePath}`);
       } catch (closeError: any) {
         logFnWarn(`[EmbeddedObjects] Error closing document ${absoluteFilePath}: ${closeError.message}`);
       }
       releaseObject(doc); // Release document object
       doc = null; // Helps GC and prevents double release
     }
     if (wordApp) {
       try {
         // Attempt to close Word only if no other documents are open
         // This is risky if the user has other documents open.
         // A safer option is to simply release the object and let Word close eventually on its own.
         // if (wordApp.Documents.Count === 0) {
         //    wordApp.Quit();
         //    logFnDebug('[EmbeddedObjects] Word application Quit() called.');
         // } else {
         //    logFnDebug('[EmbeddedObjects] Word application has other documents open, not quitting.');
         // }
       } catch (quitError: any) {
          logFnWarn(`[EmbeddedObjects] Error attempting to quit Word: ${quitError.message}`);
       }
       releaseObject(wordApp); // Release application object
       wordApp = null; // Helps GC
     }
     // Access operation and filePath here is safe
     // Check if validatedRequest exists before accessing its properties in case of a very early error
     const operationFinal = (validatedRequest as any)?.operation || 'unknown'; // Use validatedRequest
     const filePathFinal = (validatedRequest as any)?.filePath || 'unknown'; // Use validatedRequest
     logFnInfo(`[EmbeddedObjects] Finished operation '${operationFinal}' on file: ${filePathFinal}`);
   }
 };

 /**
  * McpResource definition for the 'word/embedded-objects' tool.
  * Manages embedded OLE objects in Word documents.
  */
 export const embeddedObjectsTool: McpResource = {
   path: 'word/embedded-objects', // Required path property
   description: 'Manages embedded OLE objects in Word documents (insert, modify, delete, extractAll).',
   // FastMCP expects 'parameters' and 'execute', not 'schema' and 'handler' directly in the object definition.
   // The Zod schema is passed to 'parameters'.
   schema: WordEmbeddedObjectsInputSchema, // Use the new z.object schema and rename to schema
   // outputSchema: z.any(), // Optional: define if needed
   handler: embeddedObjectsExecute, // Rename 'execute' to 'handler'
   // Remove the annotations property
 };

 // Export for index.ts
 export default embeddedObjectsTool;