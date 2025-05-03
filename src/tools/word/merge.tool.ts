import { z } from 'zod';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { validateFilePath } from '../../utils/security';
// Import FastMCPContext and remove ToolContext/ToolRequestParams if not needed elsewhere
import { ApiResponse, McpResource, ToolRequestParams } from '../../types/common.types';
import { saveResource } from '../dynamic/resources.tool'; // Importar saveResource
import { Context as FastMCPContext } from 'fastmcp'; // Import FastMCP Context
import { handleToolError, createErrorResponse } from '../../utils/errorHandler'; // Importar createErrorResponse
import path from 'path';
import logger from '../../utils/logger'; // Keep global logger as fallback if needed

// Remove local MergeToolContext interface, use FastMCPContext directly
// interface MergeToolContext extends ToolContext { ... }

// 1. Define el Schema de Entrada
const mergeSchema = z.object({
    docs: z.array(z.string().min(1)).min(2, { message: 'At least two documents are required for merging.' }),
    output: z.string().min(1).refine(val => val.toLowerCase().endsWith('.docx'), {
        message: 'Output path must end with .docx',
    }),
});

type MergeParams = z.infer<typeof mergeSchema>;

// 2. Implementa el Manejador `mergeDocuments` accepting an optional FastMCPContext
export async function mergeDocuments(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{ outputPath: string }>> {
    // Use context logger if available, otherwise fallback to global logger
    const log = context?.log ?? logger;
    const reportProgress = context?.reportProgress; // Get reportProgress function if context exists

    // Validate context has necessary functions if context is provided
    if (context && (!context.log || typeof context.reportProgress !== 'function')) {
        // Log error using the determined logger
        log.error('Merge tool received context but it is missing required properties (log, reportProgress).');
        // Avoid returning ApiResponse directly, throw error for FastMCP handler
        throw new Error('Tool context is missing required properties (log, reportProgress).');
        // return createErrorResponse('CONFIGURATION_ERROR', 'Tool context is missing required properties (log, reportProgress).');
    }

    let wordApp: any = null;
    let targetDoc: any = null;
    const sourceDocs: any[] = []; // Para llevar registro de los documentos fuente abiertos
    let officeAppInstance: any = null; // Declarar officeAppInstance fuera del try

    try {
        // Validar y parsear parámetros
        const validatedParams = mergeSchema.parse(params);
        log.info(`Validating input parameters for merge operation.`);

        // Validar rutas de archivo
        const safeOutputPath = validateFilePath(validatedParams.output);
        // validateFilePath now handles the allowed paths logic internally based on env vars.
        // No need to pass allowedPaths or perform additional checks here.

        const safeSourcePaths: string[] = [];
        for (const docPath of validatedParams.docs) {
            const safePath = validateFilePath(docPath);
            if (!safePath) {
                // validateFilePath should throw an error if validation fails,
                // so this check might be redundant, but kept for safety.
                throw new Error(`Source path validation failed for: ${docPath}`);
            }
            safeSourcePaths.push(safePath);
        }
        log.info(`All file paths validated successfully.`);

        // Obtener instancia de Word
        log.info('Getting Word application instance...');
        officeAppInstance = await getOfficeApplication('Word.Application'); // Asignar aquí
        wordApp = officeAppInstance.app;
        wordApp.Visible = false; // Ejecutar en segundo plano
        wordApp.DisplayAlerts = 0; // wdAlertsNone = 0

        // Crear documento destino
        log.info('Creating target document...');
        targetDoc = wordApp.Documents.Add();

        // Iterar y combinar documentos
        const totalDocs = safeSourcePaths.length;
        log.info(`Starting merge process for ${totalDocs} documents...`);
        // Using correct { progress, total } signature from documentation, only if reportProgress is available
        reportProgress?.({ progress: 0, total: totalDocs });

        for (let i = 0; i < totalDocs; i++) {
            const sourcePath = safeSourcePaths[i];
            const currentDocNum = i + 1;
            log.info(`Processing document ${currentDocNum}/${totalDocs}: ${sourcePath}`);
            // Using correct { progress, total } signature, only if reportProgress is available
            reportProgress?.({ progress: i, total: totalDocs });

            let sourceDoc: any = null;
            try {
                sourceDoc = wordApp.Documents.Open(sourcePath, false, true); // Open(FileName, ConfirmConversions=false, ReadOnly=true)
                sourceDocs.push(sourceDoc); // Añadir a la lista para liberación posterior

                // Seleccionar y copiar contenido
                // sourceDoc.Content.Select(); // Select puede ser problemático a veces, usar Range
                const sourceRange = sourceDoc.Content;
                sourceRange.Copy();


                // Pegar en el documento destino
                targetDoc.Activate();
                // Mover al final del documento antes de pegar
                const endOfDocRange = targetDoc.Content;
                endOfDocRange.Collapse(0); // 0 = wdCollapseEnd
                // wordApp.Selection.SetRange(endOfDocRange.End, endOfDocRange.End); // Usar Range.Paste en lugar de Selection
                endOfDocRange.Paste();


                // Insertar salto de página después de cada documento excepto el último
                if (i < totalDocs - 1) {
                    // wordApp.Selection.InsertBreak(7); // 7 = wdPageBreak - Usar Range
                    const breakRange = targetDoc.Content;
                    breakRange.Collapse(0); // Collapse to end
                    breakRange.InsertBreak(7); // Insert page break at the end
                }

                // Cerrar documento fuente (sin guardar cambios)
                sourceDoc.Close(false); // wdDoNotSaveChanges = 0
                // Quitar de la lista de documentos abiertos ya que se cerró correctamente
                const indexToRemove = sourceDocs.indexOf(sourceDoc);
                if (indexToRemove > -1) {
                    sourceDocs.splice(indexToRemove, 1);
                }
                releaseObject(sourceDoc); // Liberar el objeto COM del documento fuente cerrado
                sourceDoc = null; // Asegurarse de que la variable está limpia
                 log.info(`Document ${currentDocNum}/${totalDocs} processed and closed.`);
                 // Using correct { progress, total } signature, only if reportProgress is available
                 reportProgress?.({ progress: currentDocNum, total: totalDocs });

            } catch (sourceDocError: any) {
                 log.error(`Error processing source document ${sourcePath}: ${sourceDocError.message || sourceDocError}`);
                 // Intentar cerrar el documento fuente si aún está abierto antes de relanzar
                 if (sourceDoc) {
                     try {
                         sourceDoc.Close(false);
                         const indexToRemove = sourceDocs.indexOf(sourceDoc);
                         if (indexToRemove > -1) {
                             sourceDocs.splice(indexToRemove, 1);
                         }
                         releaseObject(sourceDoc);
                     } catch (closeError: any) {
                         log.error(`Failed to close source document ${sourcePath} after error: ${closeError.message || closeError}`);
                     }
                 }
                throw sourceDocError; // Relanzar el error para que sea capturado por el catch principal
            }
        }

        // Guardar el documento combinado
        log.info(`Saving merged document to: ${safeOutputPath}`);
        // Using correct { progress, total } signature - representing the saving step, only if reportProgress is available
        reportProgress?.({ progress: totalDocs, total: totalDocs });
        // Asegurarse de que el directorio existe antes de guardar
        // Nota: COM podría manejar esto, pero ser explícito es más seguro si es posible.
        // Sin embargo, crear directorios desde aquí podría requerir permisos adicionales
        // o lógica fuera del alcance de officeInterop. Se asume que el directorio base permitido existe.
        targetDoc.SaveAs2(safeOutputPath);
        log.info(`Merged document saved successfully.`);

        // Leer el contenido del documento combinado antes de cerrarlo
        let mergedContent = '';
        try {
             // Leer el texto del documento COM object
             mergedContent = targetDoc.Content.Text;
             log.info(`Read content from merged document.`);
         } catch (readContentError: any) {
             log.error(`Failed to read content from merged document before closing: ${readContentError.message}`);
             // No lanzar error aquí, intentar guardar el recurso vacío o con error
         }

        // Cerrar el documento destino
        targetDoc.Close(false); // wdDoNotSaveChanges = 0
        // Liberar explícitamente el objeto targetDoc ahora que está cerrado y guardado
        releaseObject(targetDoc);
        targetDoc = null;

        // Guardar el documento combinado como un recurso dinámico después de cerrarlo
        try {
            await saveResource('word/merge', path.basename(safeOutputPath), mergedContent);
            log.info(`Saved ${safeOutputPath} as a dynamic resource.`);
        } catch (resourceSaveError: any) {
            log.error(`Failed to save ${safeOutputPath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continuar la ejecución aunque falle el guardado del recurso
        }


        // Using correct { progress, total } signature - representing completion, only if reportProgress is available
        reportProgress?.({ progress: totalDocs, total: totalDocs });

        return {
            success: true,
            data: { outputPath: safeOutputPath },
            message: `Successfully merged ${totalDocs} documents into ${safeOutputPath}.`
        };

    } catch (error: any) {
        // Use context.log for error logging
        log.error(`Error in mergeDocuments: ${error.message || error}`, { error }); // Log the full error object
        // Asegurarse de cerrar el documento destino si se creó y no se cerró/liberó antes
        if (targetDoc) {
            try {
                targetDoc.Close(false);
            } catch (closeError: any) {
                 log.error(`Failed to close target document after error: ${closeError.message || closeError}`);
            }
        }
        // Let the main execute wrapper in server/index.ts handle the error conversion
        throw error;
        // return handleToolError(error); // Avoid returning ApiResponse directly
    } finally {
        // Use context.log for final logging
        log.info('Starting cleanup process...');
        // Liberar documentos fuente que pudieran haber quedado abiertos por error
        if (sourceDocs.length > 0) {
            log.warn(`Releasing ${sourceDocs.length} potentially orphaned source document objects.`);
            for (const doc of sourceDocs) {
                 try {
                     doc.Close(false); // Intentar cerrar por si acaso
                 } catch (e) { /* Ignorar errores al cerrar aquí */ }
                releaseObject(doc);
            }
        }
        // Liberar documento destino si aún existe referencia (p.ej., si hubo error antes de liberarlo explícitamente)
        if (targetDoc) {
            releaseObject(targetDoc);
             log.info('Target document object released.');
        }
        // Liberar aplicación Word
         if (wordApp) {
              // No llamar a Quit() directamente aquí, confiar en releaseObject y officeInterop
              releaseObject(wordApp);
              log.info('Word application object reference released.');
         }
         // Liberar officeAppInstance
         if (officeAppInstance) {
             officeAppInstance.release();
             log.info('Office application instance released.');
         }
    }
}

// 3. Define y Exporta el Recurso
export const wordMergeTool: McpResource[] = [
    {
        path: 'word/merge',
        // Handler signature matches McpResource expectation (params, context)
        handler: mergeDocuments,
        schema: mergeSchema,
        description: 'Merges multiple Word documents (.docx) into a single new document using COM Interop.',
        // completions: ... // Puedes añadir completions si lo ves útil
    },
];