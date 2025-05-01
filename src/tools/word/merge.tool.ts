import { z } from 'zod';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { validateFilePath } from '../../utils/security';
import { ApiResponse, McpResource, ToolContext, ToolRequestParams } from '../../types/common.types';
import { handleToolError, createErrorResponse } from '../../utils/errorHandler'; // Importar createErrorResponse
import path from 'path';
import logger from '../../utils/logger'; // Importar logger global

// Interfaz local para el contexto esperado por esta herramienta específica
interface MergeToolContext extends ToolContext {
    logger: typeof logger; // Asumiendo que logger tiene un tipo exportado o usar 'any' si no
    config: {
        allowedPaths: string[];
        // otras propiedades de config...
    };
}

// 1. Define el Schema de Entrada
const mergeSchema = z.object({
    docs: z.array(z.string().min(1)).min(2, { message: 'At least two documents are required for merging.' }),
    output: z.string().min(1).refine(val => val.toLowerCase().endsWith('.docx'), {
        message: 'Output path must end with .docx',
    }),
});

type MergeParams = z.infer<typeof mergeSchema>;

// 2. Implementa el Manejador `mergeDocuments`
async function mergeDocuments(params: ToolRequestParams, context?: ToolContext): Promise<ApiResponse<{ outputPath: string }>> {
    // Guarda inicial para el contexto y configuración necesaria
    // Usamos 'as any' temporalmente para la comprobación inicial
    if (!context || !(context as any).config || !(context as any).config.allowedPaths || !(context as any).logger) {
        logger.error('Merge tool requires context with logger, config, and allowedPaths.');
        return createErrorResponse('CONFIGURATION_ERROR', 'Tool context is missing required properties (logger, config.allowedPaths).', {
            contextProvided: !!context,
            hasConfig: !!(context as any)?.config,
            hasAllowedPaths: !!(context as any)?.config?.allowedPaths,
            hasLogger: !!(context as any)?.logger,
        });
    }
    // Aserción de tipo: ahora sabemos que context tiene la estructura de MergeToolContext
    const fullContext = context as MergeToolContext;
    const currentLogger = fullContext.logger;
    const allowedPaths = fullContext.config.allowedPaths;


    let wordApp: any = null;
    let targetDoc: any = null;
    const sourceDocs: any[] = []; // Para llevar registro de los documentos fuente abiertos

    try {
        // Validar y parsear parámetros
        const validatedParams = mergeSchema.parse(params);
        currentLogger.info(`Validating input parameters for merge operation.`);

        // Validar rutas de archivo
        const safeOutputPath = validateFilePath(validatedParams.output, allowedPaths);
        if (!safeOutputPath) {
            throw new Error(`Output path validation failed for: ${validatedParams.output}`);
        }
        // Asegurarse de que el directorio de salida existe y está permitido
        const outputDir = path.dirname(safeOutputPath);
        // First check if the exact directory is allowed (might be explicitly listed)
        if (!validateFilePath(outputDir, allowedPaths)) {
             // If not explicitly listed, check if it's within an allowed root path
             const isAllowed = allowedPaths.some((allowedRoot: string) => outputDir.startsWith(allowedRoot));
             if (!isAllowed) {
                throw new Error(`Output directory validation failed for: ${outputDir}. Not within allowed paths.`);
             }
             // If it's within an allowed path, we assume it's okay to proceed (COM might create dirs or fail)
             currentLogger.warn(`Output directory ${outputDir} is not explicitly listed but is within an allowed root. Proceeding.`);
        }


        const safeSourcePaths: string[] = [];
        for (const docPath of validatedParams.docs) {
            const safePath = validateFilePath(docPath, allowedPaths);
            if (!safePath) {
                throw new Error(`Source path validation failed for: ${docPath}`);
            }
            safeSourcePaths.push(safePath);
        }
        currentLogger.info(`All file paths validated successfully.`);

        // Obtener instancia de Word
        currentLogger.info('Getting Word application instance...');
        wordApp = await getOfficeApplication('Word.Application');
        wordApp.Visible = false; // Ejecutar en segundo plano
        wordApp.DisplayAlerts = 0; // wdAlertsNone = 0

        // Crear documento destino
        currentLogger.info('Creating target document...');
        targetDoc = wordApp.Documents.Add();

        // Iterar y combinar documentos
        currentLogger.info(`Starting merge process for ${safeSourcePaths.length} documents...`);
        for (let i = 0; i < safeSourcePaths.length; i++) {
            const sourcePath = safeSourcePaths[i];
            currentLogger.info(`Processing document ${i + 1}: ${sourcePath}`);
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
                if (i < safeSourcePaths.length - 1) {
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
                 currentLogger.info(`Document ${i + 1} processed and closed.`);

            } catch (sourceDocError: any) {
                 currentLogger.error(`Error processing source document ${sourcePath}: ${sourceDocError.message || sourceDocError}`);
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
                         currentLogger.error(`Failed to close source document ${sourcePath} after error: ${closeError.message || closeError}`);
                     }
                 }
                throw sourceDocError; // Relanzar el error para que sea capturado por el catch principal
            }
        }

        // Guardar el documento combinado
        currentLogger.info(`Saving merged document to: ${safeOutputPath}`);
        // Asegurarse de que el directorio existe antes de guardar
        // Nota: COM podría manejar esto, pero ser explícito es más seguro si es posible.
        // Sin embargo, crear directorios desde aquí podría requerir permisos adicionales
        // o lógica fuera del alcance de officeInterop. Se asume que el directorio base permitido existe.
        targetDoc.SaveAs2(safeOutputPath);
        currentLogger.info(`Merged document saved successfully.`);

        // Cerrar el documento destino
        targetDoc.Close(false); // wdDoNotSaveChanges = 0
        // Liberar explícitamente el objeto targetDoc ahora que está cerrado y guardado
        releaseObject(targetDoc);
        targetDoc = null;


        return {
            success: true,
            data: { outputPath: safeOutputPath },
            message: `Successfully merged ${safeSourcePaths.length} documents into ${safeOutputPath}.`
        };

    } catch (error: any) {
        // Usar el logger validado o el global si el contexto no estaba disponible (aunque la guarda inicial debería prevenir esto)
        const errorLogger = (context as MergeToolContext)?.logger || logger;
        errorLogger.error(`Error in mergeDocuments: ${error.message || error}`);
        // Asegurarse de cerrar el documento destino si se creó y no se cerró/liberó antes
        if (targetDoc) {
            try {
                targetDoc.Close(false);
            } catch (closeError: any) {
                 errorLogger.error(`Failed to close target document after error: ${closeError.message || closeError}`);
            }
        }
        // Pasar solo el error a handleToolError
        return handleToolError(error);
    } finally {
        // Usar el logger validado o el global
        const finalLogger = (context as MergeToolContext)?.logger || logger;
        finalLogger.info('Starting cleanup process...');
        // Liberar documentos fuente que pudieran haber quedado abiertos por error
        if (sourceDocs.length > 0) {
            finalLogger.warn(`Releasing ${sourceDocs.length} potentially orphaned source document objects.`);
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
             finalLogger.info('Target document object released.');
        }
        // Liberar aplicación Word
         if (wordApp) {
             // No llamar a Quit() directamente aquí, confiar en releaseObject y officeInterop
             releaseObject(wordApp);
             finalLogger.info('Word application object reference released.');
         }
        finalLogger.info('Cleanup process finished.');
    }
}

// 3. Define y Exporta el Recurso
export const wordMergeTool: McpResource[] = [
    {
        path: 'word/merge',
        // La firma de mergeDocuments ahora coincide con la requerida
        handler: mergeDocuments,
        schema: mergeSchema,
        description: 'Merges multiple Word documents (.docx) into a single new document using COM Interop.',
        // completions: ... // Puedes añadir completions si lo ves útil
    },
];