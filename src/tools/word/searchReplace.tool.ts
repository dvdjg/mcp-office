import { z } from 'zod';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { validateFilePath } from '../../utils/security';
import { handleToolError } from '../../utils/errorHandler';
import { saveResource } from '../dynamic/resources.tool'; // Importar saveResource
// Import FastMCPContext and remove ToolContext
import type {
    ApiResponse, // Type alias for SuccessResponse | ErrorResponse
    SuccessResponse, // Specific type for success
    ErrorResponse, // Specific type for error
    McpResource,
    // ToolContext, // Removed
    ToolRequestParams,
} from '../../types/common.types';
import { Context as FastMCPContext } from 'fastmcp'; // Import FastMCP Context
import path from 'path';
import logger from '../../utils/logger'; // Import global logger

// Define Schemas de Entrada
const searchReplaceSchema = z.object({
    filePath: z.string().min(1, 'File path cannot be empty.'),
    find: z.string().min(1, 'Search text cannot be empty.'),
    replace: z.string().default(''),
    matchCase: z.boolean().optional().default(false),
    matchWholeWord: z.boolean().optional().default(false),
    useWildcards: z.boolean().optional().default(false),
    replaceAll: z.boolean().optional().default(true),
});

type SearchReplaceParams = z.infer<typeof searchReplaceSchema>;

// Implementa el Manejador searchAndReplace accepting an optional FastMCPContext
async function searchAndReplace(
    params: ToolRequestParams,
    context?: FastMCPContext<undefined>
): Promise<ApiResponse<{ replacementsMade: boolean }>> {
    // Use context logger if available, otherwise fallback to global logger
    const log = context?.log ?? logger;
    const reportProgress = context?.reportProgress; // Get reportProgress function if context exists
    const totalSteps = 4; // Define total steps for progress

    let wordApp: any = null;
    let doc: any = null;
    let findObject: any = null;
    let replacementObject: any = null;
    let replacementsMade: boolean = false;

    const filePathParam = (params as any)?.filePath || 'unknown'; // For logging in case of early error

    try {
        reportProgress?.({ progress: 0, total: totalSteps }); // Step 0: Start
        // Valida los params
        const validatedParams = searchReplaceSchema.parse(params);

        // Valida la ruta del documento
        // Note: validateFilePath might need adjustment if it relies on allowedPaths from context
        const safeFilePath = await validateFilePath(validatedParams.filePath);
        log.info(`Validated file path: ${safeFilePath}`);

        // Obtén la instancia de Word
        log.info('Getting Word application instance...');
        wordApp = await getOfficeApplication('Word.Application');
        wordApp.Visible = false; // Keep Word hidden

        // Abre el documento
        log.info(`Opening document: ${safeFilePath}`);
        doc = await wordApp.Documents.Open(safeFilePath);
        reportProgress?.({ progress: 1, total: totalSteps }); // Step 1: Document Opened

        // Accede al objeto Find y Replacement
        findObject = await doc.Content.Find;
        replacementObject = await findObject.Replacement;

        // Limpia formato previo
        await findObject.ClearFormatting();
        await replacementObject.ClearFormatting();

        // Configura propiedades de búsqueda y reemplazo
        log.info(`Configuring search for "${validatedParams.find}" and replace with "${validatedParams.replace}"`);
        findObject.Text = validatedParams.find;
        replacementObject.Text = validatedParams.replace;
        findObject.MatchCase = validatedParams.matchCase;
        findObject.MatchWholeWord = validatedParams.matchWholeWord;
        findObject.MatchWildcards = validatedParams.useWildcards;
        findObject.Forward = true;
        findObject.Wrap = 1; // wdFindContinue
        reportProgress?.({ progress: 2, total: totalSteps }); // Step 2: Configured

        // Ejecuta la operación
        const replaceOption = validatedParams.replaceAll ? 2 : 1; // wdReplaceAll = 2, wdReplaceOne = 1
        log.info(`Executing find/replace (replaceAll: ${validatedParams.replaceAll})...`);

        replacementsMade = await findObject.Execute(
            undefined, // FindText
            validatedParams.matchCase,
            validatedParams.matchWholeWord,
            validatedParams.useWildcards,
            undefined, // MatchSoundsLike
            undefined, // MatchAllWordForms
            true,      // Forward
            1,         // Wrap
            undefined, // Format
            undefined, // ReplaceWith (set on replacementObject)
            replaceOption // Replace
        );
        log.info(`Find/replace executed. Replacements made: ${replacementsMade}`);
        reportProgress?.({ progress: 3, total: totalSteps }); // Step 3: Executed

        // Guarda si hubo éxito
        if (replacementsMade) {
            log.info(`Saving document: ${safeFilePath}`);
            await doc.Save();
            log.info(`Document saved.`);

            // Leer el contenido del documento modificado antes de cerrarlo
            let modifiedContent = '';
            try {
                 // Leer el texto del documento COM object
                 modifiedContent = doc.Content.Text;
                 log.info(`Read content from modified document.`);
             } catch (readContentError: any) {
                 log.error(`Failed to read content from modified document before closing: ${readContentError.message}`);
                 // No lanzar error aquí, intentar guardar el recurso vacío o con error
             }

            // Guardar el documento modificado como un recurso dinámico después de cerrarlo
            try {
                await saveResource('word/search-replace', path.basename(safeFilePath), modifiedContent);
                log.info(`Saved ${safeFilePath} as a dynamic resource.`);
            } catch (resourceSaveError: any) {
                log.error(`Failed to save ${safeFilePath} as a dynamic resource: ${resourceSaveError.message}`);
                // Continuar la ejecución aunque falle el guardado del recurso
            }

        } else {
            log.info(`No replacements made, document not saved.`);
        }

        // Construye la respuesta de éxito manualmente
        const successResponse: SuccessResponse<{ replacementsMade: boolean }> = {
            success: true,
            message: `Search and replace operation completed on ${path.basename(
                safeFilePath
            )}. Replacements made: ${replacementsMade}`, // Adjusted message
            data: { replacementsMade },
        };
        reportProgress?.({ progress: 4, total: totalSteps }); // Step 4: Complete
        return successResponse;

    } catch (error: unknown) {
        // Maneja errores y devuelve ErrorResponse
        log.error(`Error during search/replace on ${filePathParam}: ${String(error)}`, { error: String(error) });
        throw error; // Let the main handler manage the error response
        // return handleToolError(error, 'WORD_SEARCH_REPLACE_ERROR');
    } finally {
        // Libera objetos COM en orden inverso de creación/obtención
        log.debug('Starting COM object cleanup for search/replace...');
        if (replacementObject) {
            await releaseObject(replacementObject);
        }
        if (findObject) {
            await releaseObject(findObject);
        }
        if (doc) {
            try {
                await doc.Close(false); // No guardar cambios al cerrar
            } catch (closeError: unknown) {
                 log.warn(`Non-critical error closing document: ${String(closeError)}`);
            }
            await releaseObject(doc);
        }
        if (wordApp) {
            try {
                // Check if Word is still running and has no other docs open before quitting
                // Ensure wordApp is checked for existence before accessing properties/methods
                if (wordApp && typeof wordApp.Documents !== 'undefined' && await wordApp.Documents.Count === 0) {
                    log.debug("Attempting to quit Word application as no documents are open.");
                    await wordApp.Quit();
                } else if (wordApp && typeof wordApp.Documents !== 'undefined') {
                    log.debug(`Word application not quit (${await wordApp.Documents.Count} docs open). Releasing object.`);
                } else if (wordApp) {
                    log.debug("Word application object exists but Documents property is inaccessible. Releasing object.");
                }
            } catch (quitError: unknown) {
                 log.warn(`Non-critical error quitting Word: ${String(quitError)}`);
            }
            await releaseObject(wordApp); // Release reference regardless of quit attempt
        }
        log.debug('COM object cleanup finished for search/replace.');
    }
}

// Define y Exporta el Recurso
export const wordSearchReplaceTool: McpResource[] = [
    {
        path: 'word/search-replace',
        handler: searchAndReplace, // Correct handler signature
        schema: searchReplaceSchema,
        description:
            'Searches for text in a Word document and replaces it using COM Interop. Supports options like match case, whole word, wildcards, and replace all.',
        // examples property removed
    },
];