import { z } from 'zod';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { validateFilePath } from '../../utils/security';
import { handleToolError } from '../../utils/errorHandler';
import type {
    ApiResponse, // Type alias for SuccessResponse | ErrorResponse
    SuccessResponse, // Specific type for success
    ErrorResponse, // Specific type for error
    McpResource,
    ToolContext,
    ToolRequestParams,
} from '../../types/common.types';
import path from 'path';

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

// Implementa el Manejador searchAndReplace
async function searchAndReplace(
    params: ToolRequestParams,
    context?: ToolContext // Context is optional, logger removed
): Promise<ApiResponse<{ replacementsMade: boolean }>> {
    let wordApp: any = null;
    let doc: any = null;
    let findObject: any = null;
    let replacementObject: any = null;
    let replacementsMade: boolean = false;

    const filePathParam = (params as any)?.filePath || 'unknown'; // For logging in case of early error

    try {
        // Valida los params
        const validatedParams = searchReplaceSchema.parse(params);

        // Valida la ruta del documento
        const safeFilePath = await validateFilePath(validatedParams.filePath);

        // Obtén la instancia de Word
        wordApp = await getOfficeApplication('Word.Application');
        wordApp.Visible = false; // Keep Word hidden

        // Abre el documento
        doc = await wordApp.Documents.Open(safeFilePath);

        // Accede al objeto Find y Replacement
        findObject = await doc.Content.Find;
        replacementObject = await findObject.Replacement;

        // Limpia formato previo
        await findObject.ClearFormatting();
        await replacementObject.ClearFormatting();

        // Configura propiedades de búsqueda y reemplazo
        findObject.Text = validatedParams.find;
        replacementObject.Text = validatedParams.replace;
        findObject.MatchCase = validatedParams.matchCase;
        findObject.MatchWholeWord = validatedParams.matchWholeWord;
        findObject.MatchWildcards = validatedParams.useWildcards;
        findObject.Forward = true;
        findObject.Wrap = 1; // wdFindContinue

        // Ejecuta la operación
        const replaceOption = validatedParams.replaceAll ? 2 : 1; // wdReplaceAll = 2, wdReplaceOne = 1

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

        // Guarda si hubo éxito
        if (replacementsMade) {
            await doc.Save();
        }

        // Construye la respuesta de éxito manualmente
        const successResponse: SuccessResponse<{ replacementsMade: boolean }> = {
            success: true,
            message: `Search and replace operation completed on ${path.basename(
                safeFilePath
            )}. Success status: ${replacementsMade}`,
            data: { replacementsMade },
        };
        return successResponse;

    } catch (error: unknown) {
        // Maneja errores y devuelve ErrorResponse
        // Puedes pasar un código específico si lo deseas, p.ej., 'WORD_SEARCH_REPLACE_ERROR'
        return handleToolError(error, 'WORD_TOOL_ERROR');
    } finally {
        // Libera objetos COM en orden inverso de creación/obtención
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
                 // Log warning if logger was available, otherwise ignore non-critical error
                 console.warn(`Non-critical error closing document: ${closeError}`);
            }
            await releaseObject(doc);
        }
        if (wordApp) {
            try {
                if (await wordApp.Documents.Count === 0) {
                    await wordApp.Quit();
                }
            } catch (quitError: unknown) {
                 // Log warning if logger was available, otherwise ignore non-critical error
                 console.warn(`Non-critical error quitting Word: ${quitError}`);
            }
            await releaseObject(wordApp);
        }
    }
}

// Define y Exporta el Recurso
export const wordSearchReplaceTool: McpResource[] = [
    {
        path: 'word/search-replace',
        handler: searchAndReplace,
        schema: searchReplaceSchema,
        description:
            'Searches for text in a Word document and replaces it using COM Interop. Supports options like match case, whole word, wildcards, and replace all.',
        // examples property removed
    },
];