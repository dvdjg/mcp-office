import { z } from 'zod';
import { McpResource, ApiResponse, FastMCPContext } from '../../types/common.types'; // Import FastMCPContext, remove ToolContext
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { validateFilePath } from '../../utils/security';
import { createErrorResponse, handleToolError } from '../../utils/errorHandler'; // Added createErrorResponse
import logger from '../../utils/logger'; // Corrected import style

// 1. Define el Schema de Entrada
const analyzeSchema = z.object({
    filePath: z.string().min(1, 'File path cannot be empty.'),
});

// Mapeo de tipos de Content Control (WdContentControlType) a strings descriptivos
const ContentControlTypeMap: { [key: number]: string } = {
    0: 'RichText',
    1: 'PlainText',
    2: 'Picture',
    3: 'ComboBox',
    4: 'DropDownList',
    5: 'BuildingBlockGallery',
    6: 'Date',
    7: 'Group',
    8: 'CheckBox',
    9: 'RepeatingSection',
    // Añadir más si es necesario
};

// 2. Implementa el Manejador `analyzeTemplate`
async function analyzeTemplate(
    params: unknown,
    context?: FastMCPContext<undefined> // Use FastMCPContext<undefined>
): Promise<ApiResponse<Array<object>>> {
    let wordApp: any = null;
    let doc: any = null;
    let contentControls: any = null;
    let cc: any = null; // Para liberar en finally si hay error en el bucle

    try {
        // Validar parámetros
        const validatedParams = analyzeSchema.parse(params);
        const { filePath } = validatedParams;
        logger.info(`[word/template/analyze] Analyzing file: ${filePath}`);

        // Validar ruta de archivo - Let validateFilePath use its default ALLOWED_BASE_PATHS
        // It will resolve the relative filePath based on process.cwd() and check against allowed bases.
        let safeFilePath: string;
        try {
             safeFilePath = validateFilePath(filePath); // Pass only the filePath
        } catch (validationError: any) {
             // If validation fails (e.g., outside allowed paths), return an error
             logger.warn(`Path validation failed for "${filePath}": ${validationError.message}`);
             return createErrorResponse('INVALID_PATH', `Path validation failed: ${validationError.message}`);
        }
        // safeFilePath is guaranteed to be a string here if no error was thrown.

        const foundControls: Array<object> = [];

        // Obtener instancia de Word
        wordApp = await getOfficeApplication('Word.Application');
        if (!wordApp) {
             // Use createErrorResponse for specific errors
             return createErrorResponse('COM_ERROR', 'Could not get Word application instance.');
        }

        // Abrir documento (solo lectura)
        logger.debug(`[word/template/analyze] Opening document: ${safeFilePath}`);
        doc = wordApp.Documents.Open(safeFilePath, false, true); // ReadOnly = true
        if (!doc) {
             // Use createErrorResponse for specific errors
            return createErrorResponse('FILE_OPERATION_ERROR', `Could not open document: ${safeFilePath}`);
        }

        // Acceder a Content Controls
        contentControls = doc.ContentControls;
        const count = contentControls.Count;
        logger.debug(`[word/template/analyze] Found ${count} content controls.`);

        // Iterar sobre los controles
        for (let i = 1; i <= count; i++) {
            cc = contentControls(i); // 1-based index
            if (!cc) continue; // Seguridad extra

            try {
                const controlInfo: any = {
                    id: cc.ID,
                    title: cc.Title || null, // Puede ser null o vacío
                    tag: cc.Tag || null,     // Puede ser null o vacío
                    type: ContentControlTypeMap[cc.Type] || `Unknown (${cc.Type})`,
                    text: cc.Range?.Text || '', // Texto actual
                };

                // PlaceholderText no existe en todos los tipos (ej. Group)
                try {
                    controlInfo.placeholderText = cc.PlaceholderText || null;
                } catch (placeholderError) {
                    controlInfo.placeholderText = null; // O un valor indicativo
                    logger.warn(`[word/template/analyze] Could not get PlaceholderText for CC ID ${cc.ID} (Type: ${controlInfo.type}). Error: ${placeholderError}`);
                }

                foundControls.push(controlInfo);
                logger.debug(`[word/template/analyze] Processed CC: ID=${controlInfo.id}, Title=${controlInfo.title}, Tag=${controlInfo.tag}, Type=${controlInfo.type}`);

            } catch (loopError) {
                 logger.error(`[word/template/analyze] Error processing content control index ${i}: ${loopError}`);
                 // Continuar con el siguiente control si es posible
            } finally {
                 if (cc) {
                    releaseObject(cc);
                    cc = null; // Resetear para la siguiente iteración o el finally principal
                 }
            }
        }

        logger.info(`[word/template/analyze] Successfully analyzed ${foundControls.length} content controls.`);
        return { success: true, data: foundControls };

    } catch (error: any) {
        logger.error(`[word/template/analyze] Error: ${error.message}`, { error });
        // Use handleToolError for caught errors, it handles ZodError internally
        return handleToolError(error, 'TOOL_EXECUTION_ERROR');
    } finally {
        // Asegurar la liberación de todos los objetos COM
        logger.debug('[word/template/analyze] Releasing COM objects...');
        if (cc) releaseObject(cc); // Si salió del bucle por error
        if (contentControls) releaseObject(contentControls);
        if (doc) {
            doc.Close(0); // wdDoNotSaveChanges = 0
            releaseObject(doc);
        }
        // No cerramos la aplicación aquí, getOfficeApplication la gestiona
        // if (wordApp) {
        //     // Considerar si cerrar la app o dejarla gestionada por officeInterop
        //     // wordApp.Quit(); // Podría cerrar Word si el usuario lo está usando
        //     releaseObject(wordApp);
        // }
        logger.debug('[word/template/analyze] COM objects released.');
    }
}

// 3. Define y Exporta el Recurso
export const wordTemplateTool: McpResource[] = [
    {
        path: 'word/template/analyze',
        handler: analyzeTemplate,
        schema: analyzeSchema,
        description: 'Analyzes a Word document and lists its Content Controls (placeholders) using COM Interop.'
        // Removed category: 'word'
        // Removed examples
    },
    // Placeholder for replace operation
    {
        path: 'word/template/replace',
        handler: async () => ({
            success: false,
            error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/template/replace not implemented.' }
        }),
        description: 'Replaces Content Controls in a Word document (Not Implemented).',
        // Removed category: 'word'
        // Removed examples
        schema: z.object({ // Placeholder schema - Kept schema for replace tool
            filePath: z.string().min(1),
            replacements: z.array(z.object({
                identifier: z.string().min(1), // Could be ID, Tag, or Title
                value: z.string()
            }))
        })
    }
];