import { z } from "zod";
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { handleToolError, createErrorResponse } from '../../utils/errorHandler';
import { validateFilePath } from '../../utils/security';
import logger from '../../utils/logger'; // Importación por defecto

// --- Constantes COM (Ejemplos, verificar en documentación de Word) ---
// https://learn.microsoft.com/en-us/office/vba/api/word.wdheaderfootertype
const WdHeaderFooterTypes = {
  wdHeaderFooterPrimary: 1,
  wdHeaderFooterFirstPage: 2,
  wdHeaderFooterEvenPages: 3,
} as const;

// --- Schemas ---

// Esquema base común
const HeadersFootersBaseSchema = z.object({
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  sectionIndex: z.union([z.number().int().positive(), z.literal('all')]).optional().default(1).describe("1-based index of the section or 'all'. Defaults to 1."),
  type: z.nativeEnum(WdHeaderFooterTypes).describe("Type of header/footer (Primary, FirstPage, EvenPages)."),
});

// Esquema para Insert/Modify (requiere texto)
const HeadersFootersTextSchema = HeadersFootersBaseSchema.extend({
  text: z.string().describe("Text content to insert or modify."),
});

// Esquema para Delete/Get (no requiere texto adicional)
const HeadersFootersActionSchema = HeadersFootersBaseSchema;

// Esquema para Configure (propiedades booleanas)
const HeadersFootersConfigureSchema = z.object({
   filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  sectionIndex: z.union([z.number().int().positive(), z.literal('all')]).optional().default(1).describe("1-based index of the section or 'all'. Defaults to 1."),
  linkToPrevious: z.boolean().optional().describe("Link header/footer to the previous section."),
  // Propiedades de PageSetup que afectan a headers/footers (ya en word/page, pero pueden ser relevantes aquí)
  differentFirstPage: z.boolean().optional().describe("Different header/footer on the first page for this section."),
  oddAndEvenPages: z.boolean().optional().describe("Different headers/footers for odd and even pages for this section."),
}).refine(data => data.linkToPrevious !== undefined || data.differentFirstPage !== undefined || data.oddAndEvenPages !== undefined, {
    message: "At least one configuration option (linkToPrevious, differentFirstPage, oddAndEvenPages) must be provided for 'configure'.",
    path: [],
});


// Combinar esquemas para el handler usando discriminatedUnion
const InputSchema = z.discriminatedUnion('operation', [
  z.object({ operation: z.literal('insert'), arguments: HeadersFootersTextSchema }),
  z.object({ operation: z.literal('modify'), arguments: HeadersFootersTextSchema }),
  z.object({ operation: z.literal('delete'), arguments: HeadersFootersActionSchema }),
  z.object({ operation: z.literal('get'), arguments: HeadersFootersActionSchema }),
  z.object({ operation: z.literal('configure'), arguments: HeadersFootersConfigureSchema }),
]);
type InputSchemaType = z.infer<typeof InputSchema>;


// --- Lógica COM (Placeholder) ---

async function manageHeaderFooter(
    operation: 'insert' | 'modify' | 'delete' | 'get' | 'configure',
    args: any // Usar 'any' temporalmente, se refinará con los tipos Zod específicos
): Promise<any> { // Devolverá string para 'get', void/boolean para otros
    const { filePath, sectionIndex, type, text, linkToPrevious, differentFirstPage, oddAndEvenPages } = args;
    logger.info(`[word/headers-footers] Operation: ${operation}`, { filePath, sectionIndex, type });

    let wordApp: any = null;
    let doc: any = null;
    let shouldQuit = false;
    let resultData: any = null;

    try {
        const officeResult = await getOfficeApplication('Word.Application');
        wordApp = officeResult.app;
        shouldQuit = officeResult.shouldQuit;

        doc = await wordApp.Documents.Open(filePath);

        const sectionsToProcess: any[] = [];
        if (sectionIndex === 'all') {
            for (let i = 1; i <= doc.Sections.Count; i++) {
                sectionsToProcess.push(doc.Sections.Item(i));
            }
            logger.debug(`[word/headers-footers] Processing all ${doc.Sections.Count} sections.`);
        } else {
             if (sectionIndex > doc.Sections.Count || sectionIndex <= 0) {
                 throw new Error(`Section index ${sectionIndex} is out of bounds (1-${doc.Sections.Count}).`);
             }
             sectionsToProcess.push(doc.Sections.Item(sectionIndex));
             logger.debug(`[word/headers-footers] Processing section ${sectionIndex}.`);
        }

        for (const section of sectionsToProcess) {
            const currentSectionIndex = section.Index; // Para logs
            logger.debug(`[word/headers-footers] Accessing section ${currentSectionIndex}...`);

            // Configuración específica de la sección (diferente primera página, etc.)
            if (operation === 'configure') {
                 const pageSetup = section.PageSetup;
                 if (differentFirstPage !== undefined) {
                     pageSetup.DifferentFirstPageHeaderFooter = differentFirstPage;
                     logger.debug(`[word/headers-footers] Section ${currentSectionIndex}: Set DifferentFirstPageHeaderFooter = ${differentFirstPage}`);
                 }
                 if (oddAndEvenPages !== undefined) {
                     pageSetup.OddAndEvenPagesHeaderFooter = oddAndEvenPages;
                      logger.debug(`[word/headers-footers] Section ${currentSectionIndex}: Set OddAndEvenPagesHeaderFooter = ${oddAndEvenPages}`);
                 }
                 // LinkToPrevious se maneja en el objeto HeaderFooter directamente
                 releaseObject(pageSetup); // Liberar PageSetup después de usarlo
            }

            // Acceder a la colección Headers o Footers
            // Nota: La documentación sugiere que Headers y Footers están en Section, no en PageSetup
            const headers = section.Headers;
            const footers = section.Footers;

            if (!headers || !footers) {
                logger.warn(`[word/headers-footers] Could not access Headers/Footers for section ${currentSectionIndex}. Skipping.`);
                releaseObject(headers);
                releaseObject(footers);
                continue; // Saltar a la siguiente sección si no se pueden obtener
            }

            // Obtener el objeto HeaderFooter específico según el tipo
            let headerFooter: any = null;
            try {
                 // Intentar acceder al header/footer específico. Puede fallar si el tipo no existe
                 // en esa sección (e.g., pedir EvenPages cuando OddAndEvenPagesHeaderFooter es false)
                 headerFooter = headers.Item(type) ?? footers.Item(type); // Intentar obtener de Headers o Footers
                 // TODO: Determinar si necesitamos distinguir entre header y footer explícitamente
                 //       o si el 'type' es suficiente. Por ahora, asumimos que Item(type) funciona
                 //       para ambos en el contexto de la sección. Revisar API COM.
                 //       Si se necesita distinguir, añadir un parámetro 'location': 'header' | 'footer'.

                 if (!headerFooter) {
                     throw new Error(`Header/Footer of type ${type} not found or accessible in section ${currentSectionIndex}. Check section PageSetup properties (DifferentFirstPage, OddAndEvenPages).`);
                 }
                 logger.debug(`[word/headers-footers] Accessed Header/Footer object for type ${type} in section ${currentSectionIndex}.`);

                 // Aplicar LinkToPrevious si es parte de la operación 'configure'
                 if (operation === 'configure' && linkToPrevious !== undefined) {
                     headerFooter.LinkToPrevious = linkToPrevious;
                     logger.debug(`[word/headers-footers] Section ${currentSectionIndex}, Type ${type}: Set LinkToPrevious = ${linkToPrevious}`);
                 }

                 // Realizar la operación principal
                 switch (operation) {
                     case 'insert':
                     case 'modify':
                         headerFooter.Range.Text = text;
                         logger.debug(`[word/headers-footers] Section ${currentSectionIndex}, Type ${type}: Set text content.`);
                         break;
                     case 'delete':
                         headerFooter.Range.Text = ""; // Borrar contenido
                         // headerFooter.Delete(); // ¿Existe Delete()? Revisar API. Borrar texto es más seguro.
                         logger.debug(`[word/headers-footers] Section ${currentSectionIndex}, Type ${type}: Cleared text content.`);
                         break;
                     case 'get':
                         resultData = headerFooter.Range.Text; // Solo guardamos el último si sectionIndex='all'
                         logger.debug(`[word/headers-footers] Section ${currentSectionIndex}, Type ${type}: Retrieved text content.`);
                         break;
                     case 'configure':
                         // Ya manejado arriba (LinkToPrevious, DifferentFirstPage, OddAndEvenPages)
                         break;
                 }

            } catch (hfError: any) {
                 logger.warn(`[word/headers-footers] Error processing type ${type} in section ${currentSectionIndex}: ${hfError.message}. Skipping this type/section combination.`);
                 // Continuar con la siguiente sección/tipo si es posible
            } finally {
                 releaseObject(headerFooter); // Liberar el objeto HeaderFooter específico
            }

            // Liberar colecciones de la sección actual
            releaseObject(headers);
            releaseObject(footers);
            releaseObject(section); // Liberar la sección actual
        } // Fin del bucle for sectionsToProcess

        if (operation !== 'get') {
            await doc.Save();
            logger.info(`[word/headers-footers] Document saved after ${operation}.`);
        }
        await doc.Close(false); // No guardar de nuevo al cerrar

        return resultData; // Devuelve el texto para 'get', null para otros

    } catch (error: any) {
        logger.error(`[word/headers-footers] Error during operation ${operation}: ${error.message}`, { error });
        if (doc) await doc.Close(false).catch((e: any) => logger.warn(`Failed to close document during error handling: ${e.message}`));
        throw handleToolError(error, 'OFFICE_API_ERROR');
    } finally {
        // Liberar objetos restantes
        releaseObject(doc);
        if (wordApp && shouldQuit) {
            await wordApp.Quit();
            logger.debug("[word/headers-footers] Word application closed by tool.");
        }
        releaseObject(wordApp);
    }
}


// --- Tool Definition ---

/**
 * @tool word/headers-footers
 * @description Manages headers and footers within a Word document (.docx) using COM Interop.
 * Allows inserting, modifying, deleting, retrieving text from, and configuring properties of headers and footers
 * for specific sections or the entire document, differentiating between primary, first page,
 * and even page types (wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages).
 * Operations:
 *  - `insert`: Adds or replaces text in the specified header/footer. Requires `filePath`, `type`, `text`. Optional: `sectionIndex`.
 *  - `modify`: Alias for `insert`. Adds or replaces text. Requires `filePath`, `type`, `text`. Optional: `sectionIndex`.
 *  - `delete`: Clears the text content of the specified header/footer. Requires `filePath`, `type`. Optional: `sectionIndex`.
 *  - `get`: Retrieves the text content of the specified header/footer. Requires `filePath`, `type`. Optional: `sectionIndex`. Returns text of the *last* processed section if `sectionIndex` is 'all'.
 *  - `configure`: Sets section-level properties affecting headers/footers. Requires `filePath` and at least one of `linkToPrevious`, `differentFirstPage`, `oddAndEvenPages`. Optional: `sectionIndex`. Note: `linkToPrevious` applies to the specific header/footer type, others apply to the section's PageSetup.
 * @inputSchema See `InputSchema` (discriminated union based on `operation`). Uses `HeadersFootersTextSchema`, `HeadersFootersActionSchema`, `HeadersFootersConfigureSchema`.
 * @outputSchema `get`: Returns success and text content in `data`. Others: Returns success status, null data, and a message.
 * @dependencies Requires Microsoft Word installed and accessible via COM Interop (`winax`).
 * @security Input `filePath` is validated. Ensure Word COM security settings are appropriate.
 * @errorHandling Uses standard error handling. Catches COM errors, validation errors, file access issues. Returns standardized `ErrorResponse`. Handles cases where a specific header/footer type might not be accessible in a section (e.g., requesting EvenPages when OddAndEvenPages is false).
 * @example_insert
 * ```json
 * {
 *   "operation": "insert",
 *   "arguments": {
 *     "filePath": "C:/path/to/document.docx",
 *     "sectionIndex": 1,
 *     "type": "wdHeaderFooterPrimary",
 *     "text": "Company Confidential - Page %p"
 *   }
 * }
 * ```
 * @example_get
 * ```json
 * {
 *   "operation": "get",
 *   "arguments": {
 *     "filePath": "C:/path/to/document.docx",
 *     "type": "wdHeaderFooterFirstPage"
 *   }
 * }
 * ```
 * @example_configure
 * ```json
 * {
 *   "operation": "configure",
 *   "arguments": {
 *     "filePath": "C:/path/to/document.docx",
 *     "sectionIndex": "all",
 *     "differentFirstPage": true,
 *     "oddAndEvenPages": true
 *   }
 * }
 * ```
 */
export const wordHeadersFootersTool: McpResource = { // Implementar McpResource
  path: 'word/headers-footers', // Propiedad path requerida
  // name: 'word/headers-footers', // 'name' no es parte de McpResource
  description: 'Manage headers and footers in a Word document (insert, modify, delete, get, configure).',
  schema: InputSchema, // Usar 'schema' según McpResource
  // outputSchema: z.any(), // Opcional: definir si es necesario

  async handler(params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> {
    let validatedRequest: InputSchemaType;
    try {
      // Validar la estructura completa { operation: '...', arguments: { ... } }
      validatedRequest = InputSchema.parse(params);
    } catch (error: any) {
       if (error instanceof z.ZodError) {
           logger.warn(`[word/headers-footers] Input validation failed: ${error.message}`, { errors: error.errors, params });
           return createErrorResponse('VALIDATION_ERROR', `Input validation failed: ${error.errors.map(e => `${e.path.join('.')} - ${e.message}`).join(', ')}`);
       }
       logger.error(`[word/headers-footers] Unexpected error parsing params: ${error.message}`, { error, params });
       return createErrorResponse('INTERNAL_ERROR', 'Failed to parse tool parameters.');
    }

    const { operation, arguments: args } = validatedRequest;

    try {
      const resultData = await manageHeaderFooter(operation, args);

      // Construir respuesta de éxito
      const response: ApiResponse<any> = {
          success: true,
          message: `Operation '${operation}' completed successfully for section(s) '${args.sectionIndex || 1}'.`,
          data: resultData, // Será null para operaciones que no sean 'get'
      };
      if (operation === 'get' && resultData === null) {
          response.message = `Operation 'get' completed, but no text content found for header/footer type '${args.type}' in section(s) '${args.sectionIndex || 1}'.`;
      } else if (operation === 'get') {
           response.message = `Operation 'get' completed successfully for header/footer type '${args.type}' in section(s) '${args.sectionIndex || 1}'.`;
      }

      return response;

    } catch (error: any) {
      // Los errores lanzados desde manageHeaderFooter ya deberían ser ErrorResponse
      if (error && typeof error === 'object' && 'success' in error && error.success === false && 'error' in error) {
          return error as ApiResponse<any>;
      }
      // Manejar otros errores inesperados que ocurran en el handler mismo
      logger.error(`[word/headers-footers] Unexpected error in handler: ${error.message}`, { stack: error.stack, operation, args });
      return handleToolError(error, 'UNEXPECTED_HANDLER_ERROR');
    }
  }
};