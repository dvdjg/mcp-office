import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types'; // Ajustado según text.tool.ts
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Usar getOfficeApplication
import { handleToolError, createErrorResponse } from '../../utils/errorHandler'; // Quitar ToolError
import { validateFilePath } from '../../utils/security'; // Añadir validación de ruta
import logger from '../../utils/logger'; // Añadir logger

// --- Schemas ---

// Constantes para valores comunes de PageSetup (ejemplos, ajustar según API COM)
// Referencia: https://learn.microsoft.com/en-us/office/vba/api/word.wdpapersize
const WdPageSizes = {
  wdPaper10x14: 0, wdPaper11x17: 1, wdPaperLetter: 2, wdPaperLegal: 3, wdPaperExecutive: 4,
  wdPaperA3: 5, wdPaperA4: 6, wdPaperA5: 7, wdPaperB4: 8, wdPaperB5: 9,
  wdPaperFanfoldLegalGerman: 10, wdPaperFanfoldStdGerman: 11, wdPaperFanfoldUS: 12,
  wdPaperFolio: 14, wdPaperLedger: 15, wdPaperNote: 18, wdPaperStatement: 20,
  wdPaperTabloid: 21, wdPaperQuarto: 22, wdPaperEnvelope9: 29, wdPaperEnvelope10: 30,
  // ... Añadir más si es necesario
  wdPaperCustom: 256 // Valor común para tamaño personalizado, verificar en la documentación COM si es diferente
} as const; // Usar 'as const' para inferir tipos literales

// Referencia: https://learn.microsoft.com/en-us/office/vba/api/word.wdorientation
const WdOrientations = {
  wdOrientPortrait: 0,
  wdOrientLandscape: 1,
} as const;

// Esquema para márgenes (usar string para permitir unidades como '1in', '2cm')
// La conversión a puntos (unidad de Word) se hará en la lógica COM
const MarginSchema = z.string().regex(/^\d+(\.\d+)?\s*(in|cm|mm|pt)$/i, "Invalid margin format (e.g., '1in', '2.5cm', '72pt', '10 mm')");

// Esquema base para PageSetup
const PageSetupBaseSchema = z.object({
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  sectionIndex: z.number().int().positive().optional().describe("1-based index of the section to modify. Defaults to the first section if omitted."), // Opcional, para aplicar a secciones específicas
});

// Esquema combinado para todas las operaciones (usando z.object)
const WordPageInputSchema = z.object({
  operation: z.enum(['get', 'set', 'modify']).describe('The operation to perform (get, set, or modify).'),
  // Incluir todos los campos posibles de las operaciones get, set, modify
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  sectionIndex: z.number().int().positive().optional().describe("1-based index of the section to modify. Defaults to the first section if omitted."),
  size: z.nativeEnum(WdPageSizes).optional().describe("Page size constant (e.g., wdPaperA4, wdPaperLetter). Use wdPaperCustom with pageHeight/pageWidth for custom sizes."),
  orientation: z.nativeEnum(WdOrientations).optional().describe("Page orientation constant (wdOrientPortrait or wdOrientLandscape)."),
  pageWidth: MarginSchema.optional().describe("Custom page width (required if size is wdPaperCustom for 'set'). Include units (in, cm, mm, pt)."),
  pageHeight: MarginSchema.optional().describe("Custom page height (required if size is wdPaperCustom for 'set'). Include units (in, cm, mm, pt)."),
  topMargin: MarginSchema.optional().describe("Top margin (e.g., '1in', '2.5cm')."),
  bottomMargin: MarginSchema.optional().describe("Bottom margin (e.g., '1in', '2.5cm')."),
  leftMargin: MarginSchema.optional().describe("Left margin (e.g., '1.25in', '3cm')."),
  rightMargin: MarginSchema.optional().describe("Right margin (e.g., '1.25in', '3cm')."),
  gutter: MarginSchema.optional().describe("Gutter margin (e.g., '0.5in')."),
  headerDistance: MarginSchema.optional().describe("Distance from edge to header (e.g., '0.5in')."),
  footerDistance: MarginSchema.optional().describe("Distance from edge to footer (e.g., '0.5in')."),
  differentFirstPage: z.boolean().optional().describe("Different header/footer on the first page."),
  oddAndEvenPages: z.boolean().optional().describe("Different headers/footers for odd and even pages."),
}).refine(data => {
    // Validaciones específicas por operación dentro del refinamiento
    if (data.operation === 'set') {
        // Para 'set', validar que al menos una propiedad de configuración esté presente
        const configKeys = ['size', 'orientation', 'pageWidth', 'pageHeight', 'topMargin', 'bottomMargin', 'leftMargin', 'rightMargin', 'gutter', 'headerDistance', 'footerDistance', 'differentFirstPage', 'oddAndEvenPages'];
        if (!configKeys.some(key => data[key as keyof typeof data] !== undefined)) {
            return false; // Falló la validación: ninguna propiedad de configuración para 'set'
        }
        // Para 'set' con wdPaperCustom, validar que pageWidth y pageHeight estén presentes
        if (data.size === WdPageSizes.wdPaperCustom && (!data.pageWidth || !data.pageHeight)) {
             return false; // Falló la validación: wdPaperCustom requiere pageWidth y pageHeight para 'set'
        }
    } else if (data.operation === 'modify') {
         // Para 'modify', validar que al menos una propiedad de configuración esté presente para cambiar
         const configKeys = ['size', 'orientation', 'pageWidth', 'pageHeight', 'topMargin', 'bottomMargin', 'leftMargin', 'rightMargin', 'gutter', 'headerDistance', 'footerDistance', 'differentFirstPage', 'oddAndEvenPages'];
         if (!configKeys.some(key => data[key as keyof typeof data] !== undefined)) {
             return false; // Falló la validación: ninguna propiedad de configuración para 'modify'
         }
    }
    // Para 'get', no se requieren propiedades adicionales aparte de filePath y opcionalmente sectionIndex
    // Si operation es 'get', las validaciones anteriores no aplican.
    return true; // Pasa la validación si no es 'set' o 'modify' con problemas, o si 'set'/'modify' cumplen sus requisitos
}, {
    message: "Invalid input for the specified operation. For 'set', provide at least one configuration property and include pageWidth/pageHeight if size is wdPaperCustom. For 'modify', provide at least one configuration property to change.",
    path: [], // Apply error to the whole object
});


// Inferir el tipo combinado para usar en el handler
type WordPageInput = z.infer<typeof WordPageInputSchema>;


// --- Funciones Auxiliares ---

/** Convertidor de unidades a puntos (unidad interna de Word) */
function convertToPoints(valueWithUnit: string | undefined, wordApp: any): number | undefined {
    if (valueWithUnit === undefined) return undefined;
    const match = valueWithUnit.trim().match(/^(\d+(\.\d+)?)\s*(in|cm|mm|pt)?$/i); // Hacer unidad opcional, default pt
    if (!match) throw new Error(`Invalid unit format: ${valueWithUnit}`);
    const value = parseFloat(match[1]);
    const unit = match[3]?.toLowerCase() || 'pt'; // Default to points if unit is missing

    try {
        switch (unit) {
            case 'in': return wordApp.InchesToPoints(value);
            case 'cm': return wordApp.CentimetersToPoints(value);
            case 'mm': return wordApp.MillimetersToPoints(value);
            case 'pt': return value;
            default: throw new Error(`Unsupported unit: ${unit}`); // Should not happen due to regex
        }
    } catch (comError: any) {
        logger.error(`COM Error converting unit ${valueWithUnit}: ${comError.message}`);
        throw new Error(`Failed to convert unit ${valueWithUnit} using Word COM object.`);
    }
}

// --- Lógica COM ---

/** Obtiene la configuración de PageSetup para una sección específica */
async function getPageSetup(filePath: string, sectionIndex: number = 1): Promise<any> {
    let wordApp: any = null;
    let doc: any = null;
    let pageSetup: any = null;
    let section: any = null;
    let shouldQuit = false; // Flag para saber si debemos cerrar Word

    try {
        const result = await getOfficeApplication('Word.Application');
        wordApp = result.app;
        shouldQuit = result.shouldQuit; // Determina si la app fue abierta por nosotros

        doc = await wordApp.Documents.Open(filePath);
        if (sectionIndex > doc.Sections.Count || sectionIndex <= 0) {
             throw new Error(`Section index ${sectionIndex} is out of bounds (1-${doc.Sections.Count}).`);
        }
        section = doc.Sections.Item(sectionIndex);
        pageSetup = section.PageSetup;

        // Extraer propiedades relevantes
        const config = {
            size: pageSetup.PaperSize, // Devuelve el valor numérico de WdPaperSize
            orientation: pageSetup.Orientation, // Devuelve el valor numérico de WdOrientation
            pageWidth: pageSetup.PageWidth, // En puntos
            pageHeight: pageSetup.PageHeight, // En puntos
            topMargin: pageSetup.TopMargin, // En puntos
            bottomMargin: pageSetup.BottomMargin, // En puntos
            leftMargin: pageSetup.LeftMargin, // En puntos
            rightMargin: pageSetup.RightMargin, // En puntos
            gutter: pageSetup.Gutter, // En puntos
            headerDistance: pageSetup.HeaderDistance, // En puntos
            footerDistance: pageSetup.FooterDistance, // En puntos
            differentFirstPage: pageSetup.DifferentFirstPageHeaderFooter, // Booleano
            oddAndEvenPages: pageSetup.OddAndEvenPagesHeaderFooter, // Booleano
            // Añadir otras propiedades si es necesario
        };

        await doc.Close(false); // No guardar cambios al cerrar para 'get'
        logger.info(`Page setup retrieved successfully for ${filePath}, section ${sectionIndex}.`);
        return config;
    } catch (error: any) {
        logger.error(`Error getting page setup for ${filePath}, section ${sectionIndex}: ${error.message}`, { error }); // Loguear el error completo
        if (doc) await doc.Close(false).catch((e: any) => logger.warn(`Failed to close document during error handling: ${e.message}`));
        // Lanzar el ErrorResponse formateado
        throw handleToolError(error, 'OFFICE_API_ERROR'); // Usar código más específico y quitar tercer argumento
    } finally {
        releaseObject(pageSetup);
        releaseObject(section);
        releaseObject(doc);
        if (wordApp && shouldQuit) {
            await wordApp.Quit();
            logger.debug("Word application closed by tool.");
        }
        releaseObject(wordApp);
    }
}

/** Aplica la configuración de PageSetup a una sección específica */
async function applyPageSetup(
    filePath: string,
    // Usar el tipo combinado para set/modify
    settings: WordPageInput,
    sectionIndex: number = 1,
): Promise<void> {
    let wordApp: any = null;
    let doc: any = null;
    let pageSetup: any = null;
    let section: any = null;
    let shouldQuit = false;

    try {
        const result = await getOfficeApplication('Word.Application');
        wordApp = result.app;
        shouldQuit = result.shouldQuit;

        doc = await wordApp.Documents.Open(filePath);
        if (sectionIndex > doc.Sections.Count || sectionIndex <= 0) {
             throw new Error(`Section index ${sectionIndex} is out of bounds (1-${doc.Sections.Count}).`);
        }
        section = doc.Sections.Item(sectionIndex);
        pageSetup = section.PageSetup;

        // Aplicar configuraciones (solo si están definidas en 'settings')
        if (settings.size !== undefined) pageSetup.PaperSize = settings.size;
        if (settings.orientation !== undefined) pageSetup.Orientation = settings.orientation;

        // Convertir unidades y aplicar márgenes/dimensiones
        const pageWidthPt = convertToPoints(settings.pageWidth, wordApp);
        if (pageWidthPt !== undefined) pageSetup.PageWidth = pageWidthPt;

        const pageHeightPt = convertToPoints(settings.pageHeight, wordApp);
        if (pageHeightPt !== undefined) pageSetup.PageHeight = pageHeightPt;

        const topMarginPt = convertToPoints(settings.topMargin, wordApp);
        if (topMarginPt !== undefined) pageSetup.TopMargin = topMarginPt;

        const bottomMarginPt = convertToPoints(settings.bottomMargin, wordApp);
        if (bottomMarginPt !== undefined) pageSetup.BottomMargin = bottomMarginPt;

        const leftMarginPt = convertToPoints(settings.leftMargin, wordApp);
        if (leftMarginPt !== undefined) pageSetup.LeftMargin = leftMarginPt;

        const rightMarginPt = convertToPoints(settings.rightMargin, wordApp);
        if (rightMarginPt !== undefined) pageSetup.RightMargin = rightMarginPt;

        const gutterPt = convertToPoints(settings.gutter, wordApp);
        if (gutterPt !== undefined) pageSetup.Gutter = gutterPt;

        const headerDistancePt = convertToPoints(settings.headerDistance, wordApp);
        if (headerDistancePt !== undefined) pageSetup.HeaderDistance = headerDistancePt;

        const footerDistancePt = convertToPoints(settings.footerDistance, wordApp);
        if (footerDistancePt !== undefined) pageSetup.FooterDistance = footerDistancePt;

        // Aplicar otras propiedades booleanas si se añadieron al schema
        if (settings.differentFirstPage !== undefined) pageSetup.DifferentFirstPageHeaderFooter = settings.differentFirstPage;
        if (settings.oddAndEvenPages !== undefined) pageSetup.OddAndEvenPagesHeaderFooter = settings.oddAndEvenPages;


        await doc.Save();
        await doc.Close();
        logger.info(`Page setup applied/modified successfully for ${filePath}, section ${sectionIndex}.`);

    } catch (error: any) {
        logger.error(`Error applying page setup for ${filePath}, section ${sectionIndex}: ${error.message}`, { error, settings }); // Loguear el error completo y settings
        if (doc) await doc.Close(false).catch((e: any) => logger.warn(`Failed to close document during error handling: ${e.message}`));
        // Lanzar el ErrorResponse formateado
        throw handleToolError(error, 'OFFICE_API_ERROR'); // Usar código más específico y quitar tercer argumento
    } finally {
        releaseObject(pageSetup);
        releaseObject(section);
        releaseObject(doc);
        if (wordApp && shouldQuit) {
            await wordApp.Quit();
            logger.debug("Word application closed by tool.");
        }
        releaseObject(wordApp);
    }
}


// --- Tool Definition ---

/**
 * @tool word/page
 * @description Configures page layout settings (size, margins, orientation, headers/footers) for a specific section (default: first) in a Word document using COM Interop.
 * Operations:
 *  - `get`: Retrieves the current page setup for the specified section. Requires `filePath`. Optional: `sectionIndex`.
 *  - `set`: Sets the page setup configuration for the specified section. Requires `filePath` and at least one setting (e.g., `size`, `orientation`, `margins`, `differentFirstPage`). Overwrites existing settings for the specified properties. Optional: `sectionIndex`.
 *  - `modify`: Modifies specific page setup properties for the specified section. Requires `filePath` and at least one setting to change. Leaves other settings untouched. Optional: `sectionIndex`.
 * @inputSchema See `WordPageInputSchema` (z.object). Uses combined properties from get/set/modify operations. Margins/dimensions require units (in, cm, mm, pt). Size/Orientation use Word constants (e.g., `wdPaperA4`, `wdOrientLandscape`).
 * @outputSchema `get`: Returns an object with page setup properties (values in points or Word constants). `set`/`modify`: Returns success status with null data.
 * @dependencies Requires Microsoft Word installed and accessible via COM Interop (`winax`).
 * @security Input `filePath` is validated using `validateFilePath`. Ensure Word COM security settings are appropriate.
 * @errorHandling Uses standard error handling utility (`handleToolError`). Catches COM errors, validation errors, and file access issues. Returns standardized `ErrorResponse`.
 * @example_get
 * ```json
 * {
 *   "operation": "get",
 *   "filePath": "C:/path/to/document.docx",
 *   "sectionIndex": 1
 * }
 * ```
 * @example_set
 * ```json
 * {
 *   "operation": "set",
 *   "filePath": "C:/path/to/document.docx",
 *   "size": "wdPaperA4",
 *   "orientation": "wdOrientLandscape",
   "topMargin": "1in",
   "bottomMargin": "2.5cm",
   "leftMargin": "72pt",
   "rightMargin": "30mm",
   "differentFirstPage": true
 * }
 * ```
 * @example_modify
 * ```json
 * {
 *   "operation": "modify",
 *   "filePath": "C:/path/to/document.docx",
 *   "orientation": "wdOrientPortrait",
 *   "leftMargin": "1.5in",
 *   "rightMargin": "1.5in",
 *   "oddAndEvenPages": false
 * }
 * ```
 */
export const wordPageTool: McpResource = { // Definir como objeto directamente
  path: 'word/page', // Añadir la propiedad path requerida por McpResource
  description: 'Configures page layout settings (size, margins, orientation, headers/footers) in a Word document section.',
  schema: WordPageInputSchema, // Usar el nuevo esquema z.object
  // outputSchema: z.any(), // Opcional: definir esquema de salida si es estable
  async handler(params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> { // Hacer context opcional: context?: FastMCPContext<any>
    // Validar y parsear los params genéricos usando el nuevo WordPageInputSchema
    let validatedRequest: WordPageInput;
    try {
      validatedRequest = WordPageInputSchema.parse(params);
    } catch (error: any) {
       // Si la validación inicial falla, devolver un error de validación
       if (error instanceof z.ZodError) {
           logger.warn(`Input validation failed at handler entry for word/page: ${error.message}`, { errors: error.errors, params });
           return createErrorResponse('VALIDATION_ERROR', `Input validation failed: ${error.errors.map(e => `${e.path.join('.')} - ${e.message}`).join(', ')}`);
       }
       // Otro error inesperado durante el parseo inicial
       logger.error(`Unexpected error parsing params in word/page handler: ${error.message}`, { error, params });
       return createErrorResponse('INTERNAL_ERROR', 'Failed to parse tool parameters.');
    }

    // Ahora usar validatedRequest que está correctamente tipado
    const { operation, ...args } = validatedRequest; // Extraer operation y el resto como args
    const { filePath, sectionIndex } = args; // sectionIndex es opcional

    try {
      // La validación de filePath ya está en el schema base
      // La validación de la estructura completa (operation + args) la hace Zod al definir inputSchema

      switch (operation) {
        case 'get':
          // args ya está validado por Zod como parte de WordPageInputSchema
          const config = await getPageSetup(filePath, sectionIndex); // Pasar sectionIndex
          return { success: true, data: config };
        case 'set':
        case 'modify':
          // args ya está validado por Zod como parte de WordPageInputSchema
          await applyPageSetup(filePath, validatedRequest, sectionIndex); // Pasar validatedRequest completo y sectionIndex
          return { success: true, data: null, message: `Page setup ${operation}ed successfully for section ${sectionIndex || 1}.` }; // Añadir data: null y mensaje dinámico
        // No se necesita default case debido a Zod discriminatedUnion
      }
    } catch (error: any) {
       // Si el error ya es un ErrorResponse (lanzado por handleToolError dentro de las funciones COM), devolverlo directamente
       // Comprobamos si tiene la estructura de ErrorResponse
       if (error && typeof error === 'object' && 'success' in error && error.success === false && 'error' in error) {
           // TypeScript ahora debería reconocer 'error' como compatible con ErrorResponse aquí
           return error as ApiResponse<any>; // Devolver como ApiResponse genérico que incluye ErrorResponse
       }
       // Si es un error de validación de Zod (aunque inputSchema debería atraparlos antes, por si acaso)
       if (error instanceof z.ZodError) {
           logger.warn(`Input validation failed at handler level for word/page: ${error.message}`, { errors: error.errors });
           // Usar handleToolError para formatearlo consistentemente
           return handleToolError(error, 'VALIDATION_ERROR');
       }
       // Manejar otros errores inesperados que no pasaron por handleToolError
       logger.error(`Unexpected error in word/page handler: ${error.message}`, { stack: error.stack, filePath, sectionIndex, operation });
       // Usar handleToolError para estandarizar errores inesperados
       return handleToolError(error, 'UNEXPECTED_HANDLER_ERROR'); // Usar código específico y solo 2 args
    }
  },
};

// No necesitamos export default porque las herramientas se importan directamente en index.ts