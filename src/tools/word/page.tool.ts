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

const PageSetupBaseSchema = z.object({
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  sectionIndex: z.number().int().positive().optional().describe("1-based index of the section to modify. Defaults to the first section if omitted."), // Opcional, para aplicar a secciones específicas
});

// Esquema detallado para SET (requiere al menos una propiedad de configuración)
const PageSetupSetSchema = PageSetupBaseSchema.extend({
  size: z.nativeEnum(WdPageSizes).optional().describe("Page size constant (e.g., wdPaperA4, wdPaperLetter). Use wdPaperCustom with pageHeight/pageWidth for custom sizes."),
  orientation: z.nativeEnum(WdOrientations).optional().describe("Page orientation constant (wdOrientPortrait or wdOrientLandscape)."),
  pageWidth: MarginSchema.optional().describe("Custom page width (required if size is wdPaperCustom). Include units (in, cm, mm, pt)."),
  pageHeight: MarginSchema.optional().describe("Custom page height (required if size is wdPaperCustom). Include units (in, cm, mm, pt)."),
  topMargin: MarginSchema.optional().describe("Top margin (e.g., '1in', '2.5cm')."),
  bottomMargin: MarginSchema.optional().describe("Bottom margin (e.g., '1in', '2.5cm')."),
  leftMargin: MarginSchema.optional().describe("Left margin (e.g., '1.25in', '3cm')."),
  rightMargin: MarginSchema.optional().describe("Right margin (e.g., '1.25in', '3cm')."),
  gutter: MarginSchema.optional().describe("Gutter margin (e.g., '0.5in')."),
  headerDistance: MarginSchema.optional().describe("Distance from edge to header (e.g., '0.5in')."),
  footerDistance: MarginSchema.optional().describe("Distance from edge to footer (e.g., '0.5in')."),
  // Añadir otras propiedades de PageSetup si son necesarias (e.g., DifferentFirstPageHeaderFooter, OddAndEvenPagesHeaderFooter como booleanos)
  differentFirstPage: z.boolean().optional().describe("Different header/footer on the first page."),
  oddAndEvenPages: z.boolean().optional().describe("Different headers/footers for odd and even pages."),
}).refine(data => {
    // Validar que si size es wdPaperCustom, se proporcionen pageWidth y pageHeight
    if (data.size === WdPageSizes.wdPaperCustom && (!data.pageWidth || !data.pageHeight)) {
        return false;
    }
    // Validar que al menos una propiedad de configuración esté presente para 'set'
    const configKeys = ['size', 'orientation', 'pageWidth', 'pageHeight', 'topMargin', 'bottomMargin', 'leftMargin', 'rightMargin', 'gutter', 'headerDistance', 'footerDistance', 'differentFirstPage', 'oddAndEvenPages'];
    return configKeys.some(key => data[key as keyof typeof data] !== undefined);
}, {
    message: "For 'set' operation, at least one configuration property (size, orientation, margins, etc.) must be provided. If size is wdPaperCustom, pageWidth and pageHeight are required.",
    path: [], // Apply error to the whole object
});


// Esquema detallado para MODIFY (todas las propiedades de configuración son opcionales)
const PageSetupModifySchema = PageSetupBaseSchema.extend({
  size: z.nativeEnum(WdPageSizes).optional().describe("Page size constant (e.g., wdPaperA4, wdPaperLetter). Use wdPaperCustom with pageHeight/pageWidth for custom sizes."),
  orientation: z.nativeEnum(WdOrientations).optional().describe("Page orientation constant (wdOrientPortrait or wdOrientLandscape)."),
  pageWidth: MarginSchema.optional().describe("Custom page width. Include units (in, cm, mm, pt)."),
  pageHeight: MarginSchema.optional().describe("Custom page height. Include units (in, cm, mm, pt)."),
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
    // Validar que al menos una propiedad de configuración esté presente para 'modify'
    const configKeys = ['size', 'orientation', 'pageWidth', 'pageHeight', 'topMargin', 'bottomMargin', 'leftMargin', 'rightMargin', 'gutter', 'headerDistance', 'footerDistance', 'differentFirstPage', 'oddAndEvenPages'];
    return configKeys.some(key => data[key as keyof typeof data] !== undefined);
}, {
    message: "For 'modify' operation, at least one configuration property to change must be provided.",
    path: [], // Apply error to the whole object
});


// Esquema para GET (solo necesita filePath y opcionalmente sectionIndex)
const PageSetupGetSchema = PageSetupBaseSchema;

// Combinar esquemas para el handler
const InputSchema = z.discriminatedUnion('operation', [
  z.object({ operation: z.literal('get'), arguments: PageSetupGetSchema }),
  z.object({ operation: z.literal('set'), arguments: PageSetupSetSchema }),
  z.object({ operation: z.literal('modify'), arguments: PageSetupModifySchema }),
]);
// Inferir el tipo combinado para usar en el handler
type InputSchemaType = z.infer<typeof InputSchema>;


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
    // Usar los tipos específicos inferidos de Zod para set/modify
    settings: z.infer<typeof PageSetupSetSchema> | z.infer<typeof PageSetupModifySchema>,
    sectionIndex: number = 1,
    isModify: boolean = false // Flag para distinguir set/modify si es necesario
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
        logger.info(`Page setup ${isModify ? 'modified' : 'set'} successfully for ${filePath}, section ${sectionIndex}.`);

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
 * @inputSchema See `InputSchema` (discriminated union based on `operation`). Uses `PageSetupGetSchema`, `PageSetupSetSchema`, `PageSetupModifySchema`. Margins/dimensions require units (in, cm, mm, pt). Size/Orientation use Word constants (e.g., `wdPaperA4`, `wdOrientLandscape`).
 * @outputSchema `get`: Returns an object with page setup properties (values in points or Word constants). `set`/`modify`: Returns success status with null data.
 * @dependencies Requires Microsoft Word installed and accessible via COM Interop (`winax`).
 * @security Input `filePath` is validated using `validateFilePath`. Ensure Word COM security settings are appropriate.
 * @errorHandling Uses standard error handling utility (`handleToolError`). Catches COM errors, validation errors, and file access issues. Returns standardized `ErrorResponse`.
 * @example_get
 * ```json
 * {
 *   "operation": "get",
 *   "arguments": {
 *     "filePath": "C:/path/to/document.docx",
 *     "sectionIndex": 1
 *   }
 * }
 * ```
 * @example_set
 * ```json
 * {
 *   "operation": "set",
 *   "arguments": {
 *     "filePath": "C:/path/to/document.docx",
 *     "size": "wdPaperA4",
 *     "orientation": "wdOrientLandscape",
 *     "topMargin": "1in",
 *     "bottomMargin": "2.5cm",
 *     "leftMargin": "72pt",
 *     "rightMargin": "30mm",
 *     "differentFirstPage": true
 *   }
 * }
 * ```
 * @example_modify
 * ```json
 * {
 *   "operation": "modify",
 *   "arguments": {
 *     "filePath": "C:/path/to/document.docx",
 *     "orientation": "wdOrientPortrait",
 *     "leftMargin": "1.5in",
 *     "rightMargin": "1.5in",
 *     "oddAndEvenPages": false
 *   }
 * }
 * ```
 */
export const wordPageTool = { // Definir como objeto directamente
  path: 'word/page', // Añadir la propiedad path requerida por McpResource
  name: 'word/page',
  description: 'Configures page layout settings (size, margins, orientation, headers/footers) in a Word document section.',
  inputSchema: InputSchema, // Schema para validar los params recibidos por el handler
  // outputSchema: z.any(), // Opcional: definir esquema de salida si es estable
  async handler(params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> { // Hacer context opcional: context?: FastMCPContext<any>
    // Validar y parsear los params genéricos usando el InputSchema específico
    let validatedRequest: InputSchemaType;
    try {
      validatedRequest = InputSchema.parse(params);
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
    const { operation, arguments: args } = validatedRequest;
    const { filePath, sectionIndex } = args; // sectionIndex es opcional

    try {
      // La validación de filePath ya está en el schema base
      // La validación de la estructura completa (operation + args) la hace Zod al definir inputSchema

      switch (operation) {
        case 'get':
          // args ya está validado por Zod como PageSetupGetSchema
          const config = await getPageSetup(filePath, args.sectionIndex); // Pasar sectionIndex
          return { success: true, data: config };
        case 'set':
          // args ya está validado por Zod como PageSetupSetSchema
          await applyPageSetup(filePath, args, args.sectionIndex, false); // Pasar args completos y sectionIndex
          return { success: true, data: null, message: `Page setup applied successfully to section ${args.sectionIndex || 1}.` }; // Añadir data: null
        case 'modify':
          // args ya está validado por Zod como PageSetupModifySchema
          await applyPageSetup(filePath, args, args.sectionIndex, true); // Pasar args completos, sectionIndex y flag modify
          return { success: true, data: null, message: `Page setup modified successfully for section ${args.sectionIndex || 1}.` }; // Añadir data: null
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