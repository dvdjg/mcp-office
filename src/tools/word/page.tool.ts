/**
 * @file Implements the 'word/page' tool using COM Interop for configuring page layout settings in Word documents.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types.js'; // Normalized relative path
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js'; // Normalized relative path
import { handleToolError, createErrorResponse } from '../../utils/errorHandler.js'; // Normalized relative path
import { validateFilePath } from '../../utils/security.js'; // Normalized relative path
import logger from '../../utils/logger.js'; // Normalized relative path

// --- Schemas ---

// Constants for common PageSetup values (examples, adjust based on COM API)
// Reference: https://learn.microsoft.com/en-us/office/vba/api/word.wdpapersize
/**
 * Constants for Word Paper Sizes.
 */
const WdPageSizes = {
  wdPaper10x14: 0, wdPaper11x17: 1, wdPaperLetter: 2, wdPaperLegal: 3, wdPaperExecutive: 4,
  wdPaperA3: 5, wdPaperA4: 6, wdPaperA5: 7, wdPaperB4: 8, wdPaperB5: 9,
  wdPaperFanfoldLegalGerman: 10, wdPaperFanfoldStdGerman: 11, wdPaperFanfoldUS: 12,
  wdPaperFolio: 14, wdPaperLedger: 15, wdPaperNote: 18, wdPaperStatement: 20,
  wdPaperTabloid: 21, wdPaperQuarto: 22, wdPaperEnvelope9: 29, wdPaperEnvelope10: 30,
  // ... Add more if needed
  wdPaperCustom: 256 // Common value for custom size, verify in COM documentation if different
} as const; // Use 'as const' to infer literal types

// Reference: https://learn.microsoft.com/en-us/office/vba/api/word.wdorientation
/**
 * Constants for Word Page Orientations.
 */
const WdOrientations = {
  wdOrientPortrait: 0,
  wdOrientLandscape: 1,
} as const;

// Schema for margins (use string to allow units like '1in', '2cm')
// Conversion to points (Word's internal unit) will be done in the COM logic
/**
 * Schema for margin values, allowing units (in, cm, mm, pt).
 */
const MarginSchema = z.string().regex(/^\d+(\.\d+)?\s*(in|cm|mm|pt)?$/i, "Invalid margin format (e.g., '1in', '2.5cm', '72pt', '10 mm')");

// Combined schema for all operations (using z.object)
/**
 * Zod schema for the input parameters of the 'word/page' tool.
 * Combines properties from get, set, and modify operations with refinement for operation-specific requirements.
 */
const WordPageInputSchema = z.object({
  /** The operation to perform ('get', 'set', or 'modify'). */
  operation: z.enum(['get', 'set', 'modify', 'getPageCount']).describe('The operation to perform (get, set, modify, or getPageCount).'),
  /** The path to the Word document. */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  /** 1-based index of the section to modify. Defaults to the first section if omitted. */
  sectionIndex: z.number().int().positive('Section index must be a positive integer.').optional().describe("1-based index of the section to modify. Defaults to the first section if omitted."), // Optional, to apply to specific sections
  /** Page size constant (e.g., wdPaperA4, wdPaperLetter). Use wdPaperCustom with pageHeight/pageWidth for custom sizes. */
  size: z.nativeEnum(WdPageSizes).optional().describe("Page size constant (e.g., wdPaperA4, wdPaperLetter). Use wdPaperCustom with pageHeight/pageWidth for custom sizes."),
  /** Page orientation constant (wdOrientPortrait or wdOrientLandscape). */
  orientation: z.nativeEnum(WdOrientations).optional().describe("Page orientation constant (wdOrientPortrait or wdOrientLandscape)."),
  /** Custom page width (required if size is wdPaperCustom for 'set'). Include units (in, cm, mm, pt). */
  pageWidth: MarginSchema.optional().describe("Custom page width (required if size is wdPaperCustom for 'set'). Include units (in, cm, mm, pt)."),
  /** Custom page height (required if size is wdPaperCustom for 'set'). Include units (in, cm, mm, pt). */
  pageHeight: MarginSchema.optional().describe("Custom page height (required if size is wdPaperCustom for 'set'). Include units (in, cm, mm, pt)."),
  /** Top margin (e.g., '1in', '2.5cm'). */
  topMargin: MarginSchema.optional().describe("Top margin (e.g., '1in', '2.5cm')."),
  /** Bottom margin (e.g., '1in', '2.5cm'). */
  bottomMargin: MarginSchema.optional().describe("Bottom margin (e.g., '1in', '2.5cm')."),
  /** Left margin (e.g., '1.25in', '3cm'). */
  leftMargin: MarginSchema.optional().describe("Left margin (e.g., '1.25in', '3cm')."),
  /** Right margin (e.g., '1.25in', '3cm'). */
  rightMargin: MarginSchema.optional().describe("Right margin (e.g., '1.25in', '3cm')."),
  /** Gutter margin (e.g., '0.5in'). */
  gutter: MarginSchema.optional().describe("Gutter margin (e.g., '0.5in')."),
  /** Distance from edge to header (e.g., '0.5in'). */
  headerDistance: MarginSchema.optional().describe("Distance from edge to header (e.g., '0.5in')."),
  /** Distance from edge to footer (e.g., '0.5in'). */
  footerDistance: MarginSchema.optional().describe("Distance from edge to footer (e.g., '0.5in')."),
  /** Different header/footer on the first page. */
  differentFirstPage: z.boolean().optional().describe("Different header/footer on the first page."),
  /** Different headers/footers for odd and even pages. */
  oddAndEvenPages: z.boolean().optional().describe("Different headers/footers for odd and even pages."),
}).refine(data => {
    // Specific validations per operation within the refinement
    if (data.operation === 'set') {
        // For 'set', validate that at least one configuration property is present
        const configKeys = ['size', 'orientation', 'pageWidth', 'pageHeight', 'topMargin', 'bottomMargin', 'leftMargin', 'rightMargin', 'gutter', 'headerDistance', 'footerDistance', 'differentFirstPage', 'oddAndEvenPages'];
        if (!configKeys.some(key => data[key as keyof typeof data] !== undefined)) {
            return false; // Validation failed: no configuration property for 'set'
        }
        // For 'set' with wdPaperCustom, validate that pageWidth and pageHeight are present
        if (data.size === WdPageSizes.wdPaperCustom && (!data.pageWidth || !data.pageHeight)) {
             return false; // Validation failed: wdPaperCustom requires pageWidth and pageHeight for 'set'
        }
    } else if (data.operation === 'modify') {
         // For 'modify', validate that at least one configuration property is present to change
         const configKeys = ['size', 'orientation', 'pageWidth', 'pageHeight', 'topMargin', 'bottomMargin', 'leftMargin', 'rightMargin', 'gutter', 'headerDistance', 'footerDistance', 'differentFirstPage', 'oddAndEvenPages'];
         if (!configKeys.some(key => data[key as keyof typeof data] !== undefined)) {
             return false; // Validation failed: no configuration property for 'modify'
         }
    } else if (data.operation === 'getPageCount') {
        // For 'getPageCount', only filePath is required, which is already handled by the base schema.
        // No other properties should be present.
        const allowedKeys: (keyof WordPageInput)[] = ['operation', 'filePath'];
        for (const key in data) {
            if (!allowedKeys.includes(key as keyof WordPageInput)) {
                return false; // Found an extraneous property for getPageCount
            }
        }
    }
    // For 'get', no additional properties are required besides filePath and optional sectionIndex
    return true; // Passes validation if not 'set' or 'modify' with issues, or if 'set'/'modify' meet their requirements
}, {
    message: "Invalid input for the specified operation. For 'set', provide at least one config property and include pageWidth/pageHeight if size is wdPaperCustom. For 'modify', provide at least one config property. For 'getPageCount', only 'filePath' is allowed.",
    path: [], // Apply error to the whole object
});


/**
 * Infers the combined type for use in the handler.
 */
type WordPageInput = z.infer<typeof WordPageInputSchema>;


// --- Helper Functions ---

/** Converts units to points (Word's internal unit) */
function convertToPoints(valueWithUnit: string | undefined, wordApp: any): number | undefined {
    if (valueWithUnit === undefined) return undefined;
    const match = valueWithUnit.trim().match(/^(\d+(\.\d+)?)\s*(in|cm|mm|pt)?$/i); // Make unit optional, default pt
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

// --- COM Logic ---

/** Gets the PageSetup configuration for a specific section */
async function getPageSetup(filePath: string, sectionIndex: number = 1): Promise<any | ApiResponse<any>> {
    let wordApp: any = null;
    let doc: any = null;
    let pageSetup: any = null;
    let section: any = null;
    let officeAppInstance: any = null; // To manage the application instance lifecycle

    try {
        const result = await getOfficeApplication('Word.Application');
        officeAppInstance = result; // Assign the instance
        wordApp = result.app;

        doc = await wordApp.Documents.Open(filePath);
        if (sectionIndex > doc.Sections.Count || sectionIndex <= 0) {
             throw new Error(`Section index ${sectionIndex} is out of bounds (1-${doc.Sections.Count}).`);
        }
        section = doc.Sections.Item(sectionIndex);
        pageSetup = section.PageSetup;

        // Extract relevant properties
        const config = {
            size: pageSetup.PaperSize, // Returns the numeric value of WdPaperSize
            orientation: pageSetup.Orientation, // Returns the numeric value of WdOrientation
            pageWidth: pageSetup.PageWidth, // In points
            pageHeight: pageSetup.PageHeight, // In points
            topMargin: pageSetup.TopMargin, // In points
            bottomMargin: pageSetup.BottomMargin, // In points
            leftMargin: pageSetup.LeftMargin, // In points
            rightMargin: pageSetup.RightMargin, // In points
            gutter: pageSetup.Gutter, // In points
            headerDistance: pageSetup.HeaderDistance, // In points
            footerDistance: pageSetup.FooterDistance, // In points
            differentFirstPage: pageSetup.DifferentFirstPageHeaderFooter, // Boolean
            oddAndEvenPages: pageSetup.OddAndEvenPagesHeaderFooter, // Boolean
            // Add other properties if needed
        };

        await doc.Close(false); // Do not save changes on close for 'get'
        logger.info(`Page setup retrieved successfully for ${filePath}, section ${sectionIndex}.`);
        return config;
    } catch (error: any) {
        logger.error(`Error getting page setup for ${filePath}, section ${sectionIndex}: ${error.message}`, { error }); // Log the full error
        if (doc) await doc.Close(false).catch((e: any) => logger.warn(`Failed to close document during error handling: ${e.message}`));
        // Throw the formatted ErrorResponse
        return handleToolError(error, 'OFFICE_API_ERROR'); // Use a more specific code and remove third argument
    } finally {
        releaseObject(pageSetup);
        releaseObject(section);
        releaseObject(doc);
        if (officeAppInstance) {
            officeAppInstance.release(); // Release the application instance
            logger.debug("Word application instance released.");
        }
        releaseObject(wordApp); // Release the app object reference
    }
}

/** Gets the total page count of a Word document */
async function getDocumentPageCount(filePath: string): Promise<number | ApiResponse<any>> {
    let wordApp: any = null;
    let doc: any = null;
    let officeAppInstance: any = null;

    try {
        const result = await getOfficeApplication('Word.Application');
        officeAppInstance = result;
        wordApp = result.app;

        doc = await wordApp.Documents.Open(filePath, false, true); // Open read-only

        // Using WdStatistic.wdStatisticPages (constant value 2)
        // This is generally more reliable than BuiltInDocumentProperties for page count.
        const pageCount = await doc.ComputeStatistics(2); // 2 corresponds to wdStatisticPages

        await doc.Close(false); // Do not save changes
        logger.info(`Successfully retrieved page count (${pageCount}) for ${filePath}.`);
        return pageCount;
    } catch (error: any) {
        logger.error(`Error getting page count for ${filePath}: ${error.message}`, { error });
        if (doc) await doc.Close(false).catch((e: any) => logger.warn(`Failed to close document during error handling for page count: ${e.message}`));
        return handleToolError(error, 'OFFICE_API_ERROR');
    } finally {
        releaseObject(doc);
        if (officeAppInstance) {
            officeAppInstance.release();
            logger.debug("Word application instance released after getting page count.");
        }
        releaseObject(wordApp);
    }
}

/** Applies PageSetup configuration to a specific section */
async function applyPageSetup(
    filePath: string,
    // Use the combined type for set/modify
    settings: WordPageInput,
    sectionIndex: number = 1,
): Promise<ApiResponse<any>> { // Always return ApiResponse
    let wordApp: any = null;
    let doc: any = null;
    let pageSetup: any = null;
    let section: any = null;
    let officeAppInstance: any = null; // To manage the application instance lifecycle

    try {
        const result = await getOfficeApplication('Word.Application');
        officeAppInstance = result; // Assign the instance
        wordApp = result.app;

        doc = await wordApp.Documents.Open(filePath);
        if (sectionIndex > doc.Sections.Count || sectionIndex <= 0) {
             throw new Error(`Section index ${sectionIndex} is out of bounds (1-${doc.Sections.Count}).`);
        }
        section = doc.Sections.Item(sectionIndex);
        pageSetup = section.PageSetup;

        // Apply configurations (only if defined in 'settings')
        if (settings.size !== undefined) pageSetup.PaperSize = settings.size;
        if (settings.orientation !== undefined) pageSetup.Orientation = settings.orientation;

        // Convert units and apply margins/dimensions
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

        // Apply other boolean properties if added to the schema
        if (settings.differentFirstPage !== undefined) pageSetup.DifferentFirstPageHeaderFooter = settings.differentFirstPage;
        if (settings.oddAndEvenPages !== undefined) pageSetup.OddAndEvenPagesHeaderFooter = settings.oddAndEvenPages;


        await doc.Save();
        await doc.Close();
        logger.info(`Page setup applied/modified successfully for ${filePath}, section ${sectionIndex || 1}.`);
        return { success: true, data: null, message: `Page setup operation successful for section ${sectionIndex || 1}.` };
    } catch (error: any) {
        logger.error(`Error applying page setup for ${filePath}, section ${sectionIndex || 1}: ${error.message}`, { error, settings }); // Log the full error and settings
        if (doc) await doc.Close(false).catch((e: any) => logger.warn(`Failed to close document during error handling: ${e.message}`));
        // Return the formatted ErrorResponse
        return handleToolError(error, 'OFFICE_API_ERROR'); // Use a more specific code and remove third argument
    } finally {
        releaseObject(pageSetup);
        releaseObject(section);
        releaseObject(doc);
        if (officeAppInstance) {
            officeAppInstance.release(); // Release the application instance
            logger.debug("Word application instance released.");
        }
        releaseObject(wordApp); // Release the app object reference
    }
}


// --- Tool Definition ---

/**
 * @tool word/page
 * @description Configures page layout settings (size, margins, orientation, headers/footers) for a specific section (default: first) in a Word document using COM Interop. Also retrieves document page count.
 * Operations:
 *  - `get`: Retrieves the current page setup for the specified section. Requires `filePath`. Optional: `sectionIndex`.
 *  - `set`: Sets the page setup configuration for the specified section. Requires `filePath` and at least one setting (e.g., `size`, `orientation`, `margins`, `differentFirstPage`). Overwrites existing settings for the specified properties. Optional: `sectionIndex`.
 *  - `modify`: Modifies specific page setup properties for the specified section. Requires `filePath` and at least one setting to change. Leaves other settings untouched. Optional: `sectionIndex`.
 *  - `getPageCount`: Retrieves the total number of pages in the document. Requires `filePath`.
 * @inputSchema See `WordPageInputSchema` (z.object). Uses combined properties from get/set/modify/getPageCount operations. Margins/dimensions require units (in, cm, mm, pt). Size/Orientation use Word constants (e.g., `wdPaperA4`, `wdOrientLandscape`).
 * @outputSchema `get`: Returns an object with page setup properties. `set`/`modify`: Returns success status. `getPageCount`: Returns an object like `{ pageCount: number }`.
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
 * @example_getPageCount
 * ```json
 * {
 *   "operation": "getPageCount",
 *   "filePath": "C:/path/to/document.docx"
 * }
 * ```
 */
export const wordPageTool: McpResource = { // Define as a single object
  path: 'word/page', // Add the path property required by McpResource
  description: 'Configures page layout settings (size, margins, orientation, headers/footers) in a Word document section.',
  schema: WordPageInputSchema, // Use the new z.object schema
  // outputSchema: z.any(), // Optional: define output schema if stable
  async handler(params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> { // Make context optional: context?: FastMCPContext<any>
    const log = context?.log ?? logger; // Use context logger or fallback

    // Validate and parse the generic params using the new WordPageInputSchema
    let validatedRequest: WordPageInput;
    try {
      validatedRequest = WordPageInputSchema.parse(params);
    } catch (error: any) {
       // If initial validation fails, return a validation error
       if (error instanceof z.ZodError) {
           log.warn(`Input validation failed at handler entry for word/page: ${error.message}`, { errors: JSON.stringify(error.errors), params });
           return createErrorResponse('VALIDATION_ERROR', `Input validation failed: ${error.errors.map(e => `${e.path.join('.')} - ${e.message}`).join(', ')}`);
       }
       // Another unexpected error during initial parsing
       log.error(`Unexpected error parsing params in word/page handler: ${error.message}`, { error, params });
       return createErrorResponse('INTERNAL_ERROR', 'Failed to parse tool parameters.');
    }

    // Now use validatedRequest which is correctly typed
    const { operation, ...args } = validatedRequest; // Extract operation and the rest as args
    const { filePath, sectionIndex } = args; // sectionIndex is optional

    try {
      // filePath validation is already in the base schema
      // The validation of the complete structure (operation + args) is done by Zod when defining inputSchema

      switch (operation) {
        case 'get':
          // args is already validated by Zod as part of WordPageInputSchema
          const configResult = await getPageSetup(filePath, sectionIndex); // Pass sectionIndex
          if (typeof configResult === 'object' && configResult !== null && 'success' in configResult && configResult.success === false) {
            return configResult as ApiResponse<any>;
          }
          return { success: true, data: configResult };
        case 'set':
        case 'modify':
          // args is already validated by Zod as part of WordPageInputSchema
          const applyResult = await applyPageSetup(filePath, validatedRequest, sectionIndex); // Pass the complete validatedRequest and sectionIndex
          // applyPageSetup now directly returns ApiResponse
          return applyResult;
        case 'getPageCount':
          const pageCountResult = await getDocumentPageCount(filePath);
          // Check if getDocumentPageCount returned an ErrorResponse
          if (pageCountResult && (pageCountResult as any).success === false) { // Simplified check
            return pageCountResult as ApiResponse<any>;
          }
          return { success: true, data: { pageCount: pageCountResult as number } };
        // No default case needed due to Zod enum
      }
    } catch (error: any) { // This catch block now primarily handles truly unexpected errors or Zod errors from refine
       // If it's a Zod validation error (e.g. from .refine, though initial parse is caught above)
       if (error instanceof z.ZodError) {
           log.warn(`Input validation failed at handler level for word/page: ${error.message}`, { errors: JSON.stringify(error.errors) });
           return handleToolError(error, 'VALIDATION_ERROR');
       }
       // Handle other unexpected errors
       log.error(`Unexpected error in word/page handler: ${error.message}`, { stack: error.stack, filePath, sectionIndex, operation });
       return handleToolError(error, 'UNEXPECTED_HANDLER_ERROR');
    }
  },
};

// No need for export default because tools are imported directly in index.ts