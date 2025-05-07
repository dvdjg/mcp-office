/**
 * @file Implements the 'word/headers-footers' tool using COM Interop (winax) for managing headers and footers in Word documents.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from "zod";
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext } from '../../types/common.types.js';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js';
import { handleToolError, createErrorResponse } from '../../utils/errorHandler.js';
import { validateFilePath } from '../../utils/security.js';
import logger from '../../utils/logger.js'; // Default import

// --- COM Constants (Examples, verify in Word documentation) ---
// https://learn.microsoft.com/en-us/office/vba/api/word.wdheaderfootertype
/**
 * Constants for Word Header/Footer types.
 */
const WdHeaderFooterTypes = {
  wdHeaderFooterPrimary: 1,
  wdHeaderFooterFirstPage: 2,
  wdHeaderFooterEvenPages: 3,
} as const;

// --- Schemas ---

/**
 * Base schema for header/footer operations requiring a file path and optional section index.
 */
const HeadersFootersBaseSchema = z.object({
  /** The path to the Word file (relative to the current workspace directory). */
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, {
    message: "Invalid or potentially unsafe file path provided.",
  }),
  /** 1-based index of the section or 'all'. Defaults to 1. */
  sectionIndex: z.union([z.number().int().positive(), z.literal('all')]).optional().default(1).describe("1-based index of the section or 'all'. Defaults to 1."),
});

/**
 * Zod schema for the input parameters of the 'word/headers-footers' tool.
 * Combines properties required for all operations, with refinement for operation-specific requirements.
 */
const WordHeadersFootersInputSchema = HeadersFootersBaseSchema.extend({
    /** The operation to perform ('insert', 'modify', 'delete', 'get', or 'configure'). */
    operation: z.enum(['insert', 'modify', 'delete', 'get', 'configure']).describe('The operation to perform (insert, modify, delete, get, or configure).'),
    /** Type of header/footer (Primary, FirstPage, EvenPages). Required for insert, modify, delete, get. */
    type: z.nativeEnum(WdHeaderFooterTypes).optional().describe("Type of header/footer (Primary, FirstPage, EvenPages). Required for insert, modify, delete, get."), // Make optional here, validate in handler
    /** Text content to insert or modify. Required for insert, modify. */
    text: z.string().optional().describe("Text content to insert or modify. Required for insert, modify."), // Make optional here, validate in handler
    /** Link header/footer to the previous section. Used in configure. */
    linkToPrevious: z.boolean().optional().describe("Link header/footer to the previous section. Used in configure."),
    /** Different header/footer on the first page for this section. Used in configure. */
    differentFirstPage: z.boolean().optional().describe("Different header/footer on the first page for this section. Used in configure."),
    /** Different headers/footers for odd and even pages for this section. Used in configure. */
    oddAndEvenPages: z.boolean().optional().describe("Different headers/footers for odd and even pages for this section. Used in configure."),
}).refine(data => {
    // Specific validations per operation within the refinement
    if (data.operation === 'insert' || data.operation === 'modify') {
        return data.type !== undefined && data.text !== undefined; // Requires type and text
    } else if (data.operation === 'delete' || data.operation === 'get') {
        return data.type !== undefined; // Requires type
    } else if (data.operation === 'configure') {
        return data.linkToPrevious !== undefined || data.differentFirstPage !== undefined || data.oddAndEvenPages !== undefined; // Requires at least one configuration option
    }
    return true; // Passes validation if the operation doesn't require specific fields or if it has them
}, {
    message: "Invalid input for the specified operation. Check required fields (type, text) and configuration options.",
    path: [], // Apply error to the whole object
});


/**
 * Infers the combined type for use in the handler.
 */
type WordHeadersFootersInput = z.infer<typeof WordHeadersFootersInputSchema>;


// --- COM Logic ---

/**
 * Manages headers and footers in a Word document using COM Interop.
 * @param operation - The operation to perform ('insert', 'modify', 'delete', 'get', or 'configure').
 * @param args - The validated input arguments.
 * @returns A promise resolving to the result of the operation (string for 'get', void/boolean for others).
 * @throws {Error} If a COM error occurs, a required parameter is missing for the operation, or a section/header/footer is not found.
 */
async function manageHeaderFooter(
    operation: 'insert' | 'modify' | 'delete' | 'get' | 'configure',
    args: WordHeadersFootersInput // Use the combined type
): Promise<any> { // Will return string for 'get', void/boolean for others
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

        // Open the document in read/write mode (ReadOnly = false) and visible (Visible = true)
        doc = await wordApp.Documents.Open(filePath, false, false, false, "", "", true, "", "", "", 0, false, true); // ReadOnly=false, Visible=true

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
            const currentSectionIndex = section.Index; // For logs
            logger.debug(`[word/headers-footers] Accessing section ${currentSectionIndex}...`);

            // Section-specific configuration (different first page, etc.)
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
                 // LinkToPrevious is handled on the HeaderFooter object directly
                 releaseObject(pageSetup); // Release PageSetup after use
            }

            // Access the Headers or Footers collection
            // Note: Documentation suggests Headers and Footers are on Section, not PageSetup
            const headers = section.Headers;
            const footers = section.Footers;

            if (!headers || !footers) {
                logger.warn(`[word/headers-footers] Could not access Headers/Footers for section ${currentSectionIndex}. Skipping.`);
                releaseObject(headers);
                releaseObject(footers);
                continue; // Skip to the next section if unable to get
            }

            // Get the specific HeaderFooter object based on type
            let headerFooter: any = null;
            try {
                 // Attempt to access the specific header/footer. May fail if the type does not exist
                 // in that section (e.g., requesting EvenPages when OddAndEvenPagesHeaderFooter is false)
                 // Ensure 'type' is not undefined before using it
                 if (type === undefined && operation !== 'configure') {
                     throw new Error(`Header/Footer type is required for operation '${operation}'.`);
                 }
                 if (type !== undefined) {
                    // Try to get from Headers or Footers. Note: This might need refinement
                    // based on whether the user intends to target a header or a footer.
                    // For now, assuming Item(type) works for both in the section context.
                    // If distinction is needed, add a 'location': 'header' | 'footer' parameter.
                    try {
                        headerFooter = headers.Item(type);
                    } catch (e) {
                        // If not found in headers, try footers
                        try {
                            headerFooter = footers.Item(type);
                        } catch (e2) {
                            // Neither found, headerFooter remains null
                        }
                    }
                 }


                 if (!headerFooter && operation !== 'configure') { // headerFooter is not needed for configure
                     throw new Error(`Header/Footer of type ${type} not found or accessible in section ${currentSectionIndex}. Check section PageSetup properties (DifferentFirstPage, OddAndEvenPages).`);
                 }
                 if (headerFooter) { // Only log if the object was obtained
                    logger.debug(`[word/headers-footers] Accessed Header/Footer object for type ${type} in section ${currentSectionIndex}.`);
                 }


                 // Apply LinkToPrevious if it's part of the 'configure' operation
                 if (operation === 'configure' && linkToPrevious !== undefined && headerFooter) { // Only apply if headerFooter exists
                     headerFooter.LinkToPrevious = linkToPrevious;
                     logger.debug(`[word/headers-footers] Section ${currentSectionIndex}, Type ${type}: Set LinkToPrevious = ${linkToPrevious}`);
                 }

                 // Perform the main operation
                 switch (operation) {
                     case 'insert':
                     case 'modify':
                         if (headerFooter && text !== undefined) { // Ensure headerFooter and text exist
                            headerFooter.Range.Text = text;
                            logger.debug(`[word/headers-footers] Section ${currentSectionIndex}, Type ${type}: Set text content.`);
                         }
                         break;
                     case 'delete':
                         if (headerFooter) { // Ensure headerFooter exists
                            headerFooter.Range.Text = ""; // Clear content
                            // headerFooter.Delete(); // Does Delete() exist? Check API. Clearing text is safer.
                            logger.debug(`[word/headers-footers] Section ${currentSectionIndex}, Type ${type}: Cleared text content.`);
                         }
                         break;
                     case 'get':
                         if (headerFooter) { // Ensure headerFooter exists
                            resultData = headerFooter.Range.Text; // Only store the last one if sectionIndex='all'
                            logger.debug(`[word/headers-footers] Section ${currentSectionIndex}, Type ${type}: Retrieved text content.`);
                         } else {
                             resultData = null; // Header/footer not found
                             logger.debug(`[word/headers-footers] Section ${currentSectionIndex}, Type ${type}: Header/Footer object not found for get operation.`);
                         }
                         break;
                     case 'configure':
                         // Already handled above (LinkToPrevious, DifferentFirstPage, OddAndEvenPages)
                         break;
                 }

            } catch (hfError: any) {
                 logger.warn(`[word/headers-footers] Error processing type ${type} in section ${currentSectionIndex}: ${hfError.message}. Skipping this type/section combination.`);
                 // Continue with the next section/type if possible
            } finally {
                 releaseObject(headerFooter); // Release the specific HeaderFooter object
            }

            // Release collections of the current section
            releaseObject(headers);
            releaseObject(footers);
            releaseObject(section); // Release the current section
        } // End of the for loop for sectionsToProcess

        if (operation !== 'get') {
            await doc.Save();
            logger.info(`[word/headers-footers] Document saved after ${operation}.`);
        }
        await doc.Close(false); // Do not save again on close

        return resultData; // Returns text for 'get', null for others

    } catch (error: any) {
        logger.error(`[word/headers-footers] Error during operation ${operation}: ${error.message}`, { error });
        if (doc) await doc.Close(false).catch((e: any) => logger.warn(`Failed to close document during error handling: ${e.message}`));
        throw handleToolError(error, 'OFFICE_API_ERROR');
    } finally {
        // Release remaining objects
        releaseObject(doc);
        if (wordApp && shouldQuit) {
            try {
                await wordApp.Quit();
                logger.debug("[word/headers-footers] Word application closed by tool.");
            } catch (quitError: any) {
                logger.warn(`[word/headers-footers] Error attempting to quit Word: ${quitError.message}`);
            }
        }
        releaseObject(wordApp);
    }
}


// --- Tool Definition ---

/**
 * McpResource definition for the 'word/headers-footers' tool.
 * Manages headers and footers within a Word document (.docx) using COM Interop.
 */
export const wordHeadersFootersTool: McpResource = { // Implement McpResource
  path: 'word/headers-footers', // Required path property
  description: 'Manage headers and footers in a Word document (insert, modify, delete, get, configure).',
  schema: WordHeadersFootersInputSchema, // Use the new z.object schema
  // outputSchema: z.any(), // Optional: define if needed

  async handler(params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> {
    let validatedRequest: WordHeadersFootersInput;
    try {
      // Validate the complete structure { operation: '...', ... }
      validatedRequest = WordHeadersFootersInputSchema.parse(params);
    } catch (error: any) {
       if (error instanceof z.ZodError) {
           logger.warn(`[word/headers-footers] Input validation failed: ${error.message}`, { errors: error.errors, params });
           return createErrorResponse('VALIDATION_ERROR', `Input validation failed: ${error.errors.map(e => `${e.path.join('.')} - ${e.message}`).join(', ')}`);
       }
       logger.error(`[word/headers-footers] Unexpected error parsing params: ${error.message}`, { error, params });
       return createErrorResponse('INTERNAL_ERROR', 'Failed to parse tool parameters.');
    }

    const { operation, ...args } = validatedRequest; // Extract operation and the rest as args

    try {
      const resultData = await manageHeaderFooter(operation, validatedRequest); // Pass the complete validatedRequest

      // Build success response
      const response: ApiResponse<any> = {
          success: true,
          message: `Operation '${operation}' completed successfully for section(s) '${args.sectionIndex || 1}'.`,
          data: resultData, // Will be null for operations other than 'get'
      };
      if (operation === 'get' && resultData === null) {
          response.message = `Operation 'get' completed, but no text content found for header/footer type '${args.type}' in section(s) '${args.sectionIndex || 1}'.`;
      } else if (operation === 'get') {
           response.message = `Operation 'get' completed successfully for header/footer type '${args.type}' in section(s) '${args.sectionIndex || 1}'.`;
      }

      return response;

    } catch (error: any) {
      // Errors thrown from manageHeaderFooter should already be ErrorResponse
      if (error && typeof error === 'object' && 'success' in error && error.success === false && 'error' in error) {
          return error as ApiResponse<any>;
      }
      // Handle other unexpected errors that occur in the handler itself
      logger.error(`[word/headers-footers] Unexpected error in handler: ${error.message}`, { stack: error.stack, operation, args });
      return handleToolError(error, 'UNEXPECTED_HANDLER_ERROR');
    }
  }
};