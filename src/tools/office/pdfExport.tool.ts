/**
 * @file Tool for exporting Office documents (Word, Excel, PowerPoint) to PDF format.
 * Uses COM Interop via winax to interact with Office applications.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types.js';
import { OfficeAppName, getOfficeApplication } from '../../utils/officeInterop.js'; // Import correct names
import { saveResource } from '../dynamic/resources.tool.js'; // Import saveResource
import * as fs from 'fs-extra'; // Import fs to read the PDF file
import * as path from 'path'; // Import path

// Input schema for the office/pdf/export tool
const PdfExportInputSchema = z.object({
  filePath: z.string().describe('Path to the source Office file (Word, Excel, or PowerPoint).'),
  outputFilePath: z.string().describe('Path where the output PDF file will be saved.'),
  application: z.enum(['Word', 'Excel', 'PowerPoint']).describe('The Office application to use.'),
  configuration: z.record(z.any()).optional().describe('Optional configuration for PDF export (e.g., page range, quality).'),
});

type PdfExportInput = z.infer<typeof PdfExportInputSchema>;

/**
 * @tool office/pdf/export
 * @description Tool for exporting Office documents (Word, Excel, PowerPoint) to PDF format.
 * Uses COM Interop via winax to interact with Office applications.
 * Supports operations to convert, configure (optional), and save the PDF file.
 * @operation convert - Converts the specified Office file to PDF.
 * @operation configure - (Optional) Allows configuring export options before converting. (Basic implementation)
 * @operation save - Saves the converted file to the specified path. (Integrated into convert)
 * @param filePath - Path to the source Office file.
 * @param outputFilePath - Path where the output PDF file will be saved.
 * @param application - The Office application to use ('Word', 'Excel', 'PowerPoint').
 * @param configuration - Optional configuration for PDF export.
 */
export const pdfExportTool: McpResource = {
  path: 'office/pdf/export', // Use 'path' instead of 'name'
  description: 'Exports Office documents to PDF.',
  schema: PdfExportInputSchema, // Use 'schema' instead of 'inputSchema'
  handler: async (params: ToolRequestParams, context: any): Promise<ApiResponse<any>> => { // Adjust return type
    let officeApp: any | null = null;
    let doc = null;

    try {
      // Validate input parameters
      const validationResult = PdfExportInputSchema.safeParse(params); // Validate params directly

      if (!validationResult.success) {
        return {
          success: false,
          error: { // Adjust error structure
            code: 'VALIDATION_ERROR',
            message: 'Input validation error',
            details: validationResult.error.errors,
          },
        };
      }

      const { filePath, outputFilePath, application, configuration } = validationResult.data;

      // Get the Office application
      // Map the application name to its ProgID
      let appProgId: OfficeAppName;
      if (application === 'Word') {
        appProgId = 'Word.Application';
      } else if (application === 'Excel') {
        appProgId = 'Excel.Application';
      } else if (application === 'PowerPoint') {
        appProgId = 'PowerPoint.Application';
      } else {
         // Although schema validation should catch this, keep as a fallback
         return { // Return as ErrorResponse
           success: false,
           error: {
             code: 'UNSUPPORTED_APPLICATION',
             message: `Unsupported Office application: ${application}`,
           },
         };
      }

      officeApp = await getOfficeApplication(appProgId);
      if (!officeApp) {
        return { // Return as ErrorResponse
           success: false,
           error: {
             code: 'OFFICE_APP_ERROR',
             message: `Could not get Office application instance: ${application}`,
           },
         };
      }

      // Open the document/presentation
      if (application === 'Word') {
        doc = officeApp.Documents.Open(filePath);
      } else if (application === 'Excel') {
        doc = officeApp.Workbooks.Open(filePath);
      } else if (application === 'PowerPoint') {
        doc = officeApp.Presentations.Open(filePath); // Corrected 'Presenations' to 'Presentations'
      }

      if (!doc) {
        return { // Return as ErrorResponse
           success: false,
           error: {
             code: 'FILE_OPEN_ERROR',
             message: `Could not open file: ${filePath}`,
           },
         };
      }

      // Configure export options (basic implementation, can be expanded)
      const exportConfig = configuration || {};
      const exportFormat = 17; // wdExportFormatPDF, xlTypePDF, ppSaveAsPDF (generally 17)

      // Perform the export to PDF
      if (application === 'Word') {
        // Word uses ExportAsFixedFormat
        doc.ExportAsFixedFormat(
          outputFilePath,
          exportFormat,
          false, // OpenAfterExport
          0, // OptimizeFor (0=Print, 1=Screen)
          0, // Range (0=Whole document)
          0, // From
          0, // To
          0, // Item (0=wdExportDocumentContent)
          true, // IncludeDocProperties
          true, // KeepIRM
          0, // CreateBookmarks (0=wdExportCreateNoBookmarks)
          true, // DocStructureTags
          false, // BitmapMissingFonts
          false, // UseISO19005_1
          exportConfig.OptimizeFor || 0,
          exportConfig.BitmapMissingFonts || false,
          exportConfig.IncludeDocProperties || true,
          exportConfig.KeepIRM || true,
          exportConfig.CreateBookmarks || 0,
          exportConfig.DocStructureTags || true,
          exportConfig.UseISO19005_1 || false,
          exportConfig.Range || 0,
          exportConfig.From || 0,
          exportConfig.To || 0,
          exportConfig.Item || 0,
          exportConfig.OpenAfterExport || false
        );
        // Save the exported PDF file as a dynamic resource
        try {
            const pdfContent = await fs.readFile(outputFilePath, null); // Read as Buffer
            await saveResource('office/pdf/export', path.basename(outputFilePath), pdfContent);
            context?.log.info(`Saved ${outputFilePath} as a dynamic resource.`);
        } catch (resourceSaveError: any) {
            context?.log.error(`Failed to save ${outputFilePath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continue execution even if resource saving fails
        }
      } else if (application === 'Excel') {
        // Excel uses ExportAsFixedFormat
        doc.ExportAsFixedFormat(
          exportFormat,
          outputFilePath,
          exportConfig.Quality || 0, // xlQualityStandard (0) or xlQualityMinimum (1)
          exportConfig.IncludeDocProperties || true,
          exportConfig.IgnorePrintAreas || false,
          exportConfig.From || 0,
          exportConfig.To || 0,
          exportConfig.OpenAfterPublish || false,
          exportConfig.FixedFormatExtClassPtr || null
        );
        // Save the exported PDF file as a dynamic resource
        try {
            const pdfContent = await fs.readFile(outputFilePath, null); // Read as Buffer
            await saveResource('office/pdf/export', path.basename(outputFilePath), pdfContent);
            context?.log.info(`Saved ${outputFilePath} as a dynamic resource.`);
        } catch (resourceSaveError: any) {
            context?.log.error(`Failed to save ${outputFilePath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continue execution even if resource saving fails
        }
      } else if (application === 'PowerPoint') {
        // PowerPoint uses ExportAsFixedFormat
        doc.ExportAsFixedFormat(
          outputFilePath,
          exportFormat,
          exportConfig.Intent || 1, // ppExportIntentPrint (1) or ppExportIntentScreen (2)
          exportConfig.FrameSlides || false,
          exportConfig.IncludeHiddenSlides || false,
          exportConfig.PrintRange || null, // ppPrintAll (1), ppPrintSelection (2), ppPrintCurrent (3), ppPrintSlideRange (4)
          exportConfig.From || 0,
          exportConfig.To || 0,
          exportConfig.OutputType || 1, // ppPrintOutputSlides (1), ppPrintOutputNotesPages (2), ppPrintOutputOutline (3), ppPrintOutputHandouts (4)
          exportConfig.PrintHandoutOrder || 1, // ppPrintHandoutHorizontal (1), ppPrintHandoutVertical (2)
          exportConfig.FitToPage || false,
          exportConfig.Orientation || 1, // ppPrintOrientationPortrait (1), ppPrintOrientationLandscape (2)
          exportConfig.ViewSlideDetails || false,
          exportConfig.ShowComments || false,
          exportConfig.ShowInking || false,
          exportConfig.CreateHiddenSlides || false,
          exportConfig.OpenAfterExport || false
        );
        // Save the exported PDF file as a dynamic resource
        try {
            const pdfContent = await fs.readFile(outputFilePath, null); // Read as Buffer
            await saveResource('office/pdf/export', path.basename(outputFilePath), pdfContent);
            context?.log.info(`Saved ${outputFilePath} as a dynamic resource.`);
        } catch (resourceSaveError: any) {
            context?.log.error(`Failed to save ${outputFilePath} as a dynamic resource: ${resourceSaveError.message}`);
            // Continue execution even if resource saving fails
        }
      }


      return { // Adjust success structure
        success: true,
        data: {
          outputFilePath: outputFilePath,
          message: `File exported to PDF successfully: ${outputFilePath}`,
        },
      };

    } catch (error: any) {
      return { // Adjust error structure
        success: false,
        error: {
          code: 'EXPORT_ERROR',
          message: `Error exporting to PDF: ${error.message}`,
          details: error,
        },
      };
    } finally {
      // Close the document and release COM objects
      if (doc) {
        try {
          doc.Close();
        } catch (e) {
          console.error('Error closing the document:', e);
        }
      }
      // Do not close the Office application here, as there may be other documents open
      // Management of the Office application instance should be handled externally if necessary.
    }
  },
};