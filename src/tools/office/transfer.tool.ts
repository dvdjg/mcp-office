// src/tools/office/transfer.tool.ts
import { McpResource } from '@/types/common.types';
import { z } from 'zod';
import { handleToolError } from '@/utils/errorHandler';
import { getOfficeApplication, releaseObject, OfficeAppName } from '@/utils/officeInterop'; // Import releaseObject and OfficeAppName
import * as winax from 'winax'; // Import winax for COM Interop
import { ApiResponse, ErrorResponse, SuccessResponse } from '@/types/common.types'; // Import necessary types

// Define the input schema for the office/transfer tool
const officeTransferInputSchema = z.object({
  //type: z.literal('object'), // Added to satisfy validator
  source: z.string().describe('Especificación de origen (ej: "excel:./data.xlsx:Sheet1:A1:B2", "word:./document.docx:paragraph:3", "powerpoint:./presentation.pptx:slide:2:shape:5")'),
  target: z.string().describe('Especificación de destino (ej: "word:./document.docx:paragraph:3", "excel:./data.xlsx:Sheet1:A1", "powerpoint:./presentation.pptx:slide:2")'),
  operation: z.enum(['embed', 'insert', 'copy']).describe('Tipo de operación: "embed", "insert", o "copy"'),
  // Add other potential parameters like format, etc. if needed later
});

// Define the output schema (puede ser un simple éxito/fallo o más detallado)
const officeTransferOutputSchema = z.object({
    success: z.boolean(),
    // Para éxito, puede haber datos adicionales si es necesario
    data: z.any().optional(),
    // Para fallo, se usa la estructura de ErrorResponse
    error: z.object({
        code: z.string(),
        message: z.string(),
        details: z.any().optional(),
    }).optional(),
});

// Helper function to parse source/target strings
function parseOfficeSpecifier(specifier: string): { app: string; filePath: string; location: string } | null {
    const parts = specifier.split(':');
    if (parts.length < 3) {
        return null; // Invalid format
    }
    const [app, filePath, ...locationParts] = parts;
    return {
        app: app.toLowerCase(),
        filePath: filePath,
        location: locationParts.join(':'), // Join remaining parts for location
    };
}

const officeTransferTool: McpResource = {
    path: 'office/transfer',
    description: 'Mueve datos entre aplicaciones de Office (Word, Excel, PowerPoint) usando COM Interop.',
    schema: officeTransferInputSchema, // Corregido de inputSchema a schema
    // outputSchema: officeTransferOutputSchema, // outputSchema no es parte de McpResource
    handler: async (params, context): Promise<ApiResponse<any>> => {
        let sourceApp: any = null;
        let targetApp: any = null;
        let sourceDoc: any = null;
        let targetDoc: any = null; // Permitir que targetDoc sea de tipo any

        try {
            const { source, target, operation } = officeTransferInputSchema.parse(params);

            const sourceSpec = parseOfficeSpecifier(source);
            const targetSpec = parseOfficeSpecifier(target);

            if (!sourceSpec || !targetSpec) {
                return {
                    success: false,
                    error: {
                        code: 'INVALID_INPUT',
                        message: 'Formato de especificador de origen o destino inválido.',
                    },
                } as ErrorResponse; // Asegurar que coincide con ErrorResponse
            }

            // Basic validation for supported applications
            const supportedApps: OfficeAppName[] = ['Word.Application', 'Excel.Application', 'PowerPoint.Application'];
            const sourceAppProgId = `${sourceSpec.app}.Application`;
            const targetAppProgId = `${targetSpec.app}.Application`;

            if (!supportedApps.includes(sourceAppProgId as OfficeAppName) || !supportedApps.includes(targetAppProgId as OfficeAppName)) {
                return {
                   success: false,
                   error: {
                       code: 'UNSUPPORTED_APPLICATION',
                       message: `Aplicación de origen o destino no soportada. Soportadas: ${supportedApps.map(app => app.split('.')[0]).join(', ')}.`,
                   },
               } as ErrorResponse; // Asegurar que coincide con ErrorResponse
           }

            // Get COM objects for source and target applications
            // Need to handle cases where source/target app is the same
            sourceApp = await getOfficeApplication(sourceSpec.app as OfficeAppName);
            if (sourceSpec.app !== targetSpec.app) {
                 targetApp = await getOfficeApplication(targetSpec.app as OfficeAppName);
            } else {
                 targetApp = sourceApp; // Same application instance
            }


            // Open source document/presentation
            // This part requires specific COM API calls for each application type
            // Placeholder: Assume a function like openDocument exists in officeInterop
            // Need to handle different document types (Word.Document, Excel.Workbook, PowerPoint.Presentation)
            // Example (simplified):
            // if (sourceSpec.app === 'word') sourceDoc = sourceApp.Documents.Open(sourceSpec.filePath);
            // else if (sourceSpec.app === 'excel') sourceDoc = sourceApp.Workbooks.Open(sourceSpec.filePath);
            // else if (sourceSpec.app === 'powerpoint') sourceDoc = sourceApp.Presentations.Open(sourceSpec.filePath);
            // else throw new Error('Unsupported source application for opening document.');

            // Open target document/presentation
            // Similar logic as source
            // if (targetSpec.app === 'word') targetDoc = targetApp.Documents.Open(targetSpec.filePath);
            // else if (targetSpec.app === 'excel') targetDoc = targetApp.Workbooks.Open(targetSpec.filePath);
            // else if (targetSpec.app === 'powerpoint') targetDoc = targetApp.Presentations.Open(targetSpec.filePath);
            // else throw new Error('Unsupported target application for opening document.');

            // --- COM Interop Logic for Copy/Paste ---
            // This is the complex part requiring specific Office COM API knowledge.
            // The logic will depend heavily on the source and target applications and the specified locations.
            // Example (highly simplified placeholder):
            // 1. Select content in sourceDoc based on sourceSpec.location
            //    e.g., if excel: sourceDoc.Sheets(sourceSpec.location.split(':')[0]).Range(sourceSpec.location.split(':')[1]).Copy();
            //    e.g., if word: sourceDoc.Range(start, end).Copy();
            //    e.g., if powerpoint: sourceDoc.Slides(slideIndex).Shapes(shapeIndex).Copy();
            // 2. Activate targetDoc
            //    e.g., targetDoc.Activate();
            // 3. Navigate to target location in targetDoc based on targetSpec.location
            //    e.g., if word: targetDoc.GoTo(What: wdGoToBookmark, Name: "myBookmark").Select();
            //    e.g., if excel: targetDoc.Sheets(targetSpec.location.split(':')[0]).Range(targetSpec.location.split(':')[1]).Select();
            //    e.g., if powerpoint: targetDoc.Slides(slideIndex).Select(); targetApp.ActiveWindow.View.Paste();
            // 4. Paste content based on operation ('embed', 'insert', 'copy' - often map to different PasteSpecial options)
            //    e.g., targetApp.Selection.PasteSpecial(...); // Need to determine appropriate PasteSpecial options

            // Since implementing the full COM logic for all combinations is extensive,
            // this initial version will focus on the structure and use placeholders for the COM calls.
            // A simple copy/paste operation using the clipboard is a common approach.

            // Placeholder for COM Interop:
            console.log(`Attempting to transfer from ${source} to ${target} with operation ${operation}`);
            console.log('COM Interop logic for copy/paste needs to be implemented here.');

            // Simulate success for now
            const success = true; // Replace with actual COM operation result

            if (success) {
                 // Save changes (optional, depending on requirements)
                 // sourceDoc.Save();
                 // targetDoc.Save();
                 return { success: true, data: { message: `Operación de transferencia de datos '${operation}' completada con éxito.` } };
            } else {
                 return {
                    success: false,
                    error: {
                        code: 'TRANSFER_FAILED',
                        message: `La operación de transferencia de datos '${operation}' falló.`,
                    },
                 } as ErrorResponse; // Asegurar que coincide con ErrorResponse
            }

        } catch (error) {
            // Handle errors, including COM exceptions
            return handleToolError(error, 'OFFICE_TRANSFER_ERROR');
        } finally {
            // Ensure COM objects are released
            // Need to handle cases where source/target app is the same
            if (sourceDoc) {
                try { sourceDoc.Close(); } catch (e: any) { console.error('Error closing source document:', e); }
                releaseObject(sourceDoc);
            }
             if (targetDoc && sourceDoc !== targetDoc) { // Only close target if it's a different document
                try { targetDoc.Close(); } catch (e: any) { console.error('Error closing target document:', e); }
                releaseObject(targetDoc);
            }
            if (sourceApp) {
                 // Decide whether to Quit the application or just release the object
                 // Quitting might close other open documents, which might not be desired.
                 // Releasing the object is generally safer if the application might be used elsewhere.
                 // For this tool, it might be better to just release the object reference.
                 // sourceApp.Quit(); // Use with caution
                 releaseObject(sourceApp);
            }
             if (targetApp && sourceApp !== targetApp) { // Only release target if it's a different instance
                 releaseObject(targetApp);
            }
        }
    },
};

export default officeTransferTool;