import { McpResource, ToolRequestParams, ApiResponse } from '@/types/common.types.js';
import { z } from 'zod';
import logger from '../../utils/logger.js';
import {
  getOfficeApplication,
  getOpenWordDocuments,
  getOpenExcelWorkbooks,
  getOpenPowerPointPresentations,
  releaseObject,
  OfficeAppName,
} from '../../utils/officeInterop.js';

// Define an interface for the structure of an active Office document
export interface ActiveOfficeDocument {
  filePath: string;
  applicationType: 'Word' | 'Excel' | 'PowerPoint';
}

// Define the input interface for the tool (currently no specific inputs)
interface GetActiveOfficeDocumentsInput extends ToolRequestParams {}

// Define the output interface for the tool
interface GetActiveOfficeDocumentsOutputData {
  documents: ActiveOfficeDocument[];
}

const getActiveOfficeDocumentsTool: McpResource = {
  path: 'os/getActiveOfficeDocuments',
  description: 'Lists all currently open Microsoft Office documents (Word, Excel, PowerPoint) and their full file paths using winax COM interop. Windows only.',
  schema: z.object({}),
  async handler(input: GetActiveOfficeDocumentsInput): Promise<ApiResponse<GetActiveOfficeDocumentsOutputData>> {
    const allOpenDocuments: ActiveOfficeDocument[] = [];

    if (process.platform !== 'win32') {
      logger.warn('[os/getActiveOfficeDocuments] This tool is currently only supported on Windows. Returning empty list.');
      return { success: true, data: { documents: [] } };
    }

    const officeAppDefinitions: { name: OfficeAppName; type: ActiveOfficeDocument['applicationType']; getter: (app: any) => Promise<string[]> }[] = [
      { name: 'Word.Application', type: 'Word', getter: getOpenWordDocuments },
      { name: 'Excel.Application', type: 'Excel', getter: getOpenExcelWorkbooks },
      { name: 'PowerPoint.Application', type: 'PowerPoint', getter: getOpenPowerPointPresentations },
    ];

    for (const appDef of officeAppDefinitions) {
      let appInstance: any = null;
      try {
        logger.info(`[os/getActiveOfficeDocuments] Attempting to connect to ${appDef.name}...`);
        appInstance = await getOfficeApplication(appDef.name);

        if (appInstance) {
          logger.info(`[os/getActiveOfficeDocuments] Successfully connected to ${appDef.name}. Fetching open documents...`);
          const filePaths = await appDef.getter(appInstance);
          filePaths.forEach(filePath => {
            if (filePath && typeof filePath === 'string' && filePath.trim() !== '') {
              allOpenDocuments.push({ filePath, applicationType: appDef.type });
              logger.debug(`[os/getActiveOfficeDocuments] Found open ${appDef.type} document: ${filePath}`);
            } else {
              logger.warn(`[os/getActiveOfficeDocuments] Invalid or empty file path received for an open ${appDef.type} document.`);
            }
          });
          logger.info(`[os/getActiveOfficeDocuments] Found ${filePaths.length} open document(s) for ${appDef.name}.`);
        } else {
          logger.info(`[os/getActiveOfficeDocuments] No active instance of ${appDef.name} found or could not connect.`);
        }
      } catch (error) {
        // Log error if getOfficeApplication fails (e.g., app not installed or running with no docs)
        // Or if the specific getter (getOpenWordDocuments etc.) fails.
        const errorMessage = error instanceof Error ? error.message : String(error);
        if (errorMessage.includes('RPC_E_CALL_REJECTED') || errorMessage.includes('OLE error 8001010A')) {
             logger.info(`[os/getActiveOfficeDocuments] ${appDef.name} might be busy or not responding. Skipping. Error: ${errorMessage}`);
        } else if (errorMessage.includes('Failed to get application') && (errorMessage.includes('Object is not connected to server') || errorMessage.includes('Invalid class string'))) {
            // This often means the application is not running or no documents are open, which is not an error for this tool's purpose.
            logger.info(`[os/getActiveOfficeDocuments] No running instance of ${appDef.name} found or it has no open documents. Error: ${errorMessage}`);
        } else {
            logger.error(`[os/getActiveOfficeDocuments] Error processing ${appDef.name}: ${errorMessage}`, { error });
        }
      } finally {
        if (appInstance) {
          // Release the application object.
          // Important: Releasing the main application object might close it if it was newly created
          // and has no visible windows/documents. If it was an existing instance,
          // releasing it here generally detaches our script from it, not closes the user's app.
          // The individual document/workbook/presentation objects are released within their respective getter functions.
          releaseObject(appInstance);
          logger.info(`[os/getActiveOfficeDocuments] Released ${appDef.name} instance.`);
        }
      }
    }

    logger.info(`[os/getActiveOfficeDocuments] Final combined list of open documents: ${JSON.stringify(allOpenDocuments)}`);
    return { success: true, data: { documents: allOpenDocuments } };
  },
};

export default getActiveOfficeDocumentsTool;