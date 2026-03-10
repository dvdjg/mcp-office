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
  DocumentPathInfo, // Import the new interface
  getCloudUrlLocalPath,
} from '../../utils/officeInterop.js';
import { join as joinPath, isAbsolute } from 'path'; // For path manipulation

// Define an interface for the structure of an active Office document
export interface ActiveOfficeDocument {
  fullName: string; // Original full path or URL from COM (e.g., "C:\\path\\to\\doc.docx" or "https://tenant-my.sharepoint.com/...")
  path: string;     // Directory path from COM, if available (e.g., "C:\\path\\to\\" or "https://tenant-my.sharepoint.com/...")
  name: string;     // Filename from COM (e.g., "doc.docx")
  resolvedPath: string | null; // The local file system path.
                           // - If fullName is a local path, resolvedPath is fullName.
                           // - If fullName is a cloud URL and successfully converted, resolvedPath is the local equivalent.
                           // - If fullName is a cloud URL and conversion fails, or if fullName is empty, resolvedPath is null.
  applicationType: 'Word' | 'Excel' | 'PowerPoint';
  isLocal: boolean; // True if resolvedPath contains a local file system path (i.e., resolvedPath is not null and is not a URL).
                           // False if resolvedPath is null (cloud URL that failed to convert or empty fullName)
                           // or if resolvedPath is still a URL (should not happen with current logic but good to note).
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

    const officeAppDefinitions: { name: OfficeAppName; type: ActiveOfficeDocument['applicationType']; getter: (app: any) => Promise<DocumentPathInfo[]> }[] = [
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
          const documentInfos = await appDef.getter(appInstance);
          for (const docInfo of documentInfos) {
            if (docInfo && docInfo.name) { // Ensure we have at least a name
              const originalPath = docInfo.fullName;
              let resolvedPath: string | null = null;
              let isLocal = false;

              if (originalPath && originalPath.match(/^https?:\/\//i)) {
                // It's a cloud URL
                const localPathFromCloud = await getCloudUrlLocalPath(originalPath);
                if (localPathFromCloud) {
                  resolvedPath = localPathFromCloud;
                  isLocal = true;
                  logger.debug(`[os/getActiveOfficeDocuments] Resolved cloud URL '${originalPath}' to local path '${resolvedPath}'`);
                } else {
                  // Conversion failed for cloud URL
                  resolvedPath = null;
                  isLocal = false;
                  logger.debug(`[os/getActiveOfficeDocuments] Could not resolve cloud URL '${originalPath}' to a local path. resolvedPath is null.`);
                }
              } else if (originalPath) {
                // Not a cloud URL, assume it's a local path
                resolvedPath = originalPath;
                isLocal = true; // If it's not a cloud URL and it exists, assume it's local.
                logger.debug(`[os/getActiveOfficeDocuments] '${originalPath}' is not a cloud URL. Assuming local path. resolvedPath is '${resolvedPath}'`);
              } else {
                // originalPath (docInfo.fullName) is null or empty
                resolvedPath = null;
                isLocal = false;
                logger.debug(`[os/getActiveOfficeDocuments] docInfo.fullName is null or empty. resolvedPath is null.`);
              }
              
              allOpenDocuments.push({
                fullName: originalPath, // This is docInfo.fullName
                path: docInfo.path,
                name: docInfo.name,
                resolvedPath, // This is the newly determined resolvedPath
                applicationType: appDef.type,
                isLocal, // This is the newly determined isLocal
              });
              logger.debug(`[os/getActiveOfficeDocuments] Found open ${appDef.type} document: Name='${docInfo.name}', FullName='${originalPath}', Path='${docInfo.path}', ResolvedPath='${resolvedPath}', IsLocal=${isLocal}`);
            } else {
              logger.warn(`[os/getActiveOfficeDocuments] Invalid or empty document info received for an open ${appDef.type} document (docInfo or docInfo.name is null/undefined). Info: ${JSON.stringify(docInfo)}`);
            }
          }
          logger.info(`[os/getActiveOfficeDocuments] Processed ${documentInfos.length} document entries for ${appDef.name}.`);
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
