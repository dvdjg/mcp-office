import { exec } from 'child_process';
import { McpResource, ToolRequestParams, ApiResponse } from '@/types/common.types.js';
import { z } from 'zod';

// Define an interface for the structure of an active Office document
interface ActiveOfficeDocument {
  filePath: string;
  applicationType: 'Word' | 'Excel' | 'PowerPoint';
}

// Define the input interface for the tool (currently no specific inputs)
interface GetActiveOfficeDocumentsInput extends ToolRequestParams {}

// Define the output interface for the tool
interface GetActiveOfficeDocumentsOutputData { // Renamed to reflect it's the data part of ApiResponse
  documents: ActiveOfficeDocument[];
}

// Helper function to execute PowerShell commands
function executePowerShell(script: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const command = `powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& {${script}}"`;

    exec(command, (error, stdout, stderr) => {
      if (error) {
        return reject(new Error(`PowerShell script execution failed: ${error.message}. Stderr: ${stderr}`));
      }
      if (stderr) {
        // Stderr might contain warnings or errors from within the script if not caught by try/catch in PS.
        // For this tool, PowerShell's try/catch should handle most issues, returning empty output.
        // console.warn(`[os/getActiveOfficeDocuments] PowerShell script stderr: ${stderr}`);
      }
      resolve(stdout.trim());
    });
  });
}

const getActiveOfficeDocumentsTool: McpResource = {
  path: 'os/getActiveOfficeDocuments',
  description: 'Lists all currently open Microsoft Office documents (Word, Excel, PowerPoint) and their full file paths. Windows only.',
  schema: z.object({}), // Correctly defined as an empty Zod object schema
  // outputSchema removed as it's not part of McpResource
  async handler(input: GetActiveOfficeDocumentsInput): Promise<ApiResponse<GetActiveOfficeDocumentsOutputData>> {
    const documents: ActiveOfficeDocument[] = [];

    if (process.platform !== 'win32') {
      // console.warn('[os/getActiveOfficeDocuments] This tool is currently only supported on Windows.');
      return { success: true, data: { documents: [] } };
    }

    const officeApps = [
      { comObject: 'Word.Application', type: 'Word' as const },
      { comObject: 'Excel.Application', type: 'Excel' as const },
      { comObject: 'PowerPoint.Application', type: 'PowerPoint' as const },
    ];

    for (const app of officeApps) {
      let script = '';
      switch (app.type) {
        case 'Word':
          script = `
            try {
              $appInstance = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${app.comObject}')
              $docPaths = @()
              foreach ($document in $appInstance.Documents) {
                try {
                  if ($document.FullName -and $document.FullName.Trim() -ne '') {
                    $docPaths += $document.FullName
                  }
                } catch { /* Ignore errors for individual documents */ }
              }
              Write-Output ($docPaths -join ';')
            } catch { Write-Output "" }`; // App not running or other COM error
          break;
        case 'Excel':
          script = `
            try {
              $appInstance = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${app.comObject}')
              $wbPaths = @()
              foreach ($workbook in $appInstance.Workbooks) {
                try {
                  if ($workbook.FullName -and $workbook.FullName.Trim() -ne '') {
                    $wbPaths += $workbook.FullName
                  }
                } catch { /* Ignore errors for individual workbooks */ }
              }
              Write-Output ($wbPaths -join ';')
            } catch { Write-Output "" }`;
          break;
        case 'PowerPoint':
          script = `
            try {
              $appInstance = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${app.comObject}')
              $presPaths = @()
              foreach ($presentation in $appInstance.Presentations) {
                try {
                  if ($presentation.FullName -and $presentation.FullName.Trim() -ne '') {
                    $presPaths += $presentation.FullName
                  }
                } catch { /* Ignore errors for individual presentations */ }
              }
              Write-Output ($presPaths -join ';')
            } catch { Write-Output "" }`;
          break;
      }

      if (script) {
        try {
          const output = await executePowerShell(script);
          const filePaths = output.split(';').map(p => p.trim()).filter(p => p.length > 0);

          for (const filePath of filePaths) {
            const lcFilePath = filePath.toLowerCase();
            let isValidExtension = false;
            if (app.type === 'Word' && (lcFilePath.endsWith('.docx') || lcFilePath.endsWith('.doc') || lcFilePath.endsWith('.docm'))) isValidExtension = true;
            if (app.type === 'Excel' && (lcFilePath.endsWith('.xlsx') || lcFilePath.endsWith('.xls') || lcFilePath.endsWith('.xlsm') || lcFilePath.endsWith('.xlsb'))) isValidExtension = true;
            if (app.type === 'PowerPoint' && (lcFilePath.endsWith('.pptx') || lcFilePath.endsWith('.ppt') || lcFilePath.endsWith('.pptm'))) isValidExtension = true;

            // Basic check for an absolute path (contains ':') and a valid extension
            if (filePath.includes(':') && isValidExtension) {
                 documents.push({ filePath, applicationType: app.type });
            } else {
                // console.warn(`[os/getActiveOfficeDocuments] Filtered out potentially invalid path for ${app.type}: '${filePath}'`);
            }
          }
        } catch (err) {
          // console.warn(`[os/getActiveOfficeDocuments] Error querying ${app.type} documents: ${err.message}. App might not be running or no documents open.`);
        }
      }
    }
    return { success: true, data: { documents } };
  },
};

export default getActiveOfficeDocumentsTool;