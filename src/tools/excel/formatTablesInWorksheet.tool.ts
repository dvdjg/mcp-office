/**
 * @file Tool for formatting tables within a specific worksheet of an Excel file.
 * @author Roo
 * @date 2025-05-15
 * @copyright Copyright (c) 2025 Your Name/Company
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse, FastMCPContext } from '../../types/common.types.js'; // Added FastMCPContext
import { getOfficeApplication, releaseObject, OfficeAppName } from '../../utils/officeInterop.js'; // Added OfficeAppName
import { saveResource } from '../dynamic/resources.tool.js';
import * as fs from 'fs-extra';
import * as path from 'path';
import ExcelJS, { Table } from 'exceljs'; // Import Table type
import logger from '../../utils/logger.js';
import getActiveOfficeDocumentsTool, { ActiveOfficeDocument } from '../os/getActiveOfficeDocuments.tool.js';
import excelRangeTool from './range.tool.js'; // Import excelRangeTool
import { aiSuggestTool } from '../office/aiSuggest.tool.js'; // Import aiSuggestTool
import excelTablesTool from './tables.tool.js'; // Import excelTablesTool - CORRECTED EXTENSION

// Define the input schema for the excel/formatTablesInWorksheet tool
const FormatTablesInWorksheetInputSchema = z.object({
  inputExcelPath: z.string().optional().describe('Optional path to the Excel file. If not provided, the tool will try to find an active Excel document.'),
  worksheetIdentifier: z.union([
    z.string().min(1),
    z.number().int().positive()
  ]).optional().default(2).describe('Identifier for the worksheet (name or 1-based index). Defaults to 2 (the second worksheet).'),
  useComInterop: z.boolean().optional().default(false).describe('Use COM interop for local Excel interaction. Defaults to false (uses exceljs).'),
});

type FormatTablesInWorksheetInput = z.infer<typeof FormatTablesInWorksheetInputSchema>;

/**
 * @tool excel/formatTablesInWorksheet
 * @description Identifies tables in a specified Excel worksheet and applies AI-suggested formatting.
 * If no Excel path is provided, it attempts to use the active Excel document.
 * The target worksheet defaults to the second sheet if not specified.
 * @input FormatTablesInWorksheetInputSchema
 * @output An object indicating success or failure, with details of actions taken or errors encountered.
 */
export const formatTablesInWorksheetTool: McpResource[] = [{
  path: 'excel/formatTablesInWorksheet',
  description: 'Identifies tables in a specified Excel worksheet and applies AI-suggested formatting.',
  schema: FormatTablesInWorksheetInputSchema,
  handler: async (params: ToolRequestParams, context?: FastMCPContext<any>): Promise<ApiResponse<any>> => { // Added context
    let input: FormatTablesInWorksheetInput;
    try {
      input = FormatTablesInWorksheetInputSchema.parse(params);
    } catch (error: any) {
      if (error instanceof z.ZodError) {
        return { success: false, error: { code: 'VALIDATION_ERROR', message: 'Input validation failed', details: error.errors } };
      }
      return { success: false, error: { code: 'UNEXPECTED_PARSING_ERROR', message: `An unexpected error occurred during input parsing: ${error.message}` } };
    }

    const { inputExcelPath, worksheetIdentifier, useComInterop } = input;
    let absoluteFilePath: string | undefined = inputExcelPath ? path.resolve(inputExcelPath) : undefined;
    logger.info(`Starting formatTablesInWorksheet tool with input: ${JSON.stringify(input)}`);

    try {
      // Step 2: Identify Target Excel Document and Worksheet
      if (!absoluteFilePath) {
        logger.info('No inputExcelPath provided, attempting to find active Excel document.');
        const activeDocsResponse = await getActiveOfficeDocumentsTool.handler({});
        if (!activeDocsResponse.success) {
          // Type guard ensures activeDocsResponse is ErrorResponse here
          throw new Error(`Failed to get active office documents: ${activeDocsResponse.error?.message || 'Unknown error'}`);
        }
        // If success is true, activeDocsResponse.data should be defined.
        // Adding an explicit check for robustness, though TypeScript might infer it.
        if (!activeDocsResponse.data) {
            throw new Error('Failed to get active office documents: Response data is missing.');
        }

        const activeExcelDocs = activeDocsResponse.data.documents.filter(
          (doc: ActiveOfficeDocument) => doc.applicationType === 'Excel'
        );

        if (activeExcelDocs.length === 0) {
          throw new Error('No active Excel document found. Please open an Excel file or provide inputExcelPath.');
        } else if (activeExcelDocs.length > 1) {
          const docPaths = activeExcelDocs.map((doc: ActiveOfficeDocument) => doc.resolvedPath).join(', ');
          throw new Error(`Multiple active Excel documents found: ${docPaths}. Please specify the target file using inputExcelPath or close other Excel files.`);
        } else {
          absoluteFilePath = path.resolve(activeExcelDocs[0].resolvedPath);
          logger.info(`Using active Excel document: ${absoluteFilePath}`);
        }
      }

      if (!absoluteFilePath) { // Should not happen if logic above is correct, but as a safeguard
        throw new Error('Could not determine the target Excel file path.');
      }
      
      logger.info(`Target Excel file: ${absoluteFilePath}`);
      logger.info(`Target worksheet identifier: ${worksheetIdentifier}`);

      if (useComInterop) {
        // COM Interop Path
        logger.info('Using COM Interop path.');
        let officeAppInstance: any = null;
        let excelApp: any = null;
        let workbook: any = null;
        let worksheet: any = null;
        let listObjects: any = null;
        const identifiedTables: { name: string, rangeAddress: string }[] = [];
        let changesMadeToWorkbook = false; // Initialize here

        try {
          officeAppInstance = await getOfficeApplication('Excel.Application');
          excelApp = officeAppInstance.app;
          excelApp.Visible = false; // Keep Excel hidden

          try {
            workbook = await excelApp.Workbooks.Open(absoluteFilePath);
          } catch (e: any) {
            logger.error(`COM: Failed to open workbook ${absoluteFilePath}: ${e.message}`);
            throw new Error(`COM: Failed to open workbook ${absoluteFilePath}. Ensure the file exists. Error: ${e.message}`);
          }

          if (typeof worksheetIdentifier === 'string') {
            worksheet = await workbook.Sheets(worksheetIdentifier);
          } else { // worksheetIdentifier is a number (1-based index)
            worksheet = await workbook.Sheets(worksheetIdentifier);
          }

          if (!worksheet) {
            throw new Error(`COM: Worksheet with identifier '${worksheetIdentifier}' not found in ${absoluteFilePath}.`);
          }
          const actualWorksheetName = await worksheet.Name;
          logger.info(`COM: Successfully identified worksheet: '${actualWorksheetName}'`);

          listObjects = await worksheet.ListObjects;
          const tableCount = await listObjects.Count;

          if (tableCount === 0) {
            logger.info(`COM: No tables found in worksheet '${actualWorksheetName}'.`);
            return { success: true, data: { message: `No tables found in worksheet '${actualWorksheetName}'. No formatting applied.` } };
          }

          logger.info(`COM: Found ${tableCount} table(s) in worksheet '${actualWorksheetName}'.`);

          for (let i = 1; i <= tableCount; i++) {
            const table = await listObjects.Item(i);
            const tableName = await table.Name;
            const tableRange = await table.Range.Address();
            identifiedTables.push({ name: tableName, rangeAddress: tableRange });
            logger.debug(`COM: Table found: ${tableName}, Range: ${tableRange}`);
            releaseObject(table); // Release individual table object after use
          }

          // TODO: Steps 4 & 5 for COM path
          // For each table in identifiedTables:
          //    a. Extract sample data (range.tool.ts or direct COM)
          //    b. Call aiSuggest.tool.ts
          //    c. Parse suggestions
          //    d. Apply formatting (tables.tool.ts, range.tool.ts or direct COM)
          logger.info(`COM: Identified tables: ${JSON.stringify(identifiedTables)}`);

          const allSuggestions: { tableName: string, suggestion: any, error?: string, appliedFormats?: string[] }[] = [];

          for (const tableInfo of identifiedTables) {
            try {
              logger.info(`COM: Processing table '${tableInfo.name}' in range '${tableInfo.rangeAddress}'`);

              // Step 4a: Extract sample data (first few rows, e.g., up to 5 rows including header)
              // We need to determine the actual number of rows to fetch.
              // For simplicity, let's try to get the whole table data for now.
              // A more refined approach would be to get headers and a few data rows.
              const rangeDataResponse = await excelRangeTool.handler({
                filePath: absoluteFilePath as string, // Already resolved and checked
                sheetName: actualWorksheetName, // Use the identified worksheet name
                rangeAddress: tableInfo.rangeAddress,
                operation: 'read',
                useComInterop: true,
              });

              if (!rangeDataResponse.success) { // Type guard
                const errorMsg = `COM: Failed to read data for table '${tableInfo.name}': ${rangeDataResponse.error?.message || 'Unknown error'}`;
                logger.error(errorMsg);
                allSuggestions.push({ tableName: tableInfo.name, suggestion: null, error: errorMsg, appliedFormats: [] });
                continue;
              }
              // If success is true, data should be defined. Adding explicit check for robustness.
              if (!rangeDataResponse.data) {
                const errorMsg = `COM: Failed to read data for table '${tableInfo.name}': Response data is missing.`;
                logger.error(errorMsg);
                allSuggestions.push({ tableName: tableInfo.name, suggestion: null, error: errorMsg, appliedFormats: [] });
                continue;
              }
              const tableDataForAI = rangeDataResponse.data;
              // Convert tableDataForAI to a string format suitable for the AI prompt
              // For Excel, a simple string representation of rows/columns might be better than full JSON for the AI prompt.
              let contextTextForAI = "Table Data:\n";
              if (Array.isArray(tableDataForAI)) {
                tableDataForAI.forEach((row: any) => {
                  if (Array.isArray(row)) {
                    contextTextForAI += row.join("\t") + "\n";
                  }
                });
              } else {
                contextTextForAI += JSON.stringify(tableDataForAI);
              }


              logger.debug(`COM: Data for table '${tableInfo.name}' for AI: ${contextTextForAI.substring(0, 500)}...`);

              // Step 4b: Call aiSuggest.tool.ts
              const aiPrompt = `Given the following Excel table data from table "${tableInfo.name}" in range "${tableInfo.rangeAddress}" on sheet "${actualWorksheetName}":\n${contextTextForAI}\nSuggest specific, actionable formatting instructions to make this table look professional and easy to read. Examples: "Bold headers", "Apply banded rows", "Set column 'ColumnName' width to 20", "Apply table style 'TableStyleMedium9'". Provide each instruction on a new line.`;
              logger.debug(`COM: AI Prompt for table '${tableInfo.name}': ${aiPrompt}`);

              const aiSuggestionResponse = await aiSuggestTool.handler({
                application: 'Excel.Application' as OfficeAppName,
                operation: 'format', // This operation in aiSuggestTool might be too generic.
                                     // We are essentially using it as a generic LLM call here.
                contextText: aiPrompt, // Using our specific prompt
                filePath: absoluteFilePath,
              }, context);

              if (!aiSuggestionResponse.success) {
                const errorMsg = `COM: Failed to get AI suggestion for table '${tableInfo.name}': ${aiSuggestionResponse.error?.message || 'Unknown error'}`;
                logger.error(errorMsg);
                allSuggestions.push({ tableName: tableInfo.name, suggestion: null, error: errorMsg, appliedFormats: [] });
                continue;
              }
              if (!aiSuggestionResponse.data || typeof aiSuggestionResponse.data !== 'string') {
                const errorMsg = `COM: AI suggestion for table '${tableInfo.name}' is missing or not a string.`;
                logger.error(errorMsg);
                allSuggestions.push({ tableName: tableInfo.name, suggestion: null, error: errorMsg, appliedFormats: [] });
                continue;
              }
              
              const suggestionsText = aiSuggestionResponse.data;
              logger.info(`COM: Raw AI suggestions for table '${tableInfo.name}':\n${suggestionsText}`);
              const appliedFormatsForTable: string[] = [];

              // Step 4c & 5: Parse and Apply Formatting
              const individualSuggestions = suggestionsText.split('\n').map(s => s.trim()).filter(s => s.length > 0);
              
              const comTableObject = await listObjects.Item(tableInfo.name); // Get the COM table object again for manipulation

              for (const suggestion of individualSuggestions) {
                logger.debug(`COM: Attempting to apply suggestion: "${suggestion}" for table ${tableInfo.name}`);
                try {
                  if (suggestion.toLowerCase().includes("bold header")) {
                    await comTableObject.HeaderRowRange.Font.Bold(true);
                    appliedFormatsForTable.push("Bolded headers");
                    changesMadeToWorkbook = true;
                  } else if (suggestion.toLowerCase().includes("banded rows")) {
                    await comTableObject.ShowTableStyleRowStripes(true);
                    appliedFormatsForTable.push("Applied banded rows");
                    changesMadeToWorkbook = true;
                  } else if (suggestion.toLowerCase().match(/apply table style ['"]?(.*?)['"]?/)) {
                    const styleMatch = suggestion.toLowerCase().match(/apply table style ['"]?(.*?)['"]?/);
                    if (styleMatch && styleMatch[1]) {
                        const styleName = styleMatch[1];
                        await comTableObject.TableStyle(styleName); // Direct property access might vary; ensure this is correct.
                                                                // Often it's a string like "TableStyleMedium9"
                        appliedFormatsForTable.push(`Applied table style: ${styleName}`);
                        changesMadeToWorkbook = true;
                    }
                  } else if (suggestion.toLowerCase().match(/set column ['"]?(.*?)['"]? width to (\d+)/)) {
                    const colWidthMatch = suggestion.toLowerCase().match(/set column ['"]?(.*?)['"]? width to (\d+)/);
                    if (colWidthMatch && colWidthMatch[1] && colWidthMatch[2]) {
                        const colName = colWidthMatch[1];
                        const colWidth = parseInt(colWidthMatch[2], 10);
                        // Find column by name in HeaderRowRange and set its EntireColumn.ColumnWidth
                        const headerCells = await comTableObject.HeaderRowRange.Cells;
                        const numHeaderCells = await headerCells.Count;
                        for (let k = 1; k <= numHeaderCells; k++) {
                            const cell = await headerCells.Item(k);
                            const cellValue = String(await cell.Value2).toLowerCase();
                            if (cellValue === colName.toLowerCase()) {
                                const columnToFormat = await cell.EntireColumn;
                                await columnToFormat.ColumnWidth(colWidth);
                                appliedFormatsForTable.push(`Set column '${colName}' width to ${colWidth}`);
                                changesMadeToWorkbook = true;
                                releaseObject(columnToFormat);
                                break;
                            }
                            releaseObject(cell);
                        }
                        releaseObject(headerCells);
                    }
                  }
                  // Add more parsing and application logic here
                } catch (applyError: any) {
                    logger.warn(`COM: Failed to apply suggestion "${suggestion}" to table ${tableInfo.name}: ${applyError.message}`);
                    appliedFormatsForTable.push(`Failed: ${suggestion} (${applyError.message})`);
                }
              }
              releaseObject(comTableObject);
              allSuggestions.push({ tableName: tableInfo.name, suggestion: suggestionsText, appliedFormats: appliedFormatsForTable });

            } catch (tableProcessingError: any) {
                const errorMsg = `COM: Error processing table '${tableInfo.name}': ${tableProcessingError.message}`;
                logger.error(errorMsg, tableProcessingError);
                allSuggestions.push({ tableName: tableInfo.name, suggestion: null, error: errorMsg, appliedFormats: [] });
            }
          }
          
          if (changesMadeToWorkbook && workbook) {
            logger.info("COM: Saving workbook due to applied formatting changes.");
            await workbook.Save();
          }

          return { success: true, data: { message: "COM: Table formatting process completed.", results: allSuggestions } };

        } catch (comError: any) {
          logger.error(`Error in formatTablesInWorksheet tool (COM Interop): ${comError.message}`, comError);
          throw comError; // Re-throw to be caught by the main try-catch
        } finally {
          if (listObjects) releaseObject(listObjects);
          if (worksheet) releaseObject(worksheet);
          if (workbook) {
            // Close workbook. If changesMadeToWorkbook is true, it was already saved.
            // If false, close without saving.
            await workbook.Close(false); // false = do not save changes (already saved if needed or no changes)
            releaseObject(workbook);
          }
          if (officeAppInstance) {
            officeAppInstance.release();
            logger.debug("COM: Excel application instance released.");
          }
          // excelApp is part of officeAppInstance, should be released by officeAppInstance.release()
        }

      } else {
        // ExcelJS Path
        logger.info('Using ExcelJS path.');
        const workbook = new ExcelJS.Workbook();
        const fileExists = await fs.pathExists(absoluteFilePath);
        if (!fileExists) {
          throw new Error(`File not found: ${absoluteFilePath}`);
        }
        await workbook.xlsx.readFile(absoluteFilePath);

        let targetWorksheet: ExcelJS.Worksheet | undefined;
        if (typeof worksheetIdentifier === 'string') {
          targetWorksheet = workbook.getWorksheet(worksheetIdentifier);
        } else { // worksheetIdentifier is a number (1-based index)
          if (worksheetIdentifier > 0 && worksheetIdentifier <= workbook.worksheets.length) {
            targetWorksheet = workbook.worksheets[worksheetIdentifier - 1]; // ExcelJS is 0-indexed
          }
        }

        if (!targetWorksheet) {
          throw new Error(`Worksheet with identifier '${worksheetIdentifier}' not found in ${absoluteFilePath}.`);
        }
        logger.info(`Successfully identified worksheet: '${targetWorksheet.name}'`);

        // Step 3: Identify Tables in the Worksheet
        // TODO: List tables using tables.tool.ts (ExcelJS part) or direct ExcelJS table access
        // For now, exceljs workbook.getWorksheet('sheetName').getTables() can be used.
        const tables = targetWorksheet.getTables();
        if (!tables || tables.length === 0) {
          logger.info(`No tables found in worksheet '${targetWorksheet.name}'.`);
          return { success: true, data: { message: `No tables found in worksheet '${targetWorksheet.name}'. No formatting applied.` } };
        }
        logger.info(`Found ${tables.length} table(s) in worksheet '${targetWorksheet.name}'.`);
        tables.forEach(([table]: [Table, void]) => { // Destructure the Table object from the tuple
          if (table) { // Check if table is defined
            logger.debug(`Table found: ${table.name}, Range: ${table.ref}`);
          }
        });


        // Step 4 & 5: For Each Table, Get Formatting Suggestions (AI) & Apply Formatting
        const excelJsAllSuggestions: { tableName: string, suggestion: any, error?: string, appliedFormats?: string[] }[] = [];
        let excelJsChangesMade = false;

        for (const tableEntry of tables) { // Iterate through ExcelJS tables
          const [excelJsTable] = tableEntry; // Destructure here
          if (!excelJsTable) continue;

          const tableName = excelJsTable.name;
          const tableRef = excelJsTable.ref;
          logger.info(`ExcelJS: Processing table '${tableName}' in range '${tableRef}'`);
          const appliedFormatsForTable: string[] = [];

          try {
            // Step 4a: Extract sample data for ExcelJS table
            let excelJsContextText = `Table "${tableName}" (Range: ${tableRef}):\n`;
            // Get header row
            // const tableRange = ExcelJS.utils.decodeRange(excelJsTable.ref); // This was incorrect.
            // Manual parsing of ref as a workaround:
            const refParts = excelJsTable.ref.split(':');
            const startCellRef = refParts[0];
            const endCellRef = refParts[refParts.length - 1]; // Handles single cell like "A1" or range "A1:D10"

            const decodeCell = (cellRef: string): { row: number, col: number } => {
                const letterPart = cellRef.match(/[A-Z]+/)?.[0] || "A";
                const numberPart = cellRef.match(/\d+/)?.[0] || "1";
                let col = 0;
                for (let i = 0; i < letterPart.length; i++) {
                    col = col * 26 + (letterPart.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
                }
                return { row: parseInt(numberPart, 10), col: col };
            };

            const startCell = decodeCell(startCellRef);
            const endCell = decodeCell(endCellRef);

            const tableRange = {
                top: startCell.row,
                left: startCell.col,
                bottom: endCell.row,
                right: endCell.col,
            };


            if (excelJsTable.headerRow) {
                const headerRowObject = targetWorksheet.getRow(tableRange.top);
                let headerText = "";
                for (let col = tableRange.left; col <= tableRange.right; col++) {
                    headerText += (headerRowObject.getCell(col).value?.toString() || "") + (col === tableRange.right ? "" : "\t");
                }
                excelJsContextText += headerText + "\n";
            }
            // Get a few data rows (e.g., up to 3 data rows)
            const dataRowStartIndex = tableRange.top + (excelJsTable.headerRow ? 1 : 0);
            const actualDataRowCount = tableRange.bottom - dataRowStartIndex + 1;
            const dataRowCountToFetch = Math.min(3, actualDataRowCount > 0 ? actualDataRowCount : 0);


            for (let i = 0; i < dataRowCountToFetch; i++) {
                const dataRowObject = targetWorksheet.getRow(dataRowStartIndex + i);
                let rowText = "";
                for (let col = tableRange.left; col <= tableRange.right; col++) {
                    rowText += (dataRowObject.getCell(col).value?.toString() || "") + (col === tableRange.right ? "" : "\t");
                }
                excelJsContextText += rowText + "\n";
            }
            logger.debug(`ExcelJS: Data for table '${tableName}' for AI: ${excelJsContextText.substring(0,500)}...`);

            // Step 4b: Call aiSuggestTool for ExcelJS table
            const aiPrompt = `Given the following Excel table data from table "${tableName}" in range "${tableRef}" on sheet "${targetWorksheet.name}":\n${excelJsContextText}\nSuggest specific, actionable formatting instructions to make this table look professional and easy to read. Examples: "Apply table style 'TableStyleMedium9'", "Enable header row filter buttons". Provide each instruction on a new line.`;
            const aiSuggestionResponse = await aiSuggestTool.handler({
              application: 'Excel.Application' as OfficeAppName, // Still using Excel context for suggestions
              operation: 'format',
              contextText: aiPrompt,
              filePath: absoluteFilePath,
            }, context);

            if (!aiSuggestionResponse.success) {
              const errorMsg = `ExcelJS: Failed to get AI suggestion for table '${tableName}': ${aiSuggestionResponse.error?.message || 'Unknown error'}`;
              logger.error(errorMsg);
              excelJsAllSuggestions.push({ tableName, suggestion: null, error: errorMsg, appliedFormats: [] });
              continue;
            }
            if (!aiSuggestionResponse.data || typeof aiSuggestionResponse.data !== 'string') {
              const errorMsg = `ExcelJS: AI suggestion for table '${tableName}' is missing or not a string.`;
              logger.error(errorMsg);
              excelJsAllSuggestions.push({ tableName, suggestion: null, error: errorMsg, appliedFormats: [] });
              continue;
            }

            const suggestionsText = aiSuggestionResponse.data;
            logger.info(`ExcelJS: Raw AI suggestions for table '${tableName}':\n${suggestionsText}`);
            
            // Step 4c & 5: Parse and Apply Formatting for ExcelJS
            const individualSuggestions = suggestionsText.split('\n').map(s => s.trim()).filter(s => s.length > 0);

            for (const suggestion of individualSuggestions) {
              logger.debug(`ExcelJS: Attempting to apply suggestion: "${suggestion}" for table ${tableName}`);
              try {
                if (suggestion.toLowerCase().match(/apply table style ['"]?(.*?)['"]?/)) {
                  const styleMatch = suggestion.toLowerCase().match(/apply table style ['"]?(.*?)['"]?/);
                  if (styleMatch && styleMatch[1]) {
                    const styleName = styleMatch[1];
                    // ExcelJS table style application:
                    excelJsTable.style = {
                      theme: styleName as any, // e.g., 'TableStyleMedium9', 'TableStyleLight1'
                      showRowStripes: excelJsTable.style?.showRowStripes ?? true, // Preserve or default
                      showColumnStripes: excelJsTable.style?.showColumnStripes ?? false,
                      showFirstColumn: excelJsTable.style?.showFirstColumn ?? false,
                      showLastColumn: excelJsTable.style?.showLastColumn ?? false,
                    };
                    appliedFormatsForTable.push(`Applied table style: ${styleName}`);
                    excelJsChangesMade = true;
                  }
                } else if (suggestion.toLowerCase().includes("banded rows") || suggestion.toLowerCase().includes("row stripes")) {
                    excelJsTable.style = { ...excelJsTable.style, showRowStripes: true };
                    appliedFormatsForTable.push("Enabled banded rows (row stripes)");
                    excelJsChangesMade = true;
                } else if (suggestion.toLowerCase().includes("filter buttons") && excelJsTable.columns) {
                    excelJsTable.columns.forEach(column => column.filterButton = true);
                    appliedFormatsForTable.push("Enabled filter buttons on header row");
                    excelJsChangesMade = true;
                }
                // Note: More complex formatting like specific font bolding on headers or column widths
                // for ExcelJS tables requires manipulating individual cells within the table's range,
                // which is more involved than direct table properties.
              } catch (applyError: any) {
                logger.warn(`ExcelJS: Failed to apply suggestion "${suggestion}" to table ${tableName}: ${applyError.message}`);
                appliedFormatsForTable.push(`Failed: ${suggestion} (${applyError.message})`);
              }
            }
            excelJsAllSuggestions.push({ tableName, suggestion: suggestionsText, appliedFormats: appliedFormatsForTable });

          } catch (tableProcessingError: any) {
            const errorMsg = `ExcelJS: Error processing table '${tableName}': ${tableProcessingError.message}`;
            logger.error(errorMsg, tableProcessingError);
            excelJsAllSuggestions.push({ tableName, suggestion: null, error: errorMsg, appliedFormats: [] });
          }
        }

        if (excelJsChangesMade) {
          logger.info("ExcelJS: Saving workbook due to applied formatting changes.");
          await workbook.xlsx.writeFile(absoluteFilePath);
        }
        return { success: true, data: { message: "ExcelJS: Table formatting process completed.", results: excelJsAllSuggestions } };
      }

      // Step 6: Output/Return Value
      // return { success: true, data: { message: 'Table formatting completed successfully.' } }; // Placeholder

    } catch (error: any) {
      logger.error(`Error in formatTablesInWorksheet tool: ${error.message}`, error);
      return {
        success: false,
        error: {
          code: useComInterop ? 'FORMAT_TABLES_COM_ERROR' : 'FORMAT_TABLES_EXCELJS_ERROR',
          message: error.message || 'An error occurred while formatting tables in the Excel worksheet.',
          details: { filePath: absoluteFilePath, worksheetIdentifier, errorInfo: error.stack }
        }
      };
    }
  }
}];

// Helper function to register the tool (if you have a central registration mechanism)
export const registerFormatTablesInWorksheetTool = () => {
  // This function could be used to register the tool with a central tool registry
  // For now, it's a placeholder. The tool is exported directly.
  logger.info('formatTablesInWorksheetTool registered.');
};