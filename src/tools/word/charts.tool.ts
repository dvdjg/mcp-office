/**
 * @file Implements the 'word/charts' tool using COM Interop (winax) for managing charts in Word documents.
 * Provides functionality to insert charts. Modify and delete operations are not yet implemented.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z, ZodError } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, SuccessResponse, ErrorResponse, FastMCPContext } from '../../types/common.types'; // Importar FastMCPContext, quitar ToolContext
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { validateFilePath } from '../../utils/security';
import { handleToolError } from '../../utils/errorHandler'; // Importar handleToolError
import path from 'path';
import logger from '../../utils/logger'; // Importar logger global si context.logger no está disponible

/** Basic mapping of chart types to XlChartType values. */
const chartTypeMap: { [key: string]: number } = {
    'column': 51, // xlColumnClustered
    'bar': 57,    // xlBarClustered
    'line': 4,     // xlLine
    'pie': 5,      // xlPie
    // Add more types as needed
};

/** Input schema for the 'word/charts/insert' tool. */
const insertChartSchema = z.object({
    /** The path of the Word document to insert the chart into (relative to the current workspace directory). */
    filePath: z.string().min(1, { message: 'filePath is required.' }),
    /** The type of chart to insert (e.g., 'column', 'bar', 'line', 'pie'). */
    type: z.string().refine(type => chartTypeMap.hasOwnProperty(type.toLowerCase()), {
        message: `Invalid chart type. Supported types: ${Object.keys(chartTypeMap).join(', ')}`,
    }),
    /** Optional insertion position (e.g., 'end', 'selection', 'paragraph:N'). Defaults to 'end'. */
    position: z.string().optional(), // Ejemplo: 'end', 'selection', 'paragraph:N'
    /** Optional 2D array for initial chart data. Complex to implement via COM, might use default data. */
    data: z.array(z.array(z.union([z.string(), z.number()])))
        .optional()
        .describe('Optional 2D array for initial chart data. Complex to implement via COM, might use default data.'),
});

// No necesitamos inferir el tipo aquí si usamos ToolRequestParams
// type InsertChartParams = z.infer<typeof insertChartSchema>;

/**
 * Inserts a chart into a Word document using COM Interop.
 * @param params - The parameters for the insert chart operation, validated against `insertChartSchema`.
 * @param context - The FastMCP context (optional).
 * @returns ApiResponse indicating success or failure.
 * @throws {Error} If the document fails to open, the chart type is invalid, or chart insertion/data setting fails.
 */
async function insertChart(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> { // Usar FastMCPContext<undefined>
    let wordApp: any = null;
    let doc: any = null;
    let chart: any = null;
    let insertionRange: any = null;
    // Variables for data handling (optional)
    // let chartData: any = null;
    // let workbook: any = null;
    // let sheet: any = null;

    // Use global logger directly
    // const currentLogger = context?.logger || logger; // No usar context.logger

    try {
        const validatedParams = insertChartSchema.parse(params);
        const { filePath, type, position, data } = validatedParams;

        logger.info(`[word/charts/insert] Attempting to insert chart of type '${type}' into document: ${filePath}`);

        // 1. Validate file path
        const safeFilePath = validateFilePath(filePath);
        logger.debug(`[word/charts/insert] File path validated: ${safeFilePath}`);
        const absoluteFilePath = path.resolve(safeFilePath);
        logger.debug(`[word/charts/insert] Absolute file path: ${absoluteFilePath}`);

        // 2. Get Word application instance
        wordApp = await getOfficeApplication('Word.Application');
        wordApp.Visible = false;

        // 3. Open the document
        doc = wordApp.Documents.Open(absoluteFilePath);
        logger.debug(`[word/charts/insert] Document opened: ${absoluteFilePath}`);

        // 4. Determine insertion range
        insertionRange = doc.Content;
        insertionRange.Collapse(0); // wdCollapseEnd
        logger.debug(`[word/charts/insert] Insertion range set to end of document (position='${position || 'end'}').`);

        // 5. Map chart type
        const chartTypeValue = chartTypeMap[type.toLowerCase()];
        logger.debug(`[word/charts/insert] Mapped chart type '${type}' to XlChartType value: ${chartTypeValue}`);

        // 6. Insert the chart
        chart = doc.InlineShapes.AddChart2(-1, chartTypeValue, insertionRange);
        chart.Width = 400;
        chart.Height = 300;
        logger.info(`[word/charts/insert] Chart inserted successfully.`);

        // 7. Optional: Populate data
        if (data && data.length > 0 && data[0].length > 0) {
            logger.warn('[word/charts/insert] Setting chart data via COM is complex and currently implemented minimally or skipped.');
            try {
                const chartObj = chart.Chart;
                const chartData = chartObj.ChartData;
                chartData.Activate();
                const workbook = chartData.Workbook;
                const sheet = workbook.Worksheets(1);

                const rows = data.length;
                const cols = data[0].length;

                const dataForExcel = new Array(rows);
                for (let i = 0; i < rows; i++) {
                    dataForExcel[i] = new Array(cols);
                    for (let j = 0; j < cols; j++) {
                        dataForExcel[i][j] = data[i][j];
                    }
                }

                const targetRange = sheet.Range("A1").Resize(rows, cols);
                targetRange.Value = dataForExcel;
                logger.debug(`[word/charts/insert] Attempted to write data to chart's workbook.`);

                workbook.Close(false);

                releaseObject(targetRange);
                releaseObject(sheet);
                releaseObject(workbook);
                releaseObject(chartData);
                releaseObject(chartObj);
                logger.info('[word/charts/insert] Chart data setting attempt finished.');

            } catch (dataError: any) {
                logger.error(`[word/charts/insert] Error setting chart data: ${dataError.message}`, { stack: dataError.stack });
            }
        } else {
            logger.info('[word/charts/insert] No data provided or data array empty, chart created with default data.');
        }

        // 8. Save the document
        doc.Save();
        logger.info(`[word/charts/insert] Document saved: ${absoluteFilePath}`);

        return {
            success: true,
            data: {},
            message: `Chart '${type}' inserted successfully into ${filePath}.`
        };

    } catch (error: unknown) {
        logger.error(`[word/charts/insert] Error in insertChart: ${error instanceof Error ? error.message : String(error)}`, { stack: error instanceof Error ? error.stack : undefined });
        if (doc) {
            try {
                doc.Close(false);
            } catch (closeError: any) {
                logger.error(`[word/charts/insert] Error closing document after error: ${closeError.message}`);
            }
        }
        const errorCode = error instanceof ZodError ? 'VALIDATION_ERROR' : 'WORD_CHART_INSERT_FAILED';
        return handleToolError(error, errorCode);
    } finally {
        releaseObject(insertionRange);
        releaseObject(chart);
        releaseObject(doc);
        if (wordApp) {
             try {
                 if (wordApp.Documents.Count === 0) {
                     wordApp.Quit();
                 }
             } catch (quitError) {
                 logger.warn(`[OfficeInterop] Error trying to Quit Word application: ${quitError instanceof Error ? quitError.message : String(quitError)}`);
             }
            releaseObject(wordApp);
        }
        logger.info('[word/charts/insert] insertChart finished, COM objects released.');
    }
}

/**
 * Array of McpResource definitions for Word chart operations.
 */
export const wordChartsTool: McpResource[] = [
    {
        path: 'word/charts/insert',
        handler: insertChart,
        schema: insertChartSchema,
        description: 'Inserts a new chart into a Word document using COM Interop. Basic data setting is optional/complex.',
    },
    {
        path: 'word/charts/modify',
        handler: async (params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<any>> => { // Usar FastMCPContext<undefined>
            // Use global logger
            logger.warn('[word/charts/modify] Tool word/charts/modify not implemented.');
            return {
                success: false,
                error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/charts/modify not implemented.' }
            };
        },
        description: 'Modify chart properties/data (Not Implemented)',
        schema: z.object({}),
    },
    {
        path: 'word/charts/delete',
        handler: async (params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<any>> => { // Usar FastMCPContext<undefined>
            // Use global logger
            logger.warn('[word/charts/delete] Tool word/charts/delete not implemented.');
            return {
                success: false,
                error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/charts/delete not implemented.' }
            };
        },
        description: 'Delete a chart (Not Implemented)',
        schema: z.object({}),
    },
];

export default wordChartsTool;