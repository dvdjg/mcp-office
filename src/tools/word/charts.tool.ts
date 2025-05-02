import { z, ZodError } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, SuccessResponse, ErrorResponse, FastMCPContext } from '../../types/common.types'; // Importar FastMCPContext, quitar ToolContext
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { validateFilePath } from '../../utils/security';
import { handleToolError } from '../../utils/errorHandler'; // Importar handleToolError
import path from 'path';
import logger from '../../utils/logger'; // Importar logger global si context.logger no está disponible

// Mapeo básico de tipos de gráfico a valores XlChartType
const chartTypeMap: { [key: string]: number } = {
    'column': 51, // xlColumnClustered
    'bar': 57,    // xlBarClustered
    'line': 4,     // xlLine
    'pie': 5,      // xlPie
    // Añadir más tipos según sea necesario
};

// Schema de entrada para la inserción de gráficos
const insertChartSchema = z.object({
    filePath: z.string().min(1, { message: 'filePath is required.' }),
    type: z.string().refine(type => chartTypeMap.hasOwnProperty(type.toLowerCase()), {
        message: `Invalid chart type. Supported types: ${Object.keys(chartTypeMap).join(', ')}`,
    }),
    position: z.string().optional(), // Ejemplo: 'end', 'selection', 'paragraph:N'
    data: z.array(z.array(z.union([z.string(), z.number()])))
        .optional()
        .describe('Optional 2D array for initial chart data. Complex to implement via COM, might use default data.'),
});

// No necesitamos inferir el tipo aquí si usamos ToolRequestParams
// type InsertChartParams = z.infer<typeof insertChartSchema>;

/**
 * Inserta un gráfico en un documento Word usando COM Interop.
 * @param params - Parámetros de la solicitud (ToolRequestParams).
 * @param context - Contexto opcional de la herramienta MCP (ToolContext).
 * @returns ApiResponse indicando éxito o fracaso.
 */
async function insertChart(params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<{}>> { // Usar FastMCPContext<undefined>
    let wordApp: any = null;
    let doc: any = null;
    let chart: any = null;
    let insertionRange: any = null;
    // Variables para manejo de datos (opcional)
    // let chartData: any = null;
    // let workbook: any = null;
    // let sheet: any = null;

    // Usar logger global directamente
    // const currentLogger = context?.logger || logger; // No usar context.logger

    try {
        const validatedParams = insertChartSchema.parse(params);
        const { filePath, type, position, data } = validatedParams;

        logger.info(`Attempting to insert chart of type '${type}' into document: ${filePath}`);

        // 1. Validar ruta del archivo
        const safeFilePath = validateFilePath(filePath);
        logger.debug(`File path validated: ${safeFilePath}`);
        const absoluteFilePath = path.resolve(safeFilePath);
        logger.debug(`Absolute file path: ${absoluteFilePath}`);

        // 2. Obtener instancia de Word
        wordApp = await getOfficeApplication('Word.Application');
        wordApp.Visible = false;

        // 3. Abrir el documento
        doc = wordApp.Documents.Open(absoluteFilePath);
        logger.debug(`Document opened: ${absoluteFilePath}`);

        // 4. Determinar el rango de inserción
        insertionRange = doc.Content;
        insertionRange.Collapse(0); // wdCollapseEnd
        logger.debug(`Insertion range set to end of document (position='${position || 'end'}').`);

        // 5. Mapear tipo de gráfico
        const chartTypeValue = chartTypeMap[type.toLowerCase()];
        logger.debug(`Mapped chart type '${type}' to XlChartType value: ${chartTypeValue}`);

        // 6. Insertar el gráfico
        chart = doc.InlineShapes.AddChart2(-1, chartTypeValue, insertionRange);
        chart.Width = 400;
        chart.Height = 300;
        logger.info(`Chart inserted successfully.`);

        // 7. Opcional: Poblar datos
        if (data && data.length > 0 && data[0].length > 0) {
            logger.warn('Setting chart data via COM is complex and currently implemented minimally or skipped.');
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
                logger.debug(`Attempted to write data to chart's workbook.`);

                workbook.Close(false);

                releaseObject(targetRange);
                releaseObject(sheet);
                releaseObject(workbook);
                releaseObject(chartData);
                releaseObject(chartObj);
                logger.info('Chart data setting attempt finished.');

            } catch (dataError: any) {
                logger.error(`Error setting chart data: ${dataError.message}`, { stack: dataError.stack });
            }
        } else {
            logger.info('No data provided or data array empty, chart created with default data.');
        }

        // 8. Guardar el documento
        doc.Save();
        logger.info(`Document saved: ${absoluteFilePath}`);

        return {
            success: true,
            data: {},
            message: `Chart '${type}' inserted successfully into ${filePath}.`
        };

    } catch (error: unknown) {
        logger.error(`Error in insertChart: ${error instanceof Error ? error.message : String(error)}`, { stack: error instanceof Error ? error.stack : undefined });
        if (doc) {
            try {
                doc.Close(false);
            } catch (closeError: any) {
                logger.error(`Error closing document after error: ${closeError.message}`);
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
        logger.info('insertChart finished, COM objects released.');
    }
}

// Definición del recurso MCP para las operaciones de gráficos de Word
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
            // Usar logger global
            logger.warn('Tool word/charts/modify not implemented.');
            return {
                success: false,
                error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/charts/modify not implemented.' }
            };
        },
        description: 'Modify chart properties/data (Not Implemented)'
    },
    {
        path: 'word/charts/delete',
        handler: async (params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<any>> => { // Usar FastMCPContext<undefined>
            // Usar logger global
            logger.warn('Tool word/charts/delete not implemented.');
            return {
                success: false,
                error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/charts/delete not implemented.' }
            };
        },
        description: 'Delete a chart (Not Implemented)'
    },
];

export default wordChartsTool;