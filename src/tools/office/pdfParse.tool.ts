import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams } from '../../types/common.types'; // Importar ToolRequestParams
import { Context } from 'fastmcp'; // Importar Context
import pdfParse from 'pdf-parse';
import { fromPath } from 'pdf2pic'; // Eliminar WriteImageResponse de la importación
import * as path from 'path';
import * as fs from 'fs/promises';

// Esquema de entrada para la herramienta office/pdf/parse
const PdfParseInputSchema = z.object({
  //type: z.literal('object'), // Added to satisfy validator
  filePath: z.string().min(1, { message: 'filePath es requerido.' }),
  operation: z.enum(['parse', 'convert']),
  outputDirectory: z.string().optional(), // Requerido para 'convert'
  pageRange: z.string().optional(), // Opcional para 'convert' (e.g., "1-3", "5")
  format: z.enum(['png', 'jpeg', 'webp']).default('png'), // Opcional para 'convert'
  quality: z.number().int().min(1).max(100).default(80).optional(), // Opcional para 'convert'
});

type PdfParseInput = z.infer<typeof PdfParseInputSchema>;

// Definir un tipo para el resultado de pdf2pic.bulk cuando se guarda en archivo
interface Pdf2PicResult {
  filename: string;
  // Añadir otras propiedades de WriteImageResponse si son necesarias y conocidas
  // Por ejemplo, si la respuesta incluye el path completo o el tamaño:
  // path?: string;
  // size?: number;
}

/**
 * @tool office/pdf/parse
 * @description Herramienta para analizar y convertir archivos PDF.
 * Permite extraer texto de un PDF o convertir páginas específicas a imágenes.
 * Utiliza las librerías `pdf-parse` y `pdf2pic`.
 * @param {object} params - Parámetros de entrada.
 * @param {string} params.filePath - Ruta al archivo PDF de origen.
 * @param {'parse' | 'convert'} params.operation - Operación a realizar ('parse' o 'convert').
 * @param {string} [params.outputDirectory] - Directorio para guardar imágenes (requerido para 'convert').
 * @param {string} [params.pageRange] - Rango de páginas a convertir (e.g., "1-3", "5"). Convierte todas si no se especifica.
 * @param {'png' | 'jpeg' | 'webp'} [params.format='png'] - Formato de imagen de salida.
 * @param {number} [params.quality=80] - Calidad de la imagen (1-100).
 * @returns {Promise<string | string[]>} - Texto extraído (para 'parse') o rutas de archivos de imagen creados (para 'convert').
 * @throws {Error} - Si ocurre un error durante el proceso o los parámetros son inválidos.
 */
const pdfParseTool: McpResource = {
  path: 'office/pdf/parse', // Añadir la propiedad path
  description: 'Analiza y convierte archivos PDF.',
  schema: PdfParseInputSchema, // Cambiar inputSchema a schema
  handler: async (params: ToolRequestParams, context?: Context<any>): Promise<ApiResponse<string | string[]>> => { // Ajustar la firma del handler
    // Castear params al tipo esperado después de la validación de Zod
    const { filePath, operation, outputDirectory, pageRange, format, quality } = params as PdfParseInput;

    try {
      // Verificar si el archivo PDF existe
      await fs.access(filePath);

      if (operation === 'parse') {
        const dataBuffer = await fs.readFile(filePath);
        const data = await pdfParse(dataBuffer);
        return { success: true, data: data.text }; // Envolver en ApiResponse

      } else if (operation === 'convert') {
        if (!outputDirectory) {
          return { success: false, error: { code: 'VALIDATION_ERROR', message: 'outputDirectory es requerido para la operación "convert".' } }; // Devolver ErrorResponse
        }

        // Asegurarse de que el directorio de salida existe
        await fs.mkdir(outputDirectory, { recursive: true });

        const options = {
          density: 100,
          saveFilename: path.basename(filePath, path.extname(filePath)),
          savePath: outputDirectory,
          format: format,
          quality: quality,
          width: 1600, // Ancho de la imagen, ajustable
          height: 2300 // Alto de la imagen, ajustable
        };

        const convert = fromPath(filePath, options);

        if (!convert) {
           return { success: false, error: { code: 'PROCESSING_ERROR', message: 'Error al inicializar pdf2pic.' } };
        }

        let pagesToConvert: number[] | 'all' = 'all';
        if (pageRange) {
          pagesToConvert = [];
          const ranges = pageRange.split(',').map(r => r.trim());
          for (const range of ranges) {
            if (range.includes('-')) {
              const [start, end] = range.split('-').map(Number);
              if (!isNaN(start) && !isNaN(end) && start <= end) {
                for (let i = start; i <= end; i++) {
                  pagesToConvert.push(i);
                }
              } else {
                return { success: false, error: { code: 'VALIDATION_ERROR', message: `Rango de páginas inválido: ${range}` } }; // Devolver ErrorResponse
              }
            } else {
              const pageNum = Number(range);
              if (!isNaN(pageNum) && pageNum > 0) {
                pagesToConvert.push(pageNum);
              } else {
                return { success: false, error: { code: 'VALIDATION_ERROR', message: `Número de página inválido: ${range}` } }; // Devolver ErrorResponse
              }
            }
          }
        }

        let results: Pdf2PicResult[]; // Tipar results
        // Mover la lógica de bulk dentro del bloque if(convert)
        if (pagesToConvert === 'all') {
          results = await convert.bulk(-1) as Pdf2PicResult[]; // -1 para todas las páginas
        } else {
          results = await convert.bulk(pagesToConvert) as Pdf2PicResult[];
        }

        // pdf2pic devuelve un array de objetos con el nombre del archivo guardado
        return { success: true, data: results.map((result: Pdf2PicResult) => path.join(outputDirectory, result.filename)) }; // Tipar result y envolver en ApiResponse

      } else {
        // Esto no debería ocurrir si el esquema Zod funciona correctamente, pero es una salvaguarda
        return { success: false, error: { code: 'INVALID_OPERATION', message: `Operación no soportada: ${operation}` } }; // Devolver ErrorResponse
      }

    } catch (error: any) {
      if (error.code === 'ENOENT') {
        return { success: false, error: { code: 'FILE_NOT_FOUND', message: `Archivo no encontrado: ${filePath}` } }; // Devolver ErrorResponse
      }
      // Capturar otros errores de pdf-parse o pdf2pic
      return { success: false, error: { code: 'PROCESSING_ERROR', message: `Error al procesar el archivo PDF: ${error.message}` } }; // Devolver ErrorResponse
    }
  },
};

export default pdfParseTool;