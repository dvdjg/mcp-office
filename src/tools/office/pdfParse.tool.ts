/**
 * @file Tool for parsing and converting PDF files.
 * Allows extracting text from a PDF or converting specific pages to images.
 * Uses the `pdf-parse` and `pdf2pic` libraries.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams } from '../../types/common.types.js'; // Import ToolRequestParams
import { Context } from 'fastmcp'; // Import Context
import pdfParse from 'pdf-parse';
import { fromPath } from 'pdf2pic'; // Remove WriteImageResponse from import
import * as path from 'path';
import * as fs from 'fs/promises';

// Input schema for the office/pdf/parse tool
const PdfParseInputSchema = z.object({
  //type: z.literal('object'), // Added to satisfy validator
  filePath: z.string().min(1, { message: 'filePath is required.' }),
  operation: z.enum(['parse', 'convert']),
  outputDirectory: z.string().optional(), // Required for 'convert'
  pageRange: z.string().optional(), // Optional for 'convert' (e.g., "1-3", "5")
  format: z.enum(['png', 'jpeg', 'webp']).default('png'), // Optional for 'convert'
  quality: z.number().int().min(1).max(100).default(80).optional(), // Optional for 'convert'
});

type PdfParseInput = z.infer<typeof PdfParseInputSchema>;

// Define a type for the result of pdf2pic.bulk when saving to file
interface Pdf2PicResult {
  filename: string;
  // Add other WriteImageResponse properties if necessary and known
  // For example, if the response includes the full path or size:
  // path?: string;
  // size?: number;
}

/**
 * @tool office/pdf/parse
 * @description Tool for parsing and converting PDF files.
 * Allows extracting text from a PDF or converting specific pages to images.
 * Uses the `pdf-parse` and `pdf2pic` libraries.
 * @param {object} params - Input parameters.
 * @param {string} params.filePath - Path to the source PDF file.
 * @param {'parse' | 'convert'} params.operation - Operation to perform ('parse' or 'convert').
 * @param {string} [params.outputDirectory] - Directory to save images (required for 'convert').
 * @param {string} [params.pageRange] - Page range to convert (e.g., "1-3", "5"). Converts all if not specified.
 * @param {'png' | 'jpeg' | 'webp'} [params.format='png'] - Output image format.
 * @param {number} [params.quality=80] - Image quality (1-100).
 * @returns {Promise<string | string[]>} - Extracted text (for 'parse') or paths of created image files (for 'convert').
 * @throws {Error} - If an error occurs during the process or parameters are invalid.
 */
const pdfParseTool: McpResource = {
  path: 'office/pdf/parse', // Add the path property
  description: 'Parses and converts PDF files.',
  schema: PdfParseInputSchema, // Change inputSchema to schema
  handler: async (params: ToolRequestParams, context?: Context<any>): Promise<ApiResponse<string | string[]>> => { // Adjust handler signature
    // Cast params to the expected type after Zod validation
    const { filePath, operation, outputDirectory, pageRange, format, quality } = params as PdfParseInput;

    try {
      // Check if the PDF file exists
      await fs.access(filePath);

      if (operation === 'parse') {
        const dataBuffer = await fs.readFile(filePath);
        const data = await pdfParse(dataBuffer);
        return { success: true, data: data.text }; // Wrap in ApiResponse

      } else if (operation === 'convert') {
        if (!outputDirectory) {
          return { success: false, error: { code: 'VALIDATION_ERROR', message: 'outputDirectory is required for the "convert" operation.' } }; // Return ErrorResponse
        }

        // Ensure the output directory exists
        await fs.mkdir(outputDirectory, { recursive: true });

        const options = {
          density: 100,
          saveFilename: path.basename(filePath, path.extname(filePath)),
          savePath: outputDirectory,
          format: format,
          quality: quality,
          width: 1600, // Image width, adjustable
          height: 2300 // Image height, adjustable
        };

        const convert = fromPath(filePath, options);

        if (!convert) {
           return { success: false, error: { code: 'PROCESSING_ERROR', message: 'Error initializing pdf2pic.' } };
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
                return { success: false, error: { code: 'VALIDATION_ERROR', message: `Invalid page range: ${range}` } }; // Return ErrorResponse
              }
            } else {
              const pageNum = Number(range);
              if (!isNaN(pageNum) && pageNum > 0) {
                pagesToConvert.push(pageNum);
              } else {
                return { success: false, error: { code: 'VALIDATION_ERROR', message: `Invalid page number: ${range}` } }; // Return ErrorResponse
              }
            }
          }
        }

        let results: Pdf2PicResult[]; // Type results
        // Move bulk logic inside the if(convert) block
        if (pagesToConvert === 'all') {
          results = await convert.bulk(-1) as Pdf2PicResult[]; // -1 for all pages
        } else {
          results = await convert.bulk(pagesToConvert) as Pdf2PicResult[];
        }

        // pdf2pic returns an array of objects with the saved filename
        return { success: true, data: results.map((result: Pdf2PicResult) => path.join(outputDirectory, result.filename)) }; // Type result and wrap in ApiResponse

      } else {
        // This should not happen if the Zod schema works correctly, but it's a safeguard
        return { success: false, error: { code: 'INVALID_OPERATION', message: `Unsupported operation: ${operation}` } }; // Return ErrorResponse
      }

    } catch (error: any) {
      if (error.code === 'ENOENT') {
        return { success: false, error: { code: 'FILE_NOT_FOUND', message: `File not found: ${filePath}` } }; // Return ErrorResponse
      }
      // Catch other pdf-parse or pdf2pic errors
      return { success: false, error: { code: 'PROCESSING_ERROR', message: `Error processing PDF file: ${error.message}` } }; // Return ErrorResponse
    }
  },
};

export default pdfParseTool;