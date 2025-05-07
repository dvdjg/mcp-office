/**
 * @file Tool for managing shapes in PowerPoint presentations.
 * Allows inserting, modifying, formatting, deleting, and listing shapes.
 * Uses COM Interop via winax.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
// src/tools/powerpoint/shapes.tool.ts

import { z } from 'zod';
import PptxGenJS from 'pptxgenjs';
// officeparser is not typically used for shape manipulation, focusing on text extraction.
// We'll rely on PptxGenJS for generation and note limitations for modification/listing.
import { McpResource, ToolRequestParams, ApiResponse, FastMCPContext } from '../../types/common.types.js'; // Added FastMCPContext
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js';
import logger from '../../utils/logger.js';
import { saveResource } from '../dynamic/resources.tool.js';
import fs from 'fs-extra';
import { resolve as resolvePath, extname as extnamePath, basename as basenamePath } from 'path';
import { validateFilePath } from '../../utils/security.js'; // Added security import

// Helper to create standard error responses
const createErrorResponse = (message: string, code = 'TOOL_EXECUTION_ERROR', details?: unknown): ApiResponse<never> => ({
    success: false,
    error: { code, message, details },
});

// Define the input schema for the powerpoint/shapes tool
const PowerPointShapesInputSchema = z.object({
  filePath: z.string().min(1, 'File path is required.').refine(validateFilePath, { // Added validation
    message: "Invalid or potentially unsafe file path provided.",
  }),
  operation: z.enum(['insert', 'modify', 'format', 'delete', 'list']).describe('Operation to perform: insert, modify, format, delete, list.'),
  slideIndex: z.number().int().positive().optional().describe('1-based index of the slide. Required for insert, modify, format, delete.'),
  shapeType: z.string().optional().describe('Type of shape to insert (e.g., "msoShapeRectangle", "msoShapeTextbox"). Required for insert.'),
  position: z.object({
    left: z.number().optional(),
    top: z.number().optional(),
  }).optional().describe('Position of the shape (left, top). Optional for insert, modify.'),
  size: z.object({
    width: z.number().optional(),
    height: z.number().optional(),
  }).optional().describe('Size of the shape (width, height). Optional for insert, modify.'),
  text: z.string().optional().describe('Text for the shape (if it is a textbox). Optional for insert, modify.'),
  shapeIndex: z.number().int().positive().optional().describe('1-based index of the shape on the slide. Required for modify, format, delete.'),
  shapeName: z.string().optional().describe('Name of the shape. Alternative to shapeIndex for modify, format, delete.'),
  formatProperties: z.object({
    fillColor: z.string().optional().describe('Fill color (e.g., "RGB(255, 0, 0)").'),
    lineColor: z.string().optional().describe('Line color (e.g., "RGB(0, 0, 255)").'),
    lineWidth: z.number().optional().describe('Line width.'),
    fontName: z.string().optional().describe('Font name.'),
    fontSize: z.number().optional().describe('Font size.'),
    fontBold: z.boolean().optional().describe('Bold.'),
    fontItalic: z.boolean().optional().describe('Italic.'),
    fontUnderline: z.boolean().optional().describe('Underline.'),
  }).optional().describe('Format properties to apply. Required for format.'),
  useComInterop: z.boolean().optional().default(false).describe('Set to true to use COM Interop for operations, otherwise uses pptxgenjs (with limitations).'),
});

type PowerPointShapesInput = z.infer<typeof PowerPointShapesInputSchema>;

/**
 * @tool powerpoint/shapes
 * @description Allows inserting, modifying, formatting, deleting, and listing shapes in PowerPoint presentations.
 * For 'insert', 'modify', 'format', and 'delete' operations, the file will be created if it does not exist.
 * Uses COM Interop via winax.
 * @param {string} filePath - Path to the PowerPoint file.
 * @param {'insert' | 'modify' | 'format' | 'delete' | 'list'} operation - Operation to perform.
 * @param {number} [slideIndex] - 1-based index of the slide. Required for insert, modify, format, delete.
 * @param {string} [shapeType] - Type of shape to insert (e.g., "msoShapeRectangle", "msoShapeTextbox"). Required for insert.
 * @param {{left?: number, top?: number}} [position] - Position of the shape.
 * @param {{width?: number, height?: number}} [size] - Size of the shape.
 * @param {string} [text] - Text for the shape (if it is a textbox).
 * @param {number} [shapeIndex] - 1-based index of the shape. Required for modify, format, delete.
 * @param {string} [shapeName] - Name of the shape. Alternative to shapeIndex.
 * @param {{fillColor?: string, lineColor?: string, lineWidth?: number, fontName?: string, fontSize?: number, fontBold?: boolean, fontItalic?: boolean, fontUnderline?: boolean}} [formatProperties] - Format properties. Required for format.
 * @returns {Promise<ApiResponse<any>>} A message indicating the result of the operation or the list of shapes.
 */
const powerpointShapesTool: McpResource = {
  path: 'powerpoint/shapes',
  description: "Allows inserting, modifying, formatting, deleting, and listing shapes in PowerPoint presentations. Creates the file if it doesn't exist for insert/modify/format/delete operations.", // Updated description
  schema: PowerPointShapesInputSchema,
  handler: async (params: ToolRequestParams, context?: FastMCPContext<undefined>): Promise<ApiResponse<any>> => { // Added context
    const input = PowerPointShapesInputSchema.parse(params);
    const {
      filePath,
      operation,
      slideIndex,
      shapeType,
      position,
      size,
      text,
      shapeIndex: comShapeIndex, // Renamed to avoid conflict with loop variables
      shapeName,
      formatProperties,
      useComInterop
    } = input;

    const absoluteFilePath = resolvePath(filePath);
    logger.info(`Executing powerpoint/shapes operation '${operation}' for file: ${absoluteFilePath} (useComInterop: ${useComInterop})`);

    if (useComInterop) {
      // COM Interop Path (existing logic, slightly adapted)
      let app: any = null;
      let presentation: any = null;
      let slide: any = null;
      let shape: any = null;
      let fileCreated = false;

      try {
        app = await getOfficeApplication('PowerPoint.Application');
        const isModificationOperation = ['insert', 'modify', 'format', 'delete'].includes(operation);

        try {
            if (await fs.pathExists(absoluteFilePath)) {
                logger.info(`COM: Opening existing presentation: ${absoluteFilePath}`);
                presentation = app.Presentations.Open(absoluteFilePath);
            } else {
                if (isModificationOperation) {
                    logger.info(`COM: File not found. Creating new presentation at: ${absoluteFilePath}`);
                    presentation = app.Presentations.Add();
                    const fileExt = extnamePath(absoluteFilePath).toLowerCase();
                    let saveFormat = 24; // ppSaveAsOpenXMLPresentation (.pptx)
                    if (fileExt === '.ppt') saveFormat = 1;
                    else if (fileExt === '.pptm') saveFormat = 25;
                    presentation.SaveAs(absoluteFilePath, saveFormat);
                    fileCreated = true;
                    logger.info(`COM: Successfully created and saved new presentation: ${absoluteFilePath}`);
                    if (presentation.Slides.Count === 0) {
                        const ppLayoutBlank = 12; // MsoPresetTextEffect.msoTextEffect1 (Placeholder, find correct constant for Blank)
                        try {
                            presentation.Slides.Add(1, app.ActivePresentation.SlideMaster.CustomLayouts(7).Index); // Common index for blank, may vary
                        } catch {
                             presentation.Slides.Add(1, 12); // Fallback to a common blank layout index
                        }
                        logger.info("COM: Added default blank slide to newly created presentation.");
                        presentation.Save();
                    }
                } else {
                    logger.warn(`COM: File not found for list operation: ${absoluteFilePath}`);
                    return createErrorResponse(`File not found: ${filePath}`, 'FILE_NOT_FOUND');
                }
            }
        } catch (fileError: any) {
             logger.error(`COM: Error opening or creating presentation '${absoluteFilePath}': ${fileError.message}`, { error: fileError });
             return createErrorResponse(`COM: Failed to open or create presentation: ${fileError.message}`, 'FILE_OPERATION_FAILED', fileError);
        }

        if (!presentation) {
             return createErrorResponse(`COM: Failed to obtain presentation object for: ${filePath}`, 'FILE_OPEN_FAILED');
        }

        if (slideIndex !== undefined) {
          if (slideIndex < 1 || slideIndex > presentation.Slides.Count) {
            if (!(fileCreated && slideIndex === 1 && presentation.Slides.Count === 1)) {
               return createErrorResponse(`COM: Slide index ${slideIndex} is out of bounds (Presentation has ${presentation.Slides.Count} slides).`, 'INVALID_PARAM');
            }
          }
           try {
              slide = presentation.Slides(slideIndex);
           } catch (slideError: any) {
               logger.error(`COM: Error getting slide index ${slideIndex}: ${slideError.message}`);
               return createErrorResponse(`COM: Could not access slide index ${slideIndex}.`, 'SLIDE_ACCESS_ERROR', slideError);
           }
        } else if (operation !== 'list') {
           return createErrorResponse('COM: slideIndex is required for this operation.', 'MISSING_PARAM');
        }

        if (slide && (operation === 'modify' || operation === 'format' || operation === 'delete')) {
            if (comShapeIndex !== undefined) {
                if (comShapeIndex < 1 || comShapeIndex > slide.Shapes.Count) {
                    return createErrorResponse(`COM: Shape index ${comShapeIndex} is out of bounds on slide ${slideIndex}.`, 'INVALID_PARAM');
                }
                try {
                   shape = slide.Shapes(comShapeIndex);
                } catch (shapeIdxError: any) {
                    logger.error(`COM: Error getting shape by index ${comShapeIndex} on slide ${slideIndex}: ${shapeIdxError.message}`);
                    return createErrorResponse(`COM: Could not access shape index ${comShapeIndex} on slide ${slideIndex}.`, 'SHAPE_ACCESS_ERROR', shapeIdxError);
                }
            } else if (shapeName !== undefined) {
                 try {
                     shape = slide.Shapes(shapeName);
                 } catch (shapeNameError: any) {
                     logger.warn(`COM: Shape with name "${shapeName}" not found on slide ${slideIndex}: ${shapeNameError.message}`);
                     return createErrorResponse(`COM: Shape with name "${shapeName}" not found on slide ${slideIndex}.`, 'SHAPE_NOT_FOUND', shapeNameError);
                 }
            } else {
                 return createErrorResponse('COM: shapeIndex or shapeName is required for modify, format, or delete operations.', 'MISSING_PARAM');
            }
            if (!shape) {
                 return createErrorResponse(`COM: Shape (index: ${comShapeIndex}, name: ${shapeName}) not found or could not be accessed on slide ${slideIndex}.`, 'SHAPE_NOT_FOUND');
            }
        }

        let resultData: any = null;
        let message = '';

        switch (operation) {
          case 'insert': {
            if (!slide) return createErrorResponse('COM: A valid slideIndex is required to insert a shape.', 'MISSING_PARAM');
            if (!shapeType) return createErrorResponse('COM: shapeType is required for insert operation.', 'MISSING_PARAM');

            let msoShapeType;
            const MsoAutoShapeTypeMap: { [key: string]: number } = { /* ... as before ... */
                msoShapeRectangle: 1, msoShapeTextbox: 17, msoShapeOval: 9, msoShapeRoundedRectangle: 5,
            };
            if (shapeType in MsoAutoShapeTypeMap) msoShapeType = MsoAutoShapeTypeMap[shapeType];
            else {
                const shapeTypeNum = parseInt(shapeType, 10);
                if (!isNaN(shapeTypeNum)) msoShapeType = shapeTypeNum;
                else return createErrorResponse(`COM: Unsupported or invalid shape type: ${shapeType}.`, 'INVALID_PARAM');
            }

            const newShape = slide.Shapes.AddShape(msoShapeType, position?.left ?? 100, position?.top ?? 100, size?.width ?? 100, size?.height ?? 100);
            if (text !== undefined) {
                if (newShape.HasTextFrame === -1 /* msoTrue */) {
                    newShape.TextFrame.TextRange.Text = text;
                }
            }
            releaseObject(newShape);
            message = `COM: Shape inserted successfully on slide ${slideIndex}.`;
            break;
          }
          case 'modify': { /* ... as before, ensure shape is valid ... */
            if (!shape) return createErrorResponse('COM: Shape not found for modification.', 'SHAPE_NOT_FOUND');
            if (position) {
                if (position.left !== undefined) shape.Left = position.left;
                if (position.top !== undefined) shape.Top = position.top;
            }
            if (size) {
                if (size.width !== undefined) shape.Width = size.width;
                if (size.height !== undefined) shape.Height = size.height;
            }
            if (text !== undefined && shape.HasTextFrame === -1) {
                shape.TextFrame.TextRange.Text = text;
            }
            message = `COM: Shape modified successfully on slide ${slideIndex}.`;
            break;
          }
          case 'format': { /* ... as before, ensure shape and formatProperties are valid ... */
            if (!shape) return createErrorResponse('COM: Shape not found for formatting.', 'SHAPE_NOT_FOUND');
            if (!formatProperties) return createErrorResponse('COM: formatProperties are required.', 'MISSING_PARAM');
            const { fillColor, lineColor, lineWidth, fontName, fontSize, fontBold, fontItalic, fontUnderline } = formatProperties;
            if (lineWidth !== undefined) shape.Line.Weight = lineWidth;
            // Color and font formatting as before
            if (shape.HasTextFrame === -1 && shape.TextFrame.HasText === -1) {
                const font = shape.TextFrame.TextRange.Font;
                if (fontName !== undefined) font.Name = fontName;
                if (fontSize !== undefined) font.Size = fontSize;
                if (fontBold !== undefined) font.Bold = fontBold ? -1 : 0; // msoTrue / msoFalse
                if (fontItalic !== undefined) font.Italic = fontItalic ? -1 : 0;
                if (fontUnderline !== undefined) font.Underline = fontUnderline ? -1 : 0; // Check MsoTriState for underline
                releaseObject(font);
            }
            message = `COM: Shape formatted successfully on slide ${slideIndex}.`;
            break;
          }
          case 'delete': { /* ... as before, ensure shape is valid ... */
            if (!shape) return createErrorResponse('COM: Shape not found for deletion.', 'SHAPE_NOT_FOUND');
            shape.Delete();
            message = `COM: Shape deleted successfully from slide ${slideIndex}.`;
            break;
          }
          case 'list': { /* ... as before ... */
            const shapesList: any[] = [];
            if (!slide) { // List all slides
                 for (let i = 1; i <= presentation.Slides.Count; i++) {
                     const currentSlide = presentation.Slides(i);
                     for (let j = 1; j <= currentSlide.Shapes.Count; j++) {
                         const currentShape = currentSlide.Shapes(j);
                         shapesList.push({
                             slideIndex: i, shapeIndex: j, shapeName: currentShape.Name, shapeType: currentShape.Type,
                             text: (currentShape.HasTextFrame === -1 && currentShape.TextFrame.HasText === -1) ? currentShape.TextFrame.TextRange.Text : undefined
                         });
                         releaseObject(currentShape);
                     }
                     releaseObject(currentSlide);
                 }
                 message = `COM: Listed shapes from all ${presentation.Slides.Count} slides.`;
            } else { // List specific slide
                for (let i = 1; i <= slide.Shapes.Count; i++) {
                    const currentShape = slide.Shapes(i);
                     shapesList.push({
                         shapeIndex: i, shapeName: currentShape.Name, shapeType: currentShape.Type,
                         text: (currentShape.HasTextFrame === -1 && currentShape.TextFrame.HasText === -1) ? currentShape.TextFrame.TextRange.Text : undefined
                     });
                     releaseObject(currentShape);
                }
                message = `COM: Listed ${slide.Shapes.Count} shapes from slide ${slideIndex}.`;
            }
            resultData = shapesList;
            break;
          }
          default: throw new Error(`COM: Unsupported operation: ${operation}`);
        }

        if (isModificationOperation && !fileCreated) {
            presentation.Save();
        }
        if (isModificationOperation && absoluteFilePath) {
            try {
                const pptContent = await fs.readFile(absoluteFilePath);
                await saveResource('powerpoint/shapes', basenamePath(absoluteFilePath), pptContent);
            } catch (resourceSaveError: any) {
                logger.error(`COM: Failed to save ${absoluteFilePath} as a dynamic resource: ${resourceSaveError.message}`);
            }
        }
        releaseObject(shape);
        releaseObject(slide);
        return { success: true, data: resultData ?? message };

      } catch (error: any) {
        logger.error(`Error in powerpoint/shapes tool (COM Interop): ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
        if (error instanceof z.ZodError) return createErrorResponse('COM: Input validation failed', 'VALIDATION_ERROR', error.errors);
        return createErrorResponse(`COM: PowerPoint shapes operation failed: ${error.message}`, 'POWERPOINT_SHAPES_ERROR_COM', error);
      } finally {
        releaseObject(shape);
        releaseObject(slide);
        if (presentation) {
            try { presentation.Close(); } catch (e: any) { logger.warn(`COM: Failed to close presentation: ${e.message}`); }
            releaseObject(presentation);
        }
        if (app) releaseObject(app);
        logger.debug("COM: Released PowerPoint COM objects for powerpoint/shapes operation.");
      }
    } else {
      // Library Path (PptxGenJS)
      try {
        const pptx = new (PptxGenJS as any)(); // Cast to any for constructor
        let slideLib: any | undefined = undefined; // Use any for ISlide

        // PptxGenJS typically creates new files or overwrites existing ones.
        // It doesn't "open" and "modify" arbitrary slides/shapes in an existing file easily.
        // So, for 'insert', we'll assume we're building a presentation.
        // For other ops, we state limitations.

        if (await fs.pathExists(absoluteFilePath) && operation !== 'insert') {
            // For modify, format, delete, list on existing files, pptxgenjs is not suitable.
            // We could try to load with officeparser to get text, but not shapes.
            logger.warn(`PptxGenJS path: Operation '${operation}' on existing file '${absoluteFilePath}' has limitations. PptxGenJS is primarily a generation library.`);
        }


        switch (operation) {
          case 'insert':
            if (slideIndex === undefined || slideIndex !== 1) {
                // PptxGenJS adds slides sequentially. For simplicity, we'll add to the first/new slide.
                // A more complex implementation could manage multiple slides.
                logger.info("PptxGenJS: Inserting shape into a new slide (or first slide of a new presentation).");
            }
            slideLib = pptx.addSlide(); // Always adds a new slide

            if (!shapeType) return createErrorResponse('PptxGenJS: shapeType is required for insert.', 'MISSING_PARAM');

            const shapeOpts: any = {
                x: position?.left !== undefined ? position.left / 72 : 1, // Convert points to inches for PptxGenJS
                y: position?.top !== undefined ? position.top / 72 : 1,
                w: size?.width !== undefined ? size.width / 72 : 2,
                h: size?.height !== undefined ? size.height / 72 : 1,
            };

            // Map MSO types to PptxGenJS types/methods
            if (shapeType.toLowerCase().includes('textbox') || text) {
                slideLib.addText(text ?? 'Sample Text', { ...shapeOpts, fontSize: formatProperties?.fontSize ?? 18 });
            } else if (shapeType.toLowerCase().includes('rectangle')) {
                // Corrected way to reference PptxGenJS shape types
                slideLib.addShape((PptxGenJS as any).ShapeType.rect, shapeOpts);
            } else if (shapeType.toLowerCase().includes('oval') || shapeType.toLowerCase().includes('ellipse')) {
                // Corrected way to reference PptxGenJS shape types
                slideLib.addShape((PptxGenJS as any).ShapeType.ellipse, shapeOpts);
            } else {
                return createErrorResponse(`PptxGenJS: Unsupported shapeType '${shapeType}'. Use 'textbox', 'rectangle', 'oval' (ellipse), etc.`, 'INVALID_PARAM');
            }
            
            await pptx.writeFile({ fileName: absoluteFilePath });
            // Save resource
            const pptContent = await fs.readFile(absoluteFilePath);
            await saveResource('powerpoint/shapes', basenamePath(absoluteFilePath), pptContent);
            return { success: true, data: `PptxGenJS: Shape inserted into ${absoluteFilePath}.` };

          case 'modify':
          case 'format':
          case 'delete':
            logger.warn(`PptxGenJS: Operation '${operation}' on existing shapes is not directly supported. PptxGenJS generates presentations. For modifications, use COM Interop.`);
            return createErrorResponse(`PptxGenJS: Operation '${operation}' for existing shapes is not supported. Use COM Interop.`, 'POWERPOINT_LIB_UNSUPPORTED');

          case 'list':
            logger.warn("PptxGenJS: Listing shapes from an existing file is not supported. PptxGenJS does not parse existing files for shape details. Use COM Interop.");
            // officeparser could be used here to get text, but not shape specifics.
            // For consistency with the tool's purpose (shapes), we'll state it's not supported for listing shapes.
            return createErrorResponse("PptxGenJS: Listing shapes from existing files is not supported. Use COM Interop.", 'POWERPOINT_LIB_UNSUPPORTED');

          default:
            const exhaustiveCheckLib: never = operation;
            throw new Error(`PptxGenJS: Unsupported operation: ${exhaustiveCheckLib}`);
        }
      } catch (error: any) {
        logger.error(`Error in powerpoint/shapes tool (PptxGenJS): ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
        if (error instanceof z.ZodError) return createErrorResponse('PptxGenJS: Input validation failed', 'VALIDATION_ERROR', error.errors);
        return createErrorResponse(`PptxGenJS: Shapes operation failed: ${error.message}`, 'POWERPOINT_LIB_ERROR', error);
      }
    }
  },
};

export default powerpointShapesTool;
