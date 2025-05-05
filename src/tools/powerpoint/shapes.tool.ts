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
import { McpResource, ToolRequestParams, ApiResponse, FastMCPContext } from '../../types/common.types'; // Added FastMCPContext
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import logger from '../../utils/logger';
import { saveResource } from '../dynamic/resources.tool';
import * as fs from 'fs-extra';
import * as path from 'path';
import { validateFilePath } from '../../utils/security'; // Added security import

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
    let app: any = null;
    let presentation: any = null;
    let slide: any = null;
    let shape: any = null; // Moved declaration here
    let fileCreated = false; // Flag to track file creation
    let absoluteFilePath: string | undefined; // Define here for use in finally

    try {
      // Validate input parameters using the Zod schema
      const input = PowerPointShapesInputSchema.parse(params);
      absoluteFilePath = path.resolve(input.filePath); // Assign resolved path
      logger.info(`Executing powerpoint/shapes operation '${input.operation}' for file: ${absoluteFilePath}`);


      const {
        operation,
        slideIndex,
        shapeType,
        position,
        size,
        text,
        shapeIndex,
        shapeName,
        formatProperties
      } = input;

      app = await getOfficeApplication('PowerPoint.Application');

      // --- Create if not exists logic (for insert/modify/format/delete) ---
      const isModificationOperation = ['insert', 'modify', 'format', 'delete'].includes(operation);

      try {
          if (await fs.pathExists(absoluteFilePath)) {
              logger.info(`Opening existing presentation: ${absoluteFilePath}`);
              presentation = app.Presentations.Open(absoluteFilePath);
          } else {
              if (isModificationOperation) {
                  logger.info(`File not found. Creating new presentation at: ${absoluteFilePath}`);
                  presentation = app.Presentations.Add(); // Create new presentation
                  // Save the new presentation immediately
                  // Determine format based on extension (default to pptx)
                  const fileExt = path.extname(absoluteFilePath).toLowerCase();
                  let saveFormat = 24; // ppSaveAsOpenXMLPresentation (.pptx)
                  if (fileExt === '.ppt') saveFormat = 1; // ppSaveAsPresentation (.ppt)
                  else if (fileExt === '.pptm') saveFormat = 25; // ppSaveAsOpenXMLPresentationMacroEnabled (.pptm)

                  presentation.SaveAs(absoluteFilePath, saveFormat);
                  fileCreated = true;
                  logger.info(`Successfully created and saved new presentation: ${absoluteFilePath}`);
                  // Add a default blank slide if creating for insert/modify/format/delete
                  if (presentation.Slides.Count === 0) {
                      const ppLayoutBlank = 12; // Assuming 12 is the constant for a blank layout
                      presentation.Slides.Add(1, ppLayoutBlank);
                      logger.info("Added default blank slide to newly created presentation.");
                      // Need to save again after adding the slide
                      presentation.Save();
                  }

              } else { // 'list' operation and file doesn't exist
                  logger.warn(`File not found for list operation: ${absoluteFilePath}`);
                  return createErrorResponse(`File not found: ${input.filePath}`, 'FILE_NOT_FOUND');
              }
          }
      } catch (fileError: any) {
           logger.error(`Error opening or creating presentation '${absoluteFilePath}': ${fileError.message}`, { error: fileError });
           return createErrorResponse(`Failed to open or create presentation: ${fileError.message}`, 'FILE_OPERATION_FAILED', fileError);
      }
      // --- End create if not exists logic ---

      if (!presentation) {
           return createErrorResponse(`Failed to obtain presentation object for: ${input.filePath}`, 'FILE_OPEN_FAILED');
      }


      // Get the target slide if required
      if (slideIndex !== undefined) {
        if (slideIndex < 1 || slideIndex > presentation.Slides.Count) {
          // Allow inserting into slide 1 even if count is 0 (handled by Add slide logic above)
          if (!(fileCreated && slideIndex === 1 && presentation.Slides.Count === 1)) {
             return createErrorResponse(`Slide index ${slideIndex} is out of bounds (Presentation has ${presentation.Slides.Count} slides).`, 'INVALID_PARAM');
          }
        }
         try {
            slide = presentation.Slides(slideIndex);
         } catch (slideError: any) {
             logger.error(`Error getting slide index ${slideIndex}: ${slideError.message}`);
             return createErrorResponse(`Could not access slide index ${slideIndex}.`, 'SLIDE_ACCESS_ERROR', slideError);
         }
      } else if (operation !== 'list') {
         // slideIndex is required for operations other than list
         return createErrorResponse('slideIndex is required for this operation.', 'MISSING_PARAM');
      }


      // Get the target shape if required (for modify, format, delete)
      if (slide && (operation === 'modify' || operation === 'format' || operation === 'delete')) {
          if (shapeIndex !== undefined) {
              if (shapeIndex < 1 || shapeIndex > slide.Shapes.Count) {
                  return createErrorResponse(`Shape index ${shapeIndex} is out of bounds on slide ${slideIndex}.`, 'INVALID_PARAM');
              }
              try {
                 shape = slide.Shapes(shapeIndex);
              } catch (shapeIdxError: any) {
                  logger.error(`Error getting shape by index ${shapeIndex} on slide ${slideIndex}: ${shapeIdxError.message}`);
                  return createErrorResponse(`Could not access shape index ${shapeIndex} on slide ${slideIndex}.`, 'SHAPE_ACCESS_ERROR', shapeIdxError);
              }
          } else if (shapeName !== undefined) {
               try {
                   shape = slide.Shapes(shapeName);
               } catch (shapeNameError: any) {
                   logger.warn(`Shape with name "${shapeName}" not found on slide ${slideIndex}: ${shapeNameError.message}`);
                   return createErrorResponse(`Shape with name "${shapeName}" not found on slide ${slideIndex}.`, 'SHAPE_NOT_FOUND', shapeNameError);
               }
          } else {
               return createErrorResponse('shapeIndex or shapeName is required for modify, format, or delete operations.', 'MISSING_PARAM');
          }
          if (!shape) {
               // Should be caught by try/catch above, but safeguard
               return createErrorResponse(`Shape (index: ${shapeIndex}, name: ${shapeName}) not found or could not be accessed on slide ${slideIndex}.`, 'SHAPE_NOT_FOUND');
          }
      }

      let resultData: any = null;
      let message = '';


      switch (operation) {
        case 'insert': {
          if (!slide) { // Ensure slide is valid for insert
              return createErrorResponse('A valid slideIndex is required to insert a shape.', 'MISSING_PARAM');
          }
          if (!shapeType) {
            return createErrorResponse('shapeType is required for insert operation.', 'MISSING_PARAM');
          }

          let msoShapeType;
          // Use a more robust way to get MsoAutoShapeType constants if possible
          // This might involve accessing the PowerPoint type library constants directly
          // For now, using a hardcoded map based on common values
          const MsoAutoShapeTypeMap: { [key: string]: number } = {
              msoShapeRectangle: 1,
              msoShapeTextbox: 17,
              msoShapeOval: 9,
              msoShapeRoundedRectangle: 5,
              // Add other common shapes as needed
          };

          if (shapeType in MsoAutoShapeTypeMap) {
              msoShapeType = MsoAutoShapeTypeMap[shapeType];
          } else {
              // Try parsing as number if it's not in the map
              const shapeTypeNum = parseInt(shapeType, 10);
              if (!isNaN(shapeTypeNum)) {
                  msoShapeType = shapeTypeNum;
                  logger.warn(`Using numeric value ${msoShapeType} for shapeType. Ensure this is a valid MsoAutoShapeType constant.`);
              } else {
                 return createErrorResponse(`Unsupported or invalid shape type: ${shapeType}. Provide a known name (e.g., msoShapeRectangle) or a valid MsoAutoShapeType enum value.`, 'INVALID_PARAM');
              }
          }


          const left = position?.left ?? 100; // Default position
          const top = position?.top ?? 100;
          const width = size?.width ?? 100; // Default size
          const height = size?.height ?? 100;

          const newShape = slide.Shapes.AddShape(msoShapeType, left, top, width, height);
          releaseObject(newShape); // Release the newly created shape object reference

          if (text !== undefined) {
              // Re-acquire the shape to add text (AddShape returns an object)
              const addedShape = slide.Shapes(slide.Shapes.Count); // Assume it's the last shape added
              if (addedShape.HasTextFrame === -1 /* msoTrue */) {
                  addedShape.TextFrame.TextRange.Text = text;
              }
              releaseObject(addedShape);
          }

          message = `Shape inserted successfully on slide ${slideIndex}.`;
          logger.info(message);
          break;
        }

        case 'modify': {
            if (!shape) { // Ensure shape is valid
                return createErrorResponse('Shape not found for modification.', 'SHAPE_NOT_FOUND');
            }

            if (position) {
                if (position.left !== undefined) shape.Left = position.left;
                if (position.top !== undefined) shape.Top = position.top;
            }
            if (size) {
                if (size.width !== undefined) shape.Width = size.width;
                if (size.height !== undefined) shape.Height = size.height;
            }
            if (text !== undefined && shape.HasTextFrame === -1 /* msoTrue */) {
                shape.TextFrame.TextRange.Text = text;
            }

            message = `Shape modified successfully on slide ${slideIndex}.`;
            logger.info(message);
            break;
        }

        case 'format': {
            if (!shape) { // Ensure shape is valid
                return createErrorResponse('Shape not found for formatting.', 'SHAPE_NOT_FOUND');
            }
            if (!formatProperties) {
                return createErrorResponse('formatProperties are required for format operation.', 'MISSING_PARAM');
            }

            const { fillColor, lineColor, lineWidth, fontName, fontSize, fontBold, fontItalic, fontUnderline } = formatProperties;

            // Basic formatting - requires more robust implementation for colors etc.
            if (fillColor !== undefined) logger.warn('Fill color formatting is basic.'); // Placeholder
            if (lineColor !== undefined) logger.warn('Line color formatting is basic.'); // Placeholder
            if (lineWidth !== undefined) shape.Line.Weight = lineWidth;

            if (shape.HasTextFrame === -1 /* msoTrue */ && shape.TextFrame.HasText === -1 /* msoTrue */) {
                const font = shape.TextFrame.TextRange.Font;
                if (fontName !== undefined) font.Name = fontName;
                if (fontSize !== undefined) font.Size = fontSize;
                if (fontBold !== undefined) font.Bold = fontBold ? -1 : 0;
                if (fontItalic !== undefined) font.Italic = fontItalic ? -1 : 0;
                if (fontUnderline !== undefined) font.Underline = fontUnderline ? -1 : 0;
                releaseObject(font);
            } else if (fontName || fontSize || fontBold || fontItalic || fontUnderline) {
                 logger.warn('Font formatting requested for a shape without text or text frame.');
            }

            message = `Shape formatted successfully on slide ${slideIndex}.`;
            logger.info(message);
            break;
        }

        case 'delete': {
            if (!shape) { // Ensure shape is valid
                return createErrorResponse('Shape not found for deletion.', 'SHAPE_NOT_FOUND');
            }
            shape.Delete();
            message = `Shape deleted successfully from slide ${slideIndex}.`;
            logger.info(message);
            break;
        }

        case 'list': {
            const shapesList: any[] = [];
            if (!slide) {
                 // List shapes for all slides
                 logger.info("Listing shapes for all slides.");
                 for (let i = 1; i <= presentation.Slides.Count; i++) {
                     const currentSlide = presentation.Slides(i);
                     for (let j = 1; j <= currentSlide.Shapes.Count; j++) {
                         const currentShape = currentSlide.Shapes(j);
                         shapesList.push({
                             slideIndex: i,
                             shapeIndex: j,
                             shapeName: currentShape.Name,
                             shapeType: currentShape.Type, // MsoShapeType enum value
                             text: (currentShape.HasTextFrame === -1 && currentShape.TextFrame.HasText === -1) ? currentShape.TextFrame.TextRange.Text : undefined
                         });
                         releaseObject(currentShape);
                     }
                     releaseObject(currentSlide);
                 }
                 message = `Listed shapes from all ${presentation.Slides.Count} slides.`;
            } else {
                // List shapes for a specific slide
                logger.info(`Listing shapes for slide ${slideIndex}.`);
                for (let i = 1; i <= slide.Shapes.Count; i++) {
                    const currentShape = slide.Shapes(i);
                     shapesList.push({
                         shapeIndex: i,
                         shapeName: currentShape.Name,
                         shapeType: currentShape.Type, // MsoShapeType enum value
                         text: (currentShape.HasTextFrame === -1 && currentShape.TextFrame.HasText === -1) ? currentShape.TextFrame.TextRange.Text : undefined
                     });
                     releaseObject(currentShape);
                }
                message = `Listed ${slide.Shapes.Count} shapes from slide ${slideIndex}.`;
            }
            resultData = shapesList;
            logger.info(message);
            break; // Added break
        }

        default:
          // Should not happen due to enum validation
          throw new Error(`Unsupported operation: ${operation}`);
      }

      // Save the presentation if modified
      if (isModificationOperation && !fileCreated) {
          presentation.Save();
          logger.info(`Presentation saved: ${absoluteFilePath}`);
      }

      // Save the modified PowerPoint file as a dynamic resource
      if (isModificationOperation && absoluteFilePath) {
          try {
              const pptContent = await fs.readFile(absoluteFilePath, null); // Read as Buffer
              await saveResource('powerpoint/shapes', path.basename(absoluteFilePath), pptContent);
              logger.info(`Saved ${absoluteFilePath} as a dynamic resource.`);
          } catch (resourceSaveError: any) {
              logger.error(`Failed to save ${absoluteFilePath} as a dynamic resource: ${resourceSaveError.message}`);
              // Continue execution even if resource saving fails
          }
      }

      // Release shape object if it was obtained
      releaseObject(shape);
      // Release slide object if it was obtained
      releaseObject(slide);

      return { success: true, data: resultData ?? message }; // Return data for list, message otherwise

    } catch (error: any) {
      logger.error(`Error in powerpoint/shapes tool: ${error.message}`, { error: error instanceof z.ZodError ? error.errors : error });
       if (error instanceof z.ZodError) {
          return createErrorResponse('Input validation failed', 'VALIDATION_ERROR', error.errors);
      }
      return createErrorResponse(`PowerPoint shapes operation failed: ${error.message}`, 'POWERPOINT_SHAPES_ERROR', error);
    } finally {
        // Release slide and shape again just in case they weren't released in the try block due to error
        releaseObject(shape);
        releaseObject(slide);
        // Ensure presentation is closed if it was opened/created
        if (presentation) {
            try {
                // Close without saving if just created (already saved by SaveAs)
                // Otherwise, close normally (Save happened in try block if needed)
                presentation.Close();
            } catch (closeError: any) {
                logger.warn(`[OfficeInterop] Failed to close presentation: ${closeError.message}`);
            }
            releaseObject(presentation);
        }
        // Release app object
        if (app) {
            releaseObject(app);
        }
        logger.debug("Released PowerPoint COM objects for powerpoint/shapes operation.");
    }
  },
};

export default powerpointShapesTool;
