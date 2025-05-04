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
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import logger from '../../utils/logger';
import { saveResource } from '../dynamic/resources.tool'; // Import saveResource
import * as fs from 'fs-extra'; // Import fs to read the PowerPoint file
import * as path from 'path'; // Import path

// Define the input schema for the powerpoint/shapes tool
const PowerPointShapesInputSchema = z.object({
  filePath: z.string().describe('Path to the PowerPoint file.'),
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
 * @returns {Promise<string>} A message indicating the result of the operation.
 */
const powerpointShapesTool: McpResource = {
  path: 'powerpoint/shapes',
  description: 'Allows inserting, modifying, formatting, deleting, and listing shapes in PowerPoint presentations.',
  schema: PowerPointShapesInputSchema,
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => {
    let app: any = null;
    let presentation: any = null;
    let slide: any = null;
    let filePath: string | undefined; // Declare filePath outside the try and allow undefined

    try {
      // Validate input parameters using the Zod schema
      const input = PowerPointShapesInputSchema.parse(params);
      filePath = input.filePath; // Assign filePath here


      const {
        operation, // Remove filePath from destructuring here
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
      presentation = app.Presentations.Open(filePath);

      if (slideIndex !== undefined) {
        if (slideIndex < 1 || slideIndex > presentation.Slides.Count) {
          throw new Error(`Slide index ${slideIndex} is out of bounds.`);
        }
        slide = presentation.Slides(slideIndex);
      } else if (operation !== 'list') {
         throw new Error('slideIndex is required for this operation.');
      }


      let shape;
      if (slideIndex !== undefined) { // Only try to get shape if slideIndex was provided
          if (shapeIndex !== undefined) {
              if (slide) {
                  if (shapeIndex < 1 || shapeIndex > slide.Shapes.Count) {
                      throw new Error(`Shape index ${shapeIndex} is out of bounds on slide ${slideIndex}.`);
                  }
                  shape = slide.Shapes(shapeIndex);
              }
          } else if (shapeName !== undefined) {
               if (slide) {
                   try {
                       shape = slide.Shapes(shapeName);
                   } catch (e) {
                       throw new Error(`Shape with name "${shapeName}" not found on slide ${slideIndex}.`);
                   }
               }
          } else if (operation !== 'insert' && operation !== 'list') {
               throw new Error('shapeIndex or shapeName is required for this operation for modify, format, and delete operations.');
          }
      }


      switch (operation) {
        case 'insert': {
          if (!shapeType) {
            throw new Error('shapeType is required for insert operation.');
          }

          let msoShapeType;
          const MsoAutoShapeType = { // Partial mapping for common shapes
              msoShapeRectangle: 1,
              msoShapeTextbox: 17,
              // Add other shapes as needed based on MsoAutoShapeType enum
              // https://learn.microsoft.com/en-us/office/vba/api/office.msoautoshapetype
          };

          if (shapeType in MsoAutoShapeType) {
              msoShapeType = MsoAutoShapeType[shapeType as keyof typeof MsoAutoShapeType];
          } else {
              throw new Error(`Unsupported shape type: ${shapeType}.`);
          }


          const left = position?.left ?? 100; // Default position
          const top = position?.top ?? 100;
          const width = size?.width ?? 100; // Default size
          const height = size?.height ?? 100;

          const newShape = slide.Shapes.AddShape(msoShapeType, left, top, width, height);

          if (text !== undefined && newShape.HasTextFrame === -1 /* msoTrue */) {
              newShape.TextFrame.TextRange.Text = text;
          }

          presentation.Save();
          return { success: true, data: `Shape inserted successfully on slide ${slideIndex}.` };
        }

        case 'modify': {
            if (!shape) {
                throw new Error('Shape not found for modification.');
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

            presentation.Save();
            return { success: true, data: `Shape modified successfully on slide ${slideIndex}.` };
        }

        case 'format': {
            if (!shape) {
                throw new Error('Shape not found for formatting.');
            }
            if (!formatProperties) {
                throw new Error('formatProperties are required for format operation.');
            }

            const { fillColor, lineColor, lineWidth, fontName, fontSize, fontBold, fontItalic, fontUnderline } = formatProperties;

            if (fillColor !== undefined) {
                 logger.warn('Fill color formatting is a placeholder. Requires proper RGB parsing and COM interaction.');
            }
            if (lineColor !== undefined) {
                 logger.warn('Line color formatting is a placeholder. Requires proper RGB parsing and COM interaction.');
            }
            if (lineWidth !== undefined) {
                shape.Line.Weight = lineWidth;
            }

            if (shape.HasTextFrame === -1 /* msoTrue */ && shape.TextFrame.HasText === -1 /* msoTrue */) {
                const font = shape.TextFrame.TextRange.Font;
                if (fontName !== undefined) font.Name = fontName;
                if (fontSize !== undefined) font.Size = fontSize;
                if (fontBold !== undefined) font.Bold = fontBold ? -1 /* msoTrue */ : 0 /* msoFalse */;
                if (fontItalic !== undefined) font.Italic = fontItalic ? -1 /* msoTrue */ : 0 /* msoFalse */;
                if (fontUnderline !== undefined) font.Underline = fontUnderline ? -1 /* msoTrue */ : 0 /* msoFalse */;
            } else if ( (fontName !== undefined || fontSize !== undefined || fontBold !== undefined || fontItalic !== undefined || fontUnderline !== undefined) && !(shape.HasTextFrame === -1 /* msoTrue */ && shape.TextFrame.HasText === -1 /* msoTrue */) ) {
                 logger.warn('Font formatting requested for a shape without text or text frame.');
            }


            presentation.Save();
            return { success: true, data: `Shape formatted successfully on slide ${slideIndex}.` };
        }

        case 'delete': {
            if (!shape) {
                throw new Error('Shape not found for deletion.');
            }
            shape.Delete();
            presentation.Save();
            return { success: true, data: `Shape deleted successfully from slide ${slideIndex}.` };
        }

        case 'list': {
            if (!slide) {
                 // List shapes for all slides if slideIndex is not provided
                 const allShapes: { slideIndex: number, shapeIndex: number, shapeName: string, shapeType: string, text?: string }[] = [];
                 for (let i = 1; i <= presentation.Slides.Count; i++) {
                     const currentSlide = presentation.Slides(i);
                     for (let j = 1; j <= currentSlide.Shapes.Count; j++) {
                         const currentShape = currentSlide.Shapes(j);
                         allShapes.push({
                             slideIndex: i,
                             shapeIndex: j,
                             shapeName: currentShape.Name,
                             shapeType: currentShape.Type, // MsoShapeType enum value
                             text: (currentShape.HasTextFrame === -1 /* msoTrue */ && currentShape.TextFrame.HasText === -1 /* msoTrue */) ? currentShape.TextFrame.TextRange.Text : undefined
                         });
                     }
                 }
                 return { success: true, data: allShapes };

            } else {
                // List shapes for a specific slide
                const slideShapes: { shapeIndex: number, shapeName: string, shapeType: string, text?: string }[] = [];
                for (let i = 1; i <= slide.Shapes.Count; i++) {
                    const currentShape = slide.Shapes(i);
                     slideShapes.push({
                         shapeIndex: i,
                         shapeName: currentShape.Name,
                         shapeType: currentShape.Type, // MsoShapeType enum value
                         text: (currentShape.HasTextFrame === -1 /* msoTrue */ && currentShape.TextFrame.HasText === -1 /* msoTrue */) ? currentShape.TextFrame.TextRange.Text : undefined
                     });
                }
                return { success: true, data: slideShapes };
            }
        }

        default:
          throw new Error(`Unsupported operation: ${operation}`);
      }

    } catch (error: any) {
      logger.error(`Error in powerpoint/shapes tool: ${error.message}`);
      return { success: false, error: { code: 'POWERPOINT_SHAPES_ERROR', message: `Failed to perform PowerPoint shapes operation: ${error.message}` } };
    } finally {
        // Ensure presentation is closed if it was opened
        if (presentation) {
            try {
                // Save the presentation before closing it
                presentation.Save();
                // Save the modified PowerPoint file as a dynamic resource
                // This is done in the finally block because Save() happens here for all modification operations.
                // We don't need to check the specific operation here.
                // Ensure filePath has a value before attempting to read the file
                if (filePath) {
                    try {
                        const pptContent = await fs.readFile(filePath, null); // Read as Buffer
                        await saveResource('powerpoint/shapes', path.basename(filePath), pptContent);
                        // logger.info(`Saved ${filePath} as a dynamic resource.`);
                    } catch (resourceSaveError: any) {
                        // logger.error(`Failed to save ${filePath} as a dynamic resource: ${resourceSaveError.message}`);
                        // Continue execution even if resource saving fails
                    }
                }
                presentation.Close(); // Close after saving
            } catch (closeError: any) {
                logger.warn(`[OfficeInterop] Failed to close presentation: ${closeError.message}`);
            }
            releaseObject(presentation);
        }
        // The app object lifecycle is managed by getOfficeApplication,
        // so we don't necessarily quit it here unless we are sure we opened it
        // and no other operations are pending. Releasing the object reference is safer.
        releaseObject(app);
    }
    // Add a return at the end to cover all possible cases
    // This will only be reached if no error was thrown or returned before.
    // In an ideal scenario, all switch cases should return.
    // But to satisfy the linter, we add this fallback return.
    return { success: false, error: { code: 'UNHANDLED_CASE', message: 'PowerPoint shapes operation did not return an explicit result.' } };
  },
};

export default powerpointShapesTool;
