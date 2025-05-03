// src/tools/powerpoint/shapes.tool.ts

import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import logger from '../../utils/logger';
import { saveResource } from '../dynamic/resources.tool'; // Importar saveResource
import * as fs from 'fs-extra'; // Importar fs para leer el archivo PowerPoint
import * as path from 'path'; // Importar path

// Define el esquema de entrada para la herramienta powerpoint/shapes
const PowerPointShapesInputSchema = z.object({
  filePath: z.string().describe('Ruta al archivo PowerPoint.'),
  operation: z.enum(['insert', 'modify', 'format', 'delete', 'list']).describe('Operación a realizar: insert, modify, format, delete, list.'),
  slideIndex: z.number().int().positive().optional().describe('Índice de la diapositiva (1-basado). Requerido para insert, modify, format, delete.'),
  shapeType: z.string().optional().describe('Tipo de forma a insertar (e.g., "msoShapeRectangle", "msoShapeTextbox"). Requerido para insert.'),
  position: z.object({
    left: z.number().optional(),
    top: z.number().optional(),
  }).optional().describe('Posición de la forma (left, top). Opcional para insert, modify.'),
  size: z.object({
    width: z.number().optional(),
    height: z.number().optional(),
  }).optional().describe('Tamaño de la forma (width, height). Opcional para insert, modify.'),
  text: z.string().optional().describe('Texto para la forma (si es un cuadro de texto). Opcional para insert, modify.'),
  shapeIndex: z.number().int().positive().optional().describe('Índice de la forma en la diapositiva (1-basado). Requerido para modify, format, delete.'),
  shapeName: z.string().optional().describe('Nombre de la forma. Alternativa a shapeIndex para modify, format, delete.'),
  formatProperties: z.object({
    fillColor: z.string().optional().describe('Color de relleno (e.g., "RGB(255, 0, 0)").'),
    lineColor: z.string().optional().describe('Color de línea (e.g., "RGB(0, 0, 255)").'),
    lineWidth: z.number().optional().describe('Ancho de línea.'),
    fontName: z.string().optional().describe('Nombre de la fuente.'),
    fontSize: z.number().optional().describe('Tamaño de la fuente.'),
    fontBold: z.boolean().optional().describe('Negrita.'),
    fontItalic: z.boolean().optional().describe('Itálica.'),
    fontUnderline: z.boolean().optional().describe('Subrayado.'),
  }).optional().describe('Propiedades de formato a aplicar. Requerido para format.'),
});

type PowerPointShapesInput = z.infer<typeof PowerPointShapesInputSchema>;

/**
 * @tool powerpoint/shapes
 * @description Permite insertar, modificar, dar formato, eliminar y listar formas en presentaciones de PowerPoint.
 * Utiliza COM Interop a través de winax.
 * @param {string} filePath - Ruta al archivo PowerPoint.
 * @param {'insert' | 'modify' | 'format' | 'delete' | 'list'} operation - Operación a realizar.
 * @param {number} [slideIndex] - Índice de la diapositiva (1-basado). Requerido para insert, modify, format, delete.
 * @param {string} [shapeType] - Tipo de forma a insertar (e.g., "msoShapeRectangle", "msoShapeTextbox"). Requerido para insert.
 * @param {{left?: number, top?: number}} [position] - Posición de la forma.
 * @param {{width?: number, height?: number}} [size] - Tamaño de la forma.
 * @param {string} [text] - Texto para la forma (si es un cuadro de texto).
 * @param {number} [shapeIndex] - Índice de la forma (1-basado). Requerido para modify, format, delete.
 * @param {string} [shapeName] - Nombre de la forma. Alternativa a shapeIndex.
 * @param {{fillColor?: string, lineColor?: string, lineWidth?: number, fontName?: string, fontSize?: number, fontBold?: boolean, fontItalic?: boolean, fontUnderline?: boolean}} [formatProperties] - Propiedades de formato. Requerido para format.
 * @returns {Promise<string>} Un mensaje indicando el resultado de la operación.
 */
const powerpointShapesTool: McpResource = {
  path: 'powerpoint/shapes',
  description: 'Permite insertar, modificar, dar formato, eliminar y listar formas en presentaciones de PowerPoint.',
  schema: PowerPointShapesInputSchema,
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => {
    let app: any = null;
    let presentation: any = null;
    let slide: any = null;
    let filePath: string | undefined; // Declarar filePath fuera del try y permitir undefined

    try {
      // Validate input parameters using the Zod schema
      const input = PowerPointShapesInputSchema.parse(params);
      filePath = input.filePath; // Asignar filePath aquí


      const {
        operation, // Eliminar filePath de la desestructuración aquí
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
                // Guardar la presentación antes de cerrarla
                presentation.Save();
                // Guardar el archivo PowerPoint modificado como un recurso dinámico
                // Esto se hace en el finally porque Save() ocurre aquí para todas las operaciones de modificación.
                // No necesitamos verificar la operación específica aquí.
                // Asegurarse de que filePath tiene un valor antes de intentar leer el archivo
                if (filePath) {
                    try {
                        const pptContent = await fs.readFile(filePath, null); // Leer como Buffer
                        await saveResource('powerpoint/shapes', path.basename(filePath), pptContent);
                        // logger.info(`Saved ${filePath} as a dynamic resource.`);
                    } catch (resourceSaveError: any) {
                        // logger.error(`Failed to save ${filePath} as a dynamic resource: ${resourceSaveError.message}`);
                        // Continuar la ejecución aunque falle el guardado del recurso
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
    // Añadir un retorno al final para cubrir todos los casos posibles
    // Esto solo se alcanzará si no se lanzó un error o se retornó antes.
    // En un escenario ideal, todos los casos del switch deberían retornar.
    // Pero para satisfacer al linter, añadimos este retorno de fallback.
    return { success: false, error: { code: 'UNHANDLED_CASE', message: 'La operación de formas de PowerPoint no retornó un resultado explícito.' } };
  },
};

export default powerpointShapesTool;

// Helper function placeholder for parsing RGB string like "RGB(255, 0, 0)"
// This would need to be implemented based on how winax handles color values.
// function parseRGB(rgbString: string): number {
//     // Example parsing logic (needs refinement based on actual format and winax capabilities)
//     const match = rgbString.match(/^RGB\((\d+),\s*(\d+),\s*(\d+)\)$/);
//     if (match) {
//         const r = parseInt(match[1], 10);
//         const g = parseInt(match[2], 10);
//         const b = parseInt(match[3], 10);
//         // COM RGB values are typically BGR, so (B * 256^2) + (G * 256^1) + (R * 256^0)
//         return (b << 16) | (g << 8) | r;
//     }
//     throw new Error(`Invalid RGB string format: ${rgbString}`);
// }
