import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext as Context } from '../../types/common.types';
import logger from '../../utils/logger';
import { getOfficeApplication, openWordDocument, releaseObject } from '../../utils/officeInterop';

// Define el esquema de entrada para la herramienta word-to-powerpoint
const WordToPowerpointInputSchema = z.object({
  wordFilePath: z.string().describe('Ruta al archivo Word de origen.'),
  powerpointFilePath: z.string().optional().describe('Ruta donde guardar la presentación PowerPoint de salida. Si no se proporciona, se creará una presentación temporal.'),
  operation: z.enum(['analyze', 'create', 'transfer']).describe('Operación a realizar: analyze, create, o transfer.'),
  headingLevelForNewSlide: z.number().int().min(1).max(9).optional().describe('Nivel de encabezado de Word que iniciará una nueva diapositiva (e.g., 1 para Heading 1).'),
  // Puedes añadir más parámetros opcionales aquí, como plantilla de PowerPoint, mapeo de estilos, etc.
});

type WordToPowerpointInput = z.infer<typeof WordToPowerpointInputSchema>;

/**
 * @tool office/word-to-powerpoint
 * @description Convierte un documento Word a una presentación PowerPoint.
 * Permite analizar la estructura del documento Word, crear una nueva presentación
 * y transferir el contenido del documento Word a la presentación, creando diapositivas
 * basadas en la estructura del documento (e.g., estilos de encabezado).
 * Utiliza COM Interop a través de winax para interactuar con Word y PowerPoint.
 * @param {object} params - Parámetros de entrada.
 * @param {string} params.wordFilePath - Ruta al archivo Word de origen.
 * @param {string} [params.powerpointFilePath] - Ruta donde guardar la presentación PowerPoint de salida.
 * @param {'analyze' | 'create' | 'transfer'} params.operation - Operación a realizar: analyze, create, o transfer.
 * @param {number} [params.headingLevelForNewSlide] - Nivel de encabezado de Word que iniciará una nueva diapositiva.
 * @returns {Promise<any>} - Resultado de la operación.
 */
const wordToPowerpointTool: McpResource = {
  path: 'office/word-to-powerpoint', // Cambiado de name a path
  description: 'Convierte un documento Word a una presentación PowerPoint.',
  schema: WordToPowerpointInputSchema, // Cambiado de inputSchema a schema
  handler: async (params: ToolRequestParams, context?: Context<any>): Promise<ApiResponse<any>> => {
    // Validar y parsear los parámetros de entrada
    const parsedParams = WordToPowerpointInputSchema.safeParse(params);

    if (!parsedParams.success) {
        return {
            success: false,
            error: {
                code: 'VALIDATION_ERROR',
                message: 'Invalid input parameters',
                details: parsedParams.error.errors,
            },
        };
    }

    const { wordFilePath, powerpointFilePath, operation, headingLevelForNewSlide } = parsedParams.data;

    let wordApp: any = null;
    let powerpointApp: any = null;
    let wordDoc: any = null;
    let powerpointPresentation: any = null;


    try {
      wordApp = await getOfficeApplication('Word.Application');
      powerpointApp = await getOfficeApplication('PowerPoint.Application');

      wordDoc = await openWordDocument(wordApp, wordFilePath, true, false); // Abrir Word read-only e invisible

      let result: any = null; // Declarar result aquí y inicializarlo

      switch (operation) {
        case 'analyze':
          // Implementación de la operación analyze
          // Podría analizar los estilos de encabezado, secciones, etc.
          logger.info(`Analizando documento Word: ${wordFilePath}`);
          // Ejemplo básico: contar párrafos y encabezados
          const paragraphCount = wordDoc.Paragraphs.Count;
          let headingCounts: { [key: number]: number } = {};
          for (let i = 1; i <= paragraphCount; i++) {
            const paragraph = wordDoc.Paragraphs.Item(i);
            // Simplificación: verificar si el estilo del párrafo es un encabezado
            // En una implementación real, se necesitaría un mapeo más robusto de estilos
            if (paragraph.Style && typeof paragraph.Style.NameLocal === 'string' && paragraph.Style.NameLocal.startsWith('Heading ')) {
                const level = parseInt(paragraph.Style.NameLocal.replace('Heading ', ''), 10);
                if (!isNaN(level)) {
                    headingCounts[level] = (headingCounts[level] || 0) + 1;
                }
            }
          }
          result = {
            message: `Análisis completado para ${wordFilePath}`,
            paragraphCount,
            headingCounts,
            // Añadir más detalles del análisis aquí
          };
          break;

        case 'create':
          // Implementación de la operación create
          logger.info('Creando nueva presentación PowerPoint');
          const presentation = powerpointApp.Presentations.Add();
          if (powerpointFilePath) {
            presentation.SaveAs(powerpointFilePath);
            result = { message: `Presentación creada y guardada en ${powerpointFilePath}` };
          } else {
            // Si no se especifica ruta, se deja abierta para la siguiente operación
            result = { message: 'Nueva presentación creada.' };
          }
          // No cerrar la presentación aquí si se va a usar en 'transfer'
          break;

        case 'transfer':
          // Implementación de la operación transfer
          logger.info(`Transfiriendo contenido de ${wordFilePath} a presentación`);

          let targetPresentation: any;
          if (powerpointFilePath) {
              // Abrir presentación existente si se especifica la ruta
              try {
                  targetPresentation = powerpointApp.Presentations.Open(powerpointFilePath);
              } catch (openError) {
                  logger.warn(`No se encontró presentación en ${powerpointFilePath}. Creando una nueva.`);
                  targetPresentation = powerpointApp.Presentations.Add();
              }
          } else {
              // Usar la presentación activa si no se especifica ruta
              if (powerpointApp.Presentations.Count === 0) {
                  logger.info('No hay presentaciones abiertas. Creando una nueva.');
                  targetPresentation = powerpointApp.Presentations.Add();
              } else {
                  targetPresentation = powerpointApp.ActivePresentation;
              }
          }

          // Lógica de transferencia: iterar sobre el documento Word
          // y crear diapositivas basadas en headingLevelForNewSlide
          let currentSlide: any = null;
          const paragraphs = wordDoc.Paragraphs;

          for (let i = 1; i <= paragraphs.Count; i++) {
              const paragraph = paragraphs.Item(i);
              const styleName = paragraph.Style ? paragraph.Style.NameLocal : '';
              const isHeading = styleName && typeof styleName === 'string' && styleName.startsWith('Heading ');
              let headingLevel: number | null = null;
              if (isHeading) {
                  headingLevel = parseInt(styleName.replace('Heading ', ''), 10);
              }

              const shouldStartNewSlide = headingLevelForNewSlide !== undefined && headingLevel !== null && headingLevel <= headingLevelForNewSlide;

              if (shouldStartNewSlide || currentSlide === null) {
                  // Añadir nueva diapositiva. Usar un diseño básico por ahora.
                  // pptLayoutTitleOnly = 11, pptLayoutText = 2, pptLayoutTitleAndContent = 1
                  const layout = shouldStartNewSlide && headingLevel === 1 ? 11 : 2; // Título solo para H1, Texto para otros
                  currentSlide = targetPresentation.Slides.Add(targetPresentation.Slides.Count + 1, layout);

                  if (shouldStartNewSlide && headingLevel !== null) {
                      // Si es un encabezado que inicia nueva diapositiva, usarlo como título
                      if (currentSlide.Shapes.HasTitle) {
                           currentSlide.Shapes.Title.TextFrame.TextRange.Text = paragraph.Range.Text.trim();
                      }
                      // Si hay texto después del título en el mismo párrafo, añadirlo al cuerpo
                      const remainingText = paragraph.Range.Text.trim().substring(paragraph.Range.Text.trim().indexOf('\n') + 1);
                       if (remainingText && currentSlide.Shapes.Placeholders.Count > 1) {
                           currentSlide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = remainingText;
                       }
                  } else {
                       // Si no es un encabezado que inicia nueva diapositiva o es el primer párrafo, añadirlo al cuerpo
                       if (currentSlide.Shapes.Placeholders.Count > 1) {
                           currentSlide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = paragraph.Range.Text.trim();
                       }
                  }

              } else {
                  // Añadir contenido al slide actual
                  if (currentSlide && currentSlide.Shapes.Placeholders.Count > 1) {
                      const bodyShape = currentSlide.Shapes.Placeholders.Item(2);
                      bodyShape.TextFrame.TextRange.InsertAfter(paragraph.Range.Text + '\n');
                  }
              }
          }

          if (powerpointFilePath) {
              targetPresentation.SaveAs(powerpointFilePath);
              result = { message: `Contenido transferido y presentación guardada en ${powerpointFilePath}` };
          } else {
              result = { message: 'Contenido transferido a la presentación activa.' };
          }

          // No cerrar la presentación aquí si se guardó o se dejó activa
          break;

        default:
          throw new Error(`Operación no soportada: ${operation}`);
      }

      // Cerrar el documento Word
      wordDoc.Close();

      // No cerrar las aplicaciones de Office aquí, OfficeInterop las gestiona

      // No cerrar las aplicaciones de Office aquí, OfficeInterop las gestiona
 
      return { success: true, data: result };
 
    } catch (error: any) {
      logger.error(`Error en wordToPowerpointTool: ${error.message}`);
      // Asegurarse de cerrar el documento Word si está abierto
      if (wordDoc) {
          try {
              wordDoc.Close(0); // wdDoNotSaveChanges = 0
              releaseObject(wordDoc);
          } catch (closeError) {
              logger.error(`Error al cerrar documento Word en catch: ${closeError}`);
          }
      }
       // Asegurarse de cerrar la presentación PowerPoint si está abierta
       if (powerpointPresentation) {
            try {
                powerpointPresentation.Close();
                releaseObject(powerpointPresentation);
            } catch (closeError) {
                logger.error(`Error al cerrar presentación PowerPoint en catch: ${closeError}`);
            }
        }
 
       // No cerrar las aplicaciones de Office aquí, OfficeInterop las gestiona
      return { success: false, error: { code: 'TOOL_ERROR', message: `Fallo en la herramienta word-to-powerpoint: ${error.message}` } };
    } finally {
        // Asegurarse de cerrar el documento Word si está abierto
        if (wordDoc) {
            try {
                wordDoc.Close(0); // wdDoNotSaveChanges = 0
                releaseObject(wordDoc);
            } catch (closeError) {
                logger.error(`Error al cerrar documento Word en finally: ${closeError}`);
            }
        }
 
         // Asegurarse de cerrar la presentación PowerPoint si fue creada en este handler y no se guardó
         if (powerpointPresentation && !powerpointFilePath) {
             try {
                 powerpointPresentation.Close();
                 releaseObject(powerpointPresentation);
             } catch (closeError) {
                 logger.error(`Error al cerrar presentación PowerPoint en finally: ${closeError}`);
             }
         }
    }
  },
 };

export default wordToPowerpointTool;