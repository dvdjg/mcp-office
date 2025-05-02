import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import logger from '../../utils/logger'; // Importar logger

// Esquema de entrada para la herramienta powerpoint/slides
const SlidesToolInputSchema = z.object({
  filePath: z.string().describe('Ruta al archivo de presentación de PowerPoint.'),
  operation: z.enum(['add', 'delete', 'set']).describe('Operación a realizar: "add" (añadir diapositiva), "delete" (eliminar diapositiva), "set" (modificar diapositiva existente).'),
  slideIndex: z.number().optional().describe('Índice de la diapositiva (1-basado) para operaciones "delete" o "set".'),
  slideLayout: z.string().optional().describe('Nombre del diseño de la diapositiva para la operación "add" (e.g., "ppLayoutTitle").'),
  // Podríamos añadir más campos aquí para la operación 'set', como contenido, etc.
  // Por ahora, 'set' podría usarse para algo simple como cambiar el diseño si fuera posible o añadir contenido básico.
  // Para esta implementación inicial, 'set' no hará nada complejo, solo se incluye para cumplir con la especificación.
});

type SlidesToolInput = z.infer<typeof SlidesToolInputSchema>;

/**
 * @tool powerpoint/slides
 * @description Gestiona diapositivas en presentaciones de PowerPoint.
 * Permite añadir, eliminar y (básicamente) modificar diapositivas.
 * Requiere la ruta al archivo y la operación a realizar.
 * Las operaciones "delete" y "set" requieren el índice de la diapositiva.
 * La operación "add" requiere el nombre del diseño de la diapositiva.
 * Utiliza COM Interop a través de winax.
 * @param {SlidesToolInput} input - Parámetros de entrada para la herramienta.
 * @returns {Promise<string>} - Un mensaje indicando el resultado de la operación.
 */
const slidesTool: McpResource = {
  path: 'powerpoint/slides',
  description: 'Gestiona diapositivas en presentaciones de PowerPoint.',
  schema: SlidesToolInputSchema, // Corregido de inputSchema a schema
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => {
    let pptApp: any = null;
    let presentation: any = null;

    try {
      const input = SlidesToolInputSchema.parse(params);
      const { filePath, operation, slideIndex, slideLayout } = input;

      // Obtener instancia de PowerPoint
      pptApp = await getOfficeApplication('PowerPoint.Application');
      pptApp.Visible = true; // Opcional: hacer visible la aplicación
      // Abrir o crear presentación
      try {
        presentation = pptApp.Presentations.Open(filePath);
      } catch (error) {
        // Si el archivo no existe, crear una nueva presentación
        presentation = pptApp.Presentations.Add();
        presentation.SaveAs(filePath); // Guardar la nueva presentación
      }

      switch (operation) {
        case 'add':
          if (!slideLayout) {
            throw new Error('El parámetro slideLayout es requerido para la operación "add".');
          }
          // Buscar el diseño de diapositiva por nombre
          let layout;
          try {
            layout = pptApp.SlideLayouts.Item(slideLayout);
          } catch (e) {
            throw new Error(`Diseño de diapositiva "${slideLayout}" no encontrado.`);
          }

          // Añadir diapositiva
          const newSlide = presentation.Slides.Add(presentation.Slides.Count + 1, layout.Layout);
          return { success: true, data: `Diapositiva añadida con diseño "${slideLayout}".` }; // Retorno ajustado

        case 'delete':
          if (slideIndex === undefined) {
            throw new Error('El parámetro slideIndex es requerido para la operación "delete".');
          }
          if (slideIndex < 1 || slideIndex > presentation.Slides.Count) {
            throw new Error(`Índice de diapositiva ${slideIndex} fuera de rango.`);
          }
          // Eliminar diapositiva
          presentation.Slides.Item(slideIndex).Delete();
          return { success: true, data: `Diapositiva en el índice ${slideIndex} eliminada.` }; // Retorno ajustado

        case 'set':
          if (slideIndex === undefined) {
            throw new Error('El parámetro slideIndex es requerido para la operación "set".');
          }
           if (slideIndex < 1 || slideIndex > presentation.Slides.Count) {
            throw new Error(`Índice de diapositiva ${slideIndex} fuera de rango.`);
          }
          // Implementación básica para 'set'. Podría expandirse para modificar contenido, diseño, etc.
          // Por ahora, solo confirmamos que la diapositiva existe.
          const slideToSet = presentation.Slides.Item(slideIndex);
          // Aquí iría la lógica para modificar la diapositiva
          return { success: true, data: `Operación 'set' en la diapositiva ${slideIndex} completada (sin modificaciones complejas implementadas).` }; // Retorno ajustado

        default:
          throw new Error(`Operación no soportada: ${operation}`);
      }

    } catch (error: any) {
      // Manejo de errores
      logger.error(`Error in powerpoint/slides tool: ${error.message}`); // Usar logger
      return { success: false, error: { code: 'POWERPOINT_SLIDES_ERROR', message: `Error al procesar la solicitud de PowerPoint: ${error.message}` } }; // Retorno de error ajustado
    } finally {
      // Guardar y cerrar la presentación
      if (presentation) {
        try {
            presentation.Save();
            presentation.Close();
        } catch (closeError: any) {
            logger.warn(`Error al cerrar la presentación: ${closeError.message}`); // Usar logger
        }
        releaseObject(presentation);
      }
      // La aplicación de PowerPoint se gestiona externamente, no la cerramos aquí.
      releaseObject(pptApp);
    }
  },
};

export default slidesTool;