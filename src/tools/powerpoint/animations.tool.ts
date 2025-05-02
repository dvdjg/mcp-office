import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams } from '../../types/common.types';
import { getOfficeApplication } from '../../utils/officeInterop';

// Esquema de entrada para la herramienta powerpoint/animations
const AnimationsInputSchema = z.object({
  filePath: z.string().describe('Ruta al archivo de PowerPoint.'),
  operation: z.enum(['add', 'configure', 'remove', 'list']).describe('Operación a realizar: add (animación a forma), configure (transición a diapositiva), remove (animación/transición), list (animaciones/transiciones).'),
  slideIndex: z.number().int().positive().optional().describe('Índice de la diapositiva (1-basado). Requerido para configure, remove y list.'),
  shapeIndex: z.number().int().positive().optional().describe('Índice de la forma (1-basado). Requerido para add, remove y list (animaciones).'),
  shapeName: z.string().optional().describe('Nombre de la forma. Alternativa a shapeIndex para add, remove y list (animaciones).'),
  animationType: z.string().optional().describe('Tipo de animación (constante MsoAnimEffect). Requerido para add.'),
  transitionType: z.string().optional().describe('Tipo de transición (constante PpTransition). Requerido para configure.'),
  duration: z.number().positive().optional().describe('Duración en segundos. Opcional para add y configure.'),
  effectParameters: z.record(z.any()).optional().describe('Parámetros adicionales para la animación/transición (p. ej., dirección, orden).'),
});

type AnimationsInput = z.infer<typeof AnimationsInputSchema>;

/**
 * @tool
 * @description Gestiona animaciones de formas y transiciones de diapositivas en PowerPoint.
 * Permite añadir animaciones a formas, configurar transiciones de diapositivas,
 * eliminar animaciones/transiciones y listar las existentes.
 * Utiliza COM Interop a través de winax.
 * Requiere especificar la ruta del archivo, la operación y los índices/nombres
 * de diapositivas/formas según la operación.
 */
const animationsTool: McpResource = {
  path: 'powerpoint/animations',
  description: 'Gestiona animaciones de formas y transiciones de diapositivas en PowerPoint.',
  schema: AnimationsInputSchema,
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => {
    const input = AnimationsInputSchema.parse(params); // Validar y parsear la entrada
    const { filePath, operation, slideIndex, shapeIndex, shapeName, animationType, transitionType, duration, effectParameters } = input;
    const app = await getOfficeApplication('PowerPoint.Application');
    let presentation;

    try {
      presentation = app.Presentations.Open(filePath);

      switch (operation) {
        case 'add':
          if (!slideIndex || (!shapeIndex && !shapeName) || !animationType) {
            throw new Error('Para la operación "add", se requieren slideIndex, shapeIndex o shapeName, y animationType.');
          }
          const slideForAdd = presentation.Slides(slideIndex);
          const shapeForAdd = shapeIndex ? slideForAdd.Shapes(shapeIndex) : slideForAdd.Shapes(shapeName);
          if (!shapeForAdd) {
             throw new Error(`Forma no encontrada en la diapositiva ${slideIndex} con índice ${shapeIndex} o nombre ${shapeName}.`);
          }

          // Añadir animación
          // Nota: MsoAnimEffect es un enum en la API de PowerPoint. Necesitamos mapear el string a su valor numérico.
          // winax debería manejar esto si la constante existe en el objeto app.MsoAnimEffect
          const animEffectValue = app.MsoAnimEffect[animationType];
          if (animEffectValue === undefined) {
              throw new Error(`Tipo de animación inválido: ${animationType}.`);
          }

          const effect = slideForAdd.TimeLine.MainSequence.AddEffect(
            shapeForAdd,
            animEffectValue
          );

          if (duration !== undefined) {
            effect.Timing.Duration = duration;
          }

          // Aplicar parámetros adicionales si existen
          if (effectParameters) {
              // Esto es un ejemplo básico. La aplicación de parámetros depende del tipo de animación.
              // Se necesitaría lógica más compleja para manejar diferentes tipos de efectos y sus propiedades.
              // Por ejemplo, para una animación de entrada, podrías querer configurar la dirección.
              // effect.EffectParameters.Direction = app.MsoAnimDirection.msoAnimDirectionLeft;
              console.warn('La aplicación de effectParameters no está completamente implementada y puede requerir lógica específica por tipo de animación.');
          }


          return { success: true, data: { message: `Animación '${animationType}' añadida a la forma ${shapeIndex || shapeName} en la diapositiva ${slideIndex}.` } };

        case 'configure':
          if (!slideIndex || !transitionType) {
            throw new Error('Para la operación "configure", se requieren slideIndex y transitionType.');
          }
          const slideForConfig = presentation.Slides(slideIndex);
          const transition = slideForConfig.SlideShowTransition;

          // Configurar transición
          // Nota: PpTransition es un enum en la API de PowerPoint. Necesitamos mapear el string a su valor numérico.
          const transitionEffectValue = app.PpTransition[transitionType];
           if (transitionEffectValue === undefined) {
              throw new Error(`Tipo de transición inválido: ${transitionType}.`);
          }
          transition.EntryEffect = transitionEffectValue;

          if (duration !== undefined) {
            transition.Duration = duration;
          }

           // Aplicar parámetros adicionales si existen
          if (effectParameters) {
              // Similar a las animaciones, la aplicación de parámetros de transición depende del tipo.
              // Por ejemplo, para una transición de empuje, podrías querer configurar la dirección.
              // transition.Direction = app.PpTransitionDirection.ppTransitionDirectionLeft;
               console.warn('La aplicación de effectParameters para transiciones no está completamente implementada y puede requerir lógica específica por tipo de transición.');
          }

          return { success: true, data: { message: `Transición '${transitionType}' configurada para la diapositiva ${slideIndex}.` } };

        case 'remove':
             if (!slideIndex || (!shapeIndex && !shapeName)) {
                 throw new Error('Para la operación "remove", se requieren slideIndex y shapeIndex o shapeName.');
             }
             const slideForRemove = presentation.Slides(slideIndex);
             const shapeForRemove = shapeIndex ? slideForRemove.Shapes(shapeIndex) : slideForRemove.Shapes(shapeName);
             if (!shapeForRemove) {
                 throw new Error(`Forma no encontrada en la diapositiva ${slideIndex} con índice ${shapeIndex} o nombre ${shapeName}.`);
             }

             // Eliminar animaciones asociadas a la forma
             const effectsToRemove = [];
             // Iterar hacia atrás para evitar problemas con índices después de la eliminación
             for (let i = slideForRemove.TimeLine.MainSequence.Count; i >= 1; i--) {
                 const effect = slideForRemove.TimeLine.MainSequence(i);
                 // Comparar formas por su COM object o un identificador único si es posible.
                 // Comparar por nombre puede ser problemático si hay formas con nombres duplicados.
                 // Comparar por el objeto COM directamente es más fiable si winax lo permite.
                 // Si winax no permite la comparación directa de objetos COM, podríamos intentar comparar propiedades únicas como ID o Name (si garantizamos nombres únicos).
                 // Por ahora, asumimos que la comparación directa de objetos COM funciona o que los nombres son únicos para este caso.
                 try {
                     if (effect.Shape && effect.Shape.Name === shapeForRemove.Name) { // Comparación por nombre como fallback/ejemplo
                          effectsToRemove.push(effect);
                     }
                 } catch (compareError) {
                      console.warn(`Error comparando formas durante la eliminación de animación: ${compareError instanceof Error ? compareError.message : String(compareError)}`);
                      // Continuar con la eliminación incluso si la comparación falla para un efecto específico
                 }
             }

             effectsToRemove.forEach(effect => {
                 try {
                     effect.Delete();
                 } catch (deleteError) {
                     console.error(`Error eliminando efecto de animación: ${deleteError instanceof Error ? deleteError.message : String(deleteError)}`);
                     // Continuar eliminando otros efectos
                 }
             });


             return { success: true, data: { message: `Animaciones eliminadas para la forma ${shapeIndex || shapeName} en la diapositiva ${slideIndex}.` } };

        case 'list':
            if (!slideIndex) {
                throw new Error('Para la operación "list", se requiere slideIndex.');
            }
            const slideForList = presentation.Slides(slideIndex);
            const animations = [];

            // Listar animaciones de formas en la diapositiva
            for (let i = 1; i <= slideForList.TimeLine.MainSequence.Count; i++) {
                const effect = slideForList.TimeLine.MainSequence(i);
                animations.push({
                    shapeName: effect.Shape.Name,
                    // shapeIndex: effect.Shape.ZOrderPosition, // ZOrderPosition no es un índice fiable en la colección Shapes
                    animationType: effect.EffectType, // Esto devuelve un valor numérico, necesitaríamos mapearlo a la constante MsoAnimEffect
                    duration: effect.Timing.Duration,
                    // Otros detalles de la animación si son relevantes
                });
            }

            // Obtener detalles de la transición de la diapositiva
            const transitionForList = slideForList.SlideShowTransition;
            const transitionDetails = {
                transitionType: transitionForList.EntryEffect, // Esto devuelve un valor numérico, necesitaríamos mapearlo a la constante PpTransition
                duration: transitionForList.Duration,
                // Otros detalles de la transición si son relevantes
            };


            return { success: true, data: { animations, transition: transitionDetails, message: `Listado de animaciones y transición para la diapositiva ${slideIndex}.` } };


        default:
          throw new Error(`Operación no soportada: ${operation}`);
      }
    } catch (error: any) {
      console.error(`Error en la herramienta powerpoint/animations: ${error.message}`);
      return {
        success: false,
        error: {
          code: 'OFFICE_API_ERROR', // O un código de error más específico
          message: error.message,
          details: error, // Incluir el objeto de error original para depuración
        },
      };
    } finally {
      if (presentation) {
        try {
            presentation.Save();
            presentation.Close();
        } catch (saveCloseError: any) { // Tipar como any
             console.error(`Error al guardar/cerrar la presentación: ${saveCloseError.message}`);
             // Continuar para no bloquear el finally
        }
      }
      // No cerrar la aplicación de PowerPoint aquí, ya que otras herramientas pueden necesitarla.
      // app.Quit(); // ¡No hacer esto!
    }
  },
};

export default animationsTool;