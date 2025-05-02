import { z } from 'zod';
import { McpResource, ApiResponse, FastMCPContext as Context, ToolRequestParams } from '../../types/common.types';
import { getOfficeApplication } from '../../utils/officeInterop';
import { handleToolError, createErrorResponse } from '../../utils/errorHandler';

// Esquema de entrada para la herramienta powerpoint/properties
const PowerPointPropertiesInputSchema = z.object({
  filePath: z.string().describe('La ruta al archivo de presentación de PowerPoint.'),
  operation: z.enum(['set', 'configure', 'add']).describe('La operación a realizar: set, configure, o add.'),
  // Propiedades comunes para 'set' y 'add'
  propertyName: z.string().optional().describe('El nombre de la propiedad a establecer o añadir.'),
  propertyValue: z.any().optional().describe('El valor de la propiedad a establecer o añadir.'),
  // Propiedades específicas para 'configure'
  size: z.enum(['16:9', '4:3']).optional().describe('El tamaño de la diapositiva (e.g., "16:9", "4:3").'),
  // Puedes añadir más parámetros de configuración aquí según sea necesario
});

type PowerPointPropertiesInput = z.infer<typeof PowerPointPropertiesInputSchema>;

/**
 * @tool powerpoint/properties
 * @description Gestiona las propiedades de una presentación de PowerPoint.
 * Permite establecer propiedades integradas, configurar ajustes como el tamaño de la diapositiva,
 * y añadir o modificar propiedades personalizadas.
 * Utiliza COM Interop a través de winax para interactuar con PowerPoint.
 * @param {ToolRequestParams} params - Los parámetros de entrada para la herramienta.
 * @param {Context<any>} [context] - El contexto de FastMCP.
 * @returns {Promise<ApiResponse<any>>} Un objeto ApiResponse indicando el resultado.
 */
const handler = async (params: ToolRequestParams, context?: Context<any>): Promise<ApiResponse<any>> => {
  try {
    // Validar los parámetros de entrada
    const input = PowerPointPropertiesInputSchema.parse(params);
    const { filePath, operation, propertyName, propertyValue, size } = input;

    let app;

    try {
      app = await getOfficeApplication('PowerPoint.Application');
      const presentation = app.Presentations.Open(filePath);

      switch (operation) {
        case 'set':
          if (!propertyName) {
            return createErrorResponse('VALIDATION_ERROR', 'propertyName es requerido para la operación set.');
          }
          // Implementación para establecer propiedades integradas
          // Esto requerirá mapear propertyName a la propiedad COM correcta
          // Ejemplo básico (puede necesitar refinamiento):
          if (propertyName === 'title') {
             presentation.BuiltinDocumentProperties('Title').Value = propertyValue;
          } else if (propertyName === 'author') {
             presentation.BuiltinDocumentProperties('Author').Value = propertyValue;
          }
          // Añadir más casos según las propiedades integradas comunes
          break;

        case 'configure':
          // Implementación para configurar ajustes de la presentación
          if (size) {
            // Configurar tamaño de diapositiva
            // ppSlideSizeOnScreen (16:9) = 1, ppSlideSizeOnScreen (4:3) = 2
            if (size === '16:9') {
              presentation.PageSetup.SlideSize = 1;
            } else if (size === '4:3') {
              presentation.PageSetup.SlideSize = 2;
            }
          }
          // Añadir más opciones de configuración aquí
          break;

        case 'add':
          if (!propertyName || propertyValue === undefined) {
            return createErrorResponse('VALIDATION_ERROR', 'propertyName y propertyValue son requeridos para la operación add.');
          }
          // Implementación para añadir o modificar propiedades personalizadas
          // Verificar si la propiedad ya existe antes de añadir
          let customProp;
          try {
              customProp = presentation.CustomDocumentProperties(propertyName);
              customProp.Value = propertyValue;
          } catch (e) {
              // Propiedad no existe, añadirla
              presentation.CustomDocumentProperties.Add(propertyName, false, 4, propertyValue); // msoPropertyTypeVariant = 4
          }
          break;

        default:
          return createErrorResponse('VALIDATION_ERROR', `Operación no soportada: ${operation}`);
      }

      presentation.Save();
      presentation.Close();

      return { success: true, data: `Operación '${operation}' completada para el archivo '${filePath}'.` };

    } catch (officeError: any) {
      // Manejo de errores específicos de Office/COM
      if (app) {
          try {
              // Intentar cerrar la presentación si está abierta para evitar bloqueos
              const presentation = app.Presentations.Open(filePath);
              presentation.Close();
          } catch (closeError) {
              // Ignorar errores al cerrar si la presentación ya estaba cerrada o no se pudo abrir
          }
      }
      return handleToolError(officeError, 'POWERPOINT_PROPERTIES_ERROR');
    }

  } catch (validationError: any) {
    // Manejo de errores de validación del esquema Zod
    return handleToolError(validationError, 'VALIDATION_ERROR');
  }
};

export const powerpointPropertiesTool: McpResource[] = [{
  path: 'powerpoint/properties',
  description: 'Manages the properties of a PowerPoint presentation.',
  schema: PowerPointPropertiesInputSchema,
  handler,
}];