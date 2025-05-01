import * as winax from 'winax';
import logger from './logger'; // Asegúrate de que la ruta al logger sea correcta

type OfficeAppName = 'Word.Application' | 'Excel.Application' | 'PowerPoint.Application';

/**
 * Obtiene una instancia de una aplicación de Office (existente o nueva).
 * @param appName El nombre ProgID de la aplicación de Office (ej. 'Word.Application').
 * @returns Una promesa que resuelve con el objeto COM de la aplicación.
 */
export async function getOfficeApplication(appName: OfficeAppName): Promise<any> {
  logger.info(`[OfficeInterop] Intentando obtener/crear instancia de ${appName}...`);
  let app: any = null;

  try {
    // Winax intenta conectar a una instancia existente o crear una nueva con new ActiveXObject
    // No hay un método separado como GetActiveObject en win32ole que sea estándar en winax.
    // Crear un nuevo objeto a menudo adjunta a uno existente si está disponible.
    app = new winax.Object(appName, { activate: true }); // { activate: true } intenta traerla al frente si existe

    if (!app) {
      throw new Error(`No se pudo crear ni conectar a ${appName}.`);
    }

    logger.info(`[OfficeInterop] Instancia de ${appName} obtenida/creada.`);

    // Hacer visible la aplicación para depuración
    try {
      if (app.Visible === false || app.Visible === 0) {
         // Para PowerPoint, la ventana principal podría no ser directamente 'Visible'
         // Necesitamos verificar si la aplicación tiene una ventana principal y hacerla visible.
         if (appName === 'PowerPoint.Application') {
            // PowerPoint puede no tener ventanas si se inicia sin interfaz gráfica.
            // Si hay presentaciones, la ventana de la aplicación podría ser visible.
            // O podríamos necesitar crear una presentación para forzar la visibilidad.
            // Por simplicidad inicial, intentaremos establecer Visible, pero puede fallar.
            try {
                 app.Visible = true;
            } catch (visError) {
                 logger.warn(`[OfficeInterop] No se pudo establecer Visible=true directamente para ${appName}. Puede requerir abrir/crear un archivo.`);
                 // Podríamos intentar crear una ventana si es necesario: app.NewWindow();
            }
         } else {
             app.Visible = true;
         }
         logger.info(`[OfficeInterop] ${appName} establecida como visible.`);
      }
    } catch (visError) {
      logger.warn(`[OfficeInterop] No se pudo establecer la propiedad Visible para ${appName}. Error: ${visError instanceof Error ? visError.message : String(visError)}`);
      // Continuar incluso si no se puede hacer visible
    }

    return app;
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    logger.error(`[OfficeInterop] Error al obtener/crear ${appName}: ${errorMessage}`, { error });
    // Intentar liberar el objeto si se creó parcialmente y falló después
    if (app) {
      releaseObject(app);
    }
    throw new Error(`Fallo al obtener la aplicación ${appName}: ${errorMessage}`);
  }
}

/**
 * Libera un objeto COM.
 * @param comObject El objeto COM a liberar.
 */
export function releaseObject(comObject: any): void {
  if (!comObject) {
    return;
  }
  try {
    // winax puede usar __release o simplemente dejar que el GC lo maneje.
    // Llamar a __release si existe es más seguro para liberar recursos inmediatamente.
    if (typeof comObject.__release === 'function') {
      comObject.__release();
      logger.info('[OfficeInterop] Objeto COM liberado usando __release().');
    } else {
      logger.info('[OfficeInterop] El objeto COM no tiene método __release(). Confiando en GC.');
      // Alternativamente, simplemente asignar a null ayuda al GC
      // comObject = null;
    }
  } catch (error) {
    logger.warn(`[OfficeInterop] Advertencia al liberar objeto COM: ${error instanceof Error ? error.message : String(error)}`);
    // No relanzar el error, solo registrar la advertencia.
  }
}

// Opcional: Añadir un pequeño test aquí si es necesario, pero mejor en index.ts como se sugirió.
// async function test() {
//   try {
//     const word = await getOfficeApplication('Word.Application');
//     console.log('Word Version:', word.Version);
//     releaseObject(word);
//   } catch(e) {
//     console.error(e);
//   }
// }
// test();