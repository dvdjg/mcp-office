import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams } from '../../types/common.types'; // Importar McpResource y tipos necesarios
import path from 'path';
import fs from 'fs/promises';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop'; // Importar funciones específicas de officeInterop
import pdfParseTool from './pdfParse.tool'; // Importar la herramienta pdfParse por defecto
import logger from '../../utils/logger'; // Importar logger

// Esquema de entrada para la herramienta office/combine
const CombineToolInputSchema = z.object({
  directoryPath: z.string().describe('Ruta al directorio con archivos de origen.'),
  outputFilePath: z.string().describe('Ruta donde guardar el documento Word combinado.'),
  operation: z.enum(['read', 'normalize', 'insert']).describe('Operación a realizar: read, normalize, o insert.'),
  filePatterns: z.array(z.string()).optional().describe('Patrones glob para filtrar archivos.'),
  formatsToInclude: z.array(z.string()).optional().describe('Formatos de archivo a incluir (e.g., ["pdf", "docx"]).'),
});

type CombineToolInput = z.infer<typeof CombineToolInputSchema>;

// Documentación JSDoc concisa
/**
 * @tool office/combine
 * @description Combina múltiples archivos de un directorio en un único documento Word.
 * @operations read, normalize, insert
 * @input CombineToolInputSchema
 * @output string (mensaje de estado o ruta del archivo de salida)
 */
export const officeCombineTool: McpResource = { // Usar McpResource
  path: 'office/combine', // Añadir la propiedad path
  description: 'Combina múltiples archivos de un directorio en un único documento Word.',
  schema: CombineToolInputSchema, // Cambiar inputSchema a schema
  handler: async (params: ToolRequestParams, context?: any): Promise<ApiResponse<string>> => { // Ajustar la firma del handler y tipo de retorno
    const input = params as CombineToolInput; // Castear params al tipo esperado
    const { directoryPath, outputFilePath, operation, filePatterns, formatsToInclude } = input;

    try {
      switch (operation) {
        case 'read': {
          logger.info(`[office/combine] Ejecutando operación 'read' en ${directoryPath}`);
          const files = await fs.readdir(directoryPath);
          let filteredFiles = files;

          if (filePatterns) {
            // TODO: Implementar filtrado por patrones glob si es necesario
            logger.warn('[office/combine] Filtrado por filePatterns no implementado aún.');
          }

          if (formatsToInclude) {
            filteredFiles = filteredFiles.filter(file => {
              const ext = path.extname(file).toLowerCase().replace('.', '');
              return formatsToInclude.includes(ext);
            });
          }

          logger.info(`[office/combine] Encontrados ${filteredFiles.length} archivos.`);
          return { success: true, data: `Archivos encontrados en ${directoryPath}: ${filteredFiles.join(', ')}` };
        }

        case 'normalize': {
          logger.info(`[office/combine] Ejecutando operación 'normalize' en ${directoryPath}`);
          const filesToNormalize = await fs.readdir(directoryPath);
          let filteredFilesToNormalize = filesToNormalize;

          if (filePatterns) {
             logger.warn('[office/combine] Filtrado por filePatterns no implementado aún.');
          }

          if (formatsToInclude) {
            filteredFilesToNormalize = filteredFilesToNormalize.filter(file => {
              const ext = path.extname(file).toLowerCase().replace('.', '');
              return formatsToInclude.includes(ext);
            });
          }

          const normalizedContent: { fileName: string, content: string }[] = [];

          for (const file of filteredFilesToNormalize) {
            const fullPath = path.join(directoryPath, file);
            const ext = path.extname(file).toLowerCase().replace('.', '');

            try {
              if (ext === 'pdf') {
                logger.info(`[office/combine] Normalizando PDF: ${file}`);
                // Usar el handler de pdfParseTool
                const pdfParseResult = await pdfParseTool.handler({ filePath: fullPath, operation: 'parse' });
                if (pdfParseResult.success) {
                  normalizedContent.push({ fileName: file, content: pdfParseResult.data as string });
                } else {
                  logger.error(`[office/combine] Error al parsear PDF ${file}: ${pdfParseResult.error?.message}`);
                  // Decidir si lanzar error o continuar con el siguiente archivo
                }
              } else if (ext === 'docx' || ext === 'doc') {
                 logger.info(`[office/combine] Normalizando DOCX/DOC: ${file}`);
                 // TODO: Implementar extracción de texto y posiblemente otros elementos de DOCX/DOC usando COM Interop
                 // Esto es complejo y requiere interactuar con la API de Word a través de winax.
                 // Por ahora, solo añadiremos un placeholder.
                 normalizedContent.push({ fileName: file, content: `[Contenido de ${file} - Normalización DOCX/DOC no implementada completamente]` });
              } else {
                 logger.warn(`[office/combine] Formato no soportado para normalización: ${file}`);
                 // Opcional: leer como texto plano si es un archivo de texto conocido
                 // try {
                 //    const textContent = await fs.readFile(fullPath, 'utf-8');
                 //    normalizedContent.push({ fileName: file, content: textContent });
                 // } catch (readError) {
                 //    logger.warn(`[office/combine] No se pudo leer ${file} como texto plano.`);
                 // }
              }
            } catch (fileError: any) {
               logger.error(`[office/combine] Error al procesar archivo ${file}: ${fileError.message}`);
               // Continuar con el siguiente archivo
            }
          }

          // Aquí podrías guardar el contenido normalizado temporalmente si la operación 'insert' se llama por separado
          // Por ahora, solo devolvemos un resumen.
          logger.info(`[office/combine] Normalización completada para ${normalizedContent.length} archivos.`);
          return { success: true, data: `Normalización completada. Contenido normalizado para: ${normalizedContent.map(item => item.fileName).join(', ')}` };
        }

        case 'insert': {
          logger.info(`[office/combine] Ejecutando operación 'insert' en ${outputFilePath}`);
          // TODO: Implementar lógica de inserción usando COM Interop (winax)
          // Esto implicaría:
          // 1. Crear o abrir un documento Word en outputFilePath.
          // 2. Iterar sobre el contenido normalizado (necesitaríamos una forma de pasarlo o normalizarlo aquí de nuevo).
          // 3. Insertar el contenido de cada archivo en el documento Word.
          // 4. Guardar y cerrar el documento Word.

          // Para la inserción, probablemente necesitemos normalizar los archivos de nuevo o esperar que el contenido normalizado se pase como parámetro.
          // Por simplicidad inicial, normalizaremos aquí de nuevo (menos eficiente pero más simple para empezar).
           const filesToInsert = await fs.readdir(directoryPath);
           let filteredFilesToInsert = filesToInsert;

          if (filePatterns) {
             logger.warn('[office/combine] Filtrado por filePatterns no implementado aún.');
          }

          if (formatsToInclude) {
            filteredFilesToInsert = filteredFilesToInsert.filter(file => {
              const ext = path.extname(file).toLowerCase().replace('.', '');
              return formatsToInclude.includes(ext);
            });
          }

          let wordApp: any = null;
          let doc: any = null;

          try {
            wordApp = await getOfficeApplication('Word.Application');
            // Crear un nuevo documento o abrir uno existente? El plan dice "combinarlos en un único documento Word",
            // lo que sugiere crear uno nuevo si outputFilePath no existe, o añadir a uno existente.
            // Por ahora, crearemos uno nuevo.
            logger.debug('[office/combine] Creando nuevo documento Word.');
            doc = wordApp.Documents.Add();

            for (const file of filteredFilesToInsert) {
              const fullPath = path.join(directoryPath, file);
              const ext = path.extname(file).toLowerCase().replace('.', '');
              logger.info(`[office/combine] Insertando contenido de: ${file}`);

              // Mover el cursor al final del documento antes de insertar
              doc.Content.Collapse(0); // wdCollapseEnd = 0
              const endRange = doc.Content;
              endRange.Collapse(0); // Asegurarse de que el rango está al final

              try {
                 if (ext === 'pdf') {
                    logger.info(`[office/combine] Insertando contenido de PDF: ${file}`);
                    // Para PDFs, podríamos intentar insertar el texto parseado o insertar el PDF como objeto.
                    // Insertar como objeto puede mantener formato pero depende de los visores instalados.
                    // Insertar texto parseado pierde formato.
                    // Opción 1: Insertar texto parseado (requiere normalización previa o aquí)
                    const pdfParseResult = await pdfParseTool.handler({ filePath: fullPath, operation: 'parse' });
                    if (pdfParseResult.success) {
                       endRange.InsertAfter(pdfParseResult.data as string + '\n\n'); // Añadir saltos de línea entre archivos
                    } else {
                       logger.error(`[office/combine] Error al parsear PDF para inserción ${file}: ${pdfParseResult.error?.message}`);
                    }

                    // Opción 2: Insertar el archivo como objeto (mantiene formato pero requiere visor)
                    // logger.debug(`[office/combine] Intentando insertar PDF como objeto: ${file}`);
                    // // wdInsertObject = 0, wdFloatOverText = 0 (para InlineShape)
                    // endRange.InlineShapes.AddOLEObject(fullPath, undefined, false, false, undefined, undefined, undefined, 0);
                    // endRange.InsertParagraphAfter(); // Añadir un salto de párrafo después del objeto
                    // logger.debug(`[office/combine] PDF insertado como objeto: ${file}`);

                 } else if (ext === 'docx' || ext === 'doc') {
                    logger.info(`[office/combine] Insertando contenido de DOCX/DOC: ${file}`);
                    // Insertar el contenido de otro documento Word
                    // wdInsertFile = 4
                    endRange.InsertFile(fullPath);
                    endRange.InsertParagraphAfter(); // Añadir un salto de párrafo después del contenido insertado
                    logger.debug(`[office/combine] Contenido de DOCX/DOC insertado: ${file}`);

                 } else if (ext === 'txt') {
                     logger.info(`[office/combine] Insertando contenido de TXT: ${file}`);
                     const textContent = await fs.readFile(fullPath, 'utf-8');
                     endRange.InsertAfter(textContent + '\n\n'); // Insertar texto plano
                     logger.debug(`[office/combine] Contenido de TXT insertado: ${file}`);
                 }
                 // TODO: Añadir manejo para otros formatos si es necesario (e.g., .xlsx, .pptx - puede ser complejo)

              } catch (fileInsertError: any) {
                 logger.error(`[office/combine] Error al insertar contenido de ${file}: ${fileInsertError.message}`);
                 // Continuar con el siguiente archivo
              }
            }

            // Guardar el documento final
            logger.debug(`[office/combine] Guardando documento final en: ${outputFilePath}`);
            const absoluteOutputFilePath = path.resolve(outputFilePath);
            // wdFormatDocumentDefault = 16 (para .docx)
            doc.SaveAs2(absoluteOutputFilePath, 16); // Guardar como .docx
            logger.info(`[office/combine] Documento combinado guardado en: ${absoluteOutputFilePath}`);

            return { success: true, data: `Documento combinado guardado en: ${absoluteOutputFilePath}` };

          } catch (insertError: any) {
            logger.error(`[office/combine] Error durante la operación 'insert': ${insertError.message}`, { error: insertError });
            return { success: false, error: { code: 'INSERT_ERROR', message: `Error durante la inserción: ${insertError.message}` } };
          } finally {
            // Limpiar objetos COM
            if (doc) {
              try {
                doc.Close(0); // wdDoNotSaveChanges = 0
                logger.debug('[office/combine] Documento Word cerrado.');
              } catch (closeError: any) {
                logger.warn(`[office/combine] Error al cerrar documento Word: ${closeError.message}`);
              }
              releaseObject(doc);
            }
            if (wordApp) {
              // Considerar si cerrar la aplicación Word o dejarla abierta.
              // Si la abrimos nosotros, probablemente deberíamos cerrarla si no hay otros documentos abiertos.
              // Esto es complicado de determinar de forma segura. Por ahora, la dejaremos abierta.
              // try {
              //   if (wordApp.Documents.Count === 0) {
              //     wordApp.Quit();
              //     logger.debug('[office/combine] Aplicación Word cerrada.');
              //   }
              // } catch (quitError: any) {
              //   logger.warn(`[office/combine] Error al cerrar aplicación Word: ${quitError.message}`);
              // }
              releaseObject(wordApp);
            }
          }
        }

        default:
          // Esto no debería ocurrir si el esquema Zod funciona correctamente
          return { success: false, error: { code: 'INVALID_OPERATION', message: `Operación no soportada: ${operation}` } };
      }
    } catch (error: any) {
      // Manejo básico de errores para errores inesperados fuera de los casos de operación
      logger.error(`[office/combine] Error inesperado: ${error.message}`, { error });
      return { success: false, error: { code: 'UNEXPECTED_ERROR', message: `Error inesperado al ejecutar la herramienta office/combine: ${error.message}` } };
    }
  },
};