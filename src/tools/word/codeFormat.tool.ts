import { z } from 'zod';
import * as winax from 'winax';
import hljs from 'highlight.js';
// Importar highlight.js y guesslang si están disponibles
// import { GuessLang } from 'guesslang';

// const guesslang = new GuessLang();

const codeFormatInputSchema = z.object({
  filePath: z.string().describe('La ruta al archivo Word.'),
  operation: z.enum(['identify', 'detect', 'apply']).describe('La operación a realizar: identify (identificar bloques de código), detect (detectar lenguaje), apply (aplicar formato).'),
  range: z.object({
    start: z.number().describe('El índice de inicio del rango de texto.'),
    end: z.number().describe('El índice de fin del rango de texto.'),
  }).optional().describe('El rango de texto a procesar para las operaciones detect y apply.'),
  language: z.string().optional().describe('El lenguaje de programación si se conoce (para la operación apply).'),
  style: z.string().optional().describe('El nombre del estilo de Word a buscar (para identify) o aplicar (para apply).'),
  code: z.string().optional().describe('El bloque de código a formatear (para la operación apply).'),
});

type CodeFormatInput = z.infer<typeof codeFormatInputSchema>;

// Definir la interfaz para el contexto de la herramienta si es necesario, o usar 'any' si no se accede a propiedades específicas del contexto de FastMCP
// interface ToolContext {
//   // Define las propiedades del contexto que necesitas, por ejemplo:
//   // log: { info: (msg: string) => void };
// }

const handler = async (input: CodeFormatInput, context: any) => {
  const { filePath, operation, range, language, style, code } = input;

  let wordApp;
  let doc;

  try {
    wordApp = new winax.Object('Word.Application', { activate: true });
    doc = wordApp.Documents.Open(filePath);

    let result: any;

    switch (operation) {
      case 'identify':
        result = identifyCodeBlocks(doc, style);
        break;
      case 'detect':
        if (!range) {
          throw new Error('El rango es requerido para la operación detect.');
        }
        result = await detectLanguage(doc, range);
        break;
      case 'apply':
        if (!range || !code || !language) {
           // Permitir aplicar formato sin detectar si se proporciona lenguaje y código
           if (!range || !code || !language) {
             throw new Error('El rango, el código y el lenguaje son requeridos para la operación apply.');
           }
        }
        result = applyCodeFormatting(doc, range, code, language);
        break;
      default:
        throw new Error(`Operación no soportada: ${operation}`);
    }

    doc.Save();
    doc.Close();
    wordApp.Quit();

    return { success: true, result };

  } catch (error: any) {
    if (doc) doc.Close();
    if (wordApp) wordApp.Quit();
    return { success: false, error: error.message };
  }
};

/**
 * @tool_description Identifica bloques de código en un documento Word.
 * Busca párrafos con un estilo específico o patrones de texto para identificar bloques de código.
 * Nota: La detección basada en patrones de texto es una heurística básica y puede no ser precisa.
 * @param doc El objeto del documento Word.
 * @param style Opcional. El nombre del estilo de párrafo a buscar.
 * @returns Una lista de rangos identificados como bloques de código.
 */
function identifyCodeBlocks(doc: any, style?: string): { start: number, end: number }[] {
  const codeBlocks: { start: number, end: number }[] = [];
  const paragraphs = doc.Paragraphs;

  for (let i = 1; i <= paragraphs.Count; i++) {
    const paragraph = paragraphs.Item(i);
    let isCode = false;

    if (style) {
      // Verificar si el estilo del párrafo coincide
      try {
        if (paragraph.Style.NameLocal === style || paragraph.Style.NamePrimary === style) {
          isCode = true;
        }
      } catch (e) {
        // Ignorar errores si el estilo no se puede acceder
      }
    } else {
      // Implementar lógica de detección basada en patrones de texto si no se especifica estilo
      // Por ahora, solo un ejemplo básico: buscar líneas que empiecen con espacios o tabulaciones
      const text = paragraph.Range.Text.trim();
      if (text.startsWith(' ') || text.startsWith('\t')) {
         // Esto es una heurística muy básica, mejorar según sea necesario
         isCode = true;
      }
    }

    if (isCode) {
      codeBlocks.push({
        start: paragraph.Range.Start,
        end: paragraph.Range.End -1 // -1 para no incluir el marcador de párrafo
      });
    }
  }

  return codeBlocks;
}

/**
 * @tool_description Detecta el lenguaje de programación de un bloque de texto.
 * Utiliza highlight.js. Si no se especifica un lenguaje, intenta la detección automática.
 * Nota: La detección automática puede no ser siempre precisa.
 * @param doc El objeto del documento Word.
 * @param range El rango de texto que contiene el código.
 * @returns El lenguaje detectado por highlight.js.
 */
async function detectLanguage(doc: any, range: { start: number, end: number }): Promise<string> {
  const textRange = doc.Range(range.start, range.end);
  const codeText = textRange.Text;

  // if (guesslang) {
  //   try {
  //     const lang = await guesslang.detectLang(codeText);
  //     return lang;
  //   } catch (e) {
  //     console.error("Error al detectar lenguaje con guesslang:", e);
  //   }
  // }

  // Heurística simple si guesslang no está disponible o falla
  if (codeText.includes('function') || codeText.includes('const') || codeText.includes('let')) return 'javascript';
  if (codeText.includes('def ') || codeText.includes('import ') || codeText.includes('print(')) return 'python';
  if (codeText.includes('<html') || codeText.includes('<div')) return 'html';
  if (codeText.includes('{') && codeText.includes('}')) return 'css';


  return 'plaintext'; // Lenguaje por defecto si no se detecta nada
}

/**
 * @tool_description Aplica formato de resaltado de sintaxis a un bloque de código en Word.
 * Utiliza highlight.js para obtener el formato (basado en clases CSS) y aplica las propiedades de fuente (color, negrita, cursiva) en el rango de texto de Word.
 * Nota: La aplicación de formato se basa en un parser simple de HTML y un mapeo básico de clases CSS a estilos de Word. Puede no ser perfecta para todos los casos.
 * @param doc El objeto del documento Word.
 * @param range El rango de texto que contiene el código.
 * @param code El bloque de código a formatear.
 * @param language El lenguaje de programación del código (opcional, highlight.js intentará detectar si no se proporciona).
 * @returns Un objeto indicando el éxito de la operación.
 */
// Función auxiliar para aplicar formato basado en HTML resaltado
function applyFormattingFromHtml(textRange: any, highlightedHtml: string) {
    // Limpiar formato existente en el rango
    textRange.Font.Reset();
    textRange.Font.Name = 'Consolas';
    textRange.Font.Size = 10;

    // Simple parser de HTML para extraer texto y clases
    let currentTextIndex = 0;
    // Regex para encontrar <span> con clase y contenido, o texto fuera de <span>
    const regex = /<span class="([^"]+)">([^<]+)<\/span>|([^<]+)/g;
    let match;

    // Obtener el texto plano del rango de Word para un mapeo de índices más preciso
    const wordPlainText = textRange.Text.replace(/\r\n/g, '\n').replace(/\r/g, '\n'); // Normalizar saltos de línea

    // Iterar sobre el HTML resaltado
    while ((match = regex.exec(highlightedHtml)) !== null) {
        let text = '';
        let classes = '';

        if (match[1] && match[2]) { // Coincide con <span class="...">...</span>
            classes = match[1];
            text = match[2];
        } else if (match[3]) { // Coincide con texto fuera de span
            text = match[3];
        }

        if (text) {
            // Encontrar la posición de este texto en el texto plano de Word
            // Esto sigue siendo una simplificación. Un mapeo de caracteres sería más robusto.
            // Asumimos que el texto en el HTML aparece en el mismo orden en el texto plano de Word.
            const startIndexInWord = wordPlainText.indexOf(text, currentTextIndex);

            if (startIndexInWord !== -1) {
                 const subRange = textRange.Duplicate;
                 subRange.Start = textRange.Start + startIndexInWord;
                 subRange.End = subRange.Start + text.length;

                 // Aplicar formato basado en clases
                 const font = subRange.Font;
                 font.Bold = false;
                 font.Italic = false;
                 font.Color = 0x000000; // Color por defecto (negro)

                 const classList = classes.split(' ');
                 for (const cls of classList) {
                     switch (cls) {
                         case 'hljs-keyword':
                             font.Color = 0xFF0000; // Azul
                             font.Bold = true;
                             break;
                         case 'hljs-built_in':
                             font.Color = 0xFFFF00; // Cian
                             break;
                         case 'hljs-literal':
                             font.Color = 0x00A5FF; // Naranja
                             break;
                         case 'hljs-number':
                             font.Color = 0x0000FF; // Rojo
                             break;
                         case 'hljs-string':
                             font.Color = 0x008000; // Verde
                             break;
                         case 'hljs-comment':
                             font.Color = 0x808080; // Gris
                             font.Italic = true;
                             break;
                         case 'hljs-variable':
                             // Color por defecto (negro)
                             break;
                         case 'hljs-title':
                             font.Color = 0x800080; // Púrpura
                             font.Bold = true;
                             break;
                         case 'hljs-params':
                             font.Italic = true;
                             break;
                         case 'hljs-operator':
                             // Color por defecto (negro)
                             break;
                         case 'hljs-punctuation':
                             // Color por defecto (negro)
                             break;
                         // Añadir más casos según sea necesario para otras clases de highlight.js
                     }
                 }

                 currentTextIndex = startIndexInWord + text.length; // Actualizar el índice para la próxima búsqueda
            } else {
                // Si no se encuentra el texto, esto indica un problema con el mapeo o el parser.
                // Podríamos loggear una advertencia o lanzar un error.
                console.warn(`Texto "${text}" del HTML no encontrado en el rango de Word a partir del índice ${currentTextIndex}.`);
                // Intentar avanzar el índice basado en la longitud del texto en el HTML de todas formas,
                // aunque el formato no se aplique correctamente a este segmento.
                 currentTextIndex += text.length;
            }
        }
    }
}

function applyCodeFormatting(doc: any, range: { start: number, end: number }, code: string, language: string): { success: boolean } {
    const textRange = doc.Range(range.start, range.end);

    // Integrar highlight.js
    // Usar highlightAuto si el lenguaje no está especificado o es 'plaintext'
    const highlightedResult = language && language !== 'plaintext'
        ? hljs.highlight(code, { language: language })
        : hljs.highlightAuto(code);

    const highlightedHtml = highlightedResult.value;

    // Obtener el texto plano del HTML resaltado para reemplazar el contenido en Word
    const plainText = highlightedHtml.replace(/<[^>]*>/g, '');

    // Reemplazar el texto en el rango de Word con el texto plano
    // Esto es crucial para que los índices del HTML coincidan con el texto en Word.
    // Advertencia: Esto elimina cualquier formato preexistente en el rango.
    textRange.Text = plainText;

    // El rango puede haber cambiado de tamaño después de reemplazar el texto.
    // Obtener el rango actualizado.
    const updatedTextRange = doc.Range(range.start, range.start + plainText.length);

    // Aplicar formato basado en el HTML resaltado
    applyFormattingFromHtml(updatedTextRange, highlightedHtml);


    return { success: true };
}


// Función para registrar la herramienta en el servidor FastMCP
export const registerWordCodeFormatTool = (server: any) => {
  server.addTool({
    name: 'word/code-format',
    description: 'Herramienta para identificar, detectar y formatear bloques de código en documentos Word.',
    parameters: codeFormatInputSchema, // Usar 'parameters' en lugar de 'inputSchema'
    execute: handler, // Usar 'execute' en lugar de 'handler'
  });
};