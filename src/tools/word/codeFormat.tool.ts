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
 * Utiliza guesslang si está disponible, o una heurística simple.
 * @param doc El objeto del documento Word.
 * @param range El rango de texto que contiene el código.
 * @returns El lenguaje detectado.
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
 * Utiliza highlight.js para obtener el formato y aplica las propiedades de fuente en Word.
 * @param doc El objeto del documento Word.
 * @param range El rango de texto que contiene el código.
 * @param code El bloque de código a formatear.
 * @param language El lenguaje de programación del código.
 * @returns Un objeto indicando el éxito de la operación.
 */
function applyCodeFormatting(doc: any, range: { start: number, end: number }, code: string, language: string): { success: boolean } {
  const textRange = doc.Range(range.start, range.end);

  // Integrar highlight.js aquí
  const highlightedCode = hljs.highlight(code, { language: language || 'plaintext' }).value;

  // Por ahora, solo aplicar una fuente de ancho fijo como Consolas
  textRange.Font.Name = 'Consolas';
  textRange.Font.Size = 10; // Tamaño de fuente de ejemplo

  // Implementar la aplicación de colores y estilos basados en el output de highlight.js
  // Esto requeriría parsear el HTML/texto de highlight.js y aplicar formato a sub-rangos.
  // Esto puede ser complejo con COM Interop y rangos.
  // Por ahora, solo aplicar una fuente de ancho fijo y un color básico.
  // TODO: Implementar un parser de HTML simple para aplicar estilos más detallados.
  textRange.Font.Name = 'Consolas';
  textRange.Font.Size = 10; // Tamaño de fuente de ejemplo
  textRange.Font.Color = 0x000000; // Color negro (BGR) - ajustar según el tema de highlight.js

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