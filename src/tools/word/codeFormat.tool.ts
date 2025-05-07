/**
 * @file Implements the 'word/code-format' tool for identifying, detecting, and formatting code blocks in Word documents using COM Interop and highlight.js.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import * as winax from 'winax';
import hljs from 'highlight.js';
// Import highlight.js and guesslang if available
// import { GuessLang } from 'guesslang';

// Import the language detection module and the ModelOperations class
// Changed import for CommonJS module '@vscode/vscode-languagedetection' to resolve named export issue with ES Modules
import VscodeLanguageDetection from '@vscode/vscode-languagedetection';
const { ModelOperations } = VscodeLanguageDetection;

// const guesslang = new GuessLang();

// Instantiate ModelOperations for language detection
const modelOperations = new ModelOperations();

/** Schema for the input parameters of the 'word/code-format' tool. */
const codeFormatInputSchema = z.object({
  /** The path to the Word file (relative to the current workspace directory). */
  filePath: z.string().describe('The path to the Word file.'),
  /** The operation to perform: 'identify' (identify code blocks), 'detect' (detect language), 'apply' (apply formatting). */
  operation: z.enum(['identify', 'detect', 'apply']).describe('The operation to perform: identify (identify code blocks), detect (detect language), apply (apply formatting).'),
  /** The text range to process for 'detect' and 'apply' operations. Optional for 'identify'. */
  range: z.object({
    /** The starting index of the text range. */
    start: z.number().describe('The starting index of the text range.'),
    /** The ending index of the text range. */
    end: z.number().describe('The ending index of the text range.'),
  }).optional().describe('The text range to process for the detect and apply operations.'),
  /** The programming language if known (for the 'apply' operation). Optional, highlight.js will attempt detection if not provided. */
  language: z.string().optional().describe('The programming language if known (for the apply operation).'),
  /** The name of the Word style to search for (for 'identify') or apply (for 'apply'). Optional. */
  style: z.string().optional().describe('The name of the Word style to search for (for identify) or apply (for apply).'),
  /** The code block to format (for the 'apply' operation). Required for 'apply'. */
  code: z.string().optional().describe('The code block to format (for the apply operation).'),
});

type CodeFormatInput = z.infer<typeof codeFormatInputSchema>;

// Define the interface for the tool context if needed, or use 'any' if specific properties of the FastMCP context are not accessed
// interface ToolContext {
//   // Define the context properties you need, for example:
//   // log: { info: (msg: string) => void };
// }

/**
 * Handler function for the 'word/code-format' tool.
 * Executes the specified operation (identify, detect, or apply) on a Word document.
 * @param input - The input parameters for the tool, validated against `codeFormatInputSchema`.
 * @param context - The FastMCP context (optional, using 'any' for simplicity if specific properties are not accessed).
 * @returns A promise resolving to an object indicating success and the result of the operation.
 * @throws {Error} If the operation is not supported, required parameters are missing, or a COM error occurs.
 */
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
          throw new Error('The range is required for the detect operation.');
        }
        result = await detectLanguage(doc, range);
        break;
      case 'apply':
        // Allow applying formatting without detecting if language and code are provided
        if (!range || !code) {
           throw new Error('The range and code are required for the apply operation.');
        }
        // If language is not provided, attempt detection
        const detectedLanguage = language || await detectLanguage(doc, range);
        result = applyCodeFormatting(doc, range, code, detectedLanguage);
        break;
      default:
        throw new Error(`Unsupported operation: ${operation}`);
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
 * Identifies potential code blocks in a Word document.
 * Searches for paragraphs with a specific style or uses basic text pattern matching.
 * Note: Text pattern detection is a basic heuristic and may not be accurate.
 * @param doc - The Word document COM object.
 * @param style - Optional. The name of the paragraph style to search for.
 * @returns A list of ranges identified as code blocks. Each range is an object with 'start' and 'end' properties.
 */
function identifyCodeBlocks(doc: any, style?: string): { start: number, end: number }[] {
  const codeBlocks: { start: number, end: number }[] = [];
  const paragraphs = doc.Paragraphs;

  for (let i = 1; i <= paragraphs.Count; i++) {
    const paragraph = paragraphs.Item(i);
    let isCode = false;

    if (style) {
      // Check if the paragraph style matches
      try {
        if (paragraph.Style.NameLocal === style || paragraph.Style.NamePrimary === style) {
          isCode = true;
        }
      } catch (e) {
        // Ignore errors if the style cannot be accessed
      }
    } else {
      // Implement detection logic based on text patterns if no style is specified
      // For now, just a basic example: look for lines starting with spaces or tabs
      const text = paragraph.Range.Text.trim();
      if (text.startsWith(' ') || text.startsWith('\t')) {
         // This is a very basic heuristic, improve as needed
         isCode = true;
      }
    }

    if (isCode) {
      codeBlocks.push({
        start: paragraph.Range.Start,
        end: paragraph.Range.End -1 // -1 to not include the paragraph marker
      });
    }
  }

  return codeBlocks;
}

/**
 * Detects the programming language of a text block using @vscode/vscode-languagedetection.
 * @param doc - The Word document COM object.
 * @param range - The text range containing the code.
 * @returns A promise resolving to the detected language string (e.g., 'javascript', 'python', 'plaintext').
 */
async function detectLanguage(doc: any, range: { start: number, end: number }): Promise<string> {
  const textRange = doc.Range(range.start, range.end);
  const codeText = textRange.Text;

  // Use @vscode/vscode-languagedetection for more accurate detection
  try {
    // Use the instantiated modelOperations object and call runModel
    const languages = await modelOperations.runModel(codeText);
    if (languages && languages.length > 0) {
      // Return the language with the highest confidence
      return languages[0].languageId;
    }
  } catch (e) {
    console.error("Error detecting language with @vscode/vscode-languagedetection:", e);
  }

  // Fallback to highlight.js auto-detection if @vscode/vscode-languagedetection is not used or fails
  const autoDetected = hljs.highlightAuto(codeText);
  return autoDetected.language || 'plaintext';
}

/**
 * Applies syntax highlighting formatting to a code block in Word.
 * Uses highlight.js to get the formatting (based on CSS classes) and applies font properties (color, bold, italic) to the Word text range.
 * Note: Formatting application is based on a simple HTML parser and a basic mapping of CSS classes to Word styles. It may not be perfect for all cases.
 * @param doc - The Word document COM object.
 * @param range - The text range containing the code.
 * @param code - The code block to format.
 * @param language - The programming language of the code (optional, highlight.js will attempt to detect if not provided).
 * @returns An object indicating the success of the operation.
 */
// Helper function to apply formatting based on highlighted HTML
function applyCodeFormatting(doc: any, range: { start: number, end: number }, code: string, language: string): { success: boolean } {
    const textRange = doc.Range(range.start, range.end);

    // Integrate highlight.js
    // Use highlightAuto if the language is not specified or is 'plaintext'
    const highlightedResult = language && language !== 'plaintext'
        ? hljs.highlight(code, { language: language })
        : hljs.highlightAuto(code);

    const highlightedHtml = highlightedResult.value;

    // Get the plain text from the highlighted HTML to replace the content in Word
    const plainText = highlightedHtml.replace(/<[^>]*>/g, '');

    // Replace the text in the Word range with the plain text
    // This is crucial for the HTML indices to match the text in Word.
    // Warning: This removes any pre-existing formatting in the range.
    textRange.Text = plainText;

    // The range might have changed size after replacing the text.
    // Get the updated range.
    const updatedTextRange = doc.Range(range.start, range.start + plainText.length);

    // Apply formatting based on the highlighted HTML
    applyFormattingFromHtml(updatedTextRange, highlightedHtml);


    return { success: true };
}

/**
 * Helper function to apply formatting to a Word range based on highlight.js generated HTML.
 * Parses the HTML and applies font properties based on CSS classes.
 * @param textRange - The Word Range object to apply formatting to.
 * @param highlightedHtml - The HTML string generated by highlight.js.
 */
function applyFormattingFromHtml(textRange: any, highlightedHtml: string) {
    // Clear existing formatting in the range
    textRange.Font.Reset();
    textRange.Font.Name = 'Consolas';
    textRange.Font.Size = 10;

    // Simple HTML parser to extract text and classes
    let currentTextIndex = 0;
    // Regex to find <span> with class and content, or text outside of <span>
    const regex = /<span class="([^"]+)">([^<]+)<\/span>|([^<]+)/g;
    let match;

    // Get the plain text from the Word range for more accurate index mapping
    const wordPlainText = textRange.Text.replace(/\r\n/g, '\n').replace(/\r/g, '\n'); // Normalize line breaks

    // Iterate over the highlighted HTML
    while ((match = regex.exec(highlightedHtml)) !== null) {
        let text = '';
        let classes = '';

        if (match[1] && match[2]) { // Matches <span class="...">...</span>
            classes = match[1];
            text = match[2];
        } else if (match[3]) { // Matches text outside of span
            text = match[3];
        }

        if (text) {
            // Find the position of this text in the Word plain text
            // This is still a simplification. A character-by-character mapping would be more robust.
            // We assume the text in the HTML appears in the same order in the Word plain text.
            const startIndexInWord = wordPlainText.indexOf(text, currentTextIndex);

            if (startIndexInWord !== -1) {
                 const subRange = textRange.Duplicate;
                 subRange.Start = textRange.Start + startIndexInWord;
                 subRange.End = subRange.Start + text.length;

                 // Apply formatting based on classes
                 const font = subRange.Font;
                 font.Bold = false;
                 font.Italic = false;
                 font.Color = 0x000000; // Default color (black)

                 const classList = classes.split(' ');
                 for (const cls of classList) {
                     switch (cls) {
                         case 'hljs-keyword':
                             font.Color = 0xFF0000; // Blue
                             font.Bold = true;
                             break;
                         case 'hljs-built_in':
                             font.Color = 0xFFFF00; // Cyan
                             break;
                         case 'hljs-literal':
                             font.Color = 0x00A5FF; // Orange
                             break;
                         case 'hljs-number':
                             font.Color = 0x0000FF; // Red
                             break;
                         case 'hljs-string':
                             font.Color = 0x008000; // Green
                             break;
                         case 'hljs-comment':
                             font.Color = 0x808080; // Gray
                             font.Italic = true;
                             break;
                         case 'hljs-variable':
                             // Default color (black)
                             break;
                         case 'hljs-title':
                             font.Color = 0x800080; // Purple
                             font.Bold = true;
                             break;
                         case 'hljs-params':
                             font.Italic = true;
                             break;
                         case 'hljs-operator':
                             // Default color (black)
                             break;
                         case 'hljs-punctuation':
                             // Default color (black)
                             break;
                         // Add more cases as needed for other highlight.js classes
                     }
                 }

                 currentTextIndex = startIndexInWord + text.length; // Update the index for the next search
            } else {
                // If the text is not found, this indicates a problem with the mapping or the parser.
                // We could log a warning or throw an error.
                console.warn(`Text "${text}" from HTML not found in Word range starting from index ${currentTextIndex}.`);
                // Attempt to advance the index based on the text length in the HTML anyway,
                // even if formatting is not applied correctly to this segment.
                 currentTextIndex += text.length;
            }
        }
    }
}


/**
 * Function to register the 'word/code-format' tool with the FastMCP server.
 * @param server - The FastMCP server instance.
 */
export const registerWordCodeFormatTool = (server: any) => {
  server.addTool({
    name: 'word/code-format',
    description: 'Tool to identify, detect, and format code blocks in Word documents.',
    parameters: codeFormatInputSchema, // Use 'parameters' instead of 'inputSchema'
    execute: handler, // Use 'execute' instead of 'handler'
  });
};