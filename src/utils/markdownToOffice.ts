import MarkdownIt from 'markdown-it';
import logger from './logger';
import { releaseObject } from './officeInterop';

// Initialize Markdown parser
const md = new MarkdownIt({
  html: false,
  xhtmlOut: false,
  breaks: true,
  linkify: true,
  typographer: true,
});

/**
 * Applies Markdown formatting to a Word Range using COM Interop.
 * This is a basic implementation and will need to be expanded to handle
 * various Markdown elements and their corresponding Word formatting via COM.
 *
 * @param range The Word Range object where the formatted text should be inserted.
 * @param markdownText The Markdown text to format and insert.
 * @param wordApp The Word Application COM object.
 */
export async function applyMarkdownFormattingToWord(range: any, markdownText: string, wordApp: any): Promise<void> {
    logger.debug('Starting Markdown formatting for Word.');

    // Parse the Markdown text
    const tokens = md.parse(markdownText, {});
    logger.debug(`Parsed Markdown into ${tokens.length} tokens.`);

    let currentRange = range;
    let listLevel = 0; // Track list nesting level

    for (const token of tokens) {
        logger.debug(`Processing token: ${token.type}`);

        switch (token.type) {
            case 'heading_open':
                // Insert heading text and apply style
                const headingLevel = parseInt(token.tag.substring(1), 10);
                // Find the next inline token for the heading text
                const headingTextToken = tokens[tokens.indexOf(token) + 1];
                if (headingTextToken && headingTextToken.type === 'inline') {
                    currentRange.Text = headingTextToken.content;
                    // Move range to the end of the inserted text
                    currentRange.Collapse(0); // wdCollapseEnd
                    // Add a new paragraph after the heading
                    currentRange.InsertParagraphAfter();
                    // Move range to the new paragraph
                    currentRange.Collapse(0); // wdCollapseEnd

                    // Apply heading style to the paragraph containing the heading text
                    // Need to get the paragraph object from the range
                    let paragraph = null;
                    try {
                         paragraph = currentRange.Paragraphs(1);
                         if (paragraph) {
                             // Apply style - Word styles are typically "Heading 1", "Heading 2", etc.
                             // Need to handle potential errors if style doesn't exist
                             try {
                                 paragraph.Style = `Heading ${headingLevel}`;
                                 logger.debug(`Applied style 'Heading ${headingLevel}'`);
                             } catch (styleError: any) {
                                 logger.warn(`Could not apply style 'Heading ${headingLevel}': ${styleError.message}`);
                                 // Fallback or ignore? For now, just log and continue.
                             }
                         }
                    } catch (paraError: any) {
                         logger.error(`Error getting paragraph for heading: ${paraError.message}`);
                    } finally {
                         if (paragraph) releaseObject(paragraph);
                    }

                    // Move the current range to the start of the paragraph *after* the heading
                    currentRange = currentRange.Next(6 /* wdParagraph */); // Move to the next paragraph
                    if (currentRange) {
                         currentRange.Collapse(1); // wdCollapseStart
                    } else {
                         // If there's no next paragraph, get the end of the document
                         currentRange = wordApp.ActiveDocument.Content;
                         currentRange.Collapse(0); // wdCollapseEnd
                    }


                }
                // Skip the corresponding heading_close token as we processed the text here
                tokens.splice(tokens.indexOf(token) + 2, 1); // Remove inline and heading_close
                break;

            case 'paragraph_open':
                // Insert paragraph text
                 // Find the next inline token for the paragraph text
                const paragraphTextToken = tokens[tokens.indexOf(token) + 1];
                if (paragraphTextToken && paragraphTextToken.type === 'inline') {
                    currentRange.Text = paragraphTextToken.content;
                    // Move range to the end of the inserted text
                    currentRange.Collapse(0); // wdCollapseEnd
                    // Add a new paragraph after the current one
                    currentRange.InsertParagraphAfter();
                    // Move range to the new paragraph
                    currentRange.Collapse(0); // wdCollapseEnd
                }
                // Skip the corresponding paragraph_close token
                tokens.splice(tokens.indexOf(token) + 2, 1); // Remove inline and paragraph_close
                break;

            case 'bullet_list_open':
            case 'ordered_list_open':
                listLevel++;
                // Need to handle list formatting via COM - this is complex.
                // For now, just insert text and rely on Word's auto-formatting if enabled.
                // Proper implementation requires setting ListFormat properties on paragraphs.
                logger.warn(`Basic list handling: Inserting text, relying on Word auto-formatting. List level: ${listLevel}`);
                break;

            case 'list_item_open':
                 // Find the next inline token for the list item text
                const listItemTextToken = tokens[tokens.indexOf(token) + 1];
                if (listItemTextToken && listItemTextToken.type === 'inline') {
                    // Add indentation based on listLevel?
                    // currentRange.Text = '  '.repeat(listLevel - 1) + '- ' + listItemTextToken.content;
                     currentRange.Text = listItemTextToken.content; // Just insert text for now
                    // Move range to the end of the inserted text
                    currentRange.Collapse(0); // wdCollapseEnd
                    // Add a new paragraph after the list item
                    currentRange.InsertParagraphAfter();
                    // Move range to the new paragraph
                    currentRange.Collapse(0); // wdCollapseEnd

                     // TODO: Apply list formatting via COM here
                     // This requires getting the paragraph and setting its ListFormat property
                     // Example (simplified):
                     // let paragraph = null;
                     // try {
                     //      paragraph = currentRange.Paragraphs(1);
                     //      if (paragraph) {
                     //           paragraph.Range.ListFormat.ApplyBulletDefault(); // Or ApplyNumberDefault()
                     //           paragraph.Range.ListFormat.ListIndent(); // For nesting
                     //      }
                     // } catch (listFormatError: any) {
                     //      logger.error(`Error applying list format: ${listFormatError.message}`);
                     // } finally {
                     //      if (paragraph) releaseObject(paragraph);
                     // }
                }
                // Skip the corresponding list_item_close token
                tokens.splice(tokens.indexOf(token) + 2, 1); // Remove inline and list_item_close
                break;

            case 'bullet_list_close':
            case 'ordered_list_close':
                listLevel--;
                logger.debug(`List closed. List level: ${listLevel}`);
                break;

            case 'strong_open':
                // Turn on bold formatting for the current range
                wordApp.Selection.Font.Bold = true; // This affects the current selection, not the range directly
                logger.debug('Bold formatting ON');
                break;
            case 'strong_close':
                 // Turn off bold formatting
                wordApp.Selection.Font.Bold = false;
                logger.debug('Bold formatting OFF');
                break;

            case 'em_open':
                // Turn on italic formatting
                wordApp.Selection.Font.Italic = true;
                logger.debug('Italic formatting ON');
                break;
            case 'em_close':
                // Turn off italic formatting
                wordApp.Selection.Font.Italic = false;
                logger.debug('Italic formatting OFF');
                break;

            case 'inline':
                // This case should ideally be handled by the parent block elements (paragraph, heading, list_item)
                // If we encounter an inline token here, it might be unexpected or needs specific handling.
                // For now, just insert the text.
                logger.warn(`Encountered unexpected inline token outside of a block element: "${token.content}"`);
                currentRange.Text = token.content;
                currentRange.Collapse(0); // wdCollapseEnd
                break;

            case 'softbreak':
                 // Insert a line break
                 currentRange.InsertBreak(6); // wdLineBreak
                 currentRange.Collapse(0); // wdCollapseEnd
                 break;

            case 'hr':
                 // Insert a horizontal rule (border)
                 // This is complex via COM. For now, insert a line of dashes.
                 currentRange.Text = '---';
                 currentRange.Collapse(0); // wdCollapseEnd
                 currentRange.InsertParagraphAfter();
                 currentRange.Collapse(0); // wdCollapseEnd
                 logger.warn('Inserted "---" for horizontal rule.');
                 break;

            // Add cases for other Markdown elements as needed:
            // - blockquote_open/blockquote_close
            // - code_block
            // - fence (for code blocks)
            // - image
            // - link
            // - table_open/table_close, thead_open/thead_close, tbody_open/tbody_close, tr_open/tr_close, th_open/th_close, td_open/td_close
            // - html_block, html_inline (if html: true in parser)

            default:
                logger.debug(`Skipping unhandled token type: ${token.type}`);
                // Handle other token types or ignore
                break;
        }
    }

    logger.debug('Finished Markdown formatting for Word.');
}

// Note: This utility currently focuses on Word. Adapting for Excel and PowerPoint
// will require separate functions or significant conditional logic due to different COM APIs.