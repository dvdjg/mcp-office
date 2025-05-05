import MarkdownIt from 'markdown-it';
import logger from './logger';
import { releaseObject } from './officeInterop';

// Initialize Markdown parser
const md = new MarkdownIt({
  html: true,
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
import { parse } from 'node-html-parser'; // Assuming node-html-parser is available or can be added as a dependency

/**
 * Parses an HTML table string and returns a structured representation.
 * @param html The HTML string containing the table.
 * @returns A 2D array representing the table rows and cells, including colspan and rowspan.
 */
function parseHtmlTable(html: string): { type: string; content: string; colspan: number; rowspan: number }[][] {
    const root = parse(html);
    const table = root.querySelector('table');
    if (!table) {
        return [];
    }

    const rows: { type: string; content: string; colspan: number; rowspan: number }[][] = [];
    const trElements = table.querySelectorAll('tr');

    for (const tr of trElements) {
        const row: { type: string; content: string; colspan: number; rowspan: number }[] = [];
        const cellElements = tr.querySelectorAll('th, td');
        for (const cell of cellElements) {
            const colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
            const rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);
            row.push({
                type: cell.tagName.toLowerCase(),
                content: cell.text,
                colspan,
                rowspan,
            });
        }
        rows.push(row);
    }
    return rows;
}

/**
 * Creates a Word table from a structured table representation and applies merging.
 * @param range The Word Range where the table should be inserted.
 * @param tableData The structured table data (2D array).
 * @param wordApp The Word Application COM object.
 */
async function createWordTableFromData(range: any, tableData: { type: string; content: string; colspan: number; rowspan: number }[][], wordApp: any): Promise<void> {
    if (tableData.length === 0) {
        return;
    }

    const numRows = tableData.length;
    const numCols = Math.max(...tableData.map(row => row.length)); // Get max columns

    // Insert a new table
    const wordTable = range.Tables.Add(range, numRows, numCols);

    let currentRow = 1;
    for (const rowData of tableData) {
        let currentCol = 1;
        for (const cellData of rowData) {
            const cell = wordTable.Cell(currentRow, currentCol);
            cell.Range.Text = cellData.content;

            // Apply bold for table headers (<th>)
            if (cellData.type === 'th') {
                cell.Range.Font.Bold = true;
            }

            // Apply merging
            if (cellData.colspan > 1 || cellData.rowspan > 1) {
                let mergeRange = cell.Range;
                // Extend the range to cover the cells to be merged
                if (cellData.colspan > 1) {
                    const endCell = wordTable.Cell(currentRow, currentCol + cellData.colspan - 1);
                    mergeRange.End = endCell.Range.End;
                }
                if (cellData.rowspan > 1) {
                     const endCell = wordTable.Cell(currentRow + cellData.rowspan - 1, currentCol);
                     mergeRange.End = endCell.Range.End;
                }
                 mergeRange.Cells.Merge();
            }

            currentCol += cellData.colspan;
        }
        currentRow++;
    }
}


export async function applyMarkdownFormattingToWord(range: any, markdownText: string, wordApp: any): Promise<void> {
    logger.debug('Starting Markdown formatting for Word.');

    // Parse the Markdown text
    const tokens = md.parse(markdownText, {});
    logger.debug(`Parsed Markdown into ${tokens.length} tokens.`);

    let currentRange = range;
    let listLevel = 0; // Track list nesting level

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        logger.debug(`Processing token: ${token.type}`);

        switch (token.type) {
            case 'heading_open':
                const headingLevel = parseInt(token.tag.substring(1), 10);
                const headingTextToken = tokens[i + 1];
                if (headingTextToken && headingTextToken.type === 'inline') {
                    currentRange.Text = headingTextToken.content;
                    currentRange.Collapse(0); // wdCollapseEnd
                    currentRange.InsertParagraphAfter();
                    currentRange.Collapse(0); // wdCollapseEnd

                    let paragraph = null;
                    try {
                         paragraph = currentRange.Paragraphs(1);
                         if (paragraph) {
                              try {
                                  paragraph.Style = `Heading ${headingLevel}`;
                                  logger.debug(`Applied style 'Heading ${headingLevel}'`);
                              } catch (styleError: any) {
                                  logger.warn(`Could not apply style 'Heading ${headingLevel}': ${styleError.message}`);
                              }
                         }
                    } catch (paraError: any) {
                         logger.error(`Error getting paragraph for heading: ${paraError.message}`);
                    } finally {
                         if (paragraph) releaseObject(paragraph);
                    }

                    currentRange = currentRange.Next(6 /* wdParagraph */);
                    if (currentRange) {
                         currentRange.Collapse(1); // wdCollapseStart
                    } else {
                         currentRange = wordApp.ActiveDocument.Content;
                         currentRange.Collapse(0); // wdCollapseEnd
                    }
                    i += 2; // Skip inline and heading_close
                }
                break;

            case 'paragraph_open':
                const paragraphTextToken = tokens[i + 1];
                if (paragraphTextToken && paragraphTextToken.type === 'inline') {
                    currentRange.Text = paragraphTextToken.content;
                    currentRange.Collapse(0); // wdCollapseEnd
                    currentRange.InsertParagraphAfter();
                    currentRange.Collapse(0); // wdCollapseEnd
                    i += 2; // Skip inline and paragraph_close
                }
                break;

            case 'bullet_list_open':
            case 'ordered_list_open':
                listLevel++;
                logger.warn(`Basic list handling: Inserting text, relying on Word auto-formatting. List level: ${listLevel}`);
                break;

            case 'list_item_open':
                const listItemTextToken = tokens[i + 1];
                if (listItemTextToken && listItemTextToken.type === 'inline') {
                     currentRange.Text = listItemTextToken.content;
                    currentRange.Collapse(0); // wdCollapseEnd
                    currentRange.InsertParagraphAfter();
                    currentRange.Collapse(0); // wdCollapseEnd
                    // TODO: Apply list formatting via COM here
                    i += 2; // Skip inline and list_item_close
                }
                break;

            case 'bullet_list_close':
            case 'ordered_list_close':
                listLevel--;
                logger.debug(`List closed. List level: ${listLevel}`);
                break;

            case 'strong_open':
                wordApp.Selection.Font.Bold = true;
                logger.debug('Bold formatting ON');
                break;
            case 'strong_close':
                wordApp.Selection.Font.Bold = false;
                logger.debug('Bold formatting OFF');
                break;

            case 'em_open':
                wordApp.Selection.Font.Italic = true;
                logger.debug('Italic formatting ON');
                break;
            case 'em_close':
                wordApp.Selection.Font.Italic = false;
                logger.debug('Italic formatting OFF');
                break;

            case 'softbreak':
                 currentRange.InsertBreak(6); // wdLineBreak
                 currentRange.Collapse(0); // wdCollapseEnd
                 break;

            case 'hr':
                 currentRange.Text = '---';
                 currentRange.Collapse(0); // wdCollapseEnd
                 currentRange.InsertParagraphAfter();
                 currentRange.Collapse(0); // wdCollapseEnd
                 logger.warn('Inserted "---" for horizontal rule.');
                 break;

            case 'html_block':
            case 'html_inline':
                // Handle HTML content, specifically tables
                if (token.content.includes('<table')) {
                    logger.debug('Detected HTML table. Parsing and creating Word table.');
                    try {
                        const tableData = parseHtmlTable(token.content);
                        if (tableData.length > 0) {
                            await createWordTableFromData(currentRange, tableData, wordApp);
                            // After inserting the table, move the range past the table
                            currentRange.Collapse(0); // wdCollapseEnd
                            currentRange.InsertParagraphAfter(); // Add a paragraph after the table
                            currentRange.Collapse(0); // wdCollapseEnd
                        }
                    } catch (error: any) {
                        logger.error(`Error processing HTML table: ${error.message}`);
                        // Optionally insert the raw HTML as text if parsing fails
                        currentRange.Text = token.content;
                        currentRange.Collapse(0); // wdCollapseEnd
                        currentRange.InsertParagraphAfter();
                        currentRange.Collapse(0); // wdCollapseEnd
                    }
                } else {
                    // For other HTML, just insert as text for now
                    currentRange.Text = token.content;
                    currentRange.Collapse(0); // wdCollapseEnd
                    currentRange.InsertParagraphAfter();
                    currentRange.Collapse(0); // wdCollapseEnd
                }
                break;

            default:
                logger.debug(`Skipping unhandled token type: ${token.type}`);
                break;
        }
    }

    logger.debug('Finished Markdown formatting for Word.');
}

// Note: This utility currently focuses on Word. Adapting for Excel and PowerPoint
// will require separate functions or significant conditional logic due to different COM APIs.