import MarkdownIt from 'markdown-it';
// Assuming the plugin is installed: npm install markdown-it-strikethrough-alt
import md_s from 'markdown-it-strikethrough-alt';
import md_fn from 'markdown-it-footnote'; // Import footnote plugin
import * as path from 'path'; // Import path module
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
// Enable strikethrough plugin
md.use(md_s);
md.use(md_fn); // Enable footnote plugin

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


export async function applyMarkdownFormattingToWord(range: any, markdownText: string, wordApp: any, doc: any): Promise<void> { // Added 'doc' parameter
    logger.debug('Starting Markdown formatting for Word.');

    // State variables for inline formatting
    let isBoldActive = false;
    let isItalicActive = false;
    let isStrikeActive = false; // State for strikethrough
    let isLinkActive = false;   // State for hyperlink
    let currentLinkHref: string | null = null;
    let currentLinkRanges: { start: number; end: number }[] = [];
    // State variables for table processing
    let isTableActive = false;
    let currentTableData: { type: string; content: string; colspan: number; rowspan: number }[][] = [];
    let currentTableRow: { type: string; content: string; colspan: number; rowspan: number }[] = [];

    // Parse the Markdown text
    const tokens = md.parse(markdownText, {});
    logger.debug(`Parsed Markdown into ${tokens.length} tokens.`);

    let currentRange = range;
    let listLevel = 0; // Track list nesting level
    let blockquoteLevel = 0; // Track blockquote nesting level
    let currentListType: 'bullet' | 'ordered' | null = null; // Track current list type
    let isDefiningFootnote = false; // Track if currently processing footnote definition text

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
                // Apply list formatting if we are inside a list
                if (listLevel > 0 && currentListType) {
                    let listPara = null;
                    try {
                        // Get the current paragraph (which might be empty at this point)
                        listPara = currentRange.Paragraphs(1);
                        if (listPara) {
                            const listFormat = listPara.Range.ListFormat;
                            let listTemplate = null;
                            try {
                                // Constants for list galleries (wdBulletGallery = 1, wdNumberGallery = 2)
                                const galleryType = currentListType === 'bullet' ? 1 : 2;
                                // Use the first template in the gallery for simplicity
                                listTemplate = wordApp.ListGalleries(galleryType).ListTemplates(1);

                                if (listTemplate) {
                                    // Apply the list template at the current level
                                    // DefaultListBehavior: wdWord10ListBehavior = 1 (or 0 for wdWord9ListBehavior)
                                    listFormat.ApplyListTemplateWithLevel(listTemplate, true, 1, listLevel);
                                    logger.debug(`Applied ${currentListType} list format at level ${listLevel}`);
                                } else {
                                    logger.warn(`Could not retrieve list template for ${currentListType} list.`);
                                }
                            } catch (templateError: any) {
                                logger.error(`Error getting list template: ${templateError.message}`);
                            } finally {
                                if (listTemplate) releaseObject(listTemplate);
                                releaseObject(listFormat);
                            }
                        }
                    } catch (listParaError: any) {
                        logger.error(`Error applying list format: ${listParaError.message}`);
                    } finally {
                        if (listPara) releaseObject(listPara);
                    }
                }
                // Content insertion is handled by the 'inline' token case
                break;

            case 'paragraph_close':
                 // Insert paragraph break after the inline content has been inserted
                 currentRange.InsertParagraphAfter();
                 currentRange.Collapse(0); // wdCollapseEnd
                 // Reset inline formatting states at the end of a paragraph block
                 isBoldActive = false;
                 isItalicActive = false;
                 isStrikeActive = false; // Reset strikethrough at end of paragraph
                 // Apply blockquote indentation if active
                 if (blockquoteLevel > 0) {
                     let paragraph = null;
                     try {
                         // Get the paragraph that was just finished (before the InsertParagraphAfter)
                         paragraph = currentRange.Paragraphs(1);
                         if (paragraph) {
                             // Apply indentation - Word uses points (1 inch = 72 points)
                             const indentPoints = blockquoteLevel * 36; // 0.5 inch per level
                             paragraph.LeftIndent = indentPoints;
                             logger.debug(`Applied blockquote indent level ${blockquoteLevel} (${indentPoints} points)`);
                             // Optional: Apply a style like "Quote" or "Block Text"
                             // try { paragraph.Style = "Quote"; } catch (e) { logger.warn("Could not apply 'Quote' style."); }
                         }
                     } catch (paraError: any) {
                         logger.error(`Error applying blockquote indent: ${paraError.message}`);
                     } finally {
                         if (paragraph) releaseObject(paragraph);
                     }
                 }
                 break;

            case 'blockquote_open':
                blockquoteLevel++;
                logger.debug(`Blockquote open, level ${blockquoteLevel}`);
                break;

            case 'blockquote_close':
                if (blockquoteLevel > 0) {
                    blockquoteLevel--;
                    logger.debug(`Blockquote close, level ${blockquoteLevel}`);
                    // Indentation is applied in paragraph_close based on the level *before* the paragraph break
                }
                break;

            case 'inline': // Handle actual text content and apply formatting
                const textContent = token.content;
                if (textContent) {
                    const startPos = currentRange.Start;
                    currentRange.Text = textContent; // Insert the text
                    const endPos = currentRange.End; // Get end position *after* insertion

                    if (startPos !== endPos) { // Only format if text was actually inserted
                        let insertedRange = null;
                        try {
                            insertedRange = doc.Range(startPos, endPos); // Get the specific range of inserted text
                            if (isBoldActive) {
                                insertedRange.Font.Bold = true;
                                logger.debug(`Applied bold to range(${startPos}, ${endPos})`);
                            } else {
                                // Explicitly turn off bold if not active, Word might inherit otherwise
                                insertedRange.Font.Bold = false;
                            }
                            if (isItalicActive) {
                                insertedRange.Font.Italic = true;
                                logger.debug(`Applied italic to range(${startPos}, ${endPos})`);
                            } else {
                                // Explicitly turn off italic
                                insertedRange.Font.Italic = false;
                            }
                            // Apply Strikethrough
                            if (isStrikeActive) {
                                insertedRange.Font.StrikeThrough = true; // Correct property name
                                logger.debug(`Applied strikethrough to range(${startPos}, ${endPos})`);
                            } else {
                                insertedRange.Font.StrikeThrough = false;
                            }
                        } catch (formatError: any) {
                            logger.error(`Error applying inline formatting to range(${startPos}, ${endPos}): ${formatError.message}`);
                        } finally {
                            if (insertedRange) releaseObject(insertedRange);
                        }
                    }
                    currentRange.Collapse(0); // Collapse to end, ready for next token
                }
                break;


            case 'bullet_list_open':
                listLevel++;
                currentListType = 'bullet';
                logger.debug(`Bullet list open, level ${listLevel}`);
                break;

            case 'ordered_list_open':
                listLevel++;
                currentListType = 'ordered';
                logger.debug(`Ordered list open, level ${listLevel}`);
                break;

            case 'list_item_open':
                // List item content is handled by paragraph_open and inline tokens
                logger.debug(`List item open, level ${listLevel}`);
                break;

            case 'list_item_close':
                 // Handled by paragraph_close applying formatting before break
                 logger.debug(`List item close, level ${listLevel}`);
                 break;

            case 'bullet_list_close':
            case 'ordered_list_close':
                if (listLevel > 0) {
                    listLevel--;
                    logger.debug(`List closed, level now ${listLevel}`);
                    if (listLevel === 0) {
                        currentListType = null; // Reset type when back to top level
                        logger.debug('Reset currentListType');
                        // Optional: Apply Normal style to the paragraph following the list
                        try {
                            currentRange.Paragraphs(1).Style = "Normal";
                        } catch(e) {
                            logger.warn('Could not apply Normal style after list.');
                        }
                    }
                }
                break;

            case 'strong_open': // Toggle bold state ON
                isBoldActive = true;
                logger.debug('Bold formatting state ON');
                break;
            case 'strong_close': // Toggle bold state OFF
                isBoldActive = false;
                logger.debug('Bold formatting state OFF');
                break;

            case 'em_open': // Toggle italic state ON
                isItalicActive = true;
                logger.debug('Italic formatting state ON');
                break;
            case 'em_close': // Toggle italic state OFF
                isItalicActive = false;
                logger.debug('Italic formatting state OFF');
                break;

            case 's_open': // Strikethrough open
                isStrikeActive = true;
                logger.debug('Strikethrough state ON');
                break;
            case 's_close': // Strikethrough close
                isStrikeActive = false;
                logger.debug('Strikethrough state OFF');
                break;

            case 'code_inline': // Handle inline code
                const codeContent = token.content;
                if (codeContent) {
                    const startPos = currentRange.Start;
                    currentRange.Text = codeContent; // Insert the code text
                    const endPos = currentRange.End;

                    if (startPos !== endPos) {
                        let codeRange = null;
                        try {
                            codeRange = doc.Range(startPos, endPos);
                            codeRange.Font.Name = "Consolas"; // Apply monospace font
                            // Optional: Apply other formatting like background shading or a character style
                            logger.debug(`Applied monospace font to inline code range(${startPos}, ${endPos})`);
                        } catch (codeFormatError: any) {
                            logger.error(`Error applying inline code formatting: ${codeFormatError.message}`);
                        } finally {
                            if (codeRange) releaseObject(codeRange);
                        }
                    }
                    currentRange.Collapse(0); // Collapse to end
                }
                break;


            case 'link_open':
                currentLinkHref = token.attrGet('href');
                if (currentLinkHref) {
                    isLinkActive = true;
                    currentLinkRanges = []; // Reset ranges for the new link
                    logger.debug(`Link open: ${currentLinkHref}`);
                } else {
                    logger.warn('Link open token without href attribute.');
                }
                break;

            case 'link_close':
                if (isLinkActive && currentLinkHref && currentLinkRanges.length > 0) {
                    // Determine the full range of the link text
                    const linkStart = Math.min(...currentLinkRanges.map(r => r.start));
                    const linkEnd = Math.max(...currentLinkRanges.map(r => r.end));

                    if (linkStart < linkEnd) {
                        let linkRange = null;
                        try {
                            linkRange = doc.Range(linkStart, linkEnd);
                            doc.Hyperlinks.Add(linkRange, currentLinkHref, "", "", token.content); // Use token content for ScreenTip
                            logger.debug(`Applied hyperlink '${currentLinkHref}' to range(${linkStart}, ${linkEnd})`);
                        } catch (linkError: any) {
                            logger.error(`Error applying hyperlink: ${linkError.message}`);
                        } finally {
                             // linkRange is implicitly released when used by Hyperlinks.Add? Check COM docs.
                             // Let's release it to be safe if Add doesn't consume it.
                             // if (linkRange) releaseObject(linkRange);
                        }
                    } else {
                         logger.warn(`Could not determine valid range for link: ${currentLinkHref}`);
                    }
                } else {
                     logger.warn(`Link close encountered without active link state or text ranges.`);
                }
                // Reset link state regardless of success
                isLinkActive = false;
                currentLinkHref = null;
                currentLinkRanges = [];
                logger.debug('Link close processed.');
                break;

            case 'image':
                const imgSrc = token.attrGet('src');
                const imgAlt = token.content || ''; // Alt text is the content of the image token
                const imgTitle = token.attrGet('title') || ''; // Optional title attribute

                if (imgSrc) {
                    let resolvedImgPath = imgSrc;
                    // Basic path resolution (assuming absolute or relative to workspace)
                    // TODO: Enhance path resolution if markdown file path context is available
                    if (!path.isAbsolute(imgSrc)) {
                        // Assuming workspace root as base for relative paths for now
                        resolvedImgPath = path.resolve(imgSrc);
                        logger.debug(`Resolved relative image path '${imgSrc}' to '${resolvedImgPath}'`);
                    }

                    let inlineShape = null;
                    try {
                        // Insert the picture as an inline shape
                        // LinkToFile = False, SaveWithDocument = True
                        inlineShape = currentRange.InlineShapes.AddPicture(resolvedImgPath, false, true);

                        if (inlineShape) {
                            // Set Alt Text
                            if (imgAlt) {
                                inlineShape.AlternativeText = imgAlt;
                            }
                            // Optional: Set Title (Word doesn't have a direct 'title' property like HTML)
                            // Could potentially add it to Description or elsewhere if needed.
                            logger.debug(`Inserted image: ${resolvedImgPath} with alt text: "${imgAlt}"`);

                            // Collapse range to be after the inserted shape
                            currentRange.Collapse(0); // Collapse to the end of the range where the shape was inserted
                        } else {
                            logger.error(`Failed to insert image shape for: ${resolvedImgPath}`);
                        }
                    } catch (imgError: any) {
                        logger.error(`Error inserting image '${resolvedImgPath}': ${imgError.message}`);
                        // Optionally insert placeholder text on error
                        currentRange.Text = `[Image Error: ${imgAlt || imgSrc}]`;
                        currentRange.Collapse(0);
                    } finally {
                        if (inlineShape) releaseObject(inlineShape);
                    }
                } else {
                    logger.warn('Image token encountered without src attribute.');
                }
                break;

            case 'softbreak':
                 currentRange.InsertBreak(6); // wdLineBreak
                 currentRange.Collapse(0); // wdCollapseEnd
                 break;

            case 'hr':
                // Insert an empty paragraph for the rule
                currentRange.InsertParagraphAfter();
                // Apply border to the *previous* paragraph (the one just inserted)
                let hrParagraph = null;
                try {
                    // Go back to the paragraph just inserted
                    hrParagraph = currentRange.Paragraphs(1); // Get the paragraph at the current range (which is after the inserted one)
                    if (hrParagraph) {
                        // Constants for borders (replace with actual values if available)
                        const wdLineStyleSingle = 1;
                        const wdLineWidth050pt = 4; // 0.5 points
                        const wdColorAutomatic = -16777216; // Or specific color if needed

                        // Apply bottom border
                        const borders = hrParagraph.Borders;
                        const bottomBorder = borders(-3); // wdBorderBottom = -3

                        bottomBorder.LineStyle = wdLineStyleSingle;
                        bottomBorder.LineWidth = wdLineWidth050pt;
                        bottomBorder.Color = wdColorAutomatic;

                        // Ensure other borders are off for this paragraph
                        borders(-1).LineStyle = 0; // wdBorderTop = -1, wdLineStyleNone = 0
                        borders(-2).LineStyle = 0; // wdBorderLeft = -2
                        borders(-4).LineStyle = 0; // wdBorderRight = -4
                        borders(-5).LineStyle = 0; // wdBorderHorizontal = -5
                        borders(-6).LineStyle = 0; // wdBorderVertical = -6

                        // Clear the text content of the paragraph if any was inherited
                        hrParagraph.Range.Text = "";

                        logger.debug('Applied horizontal rule (paragraph bottom border).');
                        releaseObject(bottomBorder);
                        releaseObject(borders);
                    }
                } catch (hrError: any) {
                    logger.error(`Error applying horizontal rule border: ${hrError.message}`);
                    // Fallback: Insert '---' if border fails
                    currentRange.InsertBefore('---');
                    currentRange.InsertParagraphAfter();
                } finally {
                    if (hrParagraph) releaseObject(hrParagraph);
                }
                // Ensure range is collapsed *after* the HR paragraph
                currentRange.Collapse(0); // wdCollapseEnd
                break;

            case 'fence': // Handle fenced code blocks
                const fenceContent = token.content;
                if (fenceContent) {
                    const startPos = currentRange.Start;
                    // Insert the code content, preserving line breaks
                    currentRange.Text = fenceContent;
                    const endPos = currentRange.End;

                    if (startPos !== endPos) {
                        let codeBlockRange = null;
                        try {
                            codeBlockRange = doc.Range(startPos, endPos);
                            codeBlockRange.Font.Name = "Consolas"; // Apply monospace font
                            // Optional: Apply a specific style like "Code" or add borders/shading
                            // codeBlockRange.Style = "Code";
                            logger.debug(`Applied monospace font to code block range(${startPos}, ${endPos})`);
                        } catch (codeBlockError: any) {
                            logger.error(`Error applying code block formatting: ${codeBlockError.message}`);
                        } finally {
                            if (codeBlockRange) releaseObject(codeBlockRange);
                        }
                    }
                    // Insert a paragraph break *after* the code block
                    currentRange.Collapse(0); // Collapse to end of inserted text
                    currentRange.InsertParagraphAfter();
                    currentRange.Collapse(0); // Collapse to the start of the new paragraph
                }
                break;

            // --- Table Handling ---
            case 'table_open':
                isTableActive = true;
                currentTableData = [];
                logger.debug('Table open');
                break;

            case 'thead_open':
            case 'tbody_open':
                // We don't need specific handling for thead/tbody for basic table creation
                logger.debug(`Table section open: ${token.type}`);
                break;

            case 'tr_open':
                currentTableRow = [];
                logger.debug('Table row open');
                break;

            case 'th_open':
            case 'td_open':
                const cellType = token.type === 'th_open' ? 'th' : 'td';
                // The content is in the following 'inline' token
                if (tokens[i + 1]?.type === 'inline') {
                    const cellContent = tokens[i + 1].content;
                    currentTableRow.push({
                        type: cellType,
                        content: cellContent,
                        colspan: 1, // Basic markdown tables don't have colspan/rowspan
                        rowspan: 1,
                    });
                    logger.debug(`Table cell open (${cellType}): ${cellContent}`);
                    i += 2; // Skip the inline and the th_close/td_close tokens
                } else {
                    logger.warn(`Expected inline token after ${token.type}, but found ${tokens[i+1]?.type}`);
                    // Add an empty cell as fallback
                     currentTableRow.push({ type: cellType, content: '', colspan: 1, rowspan: 1 });
                     i++; // Skip only the th_close/td_close
                }
                break;
            // th_close and td_close are implicitly handled by skipping tokens after _open

            case 'tr_close':
                if (currentTableRow.length > 0) {
                    currentTableData.push(currentTableRow);
                }
                logger.debug('Table row close');
                break;

            case 'thead_close':
            case 'tbody_close':
                logger.debug(`Table section close: ${token.type}`);
                break;

            case 'table_close':
                if (isTableActive && currentTableData.length > 0) {
                    logger.debug('Table close. Creating Word table.');
                    try {
                        // Use existing helper, as it handles the simple structure correctly
                        await createWordTableFromData(currentRange, currentTableData, wordApp);
                        // Move range past the inserted table - createWordTableFromData doesn't return the new range
                        // We need to collapse and insert a paragraph after
                        currentRange.Collapse(0); // Collapse to end of table range
                        currentRange.InsertParagraphAfter(); // Add paragraph break after table
                        currentRange.Collapse(0); // Move to start of new paragraph
                        logger.debug('Word table created successfully.');
                    } catch (tableError: any) {
                        logger.error(`Error creating Word table from Markdown: ${tableError.message}`);
                        // Insert placeholder on error
                        currentRange.Text = `[Table Creation Error: ${tableError.message}]\n`;
                        currentRange.Collapse(0);
                    }
                } else {
                     logger.warn('Table close encountered without active table state or data.');
                }
                // Reset table state
                isTableActive = false;
                currentTableData = [];
                currentTableRow = [];
                break;
            // --- End Table Handling ---


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

            // --- Footnote Handling (Simplified) ---
            case 'footnote_ref':
                // Insert the footnote reference marker at the current position
                try {
                    // The Reference parameter is optional; Word handles numbering automatically.
                    // We might use token.meta.id or token.meta.label if needed for specific linking later.
                    const footnote = doc.Footnotes.Add(currentRange, ""); // Add footnote at current range
                    // Footnote text needs to be handled separately where definitions appear
                    logger.debug(`Inserted footnote reference (ID: ${token.meta.id}, Label: ${token.meta.label})`);
                    // Collapse range *after* the inserted footnote reference
                    currentRange.Collapse(0); // wdCollapseEnd
                    releaseObject(footnote);
                } catch (fnError: any) {
                    logger.error(`Error inserting footnote reference: ${fnError.message}`);
                    currentRange.Text = `[Footnote Ref Error: ${token.meta.label}]`; // Insert placeholder on error
                    currentRange.Collapse(0);
                }
                break;

            case 'footnote_block_open':
                // Optional: Insert a separator before footnote definitions
                currentRange.InsertParagraphAfter();
                currentRange.InsertBefore("--- Footnotes ---");
                currentRange.InsertParagraphAfter();
                currentRange.Collapse(0);
                logger.debug("Footnote block open");
                break;

            case 'footnote_open':
                isDefiningFootnote = true;
                // Insert the footnote label (e.g., "[1]: ") as regular text
                currentRange.InsertBefore(`[${token.meta.label}]: `);
                currentRange.Collapse(0); // Collapse after the label
                logger.debug(`Footnote definition open: ${token.meta.label}`);
                break;

            // paragraph_open, inline, text tokens inside footnotes will be handled normally,
            // appending text after the label inserted by footnote_open.

            case 'footnote_close':
                isDefiningFootnote = false;
                logger.debug(`Footnote definition close: ${token.meta.label}`);
                // The paragraph_close following this will add the paragraph break.
                break;

            case 'footnote_block_close':
                logger.debug("Footnote block close");
                // Optional: Add extra space after footnotes
                currentRange.InsertParagraphAfter();
                currentRange.Collapse(0);
                break;
            // --- End Footnote Handling ---


            default:
                logger.debug(`Skipping unhandled token type: ${token.type}`);
                break;
        }
    }

    logger.debug('Finished Markdown formatting for Word.');
}

// Note: This utility currently focuses on Word. Adapting for Excel and PowerPoint
// will require separate functions or significant conditional logic due to different COM APIs.