import MarkdownIt from 'markdown-it';
import md_s from 'markdown-it-strikethrough-alt';
import md_fn from 'markdown-it-footnote';
import { isAbsolute, resolve as resolvePath } from 'path';
import logger from './logger.js';
import { releaseObject } from './officeInterop.js';

// Initialize Markdown parser
const md = new MarkdownIt({
  html: true,
  xhtmlOut: false,
  breaks: true,
  linkify: true,
  typographer: true,
});
md.use(md_s);
md.use(md_fn);

/**
 * Normalizes table data to ensure consistent column counts across rows.
 * @param tableData The raw table data from Markdown parsing.
 * @returns Normalized table data with consistent columns.
 */
function normalizeTableData(tableData: { type: string; content: string; colspan: number; rowspan: number }[][]): { type: string; content: string; colspan: number; rowspan: number }[][] {
    if (!tableData.length) return [];

    // Find the maximum number of columns (accounting for colspan)
    let maxCols = 0;
    for (const row of tableData) {
        let colCount = 0;
        for (const cell of row) {
            colCount += cell.colspan;
        }
        maxCols = Math.max(maxCols, colCount);
    }

    // Normalize rows to have the same number of columns
    const normalized: { type: string; content: string; colspan: number; rowspan: number }[][] = [];
    for (const row of tableData) {
        let currentColCount = 0;
        const normalizedRow: { type: string; content: string; colspan: number; rowspan: number }[] = [];
        for (const cell of row) {
            normalizedRow.push({ ...cell });
            currentColCount += cell.colspan;
        }
        // Fill missing columns with empty cells
        while (currentColCount < maxCols) {
            normalizedRow.push({ type: 'td', content: '', colspan: 1, rowspan: 1 });
            currentColCount++;
        }
        normalized.push(normalizedRow);
    }
    return normalized;
}

/**
 * Creates a Word table from a structured table representation and applies merging.
 * @param range The Word Range where the table should be inserted.
 * @param tableData The structured table data (2D array).
 * @param wordApp The Word Application COM object.
 */
async function createWordTableFromData(range: any, tableData: { type: string; content: string; colspan: number; rowspan: number }[][], wordApp: any): Promise<void> {
    if (tableData.length === 0) {
        logger.warn('No table data to create Word table.');
        return;
    }

    // Normalize table data to ensure consistent columns
    const normalizedTableData = normalizeTableData(tableData);
    if (!normalizedTableData.length || !normalizedTableData[0].length) {
        logger.warn('Normalized table data resulted in zero rows or columns.');
        return;
    }
    const numRows = normalizedTableData.length;
    const numCols = normalizedTableData[0].length;

    let wordTable = null; // Define wordTable outside try block for potential release in finally
    try {
        // Insert a new table
        wordTable = range.Tables.Add(range, numRows, numCols);
        logger.debug(`Created Word table with ${numRows} rows and ${numCols} columns.`);

        // Track merged cells to avoid conflicts
        const mergedCells: boolean[][] = Array(numRows).fill(null).map(() => Array(numCols).fill(false));

        let currentRow = 1;
        for (const rowData of normalizedTableData) {
            let currentCol = 1;
            for (const cellData of rowData) {
                // Skip if this cell index is out of bounds (shouldn't happen with normalization, but safety check)
                 if (currentRow > numRows || currentCol > numCols) {
                     logger.warn(`Skipping cell data at calculated position (${currentRow}, ${currentCol}) which is out of table bounds (${numRows}x${numCols}).`);
                     currentCol++; // Still need to advance column counter
                     continue;
                 }

                // Skip if this cell is already part of a rowspan/colspan merge from a previous cell
                if (mergedCells[currentRow - 1][currentCol - 1]) {
                    currentCol++; // Advance column counter
                    continue;
                }

                let cell = null;
                let cellRange = null;
                try {
                    cell = wordTable.Cell(currentRow, currentCol);
                    cellRange = cell.Range;
                    cellRange.Text = cellData.content || ''; // Ensure content is string

                    // Apply bold for table headers (<th>)
                    if (cellData.type === 'th') {
                        cellRange.Font.Bold = true;
                    }

                    // Apply merging for colspan and rowspan
                    if (cellData.colspan > 1 || cellData.rowspan > 1) {
                        const endRow = currentRow + cellData.rowspan - 1;
                        const endCol = currentCol + cellData.colspan - 1;

                        // Boundary check for merge target
                        if (endRow <= numRows && endCol <= numCols) {
                            let endCell = null;
                            let mergeTargetRange = null;
                            try {
                                endCell = wordTable.Cell(endRow, endCol);
                                // Create a temporary range for merging to avoid modifying cellRange directly before merge
                                mergeTargetRange = cellRange.Duplicate;
                                mergeTargetRange.End = endCell.Range.End;

                                // Mark merged cells in our tracking array *before* merging
                                // Start from 0-based index for the array
                                for (let r = currentRow - 1; r < endRow; r++) {
                                    for (let c = currentCol - 1; c < endCol; c++) {
                                        // Additional boundary check for safety
                                        if (r < numRows && c < numCols) {
                                            mergedCells[r][c] = true;
                                        }
                                    }
                                }
                                // Mark the starting cell as merged as well
                                mergedCells[currentRow - 1][currentCol - 1] = true;


                                mergeTargetRange.Cells.Merge();
                                logger.debug(`Merged cells from (${currentRow},${currentCol}) to (${endRow},${endCol})`);
                            } catch (mergeError: any) {
                                logger.error(`Error during merge operation for cell at (${currentRow},${currentCol}): ${mergeError.message}`);
                            } finally {
                                if (endCell) releaseObject(endCell);
                                if (mergeTargetRange) releaseObject(mergeTargetRange);
                            }
                        } else {
                            logger.warn(`Invalid merge range target (${endRow},${endCol}) for cell at (${currentRow},${currentCol}). Table size (${numRows}x${numCols}). Skipping merge.`);
                        }
                    }
                } catch(cellError: any) {
                     logger.error(`Error processing cell at (${currentRow}, ${currentCol}): ${cellError.message}`);
                } finally {
                     if (cellRange) releaseObject(cellRange);
                     if (cell) releaseObject(cell);
                }

                currentCol += cellData.colspan; // Advance by colspan width
            }
            currentRow++;
        }

        // Move range past the table - Use table's range end
        let tableRange = null;
        try {
             tableRange = wordTable.Range;
             range.SetRange(tableRange.End, tableRange.End);
             range.Collapse(0); // wdCollapseEnd
             logger.debug(`Moved range to end of table: ${range.Start}`);
        } catch (rangeError: any) {
             logger.error(`Error moving range past table: ${rangeError.message}`);
             // Fallback: try moving by unit (less reliable)
             range.MoveEnd(5, 1); // wdTable = 5
             range.Collapse(0);
        } finally {
             if (tableRange) releaseObject(tableRange);
        }

    } catch (tableError: any) {
        logger.error(`Error creating Word table: ${tableError.message}`);
        range.Text = `[Table Creation Error: ${tableError.message}]\n`;
        range.Collapse(0);
    } finally {
         if (wordTable) releaseObject(wordTable);
    }
}


/**
 * Applies inline formatting recursively to a range based on Markdown inline tokens.
 * @param range The Word Range to format. **This range will be modified (collapsed).**
 * @param tokens The inline tokens to process.
 * @param doc The Word Document COM object.
 */
async function applyInlineFormatting(range: any, tokens: any[], doc: any): Promise<void> {
    for (const token of tokens) {
        let formatRange = null; // Define outside switch for release in finally
        try {
            switch (token.type) {
                case 'text':
                    range.Text = token.content;
                    range.Collapse(0); // Collapse to end after insertion
                    break;

                case 'strong_open':
                case 'em_open':
                case 's_open':
                    // Process nested content recursively
                    await applyInlineFormatting(range, token.children || [], doc);
                    // Note: Formatting is applied *after* the content is inserted by the recursive call.
                    // We need to get the range of the just-inserted content. This is tricky.
                    // A simpler approach (used here) relies on Word potentially inheriting the range state,
                    // but a more robust method would track start/end positions.
                    // Let's try applying to the current collapsed range's font - might only affect subsequent text.
                    // A better way requires tracking start/end of the recursive call.
                    // For now, let's stick to the simpler, potentially less accurate method:
                    // This part is problematic and likely won't format correctly without range tracking.
                    // Reverting to a state-based approach might be necessary if this fails.
                    // Let's try applying to the range *before* the recursive call, assuming children modify it.
                    // This is complex. Let's defer complex nested formatting for now and focus on basic inline.
                    logger.warn(`Nested formatting (${token.type}) may not be fully applied due to complexity.`);
                    break;
                // --- Simplified Inline Formatting (No Nesting Handled Well) ---
                // Apply formatting based on state might be more reliable for basic cases
                // than complex recursive range management without start/end tracking.
                // Let's comment out the recursive attempt for now and focus on getting *any* inline formatting.

                case 'strong_close':
                case 'em_close':
                case 's_close':
                     // These are handled by the state flags in the main loop (if we revert to that)
                     // Or ignored if using the simplified recursive approach above.
                     break;


                case 'link_open':
                    const href = token.attrGet('href');
                    const linkText = token.children?.map((t: any) => t.content).join('') || ''; // Get text from children
                    if (href && linkText) {
                        const linkStart = range.Start;
                        range.Text = linkText; // Insert link text
                        const linkEnd = range.End;
                        formatRange = doc.Range(linkStart, linkEnd);
                        try {
                            // Ensure ScreenTip is a string, use linkText if title is missing
                            const screenTip = token.attrGet('title') || linkText;
                            doc.Hyperlinks.Add(formatRange, href, "", screenTip, linkText); // Use linkText for TextToDisplay
                            logger.debug(`Applied hyperlink '${href}' to range(${linkStart}, ${linkEnd})`);
                        } catch (linkError: any) {
                            logger.error(`Error applying hyperlink: ${linkError.message}`);
                        }
                        range.Collapse(0); // Collapse after link insertion
                    }
                    break;
                // link_close is implicitly handled

                case 'code_inline':
                    const codeStart = range.Start;
                    range.Text = token.content;
                    const codeEnd = range.End;
                    formatRange = doc.Range(codeStart, codeEnd);
                    formatRange.Font.Name = 'Consolas';
                    logger.debug(`Applied code font to range(${codeStart}, ${codeEnd})`);
                    range.Collapse(0); // Collapse after code insertion
                    break;

                case 'softbreak':
                     range.InsertBreak(6); // wdLineBreak
                     range.Collapse(0);
                     break;

                case 'hardbreak':
                     range.InsertParagraph(); // Insert a paragraph break for hard breaks
                     range.Collapse(0);
                     break;

                case 'image':
                     const imgSrc = token.attrGet('src');
                     const imgAlt = token.content || '';
                     if (imgSrc) {
                         let resolvedImgPath = imgSrc;
                         if (!isAbsolute(imgSrc)) {
                             resolvedImgPath = resolvePath(imgSrc); // Resolve relative to workspace for now
                         }
                         let inlineShape = null;
                         try {
                             inlineShape = range.InlineShapes.AddPicture(resolvedImgPath, false, true);
                             if (imgAlt) inlineShape.AlternativeText = imgAlt;
                             logger.debug(`Inserted image: ${resolvedImgPath}`);
                             range.Collapse(0); // Collapse after image
                         } catch (imgError: any) {
                             logger.error(`Error inserting image '${resolvedImgPath}': ${imgError.message}`);
                             range.Text = `[Image Error: ${imgAlt || imgSrc}]`;
                             range.Collapse(0);
                         } finally {
                             if (inlineShape) releaseObject(inlineShape);
                         }
                     }
                     break;

                default:
                    logger.debug(`Skipping inline token: ${token.type}`);
                    break;
            }
        } catch (inlineError: any) {
             logger.error(`Error processing inline token ${token.type}: ${inlineError.message}`);
             // Attempt to recover by collapsing the range
             try { range.Collapse(0); } catch {}
        } finally {
             if (formatRange) releaseObject(formatRange);
        }
    }
}


export async function applyMarkdownFormattingToWord(range: any, markdownText: string, wordApp: any, doc: any): Promise<void> {
    logger.debug('Starting Markdown formatting for Word.');

    // Parse the Markdown text
    const tokens = md.parse(markdownText, {});
    logger.debug(`Parsed Markdown into ${tokens.length} tokens.`);

    let currentRange = range;
    let listLevel = 0;
    let blockquoteLevel = 0;
    let currentListType: 'bullet' | 'ordered' | null = null;
    let isTableActive = false;
    let currentTableData: { type: string; content: string; colspan: number; rowspan: number }[][] = [];
    let currentTableRow: { type: string; content: string; colspan: number; rowspan: number }[] = [];
    // --- State flags for simple inline formatting (re-introduced as recursive was complex) ---
    let isBoldActive = false;
    let isItalicActive = false;
    let isStrikeActive = false;
    // --- End State flags ---

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        logger.debug(`Processing token: ${token.type} at range ${currentRange.Start}-${currentRange.End}`);
        let paragraph = null; // Define outside switch for release in finally
        let tempRange = null; // Define outside switch for release in finally

        try {
            switch (token.type) {
                case 'heading_open':
                    const headingLevel = parseInt(token.tag.substring(1), 10);
                    const headingTextToken = tokens[i + 1];
                    if (headingTextToken && headingTextToken.type === 'inline') {
                        // Apply inline formatting to the heading content first
                        await applyInlineFormatting(currentRange, headingTextToken.children || [], doc);
                        // Now apply the heading style to the paragraph containing the (now collapsed) range
                        paragraph = currentRange.Paragraphs(1);
                        try {
                            paragraph.Style = `Heading ${headingLevel}`;
                            logger.debug(`Applied style 'Heading ${headingLevel}' to paragraph at ${paragraph.Range.Start}`);
                        } catch (styleError: any) {
                            logger.warn(`Could not apply style 'Heading ${headingLevel}': ${styleError.message}`);
                        }
                        // Insert paragraph break *after* styling
                        currentRange.InsertParagraphAfter();
                        currentRange.Collapse(0);
                        i += 2; // Skip inline and heading_close
                    }
                    break;

                case 'paragraph_open':
                    // Handled by inline and paragraph_close
                    break;

                case 'paragraph_close':
                    // Apply list/blockquote formatting to the paragraph *before* the break
                    paragraph = currentRange.Paragraphs(1); // Get the paragraph at the current range (should contain content)
                    if (listLevel > 0 && currentListType) {
                        const listFormat = paragraph.Range.ListFormat;
                        let listTemplate = null;
                        try {
                            const galleryType = currentListType === 'bullet' ? 1 : 2;
                            listTemplate = wordApp.ListGalleries(galleryType).ListTemplates(1);
                            listFormat.ApplyListTemplateWithLevel(listTemplate, true, 1, listLevel);
                            logger.debug(`Applied ${currentListType} list format at level ${listLevel} to paragraph ${paragraph.Range.Start}`);
                        } catch (templateError: any) {
                            logger.error(`Error applying list template: ${templateError.message}`);
                        } finally {
                            if (listTemplate) releaseObject(listTemplate);
                            releaseObject(listFormat);
                        }
                    } else if (blockquoteLevel > 0) {
                        paragraph.LeftIndent = blockquoteLevel * 36;
                        logger.debug(`Applied blockquote indent level ${blockquoteLevel} to paragraph ${paragraph.Range.Start}`);
                    }
                    // Insert the paragraph break
                    currentRange.InsertParagraphAfter();
                    currentRange.Collapse(0);
                    // Reset inline states after paragraph
                    isBoldActive = false;
                    isItalicActive = false;
                    isStrikeActive = false;
                    break;

                case 'blockquote_open':
                    blockquoteLevel++;
                    break;

                case 'blockquote_close':
                    if (blockquoteLevel > 0) blockquoteLevel--;
                    break;

                case 'inline':
                     // --- Re-applying state-based inline formatting ---
                     tempRange = currentRange.Duplicate; // Work on a copy
                     const inlineStart = tempRange.Start;
                     await applyInlineFormatting(tempRange, token.children || [], doc);
                     const inlineEnd = tempRange.End;

                     if (inlineStart < inlineEnd) {
                         let formatRange = null;
                         try {
                             formatRange = doc.Range(inlineStart, inlineEnd);
                             if (isBoldActive) formatRange.Font.Bold = true; else formatRange.Font.Bold = false;
                             if (isItalicActive) formatRange.Font.Italic = true; else formatRange.Font.Italic = false;
                             if (isStrikeActive) formatRange.Font.StrikeThrough = true; else formatRange.Font.StrikeThrough = false;
                             logger.debug(`Applied state-based formatting to inline range ${inlineStart}-${inlineEnd}`);
                         } catch (fmtError: any) {
                              logger.error(`Error applying state-based format: ${fmtError.message}`);
                         } finally {
                              if (formatRange) releaseObject(formatRange);
                         }
                     }
                     // Update main range to the end of the processed inline content
                     currentRange.SetRange(inlineEnd, inlineEnd);
                     currentRange.Collapse(0);
                     // --- End state-based inline formatting ---
                    break;

                 // --- Inline state toggles ---
                 case 'strong_open': isBoldActive = true; break;
                 case 'strong_close': isBoldActive = false; break;
                 case 'em_open': isItalicActive = true; break;
                 case 'em_close': isItalicActive = false; break;
                 case 's_open': isStrikeActive = true; break;
                 case 's_close': isStrikeActive = false; break;
                 // --- End Inline state toggles ---

                case 'bullet_list_open':
                    listLevel++;
                    currentListType = 'bullet';
                    break;

                case 'ordered_list_open':
                    listLevel++;
                    currentListType = 'ordered';
                    break;

                case 'list_item_open':
                    // Logic moved to paragraph_close
                    break;

                case 'list_item_close':
                    // Logic moved to paragraph_close
                    break;

                case 'bullet_list_close':
                case 'ordered_list_close':
                    if (listLevel > 0) {
                        listLevel--;
                        if (listLevel === 0) {
                            currentListType = null;
                            // Apply Normal style to paragraph *after* the list (the current one)
                            try {
                                paragraph = currentRange.Paragraphs(1);
                                paragraph.Style = 'Normal';
                                logger.debug(`Applied Normal style after list to paragraph ${paragraph.Range.Start}`);
                            } catch (e: any) {
                                logger.warn(`Could not apply Normal style after list: ${e.message}`);
                            }
                        }
                    }
                    break;

                case 'table_open':
                    isTableActive = true;
                    currentTableData = [];
                    logger.debug('Table open');
                    break;

                case 'thead_open': // Track header state for potential styling
                case 'tbody_open':
                    logger.debug(`Table section open: ${token.type}`);
                    break;

                case 'tr_open':
                    currentTableRow = [];
                    logger.debug('Table row open');
                    break;

                case 'th_open':
                case 'td_open':
                    const cellType = token.type === 'th_open' ? 'th' : 'td';
                    let cellContent = '';
                    const cellCloseToken = cellType === 'th' ? 'th_close' : 'td_close';
                    let k = i + 1;
                    let cellTokens = []; // Collect tokens within the cell for potential inline formatting later

                    while (k < tokens.length && tokens[k].type !== cellCloseToken) {
                         // Collect child tokens for potential future inline processing within cells
                         if (tokens[k].type === 'inline') {
                              if (tokens[k].children) { // Check if children is not null
                                   cellTokens.push(...tokens[k].children!); // Use non-null assertion
                              }
                              // For now, still collect plain text content
                              cellContent += tokens[k].content;
                         } else if (tokens[k].type === 'text') { // Handle plain text directly within cell
                              cellTokens.push(tokens[k]);
                              cellContent += tokens[k].content;
                         } else if (tokens[k].type === 'softbreak') {
                              cellTokens.push(tokens[k]);
                              cellContent += String.fromCharCode(11); // Vertical Tab
                         } else if (tokens[k].type === 'code_inline') {
                              cellTokens.push(tokens[k]);
                              cellContent += tokens[k].content;
                         } else {
                              // Store other tokens if needed later
                              cellTokens.push(tokens[k]);
                         }
                        k++;
                    }

                    if (k < tokens.length) { // Found the closing tag
                        currentTableRow.push({
                            type: cellType,
                            content: cellContent, // Use collected plain text for now
                            colspan: 1, // Basic markdown tables don't have colspan/rowspan from parser
                            rowspan: 1,
                            // TODO: Store cellTokens if inline formatting within cells is implemented
                        });
                        logger.debug(`Table cell open (${cellType}): Collected content "${cellContent}"`);
                        i = k; // Move the main loop index past the processed cell tokens and the close tag
                    } else {
                        logger.error(`Could not find closing tag ${cellCloseToken} for ${token.type}`);
                        currentTableRow.push({ type: cellType, content: '[Error: Unclosed Cell]', colspan: 1, rowspan: 1 });
                    }
                    break;
                // th_close and td_close are handled by advancing 'i' = k

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
                        // Use a temporary range to insert the table, then update currentRange
                        tempRange = currentRange.Duplicate;
                        await createWordTableFromData(tempRange, currentTableData, wordApp);
                        // createWordTableFromData should now move the range it was given (tempRange)
                        currentRange.SetRange(tempRange.Start, tempRange.End); // Update main range
                        currentRange.Collapse(0);
                        // Insert paragraph after table to ensure separation
                        currentRange.InsertParagraphAfter();
                        currentRange.Collapse(0);
                        logger.debug(`Word table created. Range at ${currentRange.Start}`);
                    } else {
                         logger.warn('Table close encountered without active table state or data.');
                    }
                    // Reset table state
                    isTableActive = false;
                    currentTableData = [];
                    currentTableRow = [];
                    break;

                case 'fence': // Fenced code blocks
                     paragraph = currentRange.Paragraphs(1);
                     paragraph.Range.Text = token.content; // Insert content into current paragraph
                     paragraph.Range.Font.Name = 'Consolas'; // Apply font
                     // Optional: Apply a specific "Code" style if it exists
                     // try { paragraph.Style = "Code"; } catch(e) { logger.warn("Code style not found"); }
                     logger.debug(`Applied code block formatting to paragraph ${paragraph.Range.Start}`);
                     // Insert paragraph break *after* this one
                     currentRange.InsertParagraphAfter();
                     currentRange.Collapse(0); // Move to the start of the new paragraph
                     break;

                case 'hr': // Horizontal Rule
                     paragraph = currentRange.Paragraphs(1);
                     // Apply border to the *current* paragraph, then insert a new one after
                     const borders = paragraph.Borders;
                     const bottomBorder = borders(-3); // wdBorderBottom
                     try {
                         bottomBorder.LineStyle = 1; // wdLineStyleSingle
                         bottomBorder.LineWidth = 4; // wdLineWidth050pt
                         bottomBorder.Color = -16777216; // wdColorAutomatic
                         // Ensure other borders are off
                         borders(-1).LineStyle = 0; borders(-2).LineStyle = 0; borders(-4).LineStyle = 0;
                         paragraph.Range.Text = ''; // Clear any text
                         logger.debug(`Applied horizontal rule to paragraph ${paragraph.Range.Start}`);
                     } catch (hrError: any) {
                          logger.error(`Error applying HR border: ${hrError.message}`);
                          paragraph.Range.Text = '---'; // Fallback
                     } finally {
                          releaseObject(bottomBorder);
                          releaseObject(borders);
                     }
                     currentRange.InsertParagraphAfter();
                     currentRange.Collapse(0);
                     break;

                // --- Footnote Handling ---
                case 'footnote_ref':
                    let footnote = null;
                    try {
                        // Add footnote at current range, Word handles numbering
                        footnote = doc.Footnotes.Add(currentRange, "");
                        logger.debug(`Inserted footnote reference (ID: ${token.meta.id}, Label: ${token.meta.label}) at ${currentRange.Start}`);
                        // Range is automatically collapsed after footnote insertion by Word
                    } catch (fnError: any) {
                        logger.error(`Error inserting footnote reference: ${fnError.message}`);
                        currentRange.Text = `[Footnote Ref Error: ${token.meta.label}]`;
                        currentRange.Collapse(0);
                    } finally {
                         if (footnote) releaseObject(footnote);
                    }
                    break;

                case 'footnote_block_open':
                    // Usually handled implicitly by Word placing footnotes at end
                    logger.debug("Footnote block open (ignored by current logic)");
                    break;

                case 'footnote_open':
                    // The actual footnote text needs to be added to the footnote object created by footnote_ref.
                    // This requires matching the ref ID/label to the footnote object.
                    // This is complex and not fully implemented here.
                    // Current simplified approach: Skip footnote definition blocks.
                    logger.warn(`Skipping footnote definition block for label: ${token.meta.label}. Manual insertion needed.`);
                    // Skip tokens until footnote_block_close
                    let j = i + 1;
                    while (j < tokens.length && tokens[j].type !== 'footnote_block_close') {
                        j++;
                    }
                    i = j; // Advance main loop index
                    break;

                case 'footnote_close': // Should be skipped by footnote_open logic
                case 'footnote_block_close': // Should be skipped by footnote_open logic
                     logger.debug(`Skipping potentially orphaned footnote token: ${token.type}`);
                     break;
                // --- End Footnote Handling ---


                default:
                    logger.debug(`Skipping unhandled token type: ${token.type}`);
                    break;
            }
        } catch (loopError: any) {
             logger.error(`Error processing token ${token.type} at index ${i}: ${loopError.message}. Attempting to continue.`);
             // Attempt to recover by collapsing the range to the end
             try { currentRange.Collapse(0); } catch {}
        } finally {
             // Release objects created within the loop iteration
             if (paragraph) releaseObject(paragraph);
             if (tempRange) releaseObject(tempRange);
        }
    }

    logger.debug('Finished Markdown formatting for Word.');
}