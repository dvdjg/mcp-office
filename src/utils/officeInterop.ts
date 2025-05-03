/**
 * @file Provides utility functions for interacting with Microsoft Office applications via COM Interop.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import * as winax from 'winax';
import * as fs from 'fs/promises'; // For async file operations
import * as path from 'path';     // For path manipulation
import * as os from 'os';         // For temporary directory
import logger from './logger';     // Ensure logger path is correct

export type OfficeAppName = 'Word.Application' | 'Excel.Application' | 'PowerPoint.Application';

/**
 * Gets an instance of an Office application (existing or new).
 * Attempts to connect to an existing instance or create a new one using COM.
 * @param appName The ProgID name of the Office application (e.g., 'Word.Application').
 * @returns A promise resolving to the application's COM object.
 * @throws {Error} If the application instance cannot be obtained or created.
 */
export async function getOfficeApplication(appName: OfficeAppName): Promise<any> {
  logger.info(`[OfficeInterop] Attempting to get/create instance of ${appName}...`);
  let app: any = null;

  try {
    // winax attempts to connect to an existing instance or create a new one with new ActiveXObject.
    // Creating a new object often attaches to an existing one if available.
    app = new winax.Object(appName, { activate: true }); // { activate: true } tries to bring it to the foreground if it exists

    if (!app) {
      throw new Error(`Could not create or connect to ${appName}.`);
    }

    logger.info(`[OfficeInterop] Instance of ${appName} obtained/created.`);

    // Make the application visible for debugging purposes
    try {
      // Check if Visible property exists and is false
      if (typeof app.Visible !== 'undefined' && (app.Visible === false || app.Visible === 0)) {
         // Specific handling for PowerPoint visibility might be needed
         if (appName === 'PowerPoint.Application') {
            // PowerPoint might start without a visible window.
            // Attempting to set Visible might require an open presentation.
            try {
                 app.Visible = true;
            } catch (visError) {
                 logger.warn(`[OfficeInterop] Could not set Visible=true directly for ${appName}. May require opening/creating a file. Error: ${visError instanceof Error ? visError.message : String(visError)}`);
                 // Could attempt app.NewWindow() if necessary.
            }
         } else {
             app.Visible = true; // For Word/Excel, this usually works
         }
         logger.info(`[OfficeInterop] ${appName} set to visible.`);
      }
    } catch (visError) {
      logger.warn(`[OfficeInterop] Could not set Visible property for ${appName}. Error: ${visError instanceof Error ? visError.message : String(visError)}`);
      // Continue even if visibility cannot be set
    }

    return app;
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    logger.error(`[OfficeInterop] Error getting/creating ${appName}: ${errorMessage}`, { error });
    // Attempt to release the object if partially created before failure
    if (app) {
      releaseObject(app);
    }
    throw new Error(`Failed to get application ${appName}: ${errorMessage}`);
  }
}

/**
 * Releases a COM object.
 * Attempts to call the __release method if available, otherwise relies on garbage collection.
 * @param comObject The COM object to release.
 */
export function releaseObject(comObject: any): void {
  if (!comObject) {
    return;
  }
  try {
    // winax might use __release or rely on GC.
    // Calling __release if it exists is safer for immediate resource cleanup.
    if (typeof comObject.__release === 'function') {
      comObject.__release();
      logger.info('[OfficeInterop] COM object released using __release().');
    } else {
      logger.info('[OfficeInterop] COM object has no __release() method. Relying on GC.');
      // Alternatively, setting to null helps GC
      // comObject = null;
    }
  } catch (error) {
    logger.warn(`[OfficeInterop] Warning releasing COM object: ${error instanceof Error ? error.message : String(error)}`);
    // Do not rethrow, just log the warning.
  }
}

/**
 * Opens a Word document.
 * @param wordApp The Word application COM object.
 * @param filePath The path to the document.
 * @param readOnly Whether to open the document as read-only. Defaults to false.
 * @param visible Whether to make the document visible upon opening. Defaults to false.
 * @returns A promise resolving to the document's COM object.
 * @throws {Error} If the document fails to open.
 */
export async function openWordDocument(wordApp: any, filePath: string, readOnly: boolean = false, visible: boolean = false): Promise<any> {
  logger.info(`[OfficeInterop] Attempting to open document: ${filePath}`);
  const absoluteFilePath = path.resolve(filePath);
  try {
    const doc = wordApp.Documents.Open(absoluteFilePath, false, readOnly, false, "", "", false, "", "", 0, visible);
    if (!doc) {
      throw new Error(`Failed to open document: ${filePath}`);
    }
    logger.info(`[OfficeInterop] Document opened successfully: ${filePath}`);
    return doc;
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    logger.error(`[OfficeInterop] Error opening document ${filePath}: ${errorMessage}`, { error });
    throw new Error(`Failed to open document: ${errorMessage}`);
  }
}


// --- Image and Element Handling ---

/**
 * Extracts an image from a Word document.
 * Note: COM Interop for image extraction is complex and might require workarounds like using the clipboard
 * or saving parts of the document. This implementation attempts a common approach using InlineShapes and saving to a temp file.
 * String identifiers are not supported in this basic implementation.
 * @param filePath Path to the Word document.
 * @param imageIdentifier Image index (1-based). String identifiers are not supported.
 * @returns A Promise resolving to a Buffer containing the image data.
 * @throws {Error} If the image identifier is invalid or extraction fails.
 */
export async function extractImageFromWord(
  filePath: string,
  imageIdentifier: string | number
): Promise<Buffer> {
  logger.info(`[OfficeInterop] Attempting to extract image '${imageIdentifier}' from ${filePath}`);
  if (typeof imageIdentifier !== 'number') {
      throw new Error('extractImageFromWord currently only supports numeric (index-based) identifiers.');
  }
  if (imageIdentifier <= 0) {
      throw new Error('Image index must be 1-based.');
  }

  let wordApp: any = null;
  let doc: any = null;
  const tempDir = os.tmpdir();
  // Using a relatively safe temporary filename pattern
  const tempFileName = `mcp_office_img_${Date.now()}_${Math.random().toString(36).substring(2, 8)}.png`;
  const tempFilePath = path.join(tempDir, tempFileName);
  let imageBuffer: Buffer | null = null;

  try {
    wordApp = await getOfficeApplication('Word.Application');
    // Open the document as read-only if possible, make it invisible during processing
    // Use absolute path for COM
    const absoluteFilePath = path.resolve(filePath);
    logger.debug(`[OfficeInterop] Opening document: ${absoluteFilePath}`);
    doc = wordApp.Documents.Open(absoluteFilePath, false, true, false); // ReadOnly=true, Visible=false

    if (!doc) {
      throw new Error(`Failed to open document: ${filePath}`);
    }
    logger.debug(`[OfficeInterop] Document opened successfully.`);

    const inlineShapes = doc.InlineShapes;
    if (!inlineShapes || typeof inlineShapes.Count === 'undefined') {
        throw new Error('Could not access InlineShapes collection.');
    }

    const count = inlineShapes.Count;
    logger.debug(`[OfficeInterop] Document has ${count} inline shapes.`);
    if (imageIdentifier > count) {
      throw new Error(`Image index ${imageIdentifier} out of bounds. Document has ${count} inline shapes.`);
    }

    const shape = inlineShapes.Item(imageIdentifier); // 1-based index

    if (!shape) {
        throw new Error(`Could not find inline shape at index ${imageIdentifier}.`);
    }
    logger.debug(`[OfficeInterop] Found inline shape at index ${imageIdentifier}.`);

    // --- Workaround: Save shape by copying and pasting into a temporary chart ---
    // Direct saving of InlineShape is often not available. This is a common workaround.
    logger.debug('[OfficeInterop] Selecting shape and copying as picture...');
    shape.Select();
    wordApp.Selection.CopyAsPicture();

    // Create a temporary chart object to paste into and export
    // The chart needs to be added to the document temporarily
    logger.debug('[OfficeInterop] Adding temporary chart...');
    // wdChart=-1, xlXYScatter=2 - These constants might vary or need definition
    const tempChartShape = inlineShapes.AddChart2(-1, 2);
    if (!tempChartShape || !tempChartShape.Chart || !tempChartShape.Chart.ChartArea) {
        releaseObject(tempChartShape); // Clean up chart if creation failed partially
        throw new Error('Failed to create temporary chart for image extraction.');
    }
    logger.debug('[OfficeInterop] Temporary chart created. Pasting image...');

    tempChartShape.Chart.ChartArea.Format.Fill.UserPicture(null); // Clears previous picture just in case
    tempChartShape.Chart.ChartArea.Paste(); // Paste the copied image
    logger.debug('[OfficeInterop] Image pasted into chart area. Exporting chart...');

    // Export the chart area (which now contains the image)
    // Ensure the temp directory exists
    await fs.mkdir(tempDir, { recursive: true });
    tempChartShape.Chart.Export(tempFilePath, "PNG"); // Export as PNG
    logger.debug(`[OfficeInterop] Chart exported to temporary file: ${tempFilePath}`);

    // Delete the temporary chart from the document
    logger.debug('[OfficeInterop] Deleting temporary chart...');
    tempChartShape.Delete();

    // Read the exported image file into a buffer
    logger.debug('[OfficeInterop] Reading temporary image file into buffer...');
    imageBuffer = await fs.readFile(tempFilePath);

    logger.info(`[OfficeInterop] Successfully extracted image ${imageIdentifier} to buffer.`);

  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    logger.error(`[OfficeInterop] Error extracting image ${imageIdentifier} from ${filePath}: ${errorMessage}`, { error });
    throw new Error(`Failed to extract image: ${errorMessage}`);
  } finally {
    // Clean up temporary file if it exists
    try {
      if (await fs.stat(tempFilePath).catch(() => false)) {
        await fs.unlink(tempFilePath);
        logger.debug(`[OfficeInterop] Deleted temporary image file: ${tempFilePath}`);
      }
    } catch (cleanupError) {
      logger.warn(`[OfficeInterop] Failed to delete temporary image file ${tempFilePath}: ${cleanupError instanceof Error ? cleanupError.message : String(cleanupError)}`);
    }

    // Close document without saving changes
    if (doc) {
      try {
        doc.Close(0); // wdDoNotSaveChanges = 0
        logger.debug('[OfficeInterop] Document closed without saving.');
      } catch (closeError) {
        logger.warn(`[OfficeInterop] Failed to close document: ${closeError instanceof Error ? closeError.message : String(closeError)}`);
      }
      releaseObject(doc); // Release document object
    }
    // Quit Word only if we opened it and there are no other documents open (tricky to determine reliably)
    // For simplicity, we might leave Word running if it was already open.
    // Consider adding logic to quit if wordApp.Documents.Count === 0, but be cautious.
    // releaseObject(wordApp); // Release app object - careful not to close user's instance unintentionally
  }

  if (!imageBuffer) {
      throw new Error('Image extraction process completed, but no buffer was generated.');
  }
  return imageBuffer;
}

/**
 * Inserts an image into a Word document.
 * @param filePath Path to the Word document.
 * @param imageBuffer Buffer containing the image data.
 * @param position Insertion position identifier (e.g., 'end', 'start', 'paragraph:N'). Bookmarks not implemented.
 * @param options Optional parameters (width, height, altText).
 * @returns A Promise resolving when the image is inserted.
 * @throws {Error} If the document fails to open, the position is invalid, or image insertion fails.
 */
export async function insertImageIntoWord(
  filePath: string,
  imageBuffer: Buffer,
  position: string | number, // Allow number for paragraph index
  options?: { width?: number; height?: number; altText?: string }
): Promise<void> {
  logger.info(`[OfficeInterop] Attempting to insert image into ${filePath} at position '${position}'`);

  let wordApp: any = null;
  let doc: any = null;
  const tempDir = os.tmpdir();
  // Using a relatively safe temporary filename pattern
  const tempFileName = `mcp_office_img_${Date.now()}_${Math.random().toString(36).substring(2, 8)}.tmp`; // Use generic extension
  const tempFilePath = path.join(tempDir, tempFileName);

  try {
    // Ensure the temp directory exists
    await fs.mkdir(tempDir, { recursive: true });
    // Write the buffer to a temporary file
    await fs.writeFile(tempFilePath, imageBuffer);
    logger.debug(`[OfficeInterop] Image buffer saved to temporary file: ${tempFilePath}`);

    wordApp = await getOfficeApplication('Word.Application');
    // Use absolute path for COM
    const absoluteFilePath = path.resolve(filePath);
    logger.debug(`[OfficeInterop] Opening document for editing: ${absoluteFilePath}`);
    doc = wordApp.Documents.Open(absoluteFilePath); // Open normally for editing

    if (!doc) {
      throw new Error(`Failed to open document: ${filePath}`);
    }
    logger.debug(`[OfficeInterop] Document opened successfully.`);

    let targetRange: any = null;

    // Determine the target range based on the position identifier
    // Word COM constants: wdCollapseEnd = 0, wdCollapseStart = 1
    if (position === 'end') {
      logger.debug('[OfficeInterop] Position is "end", targeting end of document content.');
      targetRange = doc.Content;
      targetRange.Collapse(0); // wdCollapseEnd
    } else if (position === 'start') {
      logger.debug('[OfficeInterop] Position is "start", targeting start of document content.');
      targetRange = doc.Content;
      targetRange.Collapse(1); // wdCollapseStart
    } else if (typeof position === 'number' || (typeof position === 'string' && position.startsWith('paragraph:'))) {
        const paraIndex = typeof position === 'number' ? position : parseInt(position.split(':')[1], 10);
        logger.debug(`[OfficeInterop] Position is paragraph index: ${paraIndex}`);
        if (isNaN(paraIndex) || paraIndex <= 0) {
            throw new Error(`Invalid paragraph index specified: ${position}`);
        }
        if (paraIndex > doc.Paragraphs.Count) {
            throw new Error(`Paragraph index ${paraIndex} out of bounds. Document has ${doc.Paragraphs.Count} paragraphs.`);
        }
        targetRange = doc.Paragraphs(paraIndex).Range;
        // Decide whether to insert before or at the start of the paragraph range. Start is usually safer.
        targetRange.Collapse(1); // wdCollapseStart
        logger.debug(`[OfficeInterop] Targeting start of paragraph ${paraIndex}.`);
    }
    // TODO: Add support for bookmarks if needed: e.g., if (doc.Bookmarks.Exists("bookmarkName")) targetRange = doc.Bookmarks("bookmarkName").Range;
    else {
      logger.warn(`[OfficeInterop] Unsupported position identifier: '${position}'. Inserting at the end.`);
      targetRange = doc.Content;
      targetRange.Collapse(0); // wdCollapseEnd
    }

    if (!targetRange) {
        throw new Error(`Could not determine insertion range for position: ${position}`);
    }

    // Insert the picture from the temporary file
    logger.debug(`[OfficeInterop] Inserting picture from temp file: ${tempFilePath}`);
    const inlineShape = targetRange.InlineShapes.AddPicture(tempFilePath);

    if (!inlineShape) {
        throw new Error('Failed to insert image using AddPicture.');
    }
    logger.debug('[OfficeInterop] Image inserted as InlineShape.');

    // Apply options if provided
    if (options) {
        logger.debug('[OfficeInterop] Applying options:', options);
        if (typeof options.width === 'number') {
            inlineShape.Width = options.width;
        }
        if (typeof options.height === 'number') {
            inlineShape.Height = options.height;
        }
        if (typeof options.altText === 'string') {
            inlineShape.AlternativeText = options.altText;
        }
    }

    logger.info(`[OfficeInterop] Successfully inserted image into ${filePath}`);

    // Save the document
    logger.debug('[OfficeInterop] Saving document...');
    doc.Save();
    logger.info(`[OfficeInterop] Document saved: ${filePath}`);

  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    logger.error(`[OfficeInterop] Error inserting image into ${filePath}: ${errorMessage}`, { error });
    throw new Error(`Failed to insert image: ${errorMessage}`);
  } finally {
    // Clean up temporary file if it exists
    try {
      if (await fs.stat(tempFilePath).catch(() => false)) {
        await fs.unlink(tempFilePath);
        logger.debug(`[OfficeInterop] Deleted temporary image file: ${tempFilePath}`);
      }
    } catch (cleanupError) {
      logger.warn(`[OfficeInterop] Failed to delete temporary image file ${tempFilePath}: ${cleanupError instanceof Error ? cleanupError.message : String(cleanupError)}`);
    }

    // Close document
    if (doc) {
      try {
        // Close, but don't save again (already saved if successful)
        doc.Close(0); // wdDoNotSaveChanges = 0 (or handle based on success/failure)
        logger.debug('[OfficeInterop] Document closed.');
      } catch (closeError) {
        logger.warn(`[OfficeInterop] Failed to close document: ${closeError instanceof Error ? closeError.message : String(closeError)}`);
      }
      releaseObject(doc); // Release document object
    }
    // Release app object cautiously
    // releaseObject(wordApp);
  }
}


/**
 * Gets content (text or image buffer) from a specific element in a Word document.
 * Supports retrieving text from a paragraph or extracting an image by index.
 * @param filePath Path to the Word document.
 * @param elementType Type of element ('paragraph' or 'image').
 * @param identifier Index of the element (1-based).
 * @returns A Promise resolving to the text content (string) for a paragraph or image data (Buffer) for an image.
 * @throws {Error} If the element type or index is invalid, or if content retrieval fails.
 */
export async function getWordElementContent(
    filePath: string,
    elementType: 'paragraph' | 'image',
    identifier: number // Assuming 1-based index for simplicity
): Promise<string | Buffer> {
    logger.info(`[OfficeInterop] Attempting to get content for ${elementType} ${identifier} from ${filePath}`);

    if (identifier <= 0) {
        throw new Error('Element index must be 1-based.');
    }

    let wordApp: any = null;
    let doc: any = null;

    try {
        wordApp = await getOfficeApplication('Word.Application');
        // Open read-only and invisible
        const absoluteFilePath = path.resolve(filePath);
        logger.debug(`[OfficeInterop] Opening document read-only: ${absoluteFilePath}`);
        doc = wordApp.Documents.Open(absoluteFilePath, false, true, false); // ReadOnly=true, Visible=false

        if (!doc) {
            throw new Error(`Failed to open document: ${filePath}`);
        }
        logger.debug(`[OfficeInterop] Document opened successfully.`);

        if (elementType === 'paragraph') {
            const paraCount = doc.Paragraphs.Count;
            logger.debug(`[OfficeInterop] Document has ${paraCount} paragraphs. Accessing index ${identifier}.`);
            if (identifier > paraCount) {
                throw new Error(`Paragraph index ${identifier} out of bounds. Document has ${paraCount} paragraphs.`);
            }
            const paragraph = doc.Paragraphs(identifier); // 1-based index
            const text = paragraph.Range.Text;
            logger.info(`[OfficeInterop] Successfully retrieved text for paragraph ${identifier}.`);
            return text;

        } else if (elementType === 'image') {
            // Reuse extractImageFromWord logic - Note: extractImageFromWord handles opening/closing doc internally.
            // This is slightly inefficient as we open the doc twice, but simplifies code reuse.
            // A refactor could pass the 'doc' object to extractImageFromWord.
            logger.info(`[OfficeInterop] Delegating image extraction to extractImageFromWord for index ${identifier}.`);
            // Close the doc opened in *this* function before calling extractImageFromWord
            doc.Close(0); // wdDoNotSaveChanges = 0
            releaseObject(doc);
            doc = null; // Prevent closing again in finally block
            // releaseObject(wordApp); // Release app temporarily? Risky.
            // wordApp = null;

            const imageBuffer = await extractImageFromWord(filePath, identifier);
            logger.info(`[OfficeInterop] Successfully retrieved image buffer for image ${identifier}.`);
            return imageBuffer;

        } else {
            // Should not happen due to type checking, but good practice
            throw new Error(`Unsupported element type: ${elementType}`);
        }

    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        logger.error(`[OfficeInterop] Error getting content for ${elementType} ${identifier} from ${filePath}: ${errorMessage}`, { error });
        throw new Error(`Failed to get element content: ${errorMessage}`);
    } finally {
        // Close document only if it wasn't closed earlier (e.g., for image extraction delegation)
        if (doc) {
            try {
                doc.Close(0); // wdDoNotSaveChanges = 0
                logger.debug('[OfficeInterop] Document closed without saving.');
            } catch (closeError) {
                logger.warn(`[OfficeInterop] Failed to close document: ${closeError instanceof Error ? closeError.message : String(closeError)}`);
            }
            releaseObject(doc); // Release document object
        }
        // Release app object cautiously
        // releaseObject(wordApp);
    }
}