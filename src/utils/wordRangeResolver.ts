import logger from './logger.js';
import { releaseObject } from './officeInterop.js';

/**
 * Attempts to resolve a natural language range description to a Word COM Range object.
 * This is a basic implementation and needs significant expansion to handle
 * a wide variety of natural language inputs and document structures.
 *
 * @param doc The Word Document COM object.
 * @param rangeDescription The natural language description of the range (e.g., "after the heading 'Introduction'", "the third paragraph").
 * @param wordApp The Word Application COM object (needed for selection).
 * @returns The resolved COM Range object, or null if resolution fails.
 * @throws Error if the description format is recognized but invalid (e.g., heading not found).
 */
export function resolveNaturalLanguageRange(doc: any, rangeDescription: string, wordApp: any): any | null {
    const descriptionLower = rangeDescription.toLowerCase().trim();
    let resolvedRange: any = null;

    logger.debug(`Attempting to resolve natural language range: "${rangeDescription}"`);

    try {
        // Example: "after the heading 'Introduction'"
        if (descriptionLower.startsWith('after the heading')) {
            const match = rangeDescription.match(/after the heading\s+'([^']+)'/i);
            if (match && match[1]) {
                const headingText = match[1];
                logger.debug(`Searching for heading with text: "${headingText}"`);

                let foundRange: any = null;
                let searchRange: any = null;
                let find: any = null; // Declare find here

               try {
                   // Use Find to locate the heading text
                   searchRange = doc.Content;
                   find = searchRange.Find;
                   find.ClearFormatting();
                    find.Text = headingText;
                    find.Forward = true;
                    find.Wrap = 0; // wdFindContinue
                    find.Format = true; // Look for formatting
                    find.Style = 'Heading 1'; // Start by looking for Heading 1

                    let found = find.Execute();

                    // If not found as Heading 1, try other heading levels (basic example)
                    let level = 1;
                    while (!found && level < 7) { // Check Heading 1 to Heading 6
                         releaseObject(find.Style); // Release previous style object
                         find.Style = `Heading ${level}`;
                         found = find.Execute();
                         level++;
                    }


                    if (found) {
                        foundRange = find.Parent; // Get the range of the found text
                        logger.debug(`Found heading "${headingText}" at Start: ${foundRange.Start}, End: ${foundRange.End}`);

                        // The range we want is *after* this heading.
                        // Move the range to the end of the found heading and collapse it.
                        resolvedRange = doc.Range(foundRange.End, foundRange.End);
                        logger.debug(`Resolved range to position after heading.`);

                    } else {
                        logger.warn(`Heading with text "${headingText}" not found.`);
                        // Do not throw here, return null to indicate not found
                        resolvedRange = null;
                    }
                } catch (findError: any) {
                    logger.error(`Error during heading search: ${findError.message}`);
                    resolvedRange = null; // Indicate failure
                } finally {
                    // Release COM objects created in this block
                    if (find) releaseObject(find);
                    if (searchRange) releaseObject(searchRange);
                    if (foundRange) releaseObject(foundRange);
                }
            } else {
                logger.warn(`Could not parse "after the heading" description: "${rangeDescription}"`);
                // Do not throw, return null
                resolvedRange = null;
            }
        }
        // Add more cases for other natural language descriptions here:
        // - "the Nth paragraph"
        // - "before the section '...'"
        // - "in the table '...'"
        // - "at the end of the document" (already handled by specific 'end' specifier, but could be here)
        // - "at the start of the document" (already handled by specific 'start' specifier, but could be here)
        // - "in the current selection" (already handled by specific 'selection' specifier, but could be here)
        // - Combinations: "the 5th paragraph in the second section"

        // If no specific natural language pattern matched, return null
        if (resolvedRange === undefined) { // Check if it was explicitly set to null or not touched
             resolvedRange = null;
        }


    } catch (error: any) {
        logger.error(`Unexpected error resolving natural language range "${rangeDescription}": ${error.message}`);
        // Ensure any partially created COM objects are released before re-throwing or returning null
        // This requires careful tracking of objects within the try block.
        // For now, rely on the finally block for known objects.
        resolvedRange = null; // Indicate failure
    } finally {
        // Ensure any COM objects created *within* this function that are not returned
        // are released here if they weren't released in their specific blocks.
        // This is crucial for preventing memory leaks.
        // Example: if a temporary Range object was created and not assigned to resolvedRange.
        // This requires careful management of COM object references.
    }


    if (resolvedRange) {
         logger.debug(`Successfully resolved "${rangeDescription}" to a COM Range object.`);
    } else {
         logger.debug(`Could not resolve natural language range: "${rangeDescription}". Returning null.`);
    }

    return resolvedRange; // Caller is responsible for releasing the returned Range object
}

// Note: This utility is complex and will require iterative development
// to support a wide range of natural language inputs and document structures.
// Error handling and COM object management within this utility are critical.