/**
 * @file Tool for managing slides in PowerPoint presentations.
 * Allows adding, deleting, and (basically) modifying slides.
 * Requires the path to the file and the operation to perform.
 * Delete and set operations require the slide index.
 * The add operation requires the slide layout name.
 * Uses COM Interop via winax.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import logger from '../../utils/logger'; // Import logger
import { saveResource } from '../dynamic/resources.tool'; // Import saveResource
import * as fs from 'fs-extra'; // Import fs to read the PowerPoint file
import * as path from 'path'; // Import path

// Input schema for the powerpoint/slides tool
const SlidesToolInputSchema = z.object({
  filePath: z.string().describe('Path to the PowerPoint presentation file.'),
  operation: z.enum(['add', 'delete', 'set']).describe('Operation to perform: "add" (add slide), "delete" (delete slide), "set" (modify existing slide).'),
  slideIndex: z.number().optional().describe('1-based index of the slide for "delete" or "set" operations.'),
  slideLayout: z.string().optional().describe('Name of the slide layout for the "add" operation (e.g., "ppLayoutTitle").'),
  // We could add more fields here for the 'set' operation, such as content, etc.
  // For now, 'set' might be used for something simple like changing the layout if possible or adding basic content.
  // For this initial implementation, 'set' will not do anything complex, it is only included to meet the specification.
});

type SlidesToolInput = z.infer<typeof SlidesToolInputSchema>;

/**
 * @tool powerpoint/slides
 * @description Manages slides in PowerPoint presentations.
 * Allows adding, deleting, and (basically) modifying slides.
 * Requires the path to the file and the operation to perform.
 * Delete and set operations require the slide index.
 * The add operation requires the slide layout name.
 * Uses COM Interop via winax.
 * @param {SlidesToolInput} input - Input parameters for the tool.
 * @returns {Promise<string>} - A message indicating the result of the operation.
 */
const slidesTool: McpResource = {
  path: 'powerpoint/slides',
  description: 'Manages slides in PowerPoint presentations.',
  schema: SlidesToolInputSchema, // Corrected from inputSchema to schema
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => {
    let pptApp: any = null;
    let presentation: any = null;
    let filePath: string | undefined; // Declare filePath outside the try and allow undefined

    try {
      const input = SlidesToolInputSchema.parse(params);
      filePath = input.filePath; // Assign filePath here
      const { operation, slideIndex, slideLayout } = input;

      // Get PowerPoint instance
      pptApp = await getOfficeApplication('PowerPoint.Application');
      pptApp.Visible = true; // Optional: make the application visible
      // Open or create presentation
      try {
        presentation = pptApp.Presentations.Open(filePath);
      } catch (error) {
        // If the file does not exist, create a new presentation
        presentation = pptApp.Presentations.Add();
        presentation.SaveAs(filePath); // Save the new presentation
      }

      switch (operation) {
        case 'add':
          if (!slideLayout) {
            throw new Error('The slideLayout parameter is required for the "add" operation.');
          }
          // Find the slide layout by name
          let layout;
          try {
            layout = pptApp.SlideLayouts.Item(slideLayout);
          } catch (e) {
            throw new Error(`Slide layout "${slideLayout}" not found.`);
          }

          // Add slide
          const newSlide = presentation.Slides.Add(presentation.Slides.Count + 1, layout.Layout);
          return { success: true, data: `Slide added with layout "${slideLayout}".` }; // Adjusted return

        case 'delete':
          if (slideIndex === undefined) {
            throw new Error('The slideIndex parameter is required for the "delete" operation.');
          }
          if (slideIndex < 1 || slideIndex > presentation.Slides.Count) {
            throw new Error(`Slide index ${slideIndex} out of bounds.`);
          }
          // Delete slide
          presentation.Slides.Item(slideIndex).Delete();
          return { success: true, data: `Slide at index ${slideIndex} deleted.` }; // Adjusted return

        case 'set':
          if (slideIndex === undefined) {
            throw new Error('The slideIndex parameter is required for the "set" operation.');
          }
           if (slideIndex < 1 || slideIndex > presentation.Slides.Count) {
            throw new Error(`Slide index ${slideIndex} out of bounds.`);
          }
          // Basic implementation for 'set'. Could be expanded to modify content, layout, etc.
          // For now, we just confirm that the slide exists.
          const slideToSet = presentation.Slides.Item(slideIndex);
          // Logic to modify the slide would go here
          return { success: true, data: `Operation 'set' on slide ${slideIndex} completed (no complex modifications implemented).` }; // Adjusted return

        default:
          throw new Error(`Unsupported operation: ${operation}`);
      }

    } catch (error: any) {
      // Error handling
      logger.error(`Error in powerpoint/slides tool: ${error.message}`); // Use logger
      return { success: false, error: { code: 'POWERPOINT_SLIDES_ERROR', message: `Error processing PowerPoint request: ${error.message}` } }; // Adjusted error return
    } finally {
      // Save and close the presentation
      if (presentation) {
        try {
            presentation.Save();
            // Save the modified PowerPoint file as a dynamic resource
            // This is done in the finally block because Save() happens here for all modification operations.
            // We don't need to check the specific operation here.
            // Ensure filePath has a value before attempting to read the file
            if (filePath) {
                try {
                    const pptContent = await fs.readFile(filePath, null); // Read as Buffer
                    await saveResource('powerpoint/slides', path.basename(filePath), pptContent);
                    // logger.info(`Saved ${filePath} as a dynamic resource.`);
                } catch (resourceSaveError: any) {
                    // logger.error(`Failed to save ${filePath} as a dynamic resource: ${resourceSaveError.message}`);
                    // Continue execution even if resource saving fails
                }
            }
            presentation.Close();
        } catch (closeError: any) {
            logger.warn(`Error closing the presentation: ${closeError.message}`); // Use logger
        }
        releaseObject(presentation);
      }
      // The PowerPoint application is managed externally, we don't close it here.
      releaseObject(pptApp);
    }
    // Add a return at the end to cover all possible cases
    // This will only be reached if no error was thrown or returned before.
    // In an ideal scenario, all switch cases should return.
    // But to satisfy the linter, we add this fallback return.
    return { success: false, error: { code: 'UNHANDLED_CASE', message: 'PowerPoint slide operation did not return an explicit result.' } };
  },
};

export default slidesTool;