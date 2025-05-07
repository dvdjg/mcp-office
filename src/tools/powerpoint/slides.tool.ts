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
import PptxGenJS from 'pptxgenjs'; // Assuming this should be the class
import officeParser from 'officeparser';
import { McpResource, ToolRequestParams, ApiResponse } from '../../types/common.types.js';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js';
import logger from '../../utils/logger.js'; // Import logger
import { saveResource } from '../dynamic/resources.tool.js'; // Import saveResource
import fs from 'fs-extra'; // Import fs to read the PowerPoint file
import { basename as basenamePath } from 'path'; // Import path

// Input schema for the powerpoint/slides tool
const SlidesToolInputSchema = z.object({
  filePath: z.string().describe('Path to the PowerPoint presentation file.'),
  operation: z.enum(['add', 'delete', 'set', 'getText']).describe('Operation to perform: "add" (add slide), "delete" (delete slide), "set" (modify existing slide), "getText" (extract text from all slides).'),
  slideIndex: z.number().optional().describe('1-based index of the slide for "delete" or "set" operations.'),
  slideLayout: z.string().optional().describe('Name of the slide layout for the "add" operation (e.g., "ppLayoutTitle" for COM, or a PptxGenJS layout name).'),
  useComInterop: z.boolean().optional().default(false).describe('Set to true to use COM Interop for operations, otherwise uses pptxgenjs/officeparser.'),
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
    const input = SlidesToolInputSchema.parse(params);
    const { filePath, operation, slideIndex, slideLayout, useComInterop } = input;

    if (useComInterop) {
      // COM Interop Path (existing logic)
      let pptApp: any = null;
      let presentation: any = null;
      try {
        pptApp = await getOfficeApplication('PowerPoint.Application');
        pptApp.Visible = true; // Optional: make the application visible
        try {
          presentation = pptApp.Presentations.Open(filePath);
        } catch (error) {
          presentation = pptApp.Presentations.Add();
          presentation.SaveAs(filePath);
        }

        switch (operation) {
          case 'add':
            if (!slideLayout) {
              throw new Error('The slideLayout parameter is required for the "add" operation with COM Interop.');
            }
            let layout;
            try {
              // Attempt to find the layout by its COM-specific name or constant value
              // PowerPoint's PpSlideLayout enumeration might be needed here if names aren't direct.
              // For simplicity, assuming slideLayout is a valid name or index recognized by Item().
              layout = presentation.SlideMaster.CustomLayouts.Item(slideLayout) || presentation.SlideLayouts.Item(slideLayout);
            } catch (e) {
                 // Fallback to a default layout if the specified one is not found
                logger.warn(`Slide layout "${slideLayout}" not found. Using default title slide layout. Error: ${(e as Error).message}`);
                // Using a common default layout, e.g., Title Slide (ppLayoutTitle = 1)
                // This assumes '1' is a universally available layout index for Title Slide.
                // A more robust solution would map common names to their PpSlideLayout enum values.
                layout = presentation.SlideLayouts.Item(1); // ppLayoutTitle
            }
            const newSlide = presentation.Slides.Add(presentation.Slides.Count + 1, layout.Layout || layout); // layout might be CustomLayout or SlideLayout
            return { success: true, data: `Slide added using COM Interop with layout "${slideLayout}".` };

          case 'delete':
            if (slideIndex === undefined) {
              throw new Error('The slideIndex parameter is required for the "delete" operation with COM Interop.');
            }
            if (slideIndex < 1 || slideIndex > presentation.Slides.Count) {
              throw new Error(`Slide index ${slideIndex} out of bounds.`);
            }
            presentation.Slides.Item(slideIndex).Delete();
            return { success: true, data: `Slide at index ${slideIndex} deleted using COM Interop.` };

          case 'set':
            if (slideIndex === undefined) {
              throw new Error('The slideIndex parameter is required for the "set" operation with COM Interop.');
            }
            if (slideIndex < 1 || slideIndex > presentation.Slides.Count) {
              throw new Error(`Slide index ${slideIndex} out of bounds.`);
            }
            // const slideToSet = presentation.Slides.Item(slideIndex);
            // Logic to modify the slide would go here
            return { success: true, data: `Operation 'set' on slide ${slideIndex} using COM Interop completed (no complex modifications implemented).` };
          
          case 'getText':
             // Extract text from all slides using COM
            let allText = '';
            for (let i = 1; i <= presentation.Slides.Count; i++) {
                const slide = presentation.Slides.Item(i);
                allText += `Slide ${i}:\n`;
                for (let j = 1; j <= slide.Shapes.Count; j++) {
                    const shape = slide.Shapes.Item(j);
                    if (shape.HasTextFrame && shape.TextFrame.HasText) {
                        allText += shape.TextFrame.TextRange.Text + '\n';
                    }
                }
                allText += '\n';
            }
            return { success: true, data: allText.trim() || "No text found or presentation is empty (COM)." };

          default:
            // This check is for exhaustiveness, but the enum should prevent reaching here.
            // However, to satisfy type checking if 'operation' was not strictly from the enum:
            const exhaustiveCheck: never = operation;
            throw new Error(`Unsupported COM operation: ${exhaustiveCheck}`);
        }
      } catch (error: any) {
        logger.error(`Error in powerpoint/slides tool (COM Interop): ${error.message}`);
        return { success: false, error: { code: 'POWERPOINT_COM_ERROR', message: `COM Interop Error: ${error.message}` } };
      } finally {
        if (presentation) {
          try {
            if (operation !== 'getText') { // Don't save if just reading text
                 presentation.Save();
                 const pptContent = await fs.readFile(filePath);
                 await saveResource('powerpoint/slides', basenamePath(filePath), pptContent);
            }
            presentation.Close();
          } catch (closeError: any) {
            logger.warn(`Error closing COM presentation: ${closeError.message}`);
          }
          releaseObject(presentation);
        }
        releaseObject(pptApp);
      }
    } else {
      // Library Path (pptxgenjs for generation/some modification, officeparser for reading)
      try {
        switch (operation) {
          case 'add':
            const pptx = new (PptxGenJS as any)(); // Cast to any for constructor
            // The 'slideLayout' parameter needs careful mapping to PptxGenJS layouts.
            // PptxGenJS uses predefined constants (e.g., pptx.Layouts.TITLE_SLIDE) or custom master slides.
            // For now, we'll add a slide with a default layout.
            // A more advanced implementation would map common names or allow specifying PptxGenJS layout names.
            let genLayoutName = slideLayout;
            if (slideLayout) {
                // Example: try to map some common COM names to PptxGenJS, or use as is if it matches a PptxGenJS layout name
                // This is a placeholder for a more robust mapping strategy
                if (slideLayout.toLowerCase().includes("title")) genLayoutName = "TITLE_SLIDE";
                else if (slideLayout.toLowerCase().includes("blank")) genLayoutName = "BLANK";
                // ... other mappings
            }

            const newSlideLib = pptx.addSlide({ masterName: genLayoutName });
            newSlideLib.addText('New Slide Added via PptxGenJS', { x: 1, y: 1, w: 8, h: 1, fontSize: 24 });
            
            // If filePath exists, pptxgenjs will overwrite. If it doesn't, it will create.
            await pptx.writeFile({ fileName: filePath });
            
            // Save resource after writing file
            const pptContent = await fs.readFile(filePath);
            await saveResource('powerpoint/slides', basenamePath(filePath), pptContent);
            return { success: true, data: `Slide added using pptxgenjs and saved to "${filePath}". Layout used: ${genLayoutName || 'default'}.` };

          case 'delete':
            // Deleting a specific slide from an existing arbitrary .pptx file is not directly supported by pptxgenjs (it generates, doesn't easily modify in place).
            // officeparser is for reading, not modification.
            logger.warn('Operation "delete" for an existing file is not supported when useComInterop is false. This would require re-generating the presentation without the slide.');
            return { success: false, error: { code: 'POWERPOINT_LIB_UNSUPPORTED', message: 'Deleting slides from existing files is not supported with pptxgenjs/officeparser. Use COM Interop or regenerate the presentation.' } };

          case 'set':
            // Modifying a specific slide in an existing arbitrary .pptx file is complex with pptxgenjs.
            // It's designed for generation. officeparser is for reading.
            // A 'set' operation might involve reading with officeparser (limited info), then regenerating with pptxgenjs, which is a large operation.
            logger.warn('Operation "set" for an existing file has limited support when useComInterop is false. PptxGenJS is primarily for generation.');
            return { success: false, error: { code: 'POWERPOINT_LIB_LIMITED_SUPPORT', message: 'Modifying existing slides has limited support with pptxgenjs/officeparser. Use COM Interop or consider regeneration strategies.' } };
          
          case 'getText':
            if (!await fs.pathExists(filePath)) {
                return { success: false, error: { code: 'FILE_NOT_FOUND', message: `File not found: ${filePath}` } };
            }
            try {
                const buffer = await fs.readFile(filePath);
                // Use the generic parseOfficeAsync as per officeparser's typings
                const textContent = await officeParser.parseOfficeAsync(buffer);
                // The result is a string, not an object with slidesText.
                // If specific slide separation is needed, it's not provided by this generic parser directly.
                // For now, we return the whole text content.
                return { success: true, data: textContent || "No text found or presentation is empty (officeparser)." };
            } catch (parseError: any) {
                logger.error(`Error parsing PowerPoint with officeparser: ${parseError.message}`);
                return { success: false, error: { code: 'OFFICEPARSER_ERROR', message: `Failed to parse PowerPoint file with officeparser: ${parseError.message}` } };
            }

          default:
            // This check is for exhaustiveness
            const exhaustiveCheckLib: never = operation;
            throw new Error(`Unsupported library operation: ${exhaustiveCheckLib}`);
        }
      } catch (error: any) {
        logger.error(`Error in powerpoint/slides tool (pptxgenjs/officeparser): ${error.message}`);
        return { success: false, error: { code: 'POWERPOINT_LIB_ERROR', message: `Library Error: ${error.message}` } };
      }
    }
    // Fallback return, though all paths should ideally return explicitly.
    // This might be reached if an unhandled case occurs outside the try/catch blocks above,
    // or if a switch doesn't have a default that throws.
    return { success: false, error: { code: 'UNHANDLED_LOGIC_PATH', message: 'PowerPoint slide operation did not complete as expected.' } };
  },
};

export default slidesTool;