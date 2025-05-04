/**
 * @file Tool for converting a Word document to a PowerPoint presentation.
 * Allows analyzing the Word document structure, creating a new presentation,
 * and transferring content from the Word document to the presentation, creating slides
 * based on the document structure (e.g., heading styles).
 * Uses COM Interop via winax to interact with Word and PowerPoint.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams, FastMCPContext as Context } from '../../types/common.types';
import logger from '../../utils/logger';
import { getOfficeApplication, openWordDocument, releaseObject } from '../../utils/officeInterop';

// Define the input schema for the word-to-powerpoint tool
const WordToPowerpointInputSchema = z.object({
  wordFilePath: z.string().describe('Path to the source Word file.'),
  powerpointFilePath: z.string().optional().describe('Path where the output PowerPoint presentation will be saved. If not provided, a temporary presentation will be created.'),
  operation: z.enum(['analyze', 'create', 'transfer']).describe('Operation to perform: analyze, create, or transfer.'),
  headingLevelForNewSlide: z.number().int().min(1).max(9).optional().describe('Word heading level that will start a new slide (e.g., 1 for Heading 1).'),
  // You can add more optional parameters here, such as PowerPoint template, style mapping, etc.
});

type WordToPowerpointInput = z.infer<typeof WordToPowerpointInputSchema>;

/**
 * @tool office/word-to-powerpoint
 * @description Converts a Word document to a PowerPoint presentation.
 * Allows analyzing the Word document structure, creating a new presentation,
 * and transferring content from the Word document to the presentation, creating slides
 * based on the document structure (e.g., heading styles).
 * Uses COM Interop via winax to interact with Word and PowerPoint.
 * @param {object} params - Input parameters.
 * @param {string} params.wordFilePath - Path to the source Word file.
 * @param {string} [params.powerpointFilePath] - Path where the output PowerPoint presentation will be saved.
 * @param {'analyze' | 'create' | 'transfer'} params.operation - Operation to perform: analyze, create, or transfer.
 * @param {number} [params.headingLevelForNewSlide] - Word heading level that will start a new slide.
 * @returns {Promise<any>} - Result of the operation.
 */
const wordToPowerpointTool: McpResource = {
  path: 'office/word-to-powerpoint', // Changed from name to path
  description: 'Converts a Word document to a PowerPoint presentation.',
  schema: WordToPowerpointInputSchema, // Changed from inputSchema to schema
  handler: async (params: ToolRequestParams, context?: Context<any>): Promise<ApiResponse<any>> => {
    // Validate and parse input parameters
    const parsedParams = WordToPowerpointInputSchema.safeParse(params);

    if (!parsedParams.success) {
        return {
            success: false,
            error: {
                code: 'VALIDATION_ERROR',
                message: 'Invalid input parameters',
                details: parsedParams.error.errors,
            },
        };
    }

    const { wordFilePath, powerpointFilePath, operation, headingLevelForNewSlide } = parsedParams.data;

    let wordApp: any = null;
    let powerpointApp: any = null;
    let wordDoc: any = null;
    let powerpointPresentation: any = null;


    try {
      wordApp = await getOfficeApplication('Word.Application');
      powerpointApp = await getOfficeApplication('PowerPoint.Application');

      wordDoc = await openWordDocument(wordApp, wordFilePath, true, false); // Open Word read-only and invisible

      let result: any = null; // Declare result here and initialize it

      switch (operation) {
        case 'analyze':
          // Implementation of the analyze operation
          // Could analyze heading styles, sections, etc.
          logger.info(`Analyzing Word document: ${wordFilePath}`);
          // Basic example: count paragraphs and headings
          const paragraphCount = wordDoc.Paragraphs.Count;
          let headingCounts: { [key: number]: number } = {};
          for (let i = 1; i <= paragraphCount; i++) {
            const paragraph = wordDoc.Paragraphs.Item(i);
            // Simplification: check if the paragraph style is a heading
            // In a real implementation, a more robust style mapping would be needed
            if (paragraph.Style && typeof paragraph.Style.NameLocal === 'string' && paragraph.Style.NameLocal.startsWith('Heading ')) {
                const level = parseInt(paragraph.Style.NameLocal.replace('Heading ', ''), 10);
                if (!isNaN(level)) {
                    headingCounts[level] = (headingCounts[level] || 0) + 1;
                }
            }
          }
          result = {
            message: `Analysis completed for ${wordFilePath}`,
            paragraphCount,
            headingCounts,
            // Add more analysis details here
          };
          break;

        case 'create':
          // Implementation of the create operation
          logger.info('Creating new PowerPoint presentation');
          const presentation = powerpointApp.Presentations.Add();
          if (powerpointFilePath) {
            presentation.SaveAs(powerpointFilePath);
            result = { message: `Presentation created and saved to ${powerpointFilePath}` };
          } else {
            // If no path is specified, leave it open for the next operation
            result = { message: 'New presentation created.' };
          }
          // Do not close the presentation here if it will be used in 'transfer'
          break;

        case 'transfer':
          // Implementation of the transfer operation
          logger.info(`Transferring content from ${wordFilePath} to presentation`);

          let targetPresentation: any;
          if (powerpointFilePath) {
              // Open existing presentation if path is specified
              try {
                  targetPresentation = powerpointApp.Presentations.Open(powerpointFilePath);
              } catch (openError) {
                  logger.warn(`No presentation found at ${powerpointFilePath}. Creating a new one.`);
                  targetPresentation = powerpointApp.Presentations.Add();
              }
          } else {
              // Use the active presentation if no path is specified
              if (powerpointApp.Presentations.Count === 0) {
                  logger.info('No presentations open. Creating a new one.');
                  targetPresentation = powerpointApp.Presentations.Add();
              } else {
                  targetPresentation = powerpointApp.ActivePresentation;
              }
          }

          // Transfer logic: iterate over the Word document
          // and create slides based on headingLevelForNewSlide
          let currentSlide: any = null;
          const paragraphs = wordDoc.Paragraphs;

          for (let i = 1; i <= paragraphs.Count; i++) {
              const paragraph = paragraphs.Item(i);
              const styleName = paragraph.Style ? paragraph.Style.NameLocal : '';
              const isHeading = styleName && typeof styleName === 'string' && styleName.startsWith('Heading ');
              let headingLevel: number | null = null;
              if (isHeading) {
                  headingLevel = parseInt(styleName.replace('Heading ', ''), 10);
              }

              const shouldStartNewSlide = headingLevelForNewSlide !== undefined && headingLevel !== null && headingLevel <= headingLevelForNewSlide;

              if (shouldStartNewSlide || currentSlide === null) {
                  // Add new slide. Use a basic layout for now.
                  // pptLayoutTitleOnly = 11, pptLayoutText = 2, pptLayoutTitleAndContent = 1
                  const layout = shouldStartNewSlide && headingLevel === 1 ? 11 : 2; // Title only for H1, Text for others
                  currentSlide = targetPresentation.Slides.Add(targetPresentation.Slides.Count + 1, layout);

                  if (shouldStartNewSlide && headingLevel !== null) {
                      // If it's a heading that starts a new slide, use it as the title
                      if (currentSlide.Shapes.HasTitle) {
                           currentSlide.Shapes.Title.TextFrame.TextRange.Text = paragraph.Range.Text.trim();
                      }
                      // If there is text after the title in the same paragraph, add it to the body
                      const remainingText = paragraph.Range.Text.trim().substring(paragraph.Range.Text.trim().indexOf('\n') + 1);
                       if (remainingText && currentSlide.Shapes.Placeholders.Count > 1) {
                           currentSlide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = remainingText;
                       }
                   } else {
                       // If it's not a heading that starts a new slide or it's the first paragraph, add it to the body
                       if (currentSlide.Shapes.Placeholders.Count > 1) {
                           currentSlide.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = paragraph.Range.Text.trim();
                       }
                   }

               } else {
                   // Add content to the current slide
                   if (currentSlide && currentSlide.Shapes.Placeholders.Count > 1) {
                       const bodyShape = currentSlide.Shapes.Placeholders.Item(2);
                       bodyShape.TextFrame.TextRange.InsertAfter(paragraph.Range.Text + '\n');
                   }
               }
           }

           if (powerpointFilePath) {
               targetPresentation.SaveAs(powerpointFilePath);
               result = { message: `Content transferred and presentation saved to ${powerpointFilePath}` };
           } else {
               result = { message: 'Content transferred to the active presentation.' };
           }

           // Do not close the presentation here if it was saved or left active
           break;

         default:
           throw new Error(`Unsupported operation: ${operation}`);
       }

       // Close the Word document
       wordDoc.Close();

       // Do not close Office applications here, OfficeInterop manages them


       return { success: true, data: result };

     } catch (error: any) {
       logger.error(`Error in wordToPowerpointTool: ${error.message}`);
       // Ensure Word document is closed if open
       if (wordDoc) {
           try {
               wordDoc.Close(0); // wdDoNotSaveChanges = 0
               releaseObject(wordDoc);
           } catch (closeError) {
               logger.error(`Error closing Word document in catch: ${closeError}`);
           }
       }
        // Ensure PowerPoint presentation is closed if open
        if (powerpointPresentation) {
             try {
                 powerpointPresentation.Close();
                 releaseObject(powerpointPresentation);
             } catch (closeError) {
                 logger.error(`Error closing PowerPoint presentation in catch: ${closeError}`);
             }
         }

        // Do not close Office applications here, OfficeInterop manages them
       return { success: false, error: { code: 'TOOL_ERROR', message: `Word to PowerPoint tool failed: ${error.message}` } };
     } finally {
         // Ensure Word document is closed if open
         if (wordDoc) {
             try {
                 wordDoc.Close(0); // wdDoNotSaveChanges = 0
                 releaseObject(wordDoc);
             } catch (closeError) {
                 logger.error(`Error closing Word document in finally: ${closeError}`);
             }
         }

          // Ensure PowerPoint presentation is closed if created in this handler and not saved
          if (powerpointPresentation && !powerpointFilePath) {
              try {
                  powerpointPresentation.Close();
                  releaseObject(powerpointPresentation);
              } catch (closeError) {
                  logger.error(`Error closing PowerPoint presentation in finally: ${closeError}`);
              }
          }
     }
   },
  };

export default wordToPowerpointTool;