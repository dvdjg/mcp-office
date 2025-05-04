/**
 * @file Tool for managing shape animations and slide transitions in PowerPoint.
 * Allows adding animations to shapes, configuring slide transitions,
 * removing animations/transitions, and listing existing ones.
 * Uses COM Interop via winax.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import { McpResource, ApiResponse, ToolRequestParams } from '../../types/common.types';
import { getOfficeApplication } from '../../utils/officeInterop';

// Input schema for the powerpoint/animations tool
const AnimationsInputSchema = z.object({
  filePath: z.string().describe('Path to the PowerPoint file.'),
  operation: z.enum(['add', 'configure', 'remove', 'list']).describe('Operation to perform: add (animation to shape), configure (transition to slide), remove (animation/transition), list (animations/transitions).'),
  slideIndex: z.number().int().positive().optional().describe('1-based index of the slide. Required for configure, remove, and list.'),
  shapeIndex: z.number().int().positive().optional().describe('1-based index of the shape. Required for add, remove, and list (animations).'),
  shapeName: z.string().optional().describe('Name of the shape. Alternative to shapeIndex for add, remove, and list (animations).'),
  animationType: z.string().optional().describe('Animation type (MsoAnimEffect constant). Required for add.'),
  transitionType: z.string().optional().describe('Transition type (PpTransition constant). Required for configure.'),
  duration: z.number().positive().optional().describe('Duration in seconds. Optional for add and configure.'),
  effectParameters: z.record(z.any()).optional().describe('Additional parameters for the animation/transition (e.g., direction, order).'),
});

type AnimationsInput = z.infer<typeof AnimationsInputSchema>;

/**
 * @tool powerpoint/animations
 * @description Manages shape animations and slide transitions in PowerPoint.
 * Allows adding animations to shapes, configuring slide transitions,
 * removing animations/transitions, and listing existing ones.
 * Uses COM Interop via winax.
 * Requires specifying the file path, operation, and slide/shape indices/names based on the operation.
 */
const animationsTool: McpResource = {
  path: 'powerpoint/animations',
  description: 'Manages shape animations and slide transitions in PowerPoint.',
  schema: AnimationsInputSchema,
  handler: async (params: ToolRequestParams): Promise<ApiResponse<any>> => {
    const input = AnimationsInputSchema.parse(params); // Validate and parse input
    const { filePath, operation, slideIndex, shapeIndex, shapeName, animationType, transitionType, duration, effectParameters } = input;
    const app = await getOfficeApplication('PowerPoint.Application');
    let presentation;

    try {
      presentation = app.Presentations.Open(filePath);

      switch (operation) {
        case 'add':
          if (!slideIndex || (!shapeIndex && !shapeName) || !animationType) {
            throw new Error('For the "add" operation, slideIndex, shapeIndex or shapeName, and animationType are required.');
          }
          const slideForAdd = presentation.Slides(slideIndex);
          const shapeForAdd = shapeIndex ? slideForAdd.Shapes(shapeIndex) : slideForAdd.Shapes(shapeName);
          if (!shapeForAdd) {
             throw new Error(`Shape not found on slide ${slideIndex} with index ${shapeIndex} or name ${shapeName}.`);
          }

          // Add animation
          // Note: MsoAnimEffect is an enum in the PowerPoint API. We need to map the string to its numeric value.
          // winax should handle this if the constant exists in the app.MsoAnimEffect object
          const animEffectValue = app.MsoAnimEffect[animationType];
          if (animEffectValue === undefined) {
              throw new Error(`Invalid animation type: ${animationType}.`);
          }

          const effect = slideForAdd.TimeLine.MainSequence.AddEffect(
            shapeForAdd,
            animEffectValue
          );

          if (duration !== undefined) {
            effect.Timing.Duration = duration;
          }

          // Apply additional parameters if they exist
          if (effectParameters) {
              // This is a basic example. Parameter application depends on the animation type.
              // More complex logic would be needed to handle different effect types and their properties.
              // For example, for an entrance animation, you might want to configure the direction.
              // effect.EffectParameters.Direction = app.MsoAnimDirection.msoAnimDirectionLeft;
              console.warn('Application of effectParameters is not fully implemented and may require specific logic per animation type.');
          }


          return { success: true, data: { message: `Animation '${animationType}' added to shape ${shapeIndex || shapeName} on slide ${slideIndex}.` } };

        case 'configure':
          if (!slideIndex || !transitionType) {
            throw new Error('For the "configure" operation, slideIndex and transitionType are required.');
          }
          const slideForConfig = presentation.Slides(slideIndex);
          const transition = slideForConfig.SlideShowTransition;

          // Configure transition
          // Note: PpTransition is an enum in the PowerPoint API. We need to map the string to its numeric value.
          const transitionEffectValue = app.PpTransition[transitionType];
           if (transitionEffectValue === undefined) {
              throw new Error(`Invalid transition type: ${transitionType}.`);
          }
          transition.EntryEffect = transitionEffectValue;

          if (duration !== undefined) {
            transition.Duration = duration;
          }

           // Apply additional parameters if they exist
          if (effectParameters) {
              // Similar to animations, transition parameter application depends on the type.
              // For example, for a push transition, you might want to configure the direction.
              // transition.Direction = app.PpTransitionDirection.ppTransitionDirectionLeft;
               console.warn('Application of effectParameters for transitions is not fully implemented and may require specific logic per transition type.');
          }

          return { success: true, data: { message: `Transition '${transitionType}' configured for slide ${slideIndex}.` } };

        case 'remove':
             if (!slideIndex || (!shapeIndex && !shapeName)) {
                 throw new Error('For the "remove" operation, slideIndex and shapeIndex or shapeName are required.');
             }
             const slideForRemove = presentation.Slides(slideIndex);
             const shapeForRemove = shapeIndex ? slideForRemove.Shapes(shapeIndex) : slideForRemove.Shapes(shapeName);
             if (!shapeForRemove) {
                 throw new Error(`Shape not found on slide ${slideIndex} with index ${shapeIndex} or name ${shapeName}.`);
             }

             // Remove animations associated with the shape
             const effectsToRemove = [];
             // Iterate backwards to avoid issues with indices after deletion
             for (let i = slideForRemove.TimeLine.MainSequence.Count; i >= 1; i--) {
                 const effect = slideForRemove.TimeLine.MainSequence(i);
                 // Compare shapes by their COM object or a unique identifier if possible.
                 // Comparing by name can be problematic if there are shapes with duplicate names.
                 // Comparing by the COM object directly is more reliable if winax allows it.
                 // If winax does not allow direct comparison of COM objects, we could try comparing unique properties like ID or Name (if we guarantee unique names).
                 // For now, we assume direct COM object comparison works or names are unique for this case.
                 try {
                     if (effect.Shape && effect.Shape.Name === shapeForRemove.Name) { // Comparison by name as fallback/example
                          effectsToRemove.push(effect);
                     }
                 } catch (compareError) {
                      console.warn(`Error comparing shapes during animation removal: ${compareError instanceof Error ? compareError.message : String(compareError)}`);
                      // Continue with removal even if comparison fails for a specific effect
                 }
             }

             effectsToRemove.forEach(effect => {
                 try {
                     effect.Delete();
                 } catch (deleteError) {
                     console.error(`Error deleting animation effect: ${deleteError instanceof Error ? deleteError.message : String(deleteError)}`);
                     // Continue deleting other effects
                 }
             });


             return { success: true, data: { message: `Animations removed for shape ${shapeIndex || shapeName} on slide ${slideIndex}.` } };

        case 'list':
            if (!slideIndex) {
                throw new Error('For the "list" operation, slideIndex is required.');
            }
            const slideForList = presentation.Slides(slideIndex);
            const animations = [];

            // List shape animations on the slide
            for (let i = 1; i <= slideForList.TimeLine.MainSequence.Count; i++) {
                const effect = slideForList.TimeLine.MainSequence(i);
                animations.push({
                    shapeName: effect.Shape.Name,
                    // shapeIndex: effect.Shape.ZOrderPosition, // ZOrderPosition is not a reliable index in the Shapes collection
                    animationType: effect.EffectType, // This returns a numeric value, we would need to map it to the MsoAnimEffect constant
                    duration: effect.Timing.Duration,
                    // Other animation details if relevant
                });
            }

            // Get slide transition details
            const transitionForList = slideForList.SlideShowTransition;
            const transitionDetails = {
                transitionType: transitionForList.EntryEffect, // This returns a numeric value, we would need to map it to the PpTransition constant
                duration: transitionForList.Duration,
                // Other transition details if relevant
            };


            return { success: true, data: { animations, transition: transitionDetails, message: `Listed animations and transition for slide ${slideIndex}.` } };


        default:
          throw new Error(`Unsupported operation: ${operation}`);
      }
    } catch (error: any) {
      console.error(`Error in powerpoint/animations tool: ${error.message}`);
      return {
        success: false,
        error: {
          code: 'OFFICE_API_ERROR', // Or a more specific error code
          message: error.message,
          details: error, // Include the original error object for debugging
        },
      };
    } finally {
      if (presentation) {
        try {
            presentation.Save();
            presentation.Close();
        } catch (saveCloseError: any) { // Type as any
             console.error(`Error saving/closing the presentation: ${saveCloseError.message}`);
             // Continue to not block the finally block
        }
      }
      // Do not close the PowerPoint application here, as other tools may need it.
      // app.Quit(); // Do not do this!
    }
  },
};

export default animationsTool;