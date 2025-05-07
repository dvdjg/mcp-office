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
import PptxGenJS from 'pptxgenjs';
import { McpResource, ApiResponse, ToolRequestParams } from '../../types/common.types.js';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop.js'; // Added releaseObject
import logger from '../../utils/logger.js'; // Import logger
import fs from 'fs-extra'; // Import fs for file existence check
import { resolve as resolvePath, basename as basenamePath } from 'path'; // Import path for resolving
import { saveResource } from '../dynamic/resources.tool.js'; // Import saveResource

// Helper to create standard error responses
const createErrorResponse = (message: string, code = 'TOOL_EXECUTION_ERROR', details?: unknown): ApiResponse<never> => ({
    success: false,
    error: { code, message, details },
});


// Input schema for the powerpoint/animations tool
const AnimationsInputSchema = z.object({
  filePath: z.string().describe('Path to the PowerPoint file.'),
  operation: z.enum(['add', 'configure', 'remove', 'list']).describe('Operation to perform: add (animation to shape), configure (transition to slide), remove (animation/transition), list (animations/transitions).'),
  slideIndex: z.number().int().positive().optional().describe('1-based index of the slide. Required for configure (COM), remove (COM), and list (COM). For pptxgenjs, operations are typically on newly created slides/objects.'),
  shapeIndex: z.number().int().positive().optional().describe('1-based index of the shape. Required for add (COM), remove (COM), and list (COM animations).'),
  shapeName: z.string().optional().describe('Name of the shape. Alternative to shapeIndex for COM operations.'),
  animationType: z.string().optional().describe('Animation type (MsoAnimEffect constant for COM, or PptxGenJS animation type string e.g., "fadeIn"). Required for add.'),
  transitionType: z.string().optional().describe('Transition type (PpTransition constant for COM, or PptxGenJS transition type string e.g., "fade"). Required for configure.'),
  duration: z.number().positive().optional().describe('Duration in seconds. Optional for add and configure.'),
  effectParameters: z.record(z.any()).optional().describe('Additional parameters for the animation/transition (e.g., direction, order for COM; PptxGenJS uses specific options like `delay`, `direction`).'),
  useComInterop: z.boolean().optional().default(false).describe('Set to true to use COM Interop for operations, otherwise uses pptxgenjs (with significant limitations for animations on existing files).'),
  // For pptxgenjs 'add' operation, we might need more specific inputs if not operating on an existing shape by index/name
  newObjectText: z.string().optional().describe('Text for a new shape/textbox to which animation will be added (pptxgenjs path).'),
  newObjectOptions: z.any().optional().describe('Options for creating a new shape/textbox (pptxgenjs path, e.g., x, y, w, h).'),

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
    const input = AnimationsInputSchema.parse(params);
    const {
        filePath, operation, slideIndex, shapeIndex, shapeName,
        animationType, transitionType, duration, effectParameters, useComInterop,
        newObjectText, newObjectOptions
    } = input;

    const absoluteFilePath = resolvePath(filePath);

    if (useComInterop) {
      // COM Interop Path
      let app: any = null;
      let presentation: any = null;
      let slideForCom: any = null; // Specific to COM path to avoid conflicts
      let shapeForCom: any = null; // Specific to COM path

      try {
        app = await getOfficeApplication('PowerPoint.Application');
        if (!await fs.pathExists(absoluteFilePath)) {
            // For COM, file must exist for most operations, or be creatable for 'add'/'configure' if we extend that.
            // Current COM logic assumes file exists.
            return createErrorResponse(`COM: File not found: ${absoluteFilePath}`, 'FILE_NOT_FOUND');
        }
        presentation = app.Presentations.Open(absoluteFilePath);

        if (slideIndex !== undefined) {
            if (slideIndex < 1 || slideIndex > presentation.Slides.Count) {
                return createErrorResponse(`COM: Slide index ${slideIndex} out of bounds.`, 'INVALID_PARAM');
            }
            slideForCom = presentation.Slides(slideIndex);
        } else if (operation !== 'list') { // List can operate on whole presentation if no slideIndex
             // For other operations, slideIndex is typically needed.
        }


        if (slideForCom && (shapeIndex !== undefined || shapeName !== undefined)) {
            try {
                shapeForCom = shapeIndex ? slideForCom.Shapes(shapeIndex) : slideForCom.Shapes(shapeName);
            } catch (e) {
                return createErrorResponse(`COM: Shape not found on slide ${slideIndex} (Index: ${shapeIndex}, Name: ${shapeName}).`, 'SHAPE_NOT_FOUND');
            }
        }

        switch (operation) {
          case 'add':
            if (!slideForCom || !shapeForCom || !animationType) {
              return createErrorResponse('COM: For "add" operation, slideIndex, shapeIndex/shapeName, and animationType are required.', 'MISSING_PARAM');
            }
            const animEffectValue = app.MsoAnimEffect[animationType];
            if (animEffectValue === undefined) {
                return createErrorResponse(`COM: Invalid animation type: ${animationType}.`, 'INVALID_PARAM');
            }
            const effect = slideForCom.TimeLine.MainSequence.AddEffect(shapeForCom, animEffectValue);
            if (duration !== undefined) effect.Timing.Duration = duration;
            // effectParameters handling for COM would be complex and specific to MsoAnimEffect
            if (effectParameters) logger.warn('COM: effectParameters for animations are not fully implemented.');
            return { success: true, data: { message: `COM: Animation '${animationType}' added to shape on slide ${slideIndex}.` } };

          case 'configure':
            if (!slideForCom || !transitionType) {
              return createErrorResponse('COM: For "configure" operation, slideIndex and transitionType are required.', 'MISSING_PARAM');
            }
            const transition = slideForCom.SlideShowTransition;
            const transitionEffectValue = app.PpTransitionEffect[transitionType] ?? app.PpEntryEffect[transitionType]; // Try both common enum names
            if (transitionEffectValue === undefined) {
                return createErrorResponse(`COM: Invalid transition type: ${transitionType}. Check PpTransitionEffect or PpEntryEffect constants.`, 'INVALID_PARAM');
            }
            transition.EntryEffect = transitionEffectValue;
            if (duration !== undefined) transition.Duration = duration;
            if (effectParameters) logger.warn('COM: effectParameters for transitions are not fully implemented.');
            return { success: true, data: { message: `COM: Transition '${transitionType}' configured for slide ${slideIndex}.` } };

          case 'remove':
            if (!slideForCom || !shapeForCom) {
                 return createErrorResponse('COM: For "remove" animation, slideIndex and shapeIndex/shapeName are required.', 'MISSING_PARAM');
            }
            const effectsToRemove = [];
            for (let i = slideForCom.TimeLine.MainSequence.Count; i >= 1; i--) {
                const currentEffect = slideForCom.TimeLine.MainSequence(i);
                if (currentEffect.Shape && currentEffect.Shape.Name === shapeForCom.Name) { // Or compare by ID if available and more robust
                    effectsToRemove.push(currentEffect);
                }
            }
            effectsToRemove.forEach(eff => eff.Delete());
            // To remove slide transition, one might set it to 'None'
            // slideForCom.SlideShowTransition.EntryEffect = app.PpTransitionEffect.ppEffectNone; (Example)
            return { success: true, data: { message: `COM: Animations removed for shape on slide ${slideIndex}. Transition removal needs specific handling.` } };

          case 'list':
            if (!slideForCom) {
                return createErrorResponse('COM: For "list" operation, slideIndex is required.', 'MISSING_PARAM');
            }
            const animationsList = [];
            for (let i = 1; i <= slideForCom.TimeLine.MainSequence.Count; i++) {
                const listEffect = slideForCom.TimeLine.MainSequence(i);
                animationsList.push({
                    shapeName: listEffect.Shape ? listEffect.Shape.Name : 'Unknown Shape',
                    animationType: listEffect.EffectType, // Numeric, needs mapping to string
                    duration: listEffect.Timing.Duration,
                });
            }
            const listTransition = slideForCom.SlideShowTransition;
            const transitionDetails = {
                transitionType: listTransition.EntryEffect, // Numeric, needs mapping
                duration: listTransition.Duration,
            };
            return { success: true, data: { animations: animationsList, transition: transitionDetails, message: `COM: Listed animations and transition for slide ${slideIndex}.` } };

          default:
            throw new Error(`COM: Unsupported operation: ${operation}`);
        }
      } catch (error: any) {
        logger.error(`Error in powerpoint/animations tool (COM): ${error.message}`, { error });
        return createErrorResponse(`COM: Animations operation failed: ${error.message}`, 'POWERPOINT_ANIMATIONS_COM_ERROR', error);
      } finally {
        if (presentation) {
          try {
            if (operation === 'add' || operation === 'configure' || operation === 'remove') {
                presentation.Save();
                const pptContent = await fs.readFile(absoluteFilePath);
                await saveResource('powerpoint/animations', basenamePath(absoluteFilePath), pptContent);
            }
            presentation.Close();
          } catch (saveCloseError: any) {
            logger.warn(`COM: Error saving/closing presentation: ${saveCloseError.message}`);
          }
          releaseObject(presentation);
        }
        releaseObject(shapeForCom);
        releaseObject(slideForCom);
        releaseObject(app);
      }
    } else {
      // Library Path (PptxGenJS) - Very limited for animations on existing files.
      // PptxGenJS applies animations during object/slide creation.
      logger.info(`PptxGenJS path for animations: Operation '${operation}'. Note: PptxGenJS is best for adding animations during new presentation/object generation.`);

      try {
        const pptx = new (PptxGenJS as any)(); // Cast to any for constructor
        let message = '';

        switch (operation) {
          case 'add':
            // This implies adding a NEW animated object to a NEW slide/presentation.
            // Modifying an existing shape's animation in an existing file is not a direct pptxgenjs feature.
            if (!animationType || !newObjectText) {
                return createErrorResponse('PptxGenJS: For "add" animation, animationType and newObjectText are required to create a new animated object.', 'MISSING_PARAM_LIB');
            }
            const slideLib = pptx.addSlide();
            // Define animation options based on PptxGenJS documentation/examples for TextPropsOptions.animation
            const animationObject: { type: string; duration?: number; delay?: number; [key: string]: any } = { type: animationType as string };
            if (duration) animationObject.duration = duration;
            if (effectParameters?.delay) animationObject.delay = effectParameters.delay;
            if (effectParameters?.direction) animationObject.direction = effectParameters.direction; // Common animation param

            const textOpts: any = { // Use any for TextPropsOptions
                ...(newObjectOptions || { x: 1, y: 1, w: 8, h: 1 }), // Default position/size
                animation: animationObject // Assign the correctly structured animation object
            };
            slideLib.addText(newObjectText, textOpts);
            message = `PptxGenJS: Added new text object with animation '${animationType}' to a new slide. File will be saved to ${absoluteFilePath}.`;
            await pptx.writeFile({ fileName: absoluteFilePath });
            const pptContent = await fs.readFile(absoluteFilePath);
            await saveResource('powerpoint/animations', basenamePath(absoluteFilePath), pptContent);
            return { success: true, data: { message } };

          case 'configure':
            // This implies setting a transition for a NEW slide.
            if (!transitionType) {
                return createErrorResponse('PptxGenJS: For "configure" transition, transitionType is required for a new slide.', 'MISSING_PARAM_LIB');
            }
            // Define transition options based on PptxGenJS documentation/examples for AddSlideProps.transition
            const transitionObject: { type: string; duration?: number; [key: string]: any } = { type: transitionType as string };
            if (duration) transitionObject.duration = duration;
            if (effectParameters?.direction) transitionObject.direction = effectParameters.direction; // Common transition param

            // Add a new slide with the specified transition options.
            // Use 'as any' to bypass strict type checking for AddSlideProps if 'transition' isn't explicitly in its definition,
            // assuming the runtime library handles this common pattern.
            pptx.addSlide({ transition: transitionObject } as any);

            message = `PptxGenJS: Added a new slide with transition '${transitionType}'. File will be saved to ${absoluteFilePath}.`;
            await pptx.writeFile({ fileName: absoluteFilePath });
            const pptContentConf = await fs.readFile(absoluteFilePath);
            await saveResource('powerpoint/animations', basenamePath(absoluteFilePath), pptContentConf);
            return { success: true, data: { message } };

          case 'remove':
          case 'list':
            logger.warn(`PptxGenJS: Operation '${operation}' for animations/transitions on existing files is not supported. Use COM Interop.`);
            return createErrorResponse(`PptxGenJS: Operation '${operation}' for animations/transitions on existing files is not supported. Please use COM Interop.`, 'POWERPOINT_LIB_UNSUPPORTED');

          default:
            const exhaustiveCheckLib: never = operation;
            throw new Error(`PptxGenJS: Unsupported operation: ${exhaustiveCheckLib}`);
        }
      } catch (error: any) {
        logger.error(`Error in powerpoint/animations tool (PptxGenJS): ${error.message}`, { error });
        return createErrorResponse(`PptxGenJS: Animations operation failed: ${error.message}`, 'POWERPOINT_LIB_ERROR', error);
      }
    }
  },
};

export default animationsTool;