/**
 * @file Tool for managing the properties of a PowerPoint presentation.
 * Allows setting built-in properties, configuring settings like slide size,
 * and adding or modifying custom properties.
 * Uses COM Interop via winax to interact with PowerPoint.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import PptxGenJS from 'pptxgenjs';
import { McpResource, ApiResponse, FastMCPContext as Context, ToolRequestParams } from '../../types/common.types';
import { getOfficeApplication, releaseObject } from '../../utils/officeInterop';
import { handleToolError, createErrorResponse as createErrorResponseUtil } from '../../utils/errorHandler'; // Renamed to avoid conflict
import logger from '../../utils/logger'; // Import logger
import * as fs from 'fs-extra'; // Import fs for file operations
import * as path from 'path'; // Import path for resolving
import { saveResource } from '../dynamic/resources.tool'; // Import saveResource

// Helper to create standard error responses (if not already defined or imported from a shared util)
const createErrorResponse = (code: string, message: string, details?: unknown): ApiResponse<never> => ({
    success: false,
    error: { code, message, details },
});


// Input schema for the powerpoint/properties tool
const PowerPointPropertiesInputSchema = z.object({
  filePath: z.string().describe('The path to the PowerPoint presentation file.'),
  operation: z.enum(['set', 'configure', 'add', 'get']).describe('The operation to perform: set (built-in), configure (layout), add (custom for COM, limited for PptxGenJS), get (built-in).'),
  // Common properties for 'set', 'add', 'get'
  propertyName: z.string().optional().describe('The name of the property to set, add, or get (e.g., "title", "author", "customPropName").'),
  propertyValue: z.any().optional().describe('The value of the property to set or add.'),
  // Specific properties for 'configure'
  size: z.enum(['16:9', '4:3', '16x9', '4x3', 'LAYOUT_16x9', 'LAYOUT_4x3', 'LAYOUT_WIDE', 'LAYOUT_16x10']).optional().describe('The slide size/layout (e.g., "16:9", "LAYOUT_WIDE").'),
  useComInterop: z.boolean().optional().default(false).describe('Set to true to use COM Interop for operations, otherwise uses pptxgenjs (with limitations).'),
  // You can add more configuration parameters here as needed
});

type PowerPointPropertiesInput = z.infer<typeof PowerPointPropertiesInputSchema>;

/**
 * @tool powerpoint/properties
 * @description Manages the properties of a PowerPoint presentation.
 * Allows setting built-in properties, configuring settings like slide size,
 * and adding or modifying custom properties.
 * Uses COM Interop via winax to interact with PowerPoint.
 * @param {ToolRequestParams} params - The input parameters for the tool.
 * @param {Context<any>} [context] - The FastMCP context.
 * @returns {Promise<ApiResponse<any>>} An ApiResponse object indicating the result.
 */
const handler = async (params: ToolRequestParams, context?: Context<any>): Promise<ApiResponse<any>> => {
  const input = PowerPointPropertiesInputSchema.parse(params);
  const { filePath, operation, propertyName, propertyValue, size, useComInterop } = input;
  const absoluteFilePath = path.resolve(filePath);

  if (useComInterop) {
    // COM Interop Path
    let app: any = null;
    let presentation: any = null;
    try {
      app = await getOfficeApplication('PowerPoint.Application');
      if (!await fs.pathExists(absoluteFilePath) && (operation !== 'set' && operation !== 'add' && operation !== 'configure')) {
          // For get, file must exist. For others, COM might create, but let's be explicit.
          // PptxGenJS path will create if not exists for set/configure.
          return createErrorResponse('FILE_NOT_FOUND', `COM: File not found: ${absoluteFilePath}`);
      }
      
      try {
        presentation = app.Presentations.Open(absoluteFilePath);
      } catch (e) {
        if (operation === 'set' || operation === 'add' || operation === 'configure') {
            logger.info(`COM: File ${absoluteFilePath} not found, creating new presentation.`);
            presentation = app.Presentations.Add();
            presentation.SaveAs(absoluteFilePath); // Save immediately
        } else {
            throw e; // Re-throw if not a creation operation
        }
      }


      switch (operation) {
        case 'set':
          if (!propertyName) return createErrorResponse('VALIDATION_ERROR', 'COM: propertyName is required for set.');
          // Map to common BuiltinDocumentProperties
          const propMap: { [key: string]: string } = {
            title: 'Title', author: 'Author', subject: 'Subject', company: 'Company',
            keywords: 'Keywords', comments: 'Comments', category: 'Category', manager: 'Manager',
            revision: 'Revision number' // Note: COM expects string for revision number
          };
          if (propMap[propertyName.toLowerCase()]) {
            presentation.BuiltinDocumentProperties(propMap[propertyName.toLowerCase()]).Value = String(propertyValue);
          } else {
            return createErrorResponse('INVALID_PARAM', `COM: Built-in property "${propertyName}" not directly supported or recognized. Use 'add' for custom properties.`);
          }
          break;

        case 'configure':
          if (size) {
            const ppSlideSize: { [key: string]: number } = {
                '16:9': 1, 'LAYOUT_16x9': 1, '16x9': 1,
                '4:3': 2, 'LAYOUT_4x3': 2, '4x3': 2,
                'LAYOUT_WIDE': 1, // Typically 16:9
                'LAYOUT_16x10': 10, // Example, check actual PpSlideSizeType constant
            };
            if (ppSlideSize[size] !== undefined) {
              presentation.PageSetup.SlideSize = ppSlideSize[size];
            } else {
              return createErrorResponse('INVALID_PARAM', `COM: Unsupported slide size/layout: ${size}.`);
            }
          }
          // Add more configuration options here
          break;

        case 'add': // Add/Update Custom Property
          if (!propertyName || propertyValue === undefined) {
            return createErrorResponse('VALIDATION_ERROR', 'COM: propertyName and propertyValue are required for add.');
          }
          let customProp;
          try {
            customProp = presentation.CustomDocumentProperties(propertyName);
            customProp.Value = propertyValue; // Update if exists
          } catch (e) {
            // msoPropertyTypeString = 4, msoPropertyTypeNumber = 3, msoPropertyTypeDate = 5, msoPropertyTypeBoolean = 2
            let propType = 4; // Default to string (msoPropertyTypeString)
            if (typeof propertyValue === 'number') propType = 3;
            else if (typeof propertyValue === 'boolean') propType = 2;
            else if (propertyValue instanceof Date) propType = 5;
            presentation.CustomDocumentProperties.Add(propertyName, false, propType, propertyValue);
          }
          break;
        case 'get':
            if (!propertyName) return createErrorResponse('VALIDATION_ERROR', 'COM: propertyName is required for get.');
            const getPropMap: { [key: string]: string } = { /* same as set propMap */
                title: 'Title', author: 'Author', subject: 'Subject', company: 'Company',
                keywords: 'Keywords', comments: 'Comments', category: 'Category', manager: 'Manager',
                revision: 'Revision number'
            };
            if (getPropMap[propertyName.toLowerCase()]) {
                const val = presentation.BuiltinDocumentProperties(getPropMap[propertyName.toLowerCase()]).Value;
                return { success: true, data: { propertyName, value: val } };
            } else {
                // Try custom properties
                try {
                    const customVal = presentation.CustomDocumentProperties(propertyName).Value;
                    return { success: true, data: { propertyName, value: customVal, type: 'custom' } };
                } catch (e) {
                    return createErrorResponse('PROPERTY_NOT_FOUND', `COM: Property "${propertyName}" not found as built-in or custom.`);
                }
            }

        default:
          return createErrorResponseUtil('VALIDATION_ERROR', `COM: Unsupported operation: ${operation}`);
      }

      presentation.Save();
      // Save resource for modification operations
      if (['set', 'configure', 'add'].includes(operation)) {
          const pptContent = await fs.readFile(absoluteFilePath, null);
          await saveResource('powerpoint/properties', path.basename(absoluteFilePath), pptContent);
      }
      presentation.Close();
      return { success: true, data: `COM: Operation '${operation}' completed for file '${absoluteFilePath}'.` };

    } catch (officeError: any) {
      logger.error(`Error in powerpoint/properties (COM): ${officeError.message}`, { error: officeError });
      return handleToolError(officeError, 'POWERPOINT_PROPERTIES_COM_ERROR');
    } finally {
      releaseObject(presentation);
      releaseObject(app);
    }
  } else {
    // PptxGenJS Path
    try {
      const pptx = new PptxGenJS();
      let presentationDefinitionModified = false; // Flag to track if metadata/layout was changed

      // PptxGenJS modifies properties of the pptx instance before writing the file.
      // If the file exists, it will be overwritten.
      // Reading existing properties from a file to then modify them with PptxGenJS is not its primary use case.
      // officeparser can read some basic metadata but not all, and not custom props.

      switch (operation) {
        case 'set':
          if (!propertyName || propertyValue === undefined) {
            return createErrorResponse('VALIDATION_ERROR_LIB', 'PptxGenJS: propertyName and propertyValue are required for set.');
          }
          const lowerPropName = propertyName.toLowerCase();
          if (lowerPropName === 'title') pptx.title = String(propertyValue);
          else if (lowerPropName === 'author') pptx.author = String(propertyValue);
          else if (lowerPropName === 'subject') pptx.subject = String(propertyValue);
          else if (lowerPropName === 'company') pptx.company = String(propertyValue);
          else if (lowerPropName === 'revision') pptx.revision = String(propertyValue); // Must be a whole number string
          else {
            return createErrorResponse('INVALID_PARAM_LIB', `PptxGenJS: Setting property "${propertyName}" is not directly supported. Supported: title, author, subject, company, revision.`);
          }
          presentationDefinitionModified = true;
          break;

        case 'configure':
          if (size) {
            const layoutMap: { [key: string]: string } = {
                '16:9': 'LAYOUT_16x9', '16x9': 'LAYOUT_16x9', 'LAYOUT_16x9': 'LAYOUT_16x9',
                '4:3': 'LAYOUT_4x3', '4x3': 'LAYOUT_4x3', 'LAYOUT_4x3': 'LAYOUT_4x3',
                'LAYOUT_WIDE': 'LAYOUT_WIDE',
                'LAYOUT_16x10': 'LAYOUT_16x10'
            };
            const upperSize = size.toUpperCase(); // Ensure consistent casing for map lookup
            if (layoutMap[upperSize]) {
              pptx.layout = layoutMap[upperSize];
            } else {
              return createErrorResponse('INVALID_PARAM_LIB', `PptxGenJS: Unsupported layout: ${size}. Supported: 16:9, 4:3, WIDE, 16x10 (and their LAYOUT_ variants).`);
            }
            presentationDefinitionModified = true;
          }
          // No explicit slide addition here; will be handled before writeFile if definition was modified.
          break;

        case 'add':
          // PptxGenJS does not have a generic "CustomDocumentProperties.Add" like COM.
          // It handles standard metadata properties. We can treat 'add' like 'set' for those.
          if (!propertyName || propertyValue === undefined) {
            return createErrorResponse('VALIDATION_ERROR_LIB', 'PptxGenJS: propertyName and propertyValue are required for add (treated as set).');
          }
          const addPropNameLower = propertyName.toLowerCase();
          if (addPropNameLower === 'title') pptx.title = String(propertyValue);
          else if (addPropNameLower === 'author') pptx.author = String(propertyValue);
          else if (addPropNameLower === 'subject') pptx.subject = String(propertyValue);
          else if (addPropNameLower === 'company') pptx.company = String(propertyValue);
          else if (addPropNameLower === 'revision') pptx.revision = String(propertyValue);
          else {
            logger.warn(`PptxGenJS: 'add' operation for custom property "${propertyName}" is not supported. Only standard metadata can be set.`);
            return createErrorResponse('UNSUPPORTED_OPERATION_LIB', `PptxGenJS: Adding custom property "${propertyName}" is not supported. Use COM or set standard metadata like title, author.`);
          }
          presentationDefinitionModified = true;
          break;
        
        case 'get':
            logger.warn(`PptxGenJS: 'get' operation is not supported for reading properties from an existing file using PptxGenJS. Use COM Interop or officeparser for basic metadata.`);
            return createErrorResponse('UNSUPPORTED_OPERATION_LIB', `PptxGenJS: 'get' operation is not supported. Use COM Interop for reading properties.`);

        default:
          return createErrorResponseUtil('VALIDATION_ERROR', `PptxGenJS: Unsupported operation: ${operation}`);
      }

      // If any presentation-defining properties were changed, or if it was a 'configure' operation,
      // ensure at least one slide exists before writing to create a valid PPTX file.
      if (presentationDefinitionModified) {
        // Since we create a new PptxGenJS instance each time for this library path,
        // we know no slides have been added yet unless this tool's logic were to change.
        // So, we add one default slide to make the presentation valid.
        pptx.addSlide();
        logger.info(`PptxGenJS: Added a default slide as presentation properties/layout were modified.`);
        
        await pptx.writeFile({ fileName: absoluteFilePath });
        const pptContent = await fs.readFile(absoluteFilePath, null);
        await saveResource('powerpoint/properties', path.basename(absoluteFilePath), pptContent);
        return { success: true, data: `PptxGenJS: Operation '${operation}' completed. File saved to '${absoluteFilePath}'.` };
      } else {
        return { success: true, data: `PptxGenJS: Operation '${operation}' did not result in file modification (e.g., 'get' or no valid properties set).` };
      }

    } catch (libError: any) {
      logger.error(`Error in powerpoint/properties (PptxGenJS): ${libError.message}`, { error: libError });
      return handleToolError(libError, 'POWERPOINT_PROPERTIES_LIB_ERROR');
    }
  }
};

export const powerpointPropertiesTool: McpResource[] = [{
  path: 'powerpoint/properties',
  description: 'Manages the properties of a PowerPoint presentation. Allows setting built-in properties, configuring layout, and (with COM) custom properties.',
  schema: PowerPointPropertiesInputSchema,
  handler,
}];