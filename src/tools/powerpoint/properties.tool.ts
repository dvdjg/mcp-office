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
import { McpResource, ApiResponse, FastMCPContext as Context, ToolRequestParams } from '../../types/common.types';
import { getOfficeApplication } from '../../utils/officeInterop';
import { handleToolError, createErrorResponse } from '../../utils/errorHandler';

// Input schema for the powerpoint/properties tool
const PowerPointPropertiesInputSchema = z.object({
  filePath: z.string().describe('The path to the PowerPoint presentation file.'),
  operation: z.enum(['set', 'configure', 'add']).describe('The operation to perform: set, configure, or add.'),
  // Common properties for 'set' and 'add'
  propertyName: z.string().optional().describe('The name of the property to set or add.'),
  propertyValue: z.any().optional().describe('The value of the property to set or add.'),
  // Specific properties for 'configure'
  size: z.enum(['16:9', '4:3']).optional().describe('The slide size (e.g., "16:9", "4:3").'),
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
  try {
    // Validate input parameters
    const input = PowerPointPropertiesInputSchema.parse(params);
    const { filePath, operation, propertyName, propertyValue, size } = input;

    let app;

    try {
      app = await getOfficeApplication('PowerPoint.Application');
      const presentation = app.Presentations.Open(filePath);

      switch (operation) {
        case 'set':
          if (!propertyName) {
            return createErrorResponse('VALIDATION_ERROR', 'propertyName is required for the set operation.');
          }
          // Implementation for setting built-in properties
          // This will require mapping propertyName to the correct COM property
          // Basic example (may need refinement):
          if (propertyName === 'title') {
             presentation.BuiltinDocumentProperties('Title').Value = propertyValue;
          } else if (propertyName === 'author') {
             presentation.BuiltinDocumentProperties('Author').Value = propertyValue;
          }
          // Add more cases based on common built-in properties
          break;

        case 'configure':
          // Implementation for configuring presentation settings
          if (size) {
            // Configure slide size
            // ppSlideSizeOnScreen (16:9) = 1, ppSlideSizeOnScreen (4:3) = 2
            if (size === '16:9') {
              presentation.PageSetup.SlideSize = 1;
            } else if (size === '4:3') {
              presentation.PageSetup.SlideSize = 2;
            }
          }
          // Add more configuration options here
          break;

        case 'add':
          if (!propertyName || propertyValue === undefined) {
            return createErrorResponse('VALIDATION_ERROR', 'propertyName and propertyValue are required for the add operation.');
          }
          // Implementation for adding or modifying custom properties
          // Check if the property already exists before adding
          let customProp;
          try {
              customProp = presentation.CustomDocumentProperties(propertyName);
              customProp.Value = propertyValue;
          } catch (e) {
              // Property does not exist, add it
              presentation.CustomDocumentProperties.Add(propertyName, false, 4, propertyValue); // msoPropertyTypeVariant = 4
          }
          break;

        default:
          return createErrorResponse('VALIDATION_ERROR', `Unsupported operation: ${operation}`);
      }

      presentation.Save();
      presentation.Close();

      return { success: true, data: `Operation '${operation}' completed for file '${filePath}'.` };

    } catch (officeError: any) {
      // Handle specific Office/COM errors
      if (app) {
          try {
              // Attempt to close the presentation if open to avoid blocking
              const presentation = app.Presentations.Open(filePath);
              presentation.Close();
          } catch (closeError) {
              // Ignore errors when closing if the presentation was already closed or could not be opened
          }
      }
      return handleToolError(officeError, 'POWERPOINT_PROPERTIES_ERROR');
    }

  } catch (validationError: any) {
    // Handle Zod schema validation errors
    return handleToolError(validationError, 'VALIDATION_ERROR');
  }
};

export const powerpointPropertiesTool: McpResource[] = [{
  path: 'powerpoint/properties',
  description: 'Manages the properties of a PowerPoint presentation.',
  schema: PowerPointPropertiesInputSchema,
  handler,
}];