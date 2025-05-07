/**
 * @file Provides utility functions for handling and formatting errors.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import logger from './logger.js'; // Assuming logger is available
import axios from 'axios'; // Import axios for making HTTP requests

// Define a basic interface for LLM generation options
interface LlmGenerationOptions {
  maxTokens?: number;
  temperature?: number;
  // Add other relevant options as needed
}

/**
 * Utility function to generate text using a server-side LLM API.
 * This function reads configuration from environment variables
 * and makes a direct API call to the LLM provider.
 *
 * @param prompt The prompt to send to the LLM.
 * @param options Optional generation options (e.g., maxTokens).
 * @returns A promise resolving to the generated text.
 * @throws Error if the LLM API key is not configured or the API call fails.
 */
export async function generateText(prompt: string, options?: LlmGenerationOptions): Promise<string> {
  const apiKey = process.env.LLM_PROVIDER_API_KEY;
  const apiEndpoint = process.env.LLM_PROVIDER_ENDPOINT;
  const modelName = process.env.LLM_MODEL_NAME;

  if (!apiKey || !apiEndpoint) {
    logger.error('LLM_PROVIDER_API_KEY or LLM_PROVIDER_ENDPOINT environment variables are not configured.');
    throw new Error('LLM provider is not configured on the server.');
  }

  logger.info(`Calling server-side LLM with prompt (first 100 chars): "${prompt.substring(0, 100)}..."`);

  try {
    const response = await axios.post(apiEndpoint, {
      model: modelName || 'gemini-2.5-flash-preview-04-17', // Use configured model or a default
      messages: [
        {
          role: "user",
          content: prompt
        }
      ],
      max_tokens: options?.maxTokens,
      temperature: options?.temperature,
      // Add other parameters as needed by the specific API
    }, {
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      }
    });

    // Assuming an OpenAI-like chat completion response structure
    const generatedText = response.data?.choices?.[0]?.message?.content;

    if (!generatedText) {
      logger.warn('LLM API call did not return usable text content.', { response: response.data });
      throw new Error('LLM API call failed to generate text.');
    }

    logger.info('Successfully generated text via server-side LLM.');
    return generatedText;

  } catch (error: any) {
    let errorMessage = `Error during server-side LLM API call: ${error.message}`;
    if (error.response) {
      // The request was made and the server responded with a status code
      // that falls out of the range of 2xx
      errorMessage += ` Status: ${error.response.status}. Data: ${JSON.stringify(error.response.data)}`;
      logger.error(errorMessage, { status: error.response.status, data: error.response.data });
    } else if (error.request) {
      // The request was made but no response was received
      errorMessage += ` No response received.`;
      logger.error(errorMessage, { request: error.request });
    } else {
      // Something happened in setting up the request that triggered an Error
      logger.error(errorMessage, { error });
    }
    throw new Error(`Failed to generate text from LLM: ${errorMessage}`);
  }
}

/**
 * Interface for multimodal content parts (text or image).
 * Based on common patterns like Gemini API.
 */
interface MultimodalContentPart {
  text?: string;
  inline_data?: {
    mime_type: string; // e.g., 'image/png', 'image/jpeg'
    data: string; // base64 encoded image data
  };
  // Or potentially image_url for external images, but focusing on inline data for now
}

/**
 * Interface for a multimodal message.
 */
interface MultimodalMessage {
  role: "user" | "model"; // Or other roles as needed
  parts: MultimodalContentPart[];
}


/**
 * Utility function to analyze an image with a text prompt using a server-side LLM API.
 * Assumes the LLM API supports multimodal input (text + image).
 *
 * @param prompt The text prompt accompanying the image.
 * @param imageBase64 The base64 encoded image data.
 * @param mimeType The MIME type of the image (e.g., 'image/png').
 * @param options Optional generation options.
 * @returns A promise resolving to the generated text analysis.
 * @throws Error if the LLM provider is not configured or the API call fails.
 */
export async function analyzeImageWithText(
    prompt: string,
    imageBase64: string,
    mimeType: string = 'image/png', // Default to PNG
    options?: LlmGenerationOptions
): Promise<string> {
  const apiKey = process.env.LLM_PROVIDER_API_KEY;
  const apiEndpoint = process.env.LLM_PROVIDER_ENDPOINT;
  const modelName = process.env.LLM_MODEL_NAME; // Use the configured model

  if (!apiKey || !apiEndpoint) {
    logger.error('LLM_PROVIDER_API_KEY or LLM_PROVIDER_ENDPOINT environment variables are not configured.');
    throw new Error('LLM provider is not configured on the server.');
  }
  if (!modelName) {
      logger.warn('LLM_MODEL_NAME not configured, multimodal capabilities might be limited or require a specific model.');
      // Potentially throw an error or use a known multimodal default if applicable
      // throw new Error('LLM_MODEL_NAME must be configured for multimodal analysis.');
  }

  logger.info(`Calling server-side LLM for image analysis with prompt (first 100 chars): "${prompt.substring(0, 100)}..."`);

  // Construct multimodal message payload (adjust based on actual API spec, e.g., Gemini)
  const messages: MultimodalMessage[] = [
    {
      role: "user",
      parts: [
        { text: prompt },
        {
          inline_data: {
            mime_type: mimeType,
            data: imageBase64,
          },
        },
      ],
    },
  ];

  try {
    const response = await axios.post(apiEndpoint, {
      // Adjust payload structure based on the specific LLM API (e.g., Gemini uses 'contents')
      // This assumes an OpenAI-like structure for simplicity, might need changes.
      model: modelName, // Use the configured model
      messages: messages, // Send multimodal messages
      max_tokens: options?.maxTokens,
      temperature: options?.temperature,
      // Add other parameters as needed by the specific API
    }, {
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      }
    });

    // Assuming an OpenAI-like chat completion response structure for the text analysis
    const generatedText = response.data?.choices?.[0]?.message?.content;

    if (!generatedText) {
      logger.warn('LLM API call (image analysis) did not return usable text content.', { response: response.data });
      throw new Error('LLM API call failed to generate text analysis.');
    }

    logger.info('Successfully generated text analysis via server-side LLM.');
    return generatedText;

  } catch (error: any) {
    let errorMessage = `Error during server-side LLM API call (image analysis): ${error.message}`;
     if (error.response) {
       errorMessage += ` Status: ${error.response.status}. Data: ${JSON.stringify(error.response.data)}`;
       logger.error(errorMessage, { status: error.response.status, data: error.response.data });
     } else if (error.request) {
       errorMessage += ` No response received.`;
       logger.error(errorMessage, { request: error.request });
     } else {
       logger.error(errorMessage, { error });
     }
    throw new Error(`Failed to analyze image with LLM: ${errorMessage}`);
  }
}


/**
 * Utility function to generate an image using a server-side LLM API.
 * Assumes the LLM API supports image generation and returns base64 data.
 * NOTE: This is a placeholder structure. The actual API endpoint, request payload,
 * and response parsing will depend heavily on the specific LLM provider's
 * image generation capabilities (e.g., DALL-E, Imagen).
 *
 * @param prompt The prompt to generate the image from.
 * @param options Optional generation options (might include size, quality, etc.).
 * @returns A promise resolving to the base64 encoded generated image data.
 * @throws Error if the LLM provider is not configured or the API call fails.
 */
export async function generateImageFromText(
    prompt: string,
    options?: LlmGenerationOptions // Options might need a different interface for image generation
): Promise<string> {
  const apiKey = process.env.LLM_PROVIDER_API_KEY;
  // Image generation might use a different endpoint
  const apiEndpoint = process.env.LLM_IMAGE_GENERATION_ENDPOINT || process.env.LLM_PROVIDER_ENDPOINT;
  const modelName = process.env.LLM_IMAGE_MODEL_NAME || process.env.LLM_MODEL_NAME; // Might use a specific image model

  if (!apiKey || !apiEndpoint) {
    logger.error('LLM API Key/Endpoint for image generation is not configured.');
    throw new Error('LLM provider for image generation is not configured.');
  }

  logger.info(`Calling server-side LLM for image generation with prompt: "${prompt.substring(0, 100)}..."`);

  try {
    // --- Placeholder: Adjust API call based on actual image generation API ---
    // Example assumes an OpenAI DALL-E like structure
    const response = await axios.post(apiEndpoint, {
      model: modelName, // Specify image generation model if needed
      prompt: prompt,
      n: 1, // Number of images to generate
      size: "1024x1024", // Example size
      response_format: "b64_json", // Request base64 encoded image
      // Add other parameters like style, quality etc. based on API
    }, {
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      }
    });
    // --- End Placeholder ---

    // --- Placeholder: Adjust response parsing based on actual API ---
    // Example assumes OpenAI DALL-E like response
    const generatedImageDataBase64 = response.data?.data?.[0]?.b64_json;
    // --- End Placeholder ---

    if (!generatedImageDataBase64) {
      logger.warn('LLM API call (image generation) did not return usable image data.', { response: response.data });
      throw new Error('LLM API call failed to generate image.');
    }

    logger.info('Successfully generated image via server-side LLM.');
    return generatedImageDataBase64;

  } catch (error: any) {
    let errorMessage = `Error during server-side LLM API call (image generation): ${error.message}`;
     if (error.response) {
       errorMessage += ` Status: ${error.response.status}. Data: ${JSON.stringify(error.response.data)}`;
       logger.error(errorMessage, { status: error.response.status, data: error.response.data });
     } else if (error.request) {
       errorMessage += ` No response received.`;
       logger.error(errorMessage, { request: error.request });
     } else {
       logger.error(errorMessage, { error });
     }
    throw new Error(`Failed to generate image from LLM: ${errorMessage}`);
  }
}