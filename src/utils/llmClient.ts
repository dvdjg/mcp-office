/**
 * @file Provides utility functions for handling and formatting errors.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import logger from './logger'; // Assuming logger is available
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