# Plan for Image Analysis and Generation in Office Documents

## Goal
Enable users to analyze images from Office documents or other sources and insert generated images into Office documents, prioritizing the "master" AI's capabilities and falling back to a connected LLM if necessary.

## Plan

### 1. Image Analysis Workflow:
*   **Step 1: Identify Image Source and Location:** Determine if the image is in a Word document, PDF, dynamic resource, or provided via a new path/URL, and the specified location/description.
*   **Step 2: Access/Extract Image Data:** Obtain the image data based on the source and location. This will involve using appropriate Office MCP tools (`word/image/extract`, `office/pdf/parse`, `dynamic/resources`) or handling external paths/URLs. If the image is in a document and specified by description, this step will involve extracting images and potentially using the LLM for identification (as detailed in the previous plan, but now as a sub-step within the fallback).
*   **Step 3: Attempt Image Analysis with Master AI:** The "master" AI will attempt to analyze the image based on the user's query and the extracted image data.
*   **Step 4: Fallback to Connected LLM (if Master AI cannot analyze):** If the "master" AI cannot perform the analysis, send the image data and query to the configured connected LLM (whichever is available via `llmClient.ts`).
*   **Step 5: Present Analysis Result or Indicate Incapability:** If either the master AI or the connected LLM successfully analyzes the image, present the result to the user. If neither can perform the analysis, inform the user that the task is not possible with the available tools.

### 2. Image Generation and Insertion Workflow:
*   **Step 1: Attempt Image Generation with Master AI:** The "master" AI will attempt to generate an image based on the user's text prompt.
*   **Step 2: Fallback to Connected LLM (if Master AI cannot generate):** If the "master" AI cannot perform the generation, send the text prompt to the configured connected LLM.
*   **Step 3: Receive Generated Image Data:** Obtain the generated image data (base64) from whichever AI successfully performed the generation.
*   **Step 4: Identify Target Document and Location:** Determine the Office document file path and the specific position within the document for insertion.
*   **Step 5: Insert Image into Document:** Use the `msoffice-mcp` tool `word/image/insert` (or similar for other Office apps) with the document path, image data (base64), and position.
*   **Step 6: Confirm Insertion or Indicate Incapability:** If the image is successfully inserted, confirm this to the user. If neither AI could generate the image, inform the user that the task is not possible.

## Workflow Diagram

```mermaid
graph TD
    A[User Request: Analyze/Generate Image] --> B{Image Analysis or Generation?}

    B -- Analysis --> C{Image Source?}
    C -- In Document --> D[Access/Extract Image Data from Document]
    C -- Dynamic Resource --> E[Read Dynamic Resource]
    C -- External URL/File --> F[Access External Image Data]

    D, E, F --> G[Attempt Analysis with Master AI]
    G -- Cannot Analyze --> H[Fallback to Connected LLM]
    G -- Analysis Result --> I[Present Analysis Result]
    H -- Analysis Result --> I
    H -- Cannot Analyze --> J[Indicate Incapability]

    B -- Generation --> K[Attempt Generation with Master AI]
    K -- Cannot Generate --> L[Fallback to Connected LLM]
    K -- Generated Image Data --> M[Identify Target Document & Location]
    L -- Generated Image Data --> M
    L -- Cannot Generate --> J

    M --> N[Insert Image into Document]
    N --> O[Confirm Insertion]

    %% Tools used (simplified)
    D -.-> |word/image/extract, office/pdf/parse| P[msoffice-mcp]
    E -.-> |dynamic/resources| P
    F -.-> |stagehand_extract or new tool| Q[External Access]
    H, L --> R[Connected LLM (via llmClient.ts)]
    N -.-> |word/image/insert| P

    subgraph msoffice-mcp Tools
        P
    end

    subgraph External Capabilities
        Q
        R
    end