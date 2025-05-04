# Project Plan: Enhancing Office MCP with Archive Handling, Cloud Access (Future), and Language Detection

This plan outlines the steps to integrate archive file handling (supporting multiple formats), document the future implementation of access to Office documents via Teams/Office 365 links with user permissions, incorporate language detection, and update documentation within the Office MCP server.

**Goals:**

1.  Implement tools for handling various archive formats (ZIP, RAR, 7z, gz, tar, etc.), including listing, extracting, and creating archives.
2.  Document the planned mechanism for the MCP server to access Office documents from Teams/Office 365 links on behalf of the user, respecting their permissions, as a future development phase.
3.  Document the considerations for ensuring interoperability between desktop COM/OLE functionalities and cloud-based documents for future implementation.
4.  Integrate language detection using `@vscode/vscode-languagedetection`, specifically in the `word/code-format` tool.
5.  Update relevant documentation (`README.md`, `USE_CASES.md`, `TECHNICAL_DETAILS.md`, etc.).

**Detailed Steps:**

1.  **Archive File Handling Implementation:**
    *   **Step 1.1:** Research and select a suitable Node.js library that supports multiple archive formats (e.g., ZIP, RAR, 7z, gz, tar, etc.) for file manipulation. Prioritize libraries that handle various file types and structures within archives.
    *   **Step 1.2:** Implement a new MCP tool, potentially `fs/archive/list`, to list the contents of a specified archive file. This tool will take the archive file path as input and return a list of files and directories within it.
    *   **Step 1.3:** Implement a new MCP tool, potentially `fs/archive/extract`, to extract specific files or the entire contents of an archive to a specified directory. This tool will take the archive file path, target file/directory (optional), and output directory as input.
    *   **Step 1.4:** Implement a new MCP tool, potentially `fs/archive/create`, to create an archive from a list of files and directories. This tool will take the output archive file path and a list of source file/directory paths as input.
    *   **Step 1.5:** Integrate these new tools into the MCP server's tool registry.

2.  **Office 365/Teams Document Access and Interoperability (Future Development):**
    *   **Step 2.1:** Document the research into mechanisms for accessing Office 365/Teams files programmatically with user permissions, focusing on the Microsoft Graph API.
    *   **Step 2.2:** Document a potential authentication flow, likely involving an OAuth 2.0 flow with the Microsoft Identity Platform to grant delegated permissions.
    *   **Step 2.3:** Document the concept of a service or module within the MCP server to handle authenticated requests to the Microsoft Graph API for file access.
    *   **Step 2.4:** Document the potential modifications needed for existing Office tools to accept Office 365/Teams URLs in addition to local file paths, outlining how they would interact with the future Graph API service.
    *   **Step 2.5:** Document the investigation into the Office JavaScript API and its potential role in performing complex operations on cloud documents, possibly involving a web-based component.
    *   **Step 2.6:** Document the prioritization of COM/OLE functionalities that would be most critical to replicate or enable for cloud documents.
    *   **Step 2.7:** Include diagrams in `TECHNICAL_DETAILS.md` to illustrate the proposed architecture for cloud document access.

3.  **Language Detection Integration:**
    *   **Step 3.1:** Utilize the already installed `@vscode/vscode-languagedetection` module.
    *   **Step 3.2:** Modify the `src/tools/word/codeFormat.tool.ts` tool. Before attempting to format a code snippet, use the language detection module to identify the programming language.
    *   **Step 3.3:** Use the detected language to apply appropriate syntax highlighting and formatting rules within the Word document. Handle cases where the language cannot be confidently detected.

4.  **Documentation Updates:**
    *   **Step 4.1:** Update `USE_CASES.md` to reflect the implemented archive handling capabilities (mentioning support for multiple formats).
    *   **Step 4.2:** Update `README.md` to include a high-level overview of the new archive handling and language detection features, and mention the planned future work on cloud document access.
    *   **Step 4.3:** Update `TECHNICAL_DETAILS.md` to describe the architecture for cloud document access as future work, including the authentication flow and the use of Microsoft Graph API, with diagrams. Detail the integration of language detection.
    *   **Step 4.4:** Review and update any other relevant documentation or API guides.

**Mermaid Diagram: Cloud Document Access Flow (Future)**

```mermaid
graph TD
    A[User Request with Office 365/Teams Link] --> B{MCP Server};
    B --> C{Authentication Service};
    C -- OAuth 2.0 Flow --> D[Microsoft Identity Platform];
    D -- Tokens --> C;
    C -- Authenticated Request --> E[Microsoft Graph API];
    E -- File Content/Metadata --> B;
    B -- Process Document --> F[Office Tools (Word, Excel, PPT)];
    F -- COM/OLE (Desktop) / JS API (Web) --> G[Office Application (Desktop/Web)];
    G -- Result --> B;
    B --> H[Response to User];
```

This diagram illustrates the proposed flow for future implementation, from a user request containing an Office 365/Teams link through the MCP server, authentication, interaction with Microsoft Graph API, processing by Office tools, and finally returning a response to the user. The Office Application box represents where the actual document manipulation might occur, either via COM/OLE for desktop or potentially the JS API for web, orchestrated by the MCP server.