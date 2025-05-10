# Office MCP Server: Supercharge Your Office Automation with AI 🚀

Welcome to the **Office MCP Server**, your go-to tool for automating Microsoft Office with flair! Built with **FastMCP** in TypeScript, this server makes Word, Excel, and PowerPoint automation a breeze. Whether you’re merging documents, rendering diagrams, formatting code, or exporting to PDF, Office MCP combines AI intelligence with developer-friendly tools to make Office tasks *fun*. Say goodbye to manual grunt work and hello to productivity paradise! 🎉

## Why Office MCP?
- **AI-Powered Magic**: Leverage AI to suggest styles, resolve merge conflicts, or generate text. The `word/generate-and-insert-text` tool now processes LLM output in Markdown format, applying appropriate Word styles and handling the initial summary as a heading. This uses a server-side LLM integration due to client limitations and a new internal utility for Markdown-to-Word formatting, including support for HTML tables with merged cells.
- **Enhanced Markdown Handling**: Seamlessly import and export Word content to and from Markdown, now with support for complex tables using embedded HTML.
- **Table Data Extraction**: Easily extract data from Word tables into a structured format for further processing (`word/tables/extractData`).
- **Archive Handling**: Easily manage files within ZIP, 7z, and other archive formats (listing, extracting, creating for ZIP/7z).
- **Intelligent Language Detection**: Automatically detect programming languages in code snippets for accurate formatting.
- **Granular Control**: Fine-tune every paragraph, table, or embedded object.
- **Structured Data Insertion**: Easily insert data from arrays into Word tables with styling options (`word/tables/insertFromArray`).
- **Cross-Platform**: Supports Office 365, desktop Office, and Power Automate.
- **Developer-Friendly**: Type-safe, modular, and well-documented for easy integration.
- **Rock-Solid**: >90% test coverage ensures reliability for your automation adventures.
- **Cloud Document Access (Future)**: Planned support for accessing and working with documents via Teams/Office 365 links with user permissions.

## Quick Start
Get started in minutes:

1. **Clone the Repo**:
   ```bash
   git clone https://github.com/dvdjg/mcp-office.git
   cd mcp-office
   ```

2. **Install Dependencies**:
   ```bash
   npm install
   ```

3. **Run the Server**:

   The server can be run in several ways:

   **a) Standard Start (Network Mode for Tests/Remote Access):**
   By default, if no port is specified, the server starts in STDIO mode. To run it on `http://localhost:3000` (which disables FS tools, a consideration for tests relying on them), set the `OFFICE_MCP_PORT` environment variable:

   *Nix & macOS:*
   ```bash
   OFFICE_MCP_PORT=3000 npm start
   ```

   *Windows (Command Prompt):*
   ```bash
   set OFFICE_MCP_PORT=3000 && npm start
   ```

   *Windows (PowerShell):*
   ```bash
   $env:OFFICE_MCP_PORT="3000"; npm start
   ```
   This uses the compiled output in `dist/server/index.js`.

   **b) Development Mode with `fastmcp dev` (using `mcp-cli`):**
   For interactive testing and debugging in your terminal during development, you can use `fastmcp dev`. This command runs your server with `mcp-cli`, allowing you to directly interact with it. It's excellent for quickly testing tools and seeing live responses.

   *Using the compiled JavaScript output:*
   ```bash
   npx fastmcp dev dist/server/index.js
   ```
   *Or, directly with the TypeScript source (requires `tsx` or similar to be available if not handled by `fastmcp` internally for `.ts` files):*
   ```bash
   npx fastmcp dev src/server/index.ts
   ```
   *(Note: The `INSTALLATION_AND_API.md` refers to `dist/server/index.js` for `fastmcp dev`)*

   **c) Inspecting with `MCP Inspector`:**
   To inspect your server's capabilities (tools, resources, prompts) with a Web UI, use `fastmcp inspect`. This is very useful for understanding what the server offers and for debugging definitions.

   *Typically, you inspect the source file where tools are aggregated or the main server file:*
   ```bash
   npx fastmcp inspect src/server/index.ts
   ```
   *(Note: The `INSTALLATION_AND_API.md` refers to `src/tools/index.ts` for inspection, which might be more focused if that's where all tool definitions are centralized. We'll use `src/server/index.ts` for broader inspection unless `src/tools/index.ts` is confirmed as the primary definition point for all MCP elements.)*

   You can also use `npm run dev` for development with automatic recompilation and server restart on file changes.

Try listing Word styles:
```bash
curl -X GET "http://localhost:3000/word/styles/list?document=/docs/sample.docx"
```

For detailed setup, see [Installation & API](INSTALLATION_AND_API.md).

## Context Handling and AI Integration

Office MCP tools leverage various forms of context to perform intelligent and flexible automation tasks. Understanding these contexts is key to effectively using and extending the server.

Here are the main types of context utilized:

*   **LLM Context:** Tools like `word/generate-and-insert-text` interact with a Language Model (LLM). While FastMCP supports client-side LLM sampling via the session provided by the *FastMCP environment running the server*, the `word/generate-and-insert-text` tool currently implements server-side LLM integration. This requires configuring the server with the following environment variables:
    *   `LLM_PROVIDER_API_KEY`: Your API key for the LLM provider.
    *   `LLM_PROVIDER_ENDPOINT`: The API endpoint URL for the LLM provider (e.g., `https://api.openai.com/v1/chat/completions`).
    *   `LLM_MODEL_NAME`: (Optional) The name of the specific LLM model to use (e.g., `gpt-4o`).
    This server-side approach allows the tool to perform AI-driven tasks such as text generation, content summarization, or style suggestions, bypassing client limitations. The `word/generate-and-insert-text` tool now also includes enhanced prompt engineering to guide the LLM's output format for better parsing and formatting.
*   **Markdown Formatting Utility:** A new internal utility (`src/utils/markdownToOffice.ts`) has been introduced to handle the conversion of Markdown syntax to Office application formatting (initially focused on Word styles like headings, lists, bold, and italic) using COM Interop. This utility is designed to be reusable across different Office tools for consistent Markdown integration.
*   **File Content Context:** Many tools operate directly on the content of Office files. This includes reading text from a Word document (`word/text/get`), writing data to an Excel sheet (`excel/range/write`), or extracting images from a PowerPoint presentation (`powerpoint/shapes/list`). The tools access and manipulate the file content based on the specific operation.
*   **User Data/Input Context:** The parameters provided in tool requests (e.g., `filePath`, `position`, `prompt`, `rangeAddress`) constitute user data or input context. These parameters guide the tool's execution and determine the specific target and nature of the operation.
*   **System/Environment Context:** Tools interact with the underlying system and environment, including the file system (reading/writing files), the operating system (running COM Interop for Office applications), and the availability of installed Office applications. This context is essential for the tools to function correctly within the user's environment.
*   **Session Context:** The `FastMCPContext` provided to each tool handler includes a `session` property. This session object is crucial for interacting with the FastMCP framework, particularly for requesting LLM sampling (`context.session.requestSampling`) when supported by the client, and potentially accessing other session-specific information or resources provided by the FastMCP environment.

By combining these different types of context, Office MCP tools can perform complex tasks that go beyond simple automation, enabling more intelligent and responsive interactions with Microsoft Office applications.
## Explore Office MCP
Dive into specific topics:
- **[Installation & API](INSTALLATION_AND_API.md)**: Guides for setup, Docker, environment configs, and the full RESTful API.
- **[Use Cases](USE_CASES.md)**: Fun, practical examples for automating Word, Excel, PowerPoint, and multi-app workflows, including archive management.
- **[Technical Details](TECHNICAL_DETAILS.md)**: Insights into TypeScript, FastMCP, testing, security, future plans (including cloud document access), and licensing.

*   **Office Instance Context:** The server interacts with the default or currently active Microsoft Office application instance available via COM Interop. Managing multiple installed Office versions or specific user sessions/licenses within the Office applications themselves is a complex consideration and is not currently supported.
## Testing 🧪

This project uses Jest for testing. You can find test files in the `tests/` directory, categorized into `unit` and `e2e` (end-to-end, or functional tests).

**Running All Tests:**
To run all tests, use the following command:
```bash
npm test
```

**Running Unit Tests:**
Unit tests focus on individual components and do not require the server to be running.
```bash
npm run test:unit
```

**Running Functional (End-to-End) Tests:**
Functional tests verify the integration of multiple components and often require the Office MCP server to be running. Ensure the server is started (e.g., on port 3000) before running these tests. Refer to the "Run the Server" section for instructions on starting the server.

Launch the server with authomatic recompilation:
```bash
OFFICE_MCP_PORT=3000 npm run dev
```

Launch the e2e test:
```bash
npm run test:e2e
```
**Note on COM Components for E2E Tests:** Some E2E tests might interact with Microsoft Office applications via COM. These tests are intended for environments where Office applications are installed and accessible. For tests running in environments without Office (e.g., certain CI/CD pipelines or Docker containers not configured for Office COM), ensure they are either skipped or adapted to mock COM interactions if they are to be included. Tests that do not run locally should ideally avoid COM dependencies.

**Running a Single Test File:**
To run a specific test file, append its path to the `test:single` script:
```bash
npm run test:single tests/unit/your-test-file.test.ts
```

```bash
npm run test:single tests/e2e/word/analyze.test.ts
```
**Test Coverage:**
To generate a test coverage report:
```bash
For more detailed information on the testing strategy, requesting specific tests, and interpreting results, please see the [Detailed Testing Guide](docs/TESTING_GUIDE.md).
npm run test:coverage
```
This will create a `coverage/` directory with the report.
## Join the Automation Party
Office MCP is your key to effortless Office automation. Have ideas or questions? Open an issue on [GitHub](https://github.com/dvdjg/mcp-office/issues) or check `memory://ai_assistant_guide` for tips. Let’s make Office tasks a joy! 🚀