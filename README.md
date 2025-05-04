# Office MCP Server: Supercharge Your Office Automation with AI 🚀

Welcome to the **Office MCP Server**, your go-to tool for automating Microsoft Office with flair! Built with **FastMCP** in TypeScript, this server makes Word, Excel, and PowerPoint automation a breeze. Whether you’re merging documents, rendering diagrams, formatting code, or exporting to PDF, Office MCP combines AI intelligence with developer-friendly tools to make Office tasks *fun*. Say goodbye to manual grunt work and hello to productivity paradise! 🎉

## Why Office MCP?
- **AI-Powered Magic**: Leverage AI to suggest styles, resolve merge conflicts, or generate text via FastMCP’s prompt system.
- **Archive Handling**: Easily manage files within ZIP, 7z, and other archive formats (listing, extracting, creating for ZIP/7z).
- **Intelligent Language Detection**: Automatically detect programming languages in code snippets for accurate formatting.
- **Granular Control**: Fine-tune every paragraph, table, or embedded object.
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
   ```bash
   npm start
   ```

Try listing Word styles:
```bash
curl -X GET "http://localhost:3000/word/styles/list?document=/docs/sample.docx"
```

For detailed setup, see [Installation & API](INSTALLATION_AND_API.md).

## Context Handling and AI Integration

Office MCP tools leverage various forms of context to perform intelligent and flexible automation tasks. Understanding these contexts is key to effectively using and extending the server.

Here are the main types of context utilized:

*   **LLM Context:** Tools like `word/generate-and-insert-text` interact with a Language Model (LLM) via the FastMCP session provided by the *FastMCP environment running the server*. This context allows the tools to perform AI-driven tasks such as text generation, content summarization, or style suggestions, leveraging the capabilities of the parent AI.
*   **File Content Context:** Many tools operate directly on the content of Office files. This includes reading text from a Word document (`word/text/get`), writing data to an Excel sheet (`excel/range/write`), or extracting images from a PowerPoint presentation (`powerpoint/shapes/list`). The tools access and manipulate the file content based on the specific operation.
*   **User Data/Input Context:** The parameters provided in tool requests (e.g., `filePath`, `position`, `prompt`, `rangeAddress`) constitute user data or input context. These parameters guide the tool's execution and determine the specific target and nature of the operation.
*   **System/Environment Context:** Tools interact with the underlying system and environment, including the file system (reading/writing files), the operating system (running COM Interop for Office applications), and the availability of installed Office applications. This context is essential for the tools to function correctly within the user's environment.
*   **Session Context:** The `FastMCPContext` provided to each tool handler includes a `session` property. This session object is crucial for interacting with the FastMCP framework, particularly for requesting LLM sampling (`context.session.requestSampling`) and potentially accessing other session-specific information or resources provided by the FastMCP environment.

By combining these different types of context, Office MCP tools can perform complex tasks that go beyond simple automation, enabling more intelligent and responsive interactions with Microsoft Office applications.
## Explore Office MCP
Dive into specific topics:
- **[Installation & API](INSTALLATION_AND_API.md)**: Guides for setup, Docker, environment configs, and the full RESTful API.
- **[Use Cases](USE_CASES.md)**: Fun, practical examples for automating Word, Excel, PowerPoint, and multi-app workflows, including archive management.
- **[Technical Details](TECHNICAL_DETAILS.md)**: Insights into TypeScript, FastMCP, testing, security, future plans (including cloud document access), and licensing.

*   **Office Instance Context:** The server interacts with the default or currently active Microsoft Office application instance available via COM Interop. Managing multiple installed Office versions or specific user sessions/licenses within the Office applications themselves is a complex consideration and is not currently supported.
## Join the Automation Party
Office MCP is your key to effortless Office automation. Have ideas or questions? Open an issue on [GitHub](https://github.com/dvdjg/mcp-office/issues) or check `memory://ai_assistant_guide` for tips. Let’s make Office tasks a joy! 🚀