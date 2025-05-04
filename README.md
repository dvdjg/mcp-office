# Office MCP Server: Supercharge Your Office Automation with AI 🚀

Welcome to the **Office MCP Server**, your go-to tool for automating Microsoft Office with flair! Built with **FastMCP** in TypeScript, this server makes Word, Excel, and PowerPoint automation a breeze. Whether you’re merging documents, rendering diagrams, formatting code, or exporting to PDF, Office MCP combines AI intelligence with developer-friendly tools to make Office tasks *fun*. Say goodbye to manual grunt work and hello to productivity paradise! 🎉

## Why Office MCP?
- **AI-Powered Magic**: Leverage AI to suggest styles, resolve merge conflicts, or generate text via FastMCP’s prompt system.
- **Granular Control**: Fine-tune every paragraph, table, or embedded object.
- **Cross-Platform**: Supports Office 365, desktop Office, and Power Automate.
- **Developer-Friendly**: Type-safe, modular, and well-documented for easy integration.
- **Rock-Solid**: >90% test coverage ensures reliability for your automation adventures.

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

## Explore Office MCP
Dive into specific topics:
- **[Installation & API](INSTALLATION_AND_API.md)**: Guides for setup, Docker, environment configs, and the full RESTful API.
- **[Use Cases](USE_CASES.md)**: Fun, practical examples for automating Word, Excel, PowerPoint, and multi-app workflows.
- **[Technical Details](TECHNICAL_DETAILS.md)**: Insights into TypeScript, FastMCP, testing, security, future plans, and licensing.

## Join the Automation Party
Office MCP is your key to effortless Office automation. Have ideas or questions? Open an issue on [GitHub](https://github.com/dvdjg/mcp-office/issues) or check `memory://ai_assistant_guide` for tips. Let’s make Office tasks a joy! 🚀