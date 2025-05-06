# Installation & API Documentation

The **Office MCP Server** provides a powerful RESTful API for automating Microsoft Office applications (Word, Excel, PowerPoint) and manipulating Office document files (.docx, .xlsx, .pptx). This server employs a dual approach for document handling:

1.  **COM (Component Object Model) Automation:** Utilized when the server operates locally and needs to interact with documents currently open by the user, or to create/edit local documents in a shared editing context. This allows for deep integration with the Office applications themselves.
2.  **Direct File Manipulation (JavaScript/TypeScript Libraries):** Used for most other operations, including creating, reading, and editing Office documents by parsing or generating the files directly. This approach is platform-independent and does not require Office to be installed on the server for many tasks. Libraries such as [`docx`](https://www.npmjs.com/package/docx) for Word document generation, [`exceljs`](https://www.npmjs.com/package/exceljs) or [`xlsx`](https://www.npmjs.com/package/xlsx) for Excel file handling, and [`PptxGenJS`](https://github.com/gitbrent/PptxGenJS) for PowerPoint presentation creation are leveraged for this purpose, as detailed in our [office_library_analysis.md](office_library_analysis.md:0).

This guide covers setup, configuration, and detailed API endpoints for leveraging these capabilities.

## Table of Contents
- [Installation](#installation)
- [Running the Server](#running-the-server)
- [API Overview](#api-overview)
- [File System Tools](#file-system-tools)
- [Word Tools](#word-tools)
- [Excel Tools](#excel-tools)
- [PowerPoint Tools](#powerpoint-tools)
- [Cross-Application Tools](#cross-application-tools)
- [Static and Dynamic Resources](#static-and-dynamic-resources)

## Installation

### Prerequisites
- **Node.js**: v18 or higher
- **Microsoft Office (Desktop Version)**: Required if you intend to use features relying on COM automation (e.g., interacting with currently open documents). Not strictly necessary for all direct file manipulation tasks handled by JavaScript/TypeScript libraries, but recommended for full functionality and development/testing of COM-based features.
- **Git**: For cloning

### Clone the Repository
```bash
git clone https://github.com/dvdjg/mcp-office.git
cd mcp-office
```

### Install Dependencies
```bash
npm install
```

### Docker Option
```bash
docker build -t mcp-office .
docker run -p 3000:3000 mcp-office
```

### Compile TypeScript
```bash
npm run build
```

## Running the Server
Defaults to STDIO mode. Configure file system access with environment variables:

- **File System Access**:
  - `ALLOWED_FS_PATHS`: Restrict paths (e.g., `/docs;/data` or `none`).
  - Unset: All paths allowed (use cautiously!).

Example:
```bash
ALLOWED_FS_PATHS="/docs;./data" npm start
```

Network mode (disables FS tools):
```bash
OFFICE_MCP_PORT=3000 npm start
```

Production with PM2:
```bash
npm install -g pm2
pm2 start dist/server/index.js --name mcp-office
```

Debugging:
```bash
npx fastmcp dev dist/server/index.js
npx fastmcp inspect src/tools/index.ts
```

## LLM Context and API Key Configuration

Office MCP tools that utilize AI-powered features, such as text generation or suggestions, access Language Model (LLM) capabilities through the FastMCP environment running the server. This access is provided via the `context.session.requestSampling` function available to the tool handlers.

This means that the LLM configuration (including API keys and model names) is managed by the FastMCP environment itself, not directly by the Office MCP server.

If the FastMCP environment is not configured for LLM access for a particular session, the `context.session.requestSampling` function will not be available. In such cases, tools that require LLM sampling will return an error indicating that LLM sampling is unavailable.

To enable LLM-dependent tools, ensure that the FastMCP environment you are using to run the Office MCP server is properly configured for LLM access. Consult the documentation for your specific FastMCP setup for details on how to configure LLM providers and API keys.

## API Overview
The API uses **FastMCP**, offering type-safe endpoints with `zod` validation. Structure: `/{module}/{tool}/{operation}` (e.g., `/word/styles/apply`). Responses are JSON with `success`, `data`, or `error`. Resources use `office://<document_path>?range=<range_specifier>`.

**Workflow**:
```mermaid
graph TD
  A[Client] -->|GET /word/styles/list| B[Office MCP Server]
  B -->|Validate with zod| C[Office Interaction Layer (COM/JS Libs)]
  C -->|Process request| D[Office Document/Application]
  D -->|Return data| C
  C -->|JSON response| B
  B -->|Response: { styles: [...] }| A
```

Generate TypeDoc:
```bash
npm run docs:api
```

## File System Tools
- **fs/directory**: `list`, `create`, `delete`
  - Example: `GET /fs/directory/list?path=/docs&filter=*.docx`
- **fs/file**: `read`, `write`, `delete`
  - Example: `GET /fs/file/read?path=/docs/sample.docx`
- **fs/blob**: `save`, `read`
- **fs/archive**: `list`, `extract`, `create`

## Word Tools
- **word/styles**: `list`, `apply`, `create`, `modify`, `delete`
  - Example: `POST /word/styles/apply?document=/docs/report.docx&style=Heading1&range=paragraph:1`
- **word/markdown/import**: Convert Markdown to Word
  - Example: `POST /word/markdown/import?path=/docs/input.md&template=/docs/template.docx`
- **word/merge**: Merge with AI conflict resolution
- **word/embedded-objects**: `insert`, `modify`, `delete`, `extractAll`
- **word/image**: `extract`, `insert`
- **word/generate-and-insert-text**: AI text insertion
- **word/page**: Configure layout
- **word/analyze**: Analyze and comment
- **word/code-format**: Format code snippets

## Excel Tools
- **excel/worksheets**: `add`, `delete`, `set`
- **excel/range**: `read`, `write`, `format`, `apply`
  - Example: `POST /excel/range/write?document=/docs/data.xlsx&range=A1:B2&values=[[1,2],[3,4]]`
- **excel/data-analysis**: Filter or analyze data
- **excel/tables**: `insert`, `modify`, `delete`
- **excel/charts**: Insert charts

## PowerPoint Tools
- **powerpoint/slides**: `add`, `delete`, `set`
  - Example: `POST /powerpoint/slides/add?document=/docs/presentation.pptx&layout=TitleSlide`
- **powerpoint/shapes**: Add shapes or images
- **powerpoint/charts**: Insert charts
- **powerpoint/properties**: Manage presentation settings
- **powerpoint/animations**: Manage animations and transitions

## Cross-Application Tools
- **office/pdf/export**: Convert to PDF
- **office/combine**: Combine documents
- **office/transfer**: Move or link data across apps
  - Example: `POST /office/transfer?source=excel:./data.xlsx:Sheet1:A1&target=word:./report.docx:end`
- **office/pdf/parse**: Extract content from PDFs
- **office/word-to-powerpoint**: Convert Word to PowerPoint
- **office/ai-suggest**: Provide AI-driven suggestions
- **office/workflow**: Execute multi-step automation workflows

## Static and Dynamic Resources
- **memory/ai_assistant_guide**: `read`, `list`
- **dynamic/resources**: `list`, `read`, `write`, `delete`, `metadata`, `search`

## Next Steps
- Explore [Use Cases](USE_CASES.md) for practical examples.
- See [Technical Details](TECHNICAL_DETAILS.md) for implementation and future plans.
- Return to [README](README.md).