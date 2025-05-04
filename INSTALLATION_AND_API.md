# Installation & API Documentation

Get the **Office MCP Server** running and explore its RESTful API for automating Word, Excel, PowerPoint, and more. This guide covers setup, configuration, and detailed API endpoints.

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
- **Microsoft Office**: Desktop (for VBA/COM) or Office 365 (for JavaScript APIs)
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

## API Overview
The API uses **FastMCP**, offering type-safe endpoints with `zod` validation. Structure: `/{module}/{tool}/{operation}` (e.g., `/word/styles/apply`). Responses are JSON with `success`, `data`, or `error`. Resources use `office://<document_path>?range=<range_specifier>`.

**Workflow**:
```mermaid
graph TD
  A[Client] -->|GET /word/styles/list| B[Office MCP Server]
  B -->|Validate with zod| C[Office JavaScript API]
  C -->|Retrieve styles| D[Word Document]
  D -->|Return styles| C
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
- **excel/range**: `read`, `write`, `format`, `apply`
  - Example: `POST /excel/range/write?document=/docs/data.xlsx&range=A1:B2&values=[[1,2],[3,4]]`
- **excel/data-analysis**: Filter or analyze data

## PowerPoint Tools
- **powerpoint/slides**: `add`, `delete`, `set`
  - Example: `POST /powerpoint/slides/add?document=/docs/presentation.pptx&layout=TitleSlide`
- **powerpoint/shapes**: Add shapes or images
- **powerpoint/charts**: Insert charts

## Cross-Application Tools
- **office/pdf/export**: Convert to PDF
- **office/combine**: Combine documents
- **office/transfer**: Move or link data across apps
  - Example: `POST /office/transfer?source=excel:./data.xlsx:Sheet1:A1&target=word:./report.docx:end`

## Static and Dynamic Resources
- **memory/ai_assistant_guide**: `read`, `list`
- **dynamic/resources**: `list`, `read`, `write`, `delete`, `metadata`, `search`

## Next Steps
- Explore [Use Cases](USE_CASES.md) for practical examples.
- See [Technical Details](TECHNICAL_DETAILS.md) for implementation and future plans.
- Return to [README](README.md).