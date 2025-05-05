# Office MCP Server API Specification

**Document Purpose**: This document provides a detailed specification of the Office MCP (Model Context Protocol) Server API, built with FastMCP in TypeScript for automating Microsoft Office applications (Word, Excel, PowerPoint). It describes each tool's operations, supported data types, input/output formats, example API calls, and guidance on phrasing user requests to an AI (e.g., Claude) to trigger specific tools. This ensures developers and AI agents can use the API effectively.

**Target Audience**: Developers integrating with the Office MCP Server and AI agents processing user requests.

**Date**: May 04, 2025

---

## Table of Contents
1. [Overview](#overview)
   - [API Structure](#api-structure)
   - [Data Types and Formats](#data-types-and-formats)
   - [AI Request Guidelines](#ai-request-guidelines)
2. [File System Tools](#file-system-tools)
   - [fs/directory](#fsdirectory)
   - [fs/file](#fsfile)
   - [fs/blob](#fsblob)
   - [fs/archive](#fsarchive)
3. [Word Tools](#word-tools)
   - [word/styles](#wordstyles)
   - [word/text](#wordtext)
   - [word/search-replace](#wordsearch-replace)
   - [word/page](#wordpage)
   - [word/headers-footers](#wordheaders-footers)
   - [word/tables](#wordtables)
   - [word/charts](#wordcharts)
   - [word/embedded-objects](#wordembedded-objects)
   - [word/metadata](#wordmetadata)
   - [word/batch](#wordbatch)
   - [word/markdown/import](#wordmarkdownimport)
   - [word/markdown/export](#wordmarkdownexport)
   - [word/mermaid/import](#wordmermaidimport)
   - [word/mermaid/export](#wordmermaidexport)
   - [word/merge](#wordmerge)
   - [word/template](#wordtemplate)
   - [word/reformat](#wordreformat)
   - [word/analyze](#wordanalyze)
   - [word/code-format](#wordcode-format)
   - [word/image](#wordimage)
   - [word/generate-and-insert-text](#wordgenerate-and-insert-text)
4. [Excel Tools](#excel-tools)
   - [excel/worksheets](#excelworksheets)
   - [excel/range](#excelrange)
   - [excel/tables](#exceltables)
   - [excel/charts](#excelcharts)
   - [excel/data-analysis](#exceldata-analysis)
5. [PowerPoint Tools](#powerpoint-tools)
   - [powerpoint/slides](#powerpointslides)
   - [powerpoint/shapes](#powerpointshapes)
   - [powerpoint/properties](#powerpointproperties)
   - [powerpoint/animations](#powerpointanimations)
6. [Cross-Application Tools](#cross-application-tools)
   - [office/transfer](#officetransfer)
   - [office/workflow](#officeworkflow)
   - [office/ai-suggest](#officeai-suggest)
   - [office/pdf/export](#officepdfexport)
   - [office/pdf/parse](#officepdfparse)
   - [office/combine](#officecombine)
   - [office/word-to-powerpoint](#officeword-to-powerpoint)
7. [Resource Management](#resource-management)
   - [memory/ai_assistant_guide](#memoryai_assistant_guide)
   - [dynamic/resources](#dynamicresources)

---

## Overview

### API Structure
The Office MCP Server exposes a RESTful API via FastMCP, organized as resource templates:
- **Path**: `/{module}/{tool}/{operation}` (e.g., `/word/styles/apply`)
- **Methods**: GET (read), POST (write/modify), DELETE (remove)
- **Parameters**: Query parameters (GET) or JSON body (POST)
- **Authentication**: Role-based access tokens (optional, configured in FastMCP)
- **Responses**: JSON objects with `success`, `data`, or `error` fields
- **Completions**: Dynamic suggestions for parameters (e.g., style names, file paths)
- **Resource URIs**: Supports standard file paths and the `office://<document_path>?range=<range_specifier>` URI scheme for dynamic resource access (e.g., `office://docs/report.docx?range=paragraph:5`). *Note: Full implementation for dynamic range access via COM interop is pending.*

**Example Request**:
```bash
curl -X POST "http://localhost:3000/word/styles/apply?document=/docs/sample.docx&style=Heading1&range=paragraph:1" \
  -H "Content-Type: application/json"
```

**Example Response**:
```json
{ "success": true }
```

### Data Types and Formats
- **Strings**: UTF-8 encoded (e.g., file paths, style names).
- **Numbers**: Integers or floats (e.g., row counts, margins).
- **Booleans**: `true`/`false` (e.g., enable comments).
- **Arrays**: Ordered lists (e.g., `[1,2]`, `["doc1.docx","doc2.docx"]`).
- **Objects**: Key-value pairs (e.g., `{ size: "A4", margins: 1 }`).
- **Files/Blobs**: Binary data (multipart/form-data for uploads).
- **Ranges**: Custom format (e.g., `paragraph:1`, `A1:B2`).
- **Input Validation**: Enforced with `zod` schemas.
- **Output Format**: JSON, with optional binary data (e.g., PNG for Mermaid export).

**Error Format**:
```json
{
  "success": false,
  "error": {
    "code": 400,
    "message": "Invalid style name",
    "details": "Style 'InvalidStyle' not found"
  }
}
```

### AI Request Guidelines
The Office MCP Server integrates with AI agents (e.g., Claude) via FastMCP's prompt and completion system. Users must phrase requests clearly to trigger specific tools/operations. Each tool section below includes:
- **Prompt Examples**: Natural language phrases that map to the tool/operation.
- **Completions**: Suggested parameters the AI can infer (e.g., style names).
- **Context**: Document or resource state the AI should query.

**General Tips for AI Requests**:
- Be specific about the task (e.g., “merge two Word documents” vs. “combine files”).
- Mention the application (Word, Excel, PowerPoint) or file type (docx, md, pdf).
- Include details like file paths, styles, or formats if relevant.
- Use action verbs (e.g., “apply,” “export,” “analyze”) to align with operations.

**Example**:
- User: “Merge two Word documents, draft1.docx and draft2.docx, into one.”
- AI Infers: `word/merge` with `docs=/docs/draft1.docx,/docs/doc2.docx`.
- AI Queries: `GET /dynamic/resources/list?type=docx` to confirm file existence.

---

## File System Tools

### fs/directory
**Description**: Manages directories (list, create, delete).

**Operations**:
- **list**: Lists files/subdirectories.
  - **Input**: `path` (string, directory path), `filter` (string, optional, e.g., `*.docx`).
  - **Output**: Array of file/subdirectory paths.
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/fs/directory/list?path=/docs&filter=*.docx"
    ```
    **Response**:
    ```json
    ["/docs/sample.docx", "/docs/report.docx"]
    ```
- **create**: Creates a directory.
  - **Input**: `path` (string, new directory path).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/fs/directory/create?path=/docs/new"
    ```
- **delete**: Deletes a directory.
  - **Input**: `path` (string, directory path).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/fs/directory/delete?path=/docs/old"
    ```

**AI Request Examples**:
- User: “List all Word documents in the /docs folder.”
  - AI Infers: `GET /fs/directory/list?path=/docs&filter=*.docx`.
- User: “Create a new folder called ‘reports’ in /docs.”
  - AI Infers: `POST /fs/directory/create?path=/docs/reports`.
- User: “Delete the ‘temp’ folder in /docs.”
  - AI Infers: `POST /fs/directory/delete?path=/docs/temp`.

**Completions**: File paths, file extensions (e.g., `*.docx`, `*.md`).

### fs/file
**Description**: Reads/writes/deletes files.

**Operations**:
- **read**: Reads file content.
  - **Input**: `path` (string, file path), `format` (string, `text` or `binary`, default: `text`).
  - **Output**: `{ content: string | Buffer }`.
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/fs/file/read?path=/docs/note.txt&format=text"
    ```
    **Response**:
    ```json
    { "content": "Meeting notes" }
    ```
- **write**: Writes content to a file.
  - **Input**: `path` (string, file path), `content` (string or Buffer), `append` (boolean, optional, default: `false`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/fs/file/write?path=/docs/note.txt&content=Hello&append=true" \
      -H "Content-Type: application/json"
    ```
- **delete**: Deletes a file.
  - **Input**: `path` (string, file path).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/fs/file/delete?path=/docs/note.txt"
    ```

**AI Request Examples**:
- User: “Read the content of note.txt in /docs.”
  - AI Infers: `GET /fs/file/read?path=/docs/note.txt&format=text`.
- User: “Write ‘Project plan’ to plan.txt in /docs.”
  - AI Infers: `POST /fs/file/write?path=/docs/plan.txt&content=Project plan`.
- User: “Delete temp.docx from /docs.”
  - AI Infers: `POST /fs/file/delete?path=/docs/temp.docx`.

**Completions**: File paths, content snippets.

### fs/blob
**Description**: Handles document blobs for uploads/downloads.

**Operations**:
- **save**: Saves a blob to a file.
  - **Input**: `filename` (string, target path), `data` (binary, via multipart/form-data).
  - **Output**: `{ success: true, path: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/fs/blob/save?filename=/docs/doc1.docx" \
      -F "data=@local.docx"
    ```
    **Response**:
    ```json
    { "success": true, "path": "/docs/doc1.docx" }
    ```
- **read**: Reads a file as a blob.
  - **Input**: `path` (string, file path).
  - **Output**: Binary data (Content-Type: application/octet-stream).
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/fs/blob/read?path=/docs/doc1.docx" > doc1.docx
    ```

**AI Request Examples**:
- User: “Upload my local document draft.docx to /docs.”
  - AI Infers: `POST /fs/blob/save?filename=/docs/draft.docx` with file data.
- User: “Download doc1.docx from /docs.”
  - AI Infers: `GET /fs/blob/read?path=/docs/doc1.docx`.

**Completions**: File paths, file types (e.g., `docx`, `pdf`).

### fs/archive
**Description**: Implements tools for handling archive files (ZIP, 7z, etc.).

**Operations**:
- **list**: Lists the contents of an archive file.
  - **Input**: `filePath` (string, The path to the archive file.)
  - **Output**: `{ success: boolean, contents?: string[], error?: string }`.
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/fs/archive/list?filePath=/docs/archive.zip"
    ```
    **Response**:
    ```json
    { "success": true, "contents": ["file1.txt", "folder/file2.docx"] }
    ```
- **extract**: Extracts contents from an archive file.
  - **Input**: `filePath` (string, The path to the archive file.), `target` (string, optional, The target file or directory within the archive to extract.), `outputDirectory` (string, The directory to extract the contents to.)
  - **Output**: `{ success: boolean, error?: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/fs/archive/extract?filePath=/docs/archive.zip&outputDirectory=/docs/extracted"
    ```
- **create**: Creates an archive file from a list of source paths.
  - **Input**: `outputFilePath` (string, The path for the output archive file.), `sourcePaths` (string[], A list of file or directory paths to include in the archive.), `format` (string, 'zip' or '7z', Defaults to 'zip')
  - **Output**: `{ success: boolean, error?: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/fs/archive/create?outputFilePath=/docs/new_archive.zip" \
      -d '{"sourcePaths":["/docs/file1.txt", "/docs/folder"]}'
    ```

**AI Request Examples**:
- User: “List the contents of archive.zip in /docs.”
  - AI Infers: `GET /fs/archive/list?filePath=/docs/archive.zip`.
- User: “Extract archive.zip to /docs/extracted.”
  - AI Infers: `POST /fs/archive/extract?filePath=/docs/archive.zip&outputDirectory=/docs/extracted`.
- User: “Create a zip archive named new_archive.zip in /docs containing file1.txt and the folder /docs/folder.”
  - AI Infers: `POST /fs/archive/create?outputFilePath=/docs/new_archive.zip` with `sourcePaths`.

**Completions**: File paths, archive formats (`zip`, `7z`).

---

## Word Tools

### word/styles
**Description**: Manages styles in Word documents.

**Operations**:
- **list**: Lists available styles.
  - **Input**: `document` (string, file path).
  - **Output**: `{ styles: string[] }`.
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/word/styles/list?document=/docs/sample.docx"
    ```
    **Response**:
    ```json
    { "styles": ["Heading1", "Normal", "Title"] }
    ```
- **apply**: Applies a style to a range.
  - **Input**: `document` (string), `style` (string), `range` (string, e.g., `paragraph:1` or `selection`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/styles/apply?document=/docs/sample.docx&style=Heading1&range=paragraph:1"
    ```
- **create**: Creates a new style.
  - **Input**: `document` (string), `name` (string), `props` (object, e.g., `{ font: "Arial", size: 12 }`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/styles/create?document=/docs/sample.docx&name=CustomStyle" \
      -d '{"props":{"font":"Arial","size":12}}' \
      -H "Content-Type: application/json"
    ```
- **modify**: Modifies an existing style.
  - **Input**: `document` (string), `style` (string), `props` (object).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/styles/modify?document=/docs/sample.docx&style=Heading1" \
      -d '{"props":{"color":"blue"}}'
    ```
- **delete**: Deletes a style.
  - **Input**: `document` (string), `style` (string).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/styles/delete?document=/docs/sample.docx&style=CustomStyle"
    ```

**AI Request Examples**:
- User: “List all styles in sample.docx.”
  - AI Infers: `GET /word/styles/list?document=/docs/sample.docx`.
- User: “Apply Heading1 style to the first paragraph in sample.docx.”
  - AI Infers: `POST /word/styles/apply?document=/docs/sample.docx&style=Heading1&range=paragraph:1`.
- User: “Create a new style called ‘MyStyle’ with Arial font in sample.docx.”
  - AI Infers: `POST /word/styles/create?document=/docs/sample.docx&name=MyStyle` with `props`.

**Completions**: Style names, range formats (e.g., `paragraph:1`, `selection`).

### word/text
**Description**: Manipulates text and paragraphs.

**Operations**:
- **insert**: Inserts text.
  - **Input**: `document` (string), `text` (string), `position` (string, e.g., `paragraph:1` or `end`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/text/insert?document=/docs/sample.docx&text=Hello&position=paragraph:1"
    ```
- **modify**: Modifies text in a range.
  - **Input**: `document` (string), `range` (string), `text` (string).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/text/modify?document=/docs/sample.docx&range=paragraph:1&text=Updated"
    ```
- **delete**: Deletes text in a range.
  - **Input**: `document` (string), `range` (string).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/text/delete?document=/docs/sample.docx&range=paragraph:1"
    ```

**AI Request Examples**:
- User: “Add ‘Introduction’ to the start of sample.docx.”
  - AI Infers: `POST /word/text/insert?document=/docs/sample.docx&text=Introduction&position=paragraph:1`.
- User: “Change the first paragraph in sample.docx to ‘Welcome’.”
  - AI Infers: `POST /word/text/modify?document=/docs/sample.docx&range=paragraph:1&text=Welcome`.
- User: “Delete the second paragraph in sample.docx.”
  - AI Infers: `POST /word/text/delete?document=/docs/sample.docx&range=paragraph:2`.

**Completions**: Range formats, text snippets.

### word/search-replace
**Description**: Performs advanced search and replace.

**Operations**:
- **search**: Searches for text.
  - **Input**: `document` (string), `find` (string), `criteria` (object, e.g., `{ caseSensitive: true }`).
  - **Output**: `{ matches: { range: string, text: string }[] }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/search-replace/search?document=/docs/sample.docx&find=red" \
      -d '{"criteria":{"caseSensitive":true}}'
    ```
    **Response**:
    ```json
    { "matches": [{ "range": "paragraph:1", "text": "red" }] }
    ```
- **replace**: Replaces text.
  - **Input**: `document` (string), `find` (string), `replace` (string), `criteria` (object).
  - **Output**: `{ success: true, count: number }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/search-replace/replace?document=/docs/sample.docx&find=red&replace=blue" \
      -d '{"criteria":{"wholeWord":true}}'
    ```

**AI Request Examples**:
- User: “Find all instances of ‘error’ in sample.docx.”
  - AI Infers: `POST /word/search-replace/search?document=/docs/sample.docx&find=error`.
- User: “Replace ‘old’ with ‘new’ in sample.docx, case-sensitive.”
  - AI Infers: `POST /word/search-replace/replace?document=/docs/sample.docx&find=old&replace=new` with `criteria`.

**Completions**: Search terms, criteria options (e.g., `caseSensitive`, `wholeWord`).

### word/page
**Description**: Configures page layout.

**Operations**:
- **set**: Sets page properties.
  - **Input**: `document` (string), `size` (string, e.g., `A4`), `margins` (number, inches), `orientation` (string, `portrait` or `landscape`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/page/set?document=/docs/sample.docx&size=A4&margins=1&orientation=portrait"
    ```

**AI Request Examples**:
- User: “Set sample.docx to A4 with 1-inch margins.”
  - AI Infers: `POST /word/page/set?document=/docs/sample.docx&size=A4&margins=1`.

**Completions**: Page sizes (e.g., `A4`, `Letter`), orientations.

### word/headers-footers
**Description**: Manages headers and footers.

**Operations**:
- **insert**: Inserts header/footer.
  - **Input**: `document` (string), `section` (number), `text` (string), `type` (string, `header` or `footer`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/headers-footers/insert?document=/docs/sample.docx&section=1&text=Confidential&type=header"
    ```

**AI Request Examples**:
- User: “Add a header with ‘Confidential’ to sample.docx.”
  - AI Infers: `POST /word/headers-footers/insert?document=/docs/sample.docx&section=1&text=Confidential&type=header`.

**Completions**: Section numbers, header/footer types.

### word/tables
**Description**: Creates and manages tables.

**Operations**:
- **insert**: Inserts a table.
  - **Input**: `document` (string), `rows` (number), `columns` (number), `style` (string, e.g., `Grid`), `position` (string, e.g., `paragraph:1`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/tables/insert?document=/docs/sample.docx&rows=3&columns=4&style=Grid&position=paragraph:1"
    ```

**AI Request Examples**:
- User: “Insert a 3x4 table in sample.docx after the first paragraph.”
  - AI Infers: `POST /word/tables/insert?document=/docs/sample.docx&rows=3&columns=4&style=Grid&position=paragraph:1`.

**Completions**: Table styles, position ranges.

### word/charts
**Description**: Inserts and manages charts.

**Operations**:
- **insert**: Inserts a chart.
  - **Input**: `document` (string), `type` (string, e.g., `bar`), `data` (array of arrays, e.g., `[[1,2],[3,4]]`), `position` (string).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/charts/insert?document=/docs/sample.docx&type=bar&position=paragraph:1" \
      -d '{"data":[[1,2],[3,4]]}'
    ```

**AI Request Examples**:
- User: “Add a bar chart with data [[1,2],[3,4]] to sample.docx.”
  - AI Infers: `POST /word/charts/insert?document=/docs/sample.docx&type=bar&data=[[1,2],[3,4]]`.

**Completions**: Chart types (e.g., `bar`, `line`), data formats.

### word/embedded-objects
**Description**: Manages embedded objects (e.g., PDFs, Excel files).

**Operations**:
- **extractAll**: Extracts all embedded objects.
  - **Input**: `document` (string), `outputDir` (string, directory path).
  - **Output**: `{ success: true, files: string[] }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/embedded-objects/extractAll?document=/docs/Composición.docx&outputDir=/docs/extracted"
    ```
    **Response**:
    ```json
    { "success": true, "files": ["/docs/extracted/embedded_1.pdf"] }
    ```

**AI Request Examples**:
- User: “Extract all embedded files from Composición.docx to /docs/extracted.”
  - AI Infers: `POST /word/embedded-objects/extractAll?document=/docs/Composición.docx&outputDir=/docs/extracted`.

**Completions**: File paths, output directories.

### word/metadata
**Description**: Manages document properties and comments.

**Operations**:
- **set**: Sets a property.
  - **Input**: `document` (string), `property` (string, e.g., `author`), `value` (string).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/metadata/set?document=/docs/sample.docx&property=author&value=JohnDoe"
    ```

**AI Request Examples**:
- User: “Set the author of sample.docx to John Doe.”
  - AI Infers: `POST /word/metadata/set?document=/docs/sample.docx&property=author&value=John Doe`.

**Completions**: Property names (e.g., `author`, `title`).

### word/batch
**Description**: Executes multiple operations in a transaction.

**Operations**:
- **run**: Runs a batch of operations.
  - **Input**: `document` (string), `ops` (array of objects, e.g., `[{ tool: "styles/apply", params: { style: "Heading1" } }]`).
  - **Output**: `{ success: true, results: any[] }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/batch?document=/docs/sample.docx" \
      -d '{"ops":[{"tool":"styles/apply","params":{"style":"Heading1","range":"paragraph:1"}},{"tool":"text/insert","params":{"text":"Hello","position":"paragraph:2"}}]}'
    ```

**AI Request Examples**:
- User: “Apply Heading1 to the first paragraph and insert ‘Hello’ in the second paragraph of sample.docx.”
  - AI Infers: `POST /word/batch?document=/docs/sample.docx` with `ops`.

**Completions**: Tool names, operation parameters.

### word/markdown/import
**Description**: Converts Markdown to Word, applying basic formatting (headings, lists, bold, italic) based on Markdown syntax using COM Interop, with optional template support.

**Operations**:
- **parse**: Converts Markdown to Word.
  - **Input**: `filePath` (string, path to the Markdown file), `template` (string, optional, template path), `output` (string, output path).
  - **Output**: `{ success: true, output: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/markdown/import" \
      -H "Content-Type: application/json" \
      -d '{
        "filePath": "/docs/input.md",
        "template": "/docs/template.docx",
        "output": "/docs/output.docx"
      }'
    ```

**AI Request Examples**:
- User: “Convert input.md to a Word document using template.docx.”
  - AI Infers: `POST /word/markdown/import` with `filePath=/docs/input.md`, `template=/docs/template.docx`, and `output=/docs/output.docx`.

**Completions**: File paths, template paths.

### word/markdown/export
**Description**: Converts Word to Markdown with comment handling.

**Operations**:
- **extract**: Exports to Markdown.
  - **Input**: `document` (string), `output` (string), `comments` (string, `append` or `inline`, default: `append`).
  - **Output**: `{ success: true, output: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/markdown/export?document=/docs/CV.docx&output=/docs/CV.md&comments=append"
    ```

**AI Request Examples**:
- User: “Export CV.docx to Markdown with comments appended.”
  - AI Infers: `POST /word/markdown/export?document=/docs/CV.docx&output=/docs/CV.md&comments=append`.

**Completions**: File paths, comment modes.

### word/mermaid/import
**Description**: Imports and renders Mermaid diagrams.

**Operations**:
- **render**: Renders and inserts a diagram.
  - **Input**: `document` (string), `syntax` (string, Mermaid syntax), `format` (string, `svg` or `png`), `position` (string).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/mermaid/import?document=/docs/tech.docx&syntax=graph TD; A-->B&format=svg&position=paragraph:5"
    ```

**AI Request Examples**:
- User: “Add a Mermaid flowchart to tech.docx at paragraph 5.”
  - AI Infers: `POST /word/mermaid/import?document=/docs/tech.docx&syntax=...&format=svg&position=paragraph:5`.

**Completions**: Mermaid syntax, formats.

### word/mermaid/export
**Description**: Exports Mermaid diagrams from Word.

**Operations**:
- **save**: Exports a diagram.
  - **Input**: `document` (string), `diagram` (number, diagram index), `format` (string, `png` or `svg`), `output` (string).
  - **Output**: `{ success: true, output: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/mermaid/export?document=/docs/tech.docx&diagram=1&format=png&output=/docs/diagram.png"
    ```

**AI Request Examples**:
- User: “Export the first Mermaid diagram in tech.docx as a PNG.”
  - AI Infers: `POST /word/mermaid/export?document=/docs/tech.docx&diagram=1&format=png`.

**Completions**: Diagram indices, formats.

### word/merge
**Description**: Merges multiple Word documents.

**Operations**:
- **merge**: Merges documents with AI conflict resolution.
  - **Input**: `docs` (string, comma-separated paths), `output` (string).
  - **Output**: `{ success: true, output: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/merge?docs=/docs/doc1.docx,/docs/doc2.docx&output=/docs/merged.docx"
    ```

**AI Request Examples**:
- User: “Merge doc1.docx and doc2.docx into one document.”
  - AI Infers: `POST /word/merge?docs=/docs/doc1.docx,/docs/doc2.docx&output=/docs/merged.docx`.

**Completions**: File paths.

### word/template
**Description**: Creates templates with placeholders.

**Operations**:
- **generate**: Generates a template.
  - **Input**: `document` (string), `output` (string).
  - **Output**: `{ success: true, output: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/template?document=/docs/budget.docx&output=/docs/template.docx"
    ```

**AI Request Examples**:
- User: “Create a template from budget.docx with placeholders.”
  - AI Infers: `POST /word/template?document=/docs/budget.docx&output=/docs/template.docx`.

**Completions**: File paths.

### word/reformat
**Description**: Reformats documents for consistency.

**Operations**:
- **apply**: Applies a style set.
  - **Input**: `document` (string), `styleSet` (string, e.g., `Professional`), `output` (string).
  - **Output**: `{ success: true, output: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/reformat?document=/docs/PropuestasRandom.docx&styleSet=Professional&output=/docs/Formatted.docx"
    ```

**AI Request Examples**:
- User: “Reformat PropuestasRandom.docx to look professional.”
  - AI Infers: `POST /word/reformat?document=/docs/PropuestasRandom.docx&styleSet=Professional`.

**Completions**: Style sets (e.g., `Professional`, `Modern`).

### word/analyze
**Description**: Analyzes documents and adds comments.

**Operations**:
- **analyze**: Analyzes and comments.
  - **Input**: `document` (string), `criteria` (string, e.g., `technical`), `output` (string).
  - **Output**: `{ success: true, output: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/analyze?document=/docs/PropuestaTécnica.docx&criteria=technical&output=/docs/Commented.docx"
    ```

**AI Request Examples**:
- User: “Analyze PropuestaTécnica.docx for technical clarity.”
  - AI Infers: `POST /word/analyze?document=/docs/PropuestaTécnica.docx&criteria=technical`.

**Completions**: Criteria (e.g., `technical`, `clarity`).

### word/code-format
**Description**: Formats code with syntax highlighting.

**Operations**:
- **apply**: Formats code snippets.
  - **Input**: `document` (string), `style` (string, e.g., `Código`), `font` (string, e.g., `Consolas`), `output` (string).
  - **Output**: `{ success: true, output: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/code-format?document=/docs/BuenasPrácticas.docx&style=Código&font=Consolas&output=/docs/Formatted.docx"
    ```

**AI Request Examples**:
- User: “Format code in BuenasPrácticas.docx with Consolas font.”
  - AI Infers: `POST /word/code-format?document=/docs/BuenasPrácticas.docx&style=Código&font=Consolas`.

**Completions**: Style names, font names.
### word/image
**Description**: Extracts and inserts images in Word documents. *Note: Requires COM interop; implementation status may vary.*

**Operations**:
- **extract**: Extracts images from a document.
  - **Input**: `document` (string, file path), `outputDir` (string, directory path).
  - **Output**: `{ success: true, files: string[] }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/image/extract?document=/docs/report.docx&outputDir=/images/"
    ```
    **Response**:
    ```json
    { "success": true, "files": ["/images/image1.png"] }
    ```
- **insert**: Inserts an image into a document.
  - **Input**: `document` (string), `imagePath` (string, path to image file), `position` (string, e.g., `paragraph:3`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/image/insert?document=/docs/report.docx&imagePath=/images/logo.png&position=paragraph:3"
    ```

**AI Request Examples**:
- User: "Extract all images from report.docx into the /images folder."
  - AI Infers: `POST /word/image/extract?document=/docs/report.docx&outputDir=/images/`.
- User: "Insert logo.png into report.docx after the third paragraph."
  - AI Infers: `POST /word/image/insert?document=/docs/report.docx&imagePath=/images/logo.png&position=paragraph:3`.

**Completions**: File paths, position ranges.

### word/generate-and-insert-text
**Description**: Uses a server-side LLM to generate text based on a prompt, processes the Markdown output, and inserts it into a Word document with appropriate formatting. The initial summary from the LLM is treated as a Heading 1. *Note: Requires COM interop for insertion and server-side LLM configuration.*

**Operations**:
- **generate**: Generates text based on a prompt and inserts it.
  - **Input**: `filePath` (string, path to the Word file), `position` (string, insertion point, e.g., `start`, `end`, `paragraph:N:start`, `selection`), `prompt` (string, the prompt for text generation), `maxTokens` (number, optional, maximum tokens for generated text).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/generate-and-insert-text" \
      -H "Content-Type: application/json" \
      -d '{
        "filePath": "/docs/report.docx",
        "position": "end",
        "prompt": "Write a concluding paragraph for this report.",
        "maxTokens": 150
      }'
    ```

**Configuration**: This tool requires server-side configuration to connect to an LLM provider. The following environment variables must be set where the Office MCP server is running:
- `LLM_PROVIDER_API_KEY`: Your API key for the LLM provider.
- `LLM_PROVIDER_ENDPOINT`: The API endpoint URL for the LLM provider (e.g., `https://api.openai.com/v1/chat/completions` for OpenAI, or a similar endpoint for other providers).
- `LLM_MODEL_NAME`: (Optional) The name of the specific LLM model to use (e.g., `gpt-4o`, `gemini-1.5-pro`). If not set, a default model may be used by the server-side LLM utility. The tool also enhances the user's prompt to guide the LLM in generating content suitable for Markdown processing.

**AI Request Examples**:
- User: "Generate a concluding paragraph for report.docx and insert it at the end."
  - AI Infers: `POST /word/generate-and-insert-text` with `filePath=/docs/report.docx`, `position=end`, and a suitable `prompt`.
- User: "Write an introductory sentence for report.docx and insert it at the beginning."
  - AI Infers: `POST /word/generate-and-insert-text` with `filePath=/docs/report.docx`, `position=start`, and a suitable `prompt`.

**Completions**: Document paths, position specifiers, prompt text.

---


## Excel Tools

### excel/worksheets
**Description**: Manages worksheets.

**Operations**:
- **add**: Adds a worksheet.
  - **Input**: `document` (string), `name` (string).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/excel/worksheets/add?document=/docs/data.xlsx&name=Data"
    ```

**AI Request Examples**:
- User: “Add a worksheet named ‘Data’ to data.xlsx.”
  - AI Infers: `POST /excel/worksheets/add?document=/docs/data.xlsx&name=Data`.

**Completions**: Worksheet names.

### excel/range
**Description**: Manipulates cell ranges.

**Operations**:
- **write**: Writes values to a range.
  - **Input**: `document` (string), `range` (string, e.g., `A1:B2`), `values` (array of arrays).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/excel/range/write?document=/docs/data.xlsx&range=A1:B2" \
      -d '{"values":[[1,2],[3,4]]}'
    ```

**AI Request Examples**:
- User: “Write [[1,2],[3,4]] to A1:B2 in data.xlsx.”
  - AI Infers: `POST /excel/range/write?document=/docs/data.xlsx&range=A1:B2&values=[[1,2],[3,4]]`.

**Completions**: Range formats, value arrays.

### excel/tables
**Description**: Creates and manages tables.

**Operations**:
- **insert**: Inserts a table.
  - **Input**: `document` (string), `range` (string), `name` (string, optional).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/excel/tables/insert?document=/docs/data.xlsx&range=A1:D10&name=Sales"
    ```

**AI Request Examples**:
- User: “Create a table named ‘Sales’ in A1:D10 of data.xlsx.”
  - AI Infers: `POST /excel/tables/insert?document=/docs/data.xlsx&range=A1:D10&name=Sales`.

**Completions**: Range formats, table names.

### excel/charts
**Description**: Inserts and manages charts.

**Operations**:
- **insert**: Inserts a chart.
  - **Input**: `document` (string), `type` (string, e.g., `column`), `range` (string).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/excel/charts/insert?document=/docs/data.xlsx&type=column&range=A1:B10"
    ```

**AI Request Examples**:
- User: “Add a column chart for A1:B10 in data.xlsx.”
  - AI Infers: `POST /excel/charts/insert?document=/docs/data.xlsx&type=column&range=A1:B10`.

**Completions**: Chart types, range formats.

### excel/data-analysis
**Description**: Processes data (sort, filter, pivot).

**Operations**:
- **sort**: Sorts a range.
  - **Input**: `document` (string), `range` (string), `column` (number), `order` (string, `asc` or `desc`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/excel/data-analysis/sort?document=/docs/data.xlsx&range=A1:D100&column=1&order=asc"
    ```

**AI Request Examples**:
- User: “Sort A1:D100 in data.xlsx by the first column ascending.”
  - AI Infers: `POST /excel/data-analysis/sort?document=/docs/data.xlsx&range=A1:D100&column=1&order=asc`.

**Completions**: Range formats, sort orders.

---

## PowerPoint Tools

### powerpoint/slides
**Description**: Manages slides.

**Operations**:
- **add**: Adds a slide.
  - **Input**: `document` (string), `layout` (string, e.g., `TitleSlide`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/powerpoint/slides/add?document=/docs/presentation.pptx&layout=TitleSlide"
    ```

**AI Request Examples**:
- User: “Add a title slide to presentation.pptx.”
  - AI Infers: `POST /powerpoint/slides/add?document=/docs/presentation.pptx&layout=TitleSlide`.

**Completions**: Layout names.

### powerpoint/shapes
**Description**: Manipulates shapes and text.

**Operations**:
- **insert**: Inserts a shape.
  - **Input**: `document` (string), `type` (string, e.g., `textbox`), `text` (string, optional), `slide` (number).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/powerpoint/shapes/insert?document=/docs/presentation.pptx&type=textbox&text=Hello&slide=1"
    ```

**AI Request Examples**:
- User: “Add a textbox with ‘Hello’ to the first slide of presentation.pptx.”
  - AI Infers: `POST /powerpoint/shapes/insert?document=/docs/presentation.pptx&type=textbox&text=Hello&slide=1`.

**Completions**: Shape types, slide numbers.

### powerpoint/properties
**Description**: Manages presentation settings.

**Operations**:
- **set**: Sets properties.
  - **Input**: `document` (string), `size` (string, e.g., `16:9`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/powerpoint/properties/set?document=/docs/presentation.pptx&size=16:9"
    ```

**AI Request Examples**:
- User: “Set presentation.pptx to 16:9 aspect ratio.”
  - AI Infers: `POST /powerpoint/properties/set?document=/docs/presentation.pptx&size=16:9”.

**Completions**: Size formats.

### powerpoint/animations
**Description**: Manages animations and transitions.

**Operations**:
- **add**: Adds an animation/transition.
  - **Input**: `document` (string), `slide` (number), `type` (string, e.g., `fade`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/powerpoint/animations/add?document=/docs/presentation.pptx&slide=1&type=fade"
    ```

**AI Request Examples**:
- User: “Add a fade transition to the first slide of presentation.pptx.”
  - AI Infers: `POST /powerpoint/animations/add?document=/docs/presentation.pptx&slide=1&type=fade`.

**Completions**: Animation types, slide numbers.

---

## Cross-Application Tools

### office/transfer
**Description**: Transfers data between Office apps.

**Operations**:
- **copy**: Copies data from one app to another.
  - **Input**: `source` (string, e.g., `excel:range:A1:B2`), `target` (string, e.g., `word:paragraph:1`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/office/transfer/copy?source=excel:range:A1:B2&target=word:paragraph:1"
    ```

**AI Request Examples**:
- User: “Copy A1:B2 from data.xlsx to the first paragraph of sample.docx.”
  - AI Infers: `POST /office/transfer/copy?source=excel:range:A1:B2&target=word:paragraph:1`.

**Completions**: Source/target formats.

### office/workflow
**Description**: Executes multi-step automation workflows.

**Operations**:
- **run**: Runs a workflow.
  - **Input**: `steps` (array of objects, e.g., `[{ tool: "excel/tables/insert", params: { range: "A1:D10" } }]`).
  - **Output**: `{ success: true, results: any[] }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/office/workflow/run" \
      -d '{"steps":[{"tool":"excel/tables/insert","params":{"range":"A1:D10"}}]}'
    ```

**AI Request Examples**:
- User: “Create a table in data.xlsx and add a chart.”
  - AI Infers: `POST /office/workflow/run` with `steps`.

**Completions**: Tool names, step parameters.

### office/ai-suggest
**Description**: Provides AI-driven suggestions.

**Operations**:
- **format**: Suggests formatting.
  - **Input**: `context` (string, e.g., `word:document`), `task` (string, e.g., `format`).
  - **Output**: `{ suggestions: any }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/office/ai-suggest?context=word:document&task=format"
    ```

**AI Request Examples**:
- User: “Suggest formatting for sample.docx.”
  - AI Infers: `POST /office/ai-suggest?context=word:document&task=format`.

**Completions**: Task types, context formats.

### office/pdf/export
**Description**: Exports Office documents to PDF.

**Operations**:
- **convert**: Converts to PDF.
  - **Input**: `document` (string), `output` (string).
  - **Output**: `{ success: true, output: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/office/pdf/export?document=/docs/sample.docx&output=/docs/sample.pdf"
    ```

**AI Request Examples**:
- User: “Export sample.docx to PDF.”
  - AI Infers: `POST /office/pdf/export?document=/docs/sample.docx&output=/docs/sample.pdf`.

**Completions**: File paths.

### office/pdf/parse
**Description**: Extracts content from PDFs.

**Operations**:
- **parse**: Extracts text/images.
  - **Input**: `file` (string, PDF path), `type` (string, `text` or `images`).
  - **Output**: `{ success: true, content: any }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/office/pdf/parse?file=/docs/offer.pdf&type=text"
    ```

**AI Request Examples**:
- User: “Extract text from offer.pdf.”
  - AI Infers: `POST /office/pdf/parse?file=/docs/offer.pdf&type=text`.

**Completions**: File paths, content types.

### office/combine
**Description**: Combines Word, Excel, PowerPoint, and PDF into Word.

**Operations**:
- **combine**: Combines files.
  - **Input**: `directory` (string), `output` (string).
  - **Output**: `{ success: true, output: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/office/combine?directory=/docs/Oferta&output=/docs/combined.docx"
    ```

**AI Request Examples**:
- User: “Combine all files in /docs/Oferta into a Word document.”
  - AI Infers: `POST /office/combine?directory=/docs/Oferta&output=/docs/combined.docx`.

**Completions**: Directory paths.

### office/word-to-powerpoint
**Description**: Converts Word to PowerPoint.

**Operations**:
- **convert**: Converts document to slides.
  - **Input**: `document` (string), `output` (string).
  - **Output**: `{ success: true, output: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/office/word-to-powerpoint?document=/docs/PasoAPaso.docx&output=/docs/PasoAPaso.pptx"
    ```

**AI Request Examples**:
- User: “Convert PasoAPaso.docx to a PowerPoint presentation.”
  - AI Infers: `POST /office/word-to-powerpoint?document=/docs/PasoAPaso.docx&output=/docs/PasoAPaso.pptx`.

**Completions**: File paths.

---

## Resource Management

### memory/ai_assistant_guide
**Description**: Provides a static guide for API usage.

**Operations**:
- **read**: Retrieves guide or section.
  - **Input**: `section` (string, optional, e.g., `tool_usage`).
  - **Output**: `{ content: object }`.
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/memory/ai_assistant_guide?section=tool_usage"
    ```
    **Response**:
    ```json
    {
      "content": {
        "word/styles": {
          "description": "Manages styles in Word documents.",
          "example": "POST /word/styles/apply?style=Heading1"
        }
      }
    }
    ```

**AI Request Examples**:
- User: “Show me how to use the Word styles tool.”
  - AI Infers: `GET /memory/ai_assistant_guide?section=tool_usage`.

**Completions**: Section names (e.g., `tool_usage`, `use_cases`).

### dynamic/resources
**Description**: Manages user-generated content (documents, images, etc.).

**Operations**:
- **list**: Lists resources.
  - **Input**: `type` (string, optional, e.g., `docx`), `metadata` (object, optional).
  - **Output**: `{ resources: { id: string, path: string, type: string, metadata: object }[] }`.
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/dynamic/resources/list?type=docx"
    ```
    **Response**:
    ```json
    {
      "resources": [
        { "id": "123", "path": "/docs/doc1.docx", "type": "docx", "metadata": { "author": "John" } }
      ]
    }
    ```
- **read**: Reads a resource.
  - **Input**: `id` (string) or `path` (string).
  - **Output**: `{ path: string, data: Buffer, metadata: object }`.
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/dynamic/resources/read?id=123"
    ```
- **write**: Uploads or modifies a resource.
  - **Input**: `path` (string), `data` (binary or string), `metadata` (object, optional).
  - **Output**: `{ success: true, id: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/dynamic/resources/write?path=/docs/new.docx" \
      -F "data=@local.docx" \
      -F "metadata={\"author\":\"Jane\"}"
    ```

**AI Request Examples**:
- User: “List all Word documents in storage.”
  - AI Infers: `GET /dynamic/resources/list?type=docx`.
- User: “Upload my document new.docx with author Jane.”
  - AI Infers: `POST /dynamic/resources/write?path=/docs/new.docx` with file and metadata.

**Completions**: Resource types, metadata fields.

---

This API specification provides a comprehensive guide for developers and AI agents to interact with the Office MCP Server. Each tool is detailed with clear examples and AI prompt mappings, ensuring seamless integration and automation of Office tasks.
