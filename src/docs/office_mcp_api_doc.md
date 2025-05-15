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
   - [fs/fileContent](#fsfilecontent)
   - [fs/directoryOperations](#fsdirectoryoperations)
   - [fs/structuredData](#fsstructureddata)
3. [Word Tools](#word-tools)
   - [word/styles](#wordstyles)
   - [word/text](#wordtext)
   - [word/search-replace](#wordsearch-replace)
   - [word/page](#wordpage) <!-- Enhanced for getPageCount -->
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
   - [word/applyAutoTitles](#wordapplyautotitles)
   - [word/saveActiveWordAsMarkdown](#wordsaveactivewordasmarkdown)
   - [word/concludeStoryInDocument](#wordconcludestoryindocument)
4. [Excel Tools](#excel-tools)
   - [excel/worksheets](#excelworksheets) <!-- Enhanced for getWorksheetCount -->
   - [excel/range](#excelrange)
   - [excel/tables](#exceltables)
   - [excel/charts](#excelcharts)
   - [excel/data-analysis](#exceldata-analysis)
   - [excel/formatTablesInWorksheet](#excelformattablesinworksheet)
5. [PowerPoint Tools](#powerpoint-tools)
   - [powerpoint/slides](#powerpointslides) <!-- Enhanced for getSlideCount -->
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
   - [office/adaptWordToPowerpoint](#officeadaptwordtopowerpoint)
   - [Feature: Get Document Size](#feature-get-document-size)
7. [Resource Management](#resource-management)
   - [memory/ai_assistant_guide](#memoryai_assistant_guide)
   - [dynamic/resources](#dynamicresources)
8. [OS Tools](#os-tools)
   - [os/getActiveOfficeDocuments](#osgetactiveofficedocuments)

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

<![CDATA[
### <a name="fsfilecontent"></a>fs/fileContent
**Description**: Provides tools for fine-grained manipulation of text and binary file content.

#### fs/insert_text_lines
**Description**: Inserts an array of strings as new lines into a text file at a specified line number, respecting file encoding.
**Parameters**:
- `path` (string, required): The path to the file.
- `lineNumber` (number, integer, required): The 1-based line number before which to insert the new lines. If 0 or less, or greater than the number of lines, lines are appended.
- `lines` (string[], required): An array of strings to insert.
- `encoding` (string, optional, default: `utf8`): The file encoding (e.g., `ascii`, `utf8`, `utf-8`, `utf16le`, `ucs2`, `ucs-2`, `base64`, `latin1`, `binary`, `hex`).
**Returns**: `Promise<void>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/insert_text_lines" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/sample.txt",
    "lineNumber": 2,
    "lines": ["New line 1", "New line 2"],
    "encoding": "utf8"
  }'
```
**Example Response**:
```json
{ "success": true }
```
**AI Request Examples**:
- User: "Insert 'Hello World' as the first line in `data/my_notes.txt`."
- User: "Add these two lines to the end of `config.ini`: `enabled=true` and `mode=test`."
**Completions**: File paths, line numbers, common encodings.

#### fs/delete_text_lines
**Description**: Deletes a specified range of lines from a text file, respecting file encoding.
**Parameters**:
- `path` (string, required): The path to the file.
- `startLine` (number, integer, required): The 1-based starting line number of the range to delete.
- `endLine` (number, integer, required): The 1-based ending line number of the range to delete.
- `encoding` (string, optional, default: `utf8`): The file encoding.
**Returns**: `Promise<void>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/delete_text_lines" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/sample.txt",
    "startLine": 2,
    "endLine": 3,
    "encoding": "utf8"
  }'
```
**Example Response**:
```json
{ "success": true }
```
**AI Request Examples**:
- User: "Delete lines 5 to 10 from `log.txt`."
**Completions**: File paths, line numbers.

#### fs/replace_text_lines
**Description**: Replaces a range of lines in a text file with new lines, respecting file encoding.
**Parameters**:
- `path` (string, required): The path to the file.
- `startLine` (number, integer, required): The 1-based starting line number of the range to replace.
- `endLine` (number, integer, required): The 1-based ending line number of the range to replace.
- `newLines` (string[], required): An array of strings to replace the specified lines.
- `encoding` (string, optional, default: `utf8`): The file encoding.
**Returns**: `Promise<void>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/replace_text_lines" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/sample.txt",
    "startLine": 2,
    "endLine": 2,
    "newLines": ["This is a replaced line."],
    "encoding": "utf8"
  }'
```
**Example Response**:
```json
{ "success": true }
```
**AI Request Examples**:
- User: "Replace the first line of `config.json` with `{\\"version\\": \\"2.0\\"}`."
**Completions**: File paths, line numbers.

#### fs/extract_text_from_range
**Description**: Extracts a text segment from a file defined by character positions within line ranges, respecting file encoding.
**Parameters**:
- `path` (string, required): The path to the file.
- `startLine` (number, integer, required): The 1-based line number where the extraction begins.
- `startChar` (number, integer, required): The 0-based character offset on the `startLine` where extraction begins.
- `endLine` (number, integer, required): The 1-based line number where the extraction ends.
- `endChar` (number, integer, required): The 0-based character offset on the `endLine` where extraction ends (exclusive).
- `encoding` (string, optional, default: `utf8`): The file encoding.
**Returns**: `Promise<string>` - The extracted text.
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/extract_text_from_range" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/sample.txt",
    "startLine": 1,
    "startChar": 6,
    "endLine": 2,
    "endChar": 5
  }'
```
**Example Response**:
```json
{ "success": true, "data": "text from file..." }
```
**AI Request Examples**:
- User: "Get the text from line 2, character 5 to line 3, character 10 of `report.txt`."
**Completions**: File paths, line numbers, character offsets.

#### fs/search_replace_in_text_range
**Description**: Searches (string/regex) and replaces text within a specified line range of a file, respecting file encoding.
**Parameters**:
- `path` (string, required): The path to the file.
- `searchTerm` (string, required): The text or regex pattern to search for.
- `replacement` (string, required): The text to replace matches with.
- `startLine` (number, integer, optional, default: 1): The 1-based starting line of the range for search/replace.
- `endLine` (number, integer, optional, default: end of file): The 1-based ending line of the range.
- `isRegex` (boolean, optional, default: `false`): Whether `searchTerm` is a regex.
- `replaceAll` (boolean, optional, default: `true`): Whether to replace all occurrences or just the first on each line (or overall if regex global flag is used).
- `encoding` (string, optional, default: `utf8`): The file encoding.
**Returns**: `Promise<{ replacementsMade: number }>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/search_replace_in_text_range" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/sample.txt",
    "searchTerm": "old_text",
    "replacement": "new_text",
    "startLine": 1,
    "endLine": 10
  }'
```
**Example Response**:
```json
{ "success": true, "data": { "replacementsMade": 5 } }
```
**AI Request Examples**:
- User: "In `main.py`, replace all instances of 'foo' with 'bar' between lines 10 and 20."
**Completions**: File paths, search terms, line numbers.

#### fs/sort_text_lines
**Description**: Sorts the lines of a text file.
**Parameters**:
- `path` (string, required): The path to the file.
- `options` (object, optional):
    - `reverse` (boolean, optional, default: `false`): Sort in descending order.
    - `caseSensitive` (boolean, optional, default: `true`): Perform case-sensitive sort.
    - `locale` (string, optional, default: `en`): Locale for string comparison (e.g., `en-US`).
- `encoding` (string, optional, default: `utf8`): The file encoding.
**Returns**: `Promise<void>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/sort_text_lines" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/unsorted.txt",
    "options": { "reverse": true, "caseSensitive": false }
  }'
```
**Example Response**:
```json
{ "success": true }
```
**AI Request Examples**:
- User: "Sort the lines in `names.txt` alphabetically, ignoring case."
**Completions**: File paths, sort options.

#### fs/deduplicate_consecutive_lines
**Description**: Removes consecutive duplicate lines from a text file.
**Parameters**:
- `path` (string, required): The path to the file.
- `caseSensitive` (boolean, optional, default: `true`): Perform case-sensitive comparison.
- `encoding` (string, optional, default: `utf8`): The file encoding.
**Returns**: `Promise<void>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/deduplicate_consecutive_lines" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/log_with_duplicates.txt",
    "caseSensitive": false
  }'
```
**Example Response**:
```json
{ "success": true }
```
**AI Request Examples**:
- User: "Remove duplicate adjacent lines from `output.log`."
**Completions**: File paths, case sensitivity option.

#### fs/trim_line_whitespace
**Description**: Trims leading and/or trailing whitespace from each line in a text file.
**Parameters**:
- `path` (string, required): The path to the file.
- `options` (object, optional):
    - `leading` (boolean, optional, default: `true`): Trim leading whitespace.
    - `trailing` (boolean, optional, default: `true`): Trim trailing whitespace.
- `encoding` (string, optional, default: `utf8`): The file encoding.
**Returns**: `Promise<void>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/trim_line_whitespace" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/messy_code.txt",
    "options": { "leading": true, "trailing": true }
  }'
```
**Example Response**:
```json
{ "success": true }
```
**AI Request Examples**:
- User: "Trim all whitespace from the beginning and end of each line in `data.csv`."
**Completions**: File paths, trim options.

#### fs/read_binary_as_hex
**Description**: Reads a binary file and returns its content as a hexadecimal string.
**Parameters**:
- `path` (string, required): The path to the binary file.
**Returns**: `Promise<string>` - The hex string.
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/read_binary_as_hex" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/image.png"
  }'
```
**Example Response**:
```json
{ "success": true, "data": "89504e470d0a1a0a..." }
```
**AI Request Examples**:
- User: "Read `firmware.bin` as a hex string."
**Completions**: File paths.

#### fs/write_hex_as_binary
**Description**: Writes a hexadecimal string to a file as binary data.
**Parameters**:
- `path` (string, required): The path to the file to write.
- `hexString` (string, required): The hexadecimal string to write.
**Returns**: `Promise<void>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/write_hex_as_binary" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/output.bin",
    "hexString": "48656c6c6f20576f726c64"
  }'
```
**Example Response**:
```json
{ "success": true }
```
**AI Request Examples**:
- User: "Write the hex string '01020304' to `data.dat`."
**Completions**: File paths.

#### fs/read_base64_file
**Description**: Reads a Base64 encoded file and returns its decoded content as a string or binary buffer.
**Parameters**:
- `path` (string, required): The path to the Base64 encoded file.
- `outputEncoding` (string, optional, default: `utf8`): The encoding for the output string (e.g., `utf8`, `ascii`, `latin1`), or `'binary'` to return a Buffer.
**Returns**: `Promise<string | Buffer>` - The decoded content.
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/read_base64_file" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/encoded_text.b64",
    "outputEncoding": "utf8"
  }'
```
**Example Response**:
```json
{ "success": true, "data": "Decoded text content..." }
```
**AI Request Examples**:
- User: "Read and decode `secret.b64`."
**Completions**: File paths, output encodings.

#### fs/write_to_base64_file
**Description**: Encodes string data or binary data to Base64 and writes it to a file.
**Parameters**:
- `path` (string, required): The path to the file to write the Base64 string.
- `data` (string | Buffer, required): The string data or Buffer to encode.
- `inputEncoding` (string, optional, default: `utf8`): If `data` is a string, this is its encoding.
**Returns**: `Promise<void>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/write_to_base64_file" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "docs/data.b64",
    "data": "This will be base64 encoded",
    "inputEncoding": "utf8"
  }'
```
**Example Response**:
```json
{ "success": true }
```
**AI Request Examples**:
- User: "Encode the string 'My Data' and save it to `mydata.b64`."
**Completions**: File paths, input encodings.

### <a name="fsdirectoryoperations"></a>fs/directoryOperations
**Description**: Tools for listing directory contents, finding files, and generating directory tree structures.

#### fs/list_directory_contents
**Description**: Lists the contents of a specified directory, optionally including file sizes and recursion.
**Parameters**:
- `path` (string, required): Path to the directory.
- `options` (object, optional):
    - `recursive` (boolean, optional, default: `false`): List contents recursively.
    - `includeSize` (boolean, optional, default: `false`): Include file sizes in the output.
    - `maxDepth` (number, integer, optional): Maximum depth for recursion (0 for top-level only).
**Returns**: `Promise<ListDirectoryItem[]>` where `ListDirectoryItem` is `{ name: string, type: 'file' | 'directory' | 'other', path: string, size?: number, depth: number }`.
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/list_directory_contents" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "./src/tools",
    "options": { "recursive": true, "maxDepth": 1, "includeSize": true }
  }'
```
**Example Response**:
```json
{
  "success": true,
  "data": [
    { "name": "fs", "type": "directory", "path": "c:\\Users\\David\\Documents\\MCP\\mcp-office\\src\\tools\\fs", "size": 0, "depth": 0 },
    { "name": "index.ts", "type": "file", "path": "c:\\Users\\David\\Documents\\MCP\\mcp-office\\src\\tools\\index.ts", "size": 1024, "depth": 0 }
  ]
}
```
**AI Request Examples**:
- User: "List all files and folders in `/myproject/src` recursively up to 2 levels deep."
**Completions**: Directory paths, recursive options.

#### fs/find_files
**Description**: Searches for files within a directory (and its subdirectories) based on a name pattern (glob or regex) and optionally content.
**Parameters**:
- `directoryPath` (string, required): The directory to search within.
- `namePattern` (string, required): Glob or regex pattern for file names (e.g., `*.ts`, `^test_.*\\.js$`).
- `options` (object, optional):
    - `isRegexName` (boolean, optional, default: `false`): Treat `namePattern` as regex.
    - `recursive` (boolean, optional, default: `true`): Search recursively.
    - `contentPattern` (string, optional): Regex pattern to search within file content.
    - `contentIsRegex` (boolean, optional, default: `false`): Treat `contentPattern` as regex (currently implies regex, simple string search not main focus).
    - `maxDepth` (number, integer, optional): Maximum depth for recursion.
    - `encoding` (string, optional, default: `utf8`): Encoding for reading file content.
**Returns**: `Promise<FindFilesResultItem[]>` where `FindFilesResultItem` is `{ path: string, matchesContent?: boolean }`. `matchesContent` is true if `contentPattern` matched, false if provided but not matched, undefined if `contentPattern` was not provided.
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/find_files" \
  -H "Content-Type: application/json" \
  -d '{
    "directoryPath": "./tests",
    "namePattern": "*.test.ts",
    "options": { "recursive": true, "contentPattern": "describe" }
  }'
```
**Example Response**:
```json
{
  "success": true,
  "data": [
    { "path": "c:\\Users\\David\\Documents\\MCP\\mcp-office\\tests\\unit\\fs\\fileContent.test.ts", "matchesContent": true }
  ]
}
```
**AI Request Examples**:
- User: "Find all TypeScript files in `src` that contain the word 'interface'."
**Completions**: Directory paths, name patterns, content patterns.

#### fs/get_directory_tree
**Description**: Generates a string representation of a directory tree structure.
**Parameters**:
- `path` (string, required): The root directory path for the tree.
- `options` (object, optional):
    - `maxDepth` (number, integer, optional, default: 3): Maximum depth of the tree.
    - `includeFiles` (boolean, optional, default: `true`): Include files in the tree.
    - `includeSize` (boolean, optional, default: `false`): Include file sizes.
**Returns**: `Promise<string>` - The directory tree as a string.
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/get_directory_tree" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "./src",
    "options": { "maxDepth": 2, "includeFiles": true }
  }'
```
**Example Response**:
```json
{
  "success": true,
  "data": "src/\\n├── docs/\\n│   ├── ai_assistant_guide.md\\n│   └── ...\\n├── server/\\n│   └── index.ts\\n└── tools/\\n    ├── fs/\\n    └── index.ts"
}
```
**AI Request Examples**:
- User: "Show me the directory structure of `./project/assets` up to 3 levels."
**Completions**: Directory paths, depth options.

### <a name="fsstructureddata"></a>fs/structuredData
**Description**: Tools for reading and writing structured data files like CSV, JSON, and manipulating HTML/XML content.

#### fs/read_csv_data
**Description**: Reads data from a CSV file, allowing selection of specific columns and row ranges.
**Parameters**:
- `path` (string, required): Path to the CSV file.
- `options` (object, optional):
    - `delimiter` (string, optional): CSV delimiter character.
    - `columns` (string[] | number[], optional): Array of column names (if `hasHeaders` is true) or 0-based indices to select.
    - `startRow` (number, integer, optional, default: 1): 1-based row number to start reading from.
    - `endRow` (number, integer, optional): 1-based row number to end reading at (exclusive).
    - `hasHeaders` (boolean, optional, default: `true`): Whether the CSV has a header row.
    - `encoding` (string, optional, default: `utf8`): File encoding.
**Returns**: `Promise<Array<Record<string, any> | any[]>>` - Array of objects (if headers) or array of arrays.
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/read_csv_data" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "data/sales.csv",
    "options": { "hasHeaders": true, "columns": ["Product", "Revenue"], "startRow": 2 }
  }'
```
**Example Response**:
```json
{
  "success": true,
  "data": [
    { "Product": "Widget A", "Revenue": 1000 },
    { "Product": "Widget B", "Revenue": 1500 }
  ]
}
```
**AI Request Examples**:
- User: "Read `products.csv`, get only the 'ID' and 'Price' columns, starting from the second row."
**Completions**: File paths, CSV options.

#### fs/write_csv_data
**Description**: Writes an array of objects or arrays of data to a CSV file.
**Parameters**:
- `path` (string, required): Path to the CSV file to write.
- `data` (Array<Record<string, any> | any[]>, required): Data to write.
- `options` (object, optional):
    - `delimiter` (string, optional, default: `,`): CSV delimiter.
    - `headers` (string[], optional): Array of header names. If not provided and data is objects, keys of the first object are used.
    - `includeHeaders` (boolean, optional, default: `true` if headers can be determined, else `false`): Whether to write a header row.
    - `encoding` (string, optional, default: `utf8`): File encoding.
    - `mode` (enum: `'overwrite'` | `'append'`, optional, default: `'overwrite'`): Write mode.
**Returns**: `Promise<void>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/write_csv_data" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "output/report.csv",
    "data": [{ "id": 1, "name": "Alice" }, { "id": 2, "name": "Bob" }],
    "options": { "headers": ["id", "name"], "mode": "overwrite" }
  }'
```
**Example Response**:
```json
{ "success": true }
```
**AI Request Examples**:
- User: "Write this data to `new_data.csv`: `[[1,'A'],[2,'B']]` with headers 'Col1', 'Col2'."
**Completions**: File paths, CSV options, data structure.

#### fs/read_json_path
**Description**: Reads a specific value from a JSON file using a JSONPath expression.
**Parameters**:
- `path` (string, required): Path to the JSON file.
- `jsonPath` (string, required): JSONPath expression (e.g., `$.store.book[0].title`).
- `options` (object, optional):
    - `encoding` (string, optional, default: `utf8`): File encoding.
**Returns**: `Promise<any | undefined>` - The extracted value, or undefined if not found.
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/read_json_path" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "config/settings.json",
    "jsonPath": "$.user.preferences.theme"
  }'
```
**Example Response**:
```json
{ "success": true, "data": "dark" }
```
**AI Request Examples**:
- User: "From `package.json`, get the value of `dependencies.lodash`."
**Completions**: File paths, JSONPath expressions.

#### fs/write_json_path
**Description**: Writes or updates a value at a specific path within a JSON file. Creates missing path segments if `createMissing` is true.
**Parameters**:
- `path` (string, required): Path to the JSON file.
- `jsonPath` (string, required): JSONPath expression to the location to write.
- `value` (any, required): The value to write.
- `options` (object, optional):
    - `createMissing` (boolean, optional, default: `true`): Create parent objects/arrays if they don't exist.
    - `encoding` (string, optional, default: `utf8`): File encoding.
**Returns**: `Promise<void>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/write_json_path" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "config/settings.json",
    "jsonPath": "$.user.preferences.fontSize",
    "value": 14
  }'
```
**Example Response**:
```json
{ "success": true }
```
**AI Request Examples**:
- User: "In `manifest.json`, set `version` to `1.2.3`."
**Completions**: File paths, JSONPath expressions, values.

#### fs/extract_from_markup
**Description**: Extracts data (text, attribute, or HTML) from an HTML or XML file using a CSS selector.
**Parameters**:
- `path` (string, required): Path to the HTML/XML file.
- `selector` (string, required): CSS selector (e.g., `h1`, `.item > span`, `div#main`).
- `options` (object, optional):
    - `extract` (string | object, optional, default: `'text'`): What to extract:
        - `'text'`: Inner text of selected elements.
        - `'html'`: Outer HTML of selected elements.
        - `{ attribute: string }`: Value of a specific attribute (e.g., `{ attribute: 'href' }`).
    - `encoding` (string, optional, default: `utf8`): File encoding.
    - `isXml` (boolean, optional, default: `false`): Parse as XML.
**Returns**: `Promise<string | string[] | null>` - A single string if one match, array of strings for multiple matches, or null if no matches.
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/extract_from_markup" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "web/index.html",
    "selector": "title",
    "options": { "extract": "text" }
  }'
```
**Example Response**:
```json
{ "success": true, "data": "My Web Page" }
```
**AI Request Examples**:
- User: "Get the text of all `<h2>` tags from `article.html`."
- User: "Extract the `src` attribute of the image with id `logo` in `page.xml`."
**Completions**: File paths, CSS selectors, extraction options.

#### fs/update_markup_content
**Description**: Updates the content (text or HTML) or an attribute of elements matching a CSS selector in an HTML/XML file.
**Parameters**:
- `path` (string, required): Path to the HTML/XML file.
- `selector` (string, required): CSS selector for elements to update.
- `newContent` (string, required): The new text, HTML, or attribute value.
- `options` (object, optional):
    - `updateType` (string | object, optional, default: `'text'`): What to update:
        - `'text'`: Set inner text.
        - `'html'`: Set inner HTML.
        - `{ attribute: string }`: Set a specific attribute (e.g., `{ attribute: 'class' }`).
    - `encoding` (string, optional, default: `utf8`): File encoding.
    - `isXml` (boolean, optional, default: `false`): Parse as XML.
**Returns**: `Promise<{ modifiedCount: number }>`
**Example**:
```bash
curl -X POST "http://localhost:3000/fs/update_markup_content" \
  -H "Content-Type: application/json" \
  -d '{
    "path": "web/index.html",
    "selector": "h1#main-title",
    "newContent": "Welcome to the New Site!",
    "options": { "updateType": "text" }
  }'
```
**Example Response**:
```json
{ "success": true, "data": { "modifiedCount": 1 } }
```
**AI Request Examples**:
- User: "Change the text of the paragraph with class `intro` in `about.html` to 'New introduction'."
- User: "Set the `href` attribute of all links with class `external-link` in `links.html` to `https://example.com`."
**Completions**: File paths, CSS selectors, update types.

]]>
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
**Description**: Configures page layout and provides page count information.

**Operations**:
- **set**: Sets page properties.
  - **Input**: `document` (string), `size` (string, e.g., `A4`), `margins` (number, inches), `orientation` (string, `portrait` or `landscape`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/page/set?document=/docs/sample.docx&size=A4&margins=1&orientation=portrait"
    ```
- **getPageCount**: Retrieves the total number of pages in the document.
  - **Input**: `filePath` (string, path to the Word document). `useComInterop` (boolean, optional, default: `true`).
  - **Output**: `{ success: boolean, count?: number, error?: string }`.
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/word/page/getPageCount?filePath=/docs/my_epic_novel.docx"
    ```
    **Response**:
    ```json
    { "success": true, "count": 342 }
    ```

**AI Request Examples**:
- User: “Set sample.docx to A4 with 1-inch margins.”
  - AI Infers: `POST /word/page/set?document=/docs/sample.docx&size=A4&margins=1`.
- User: "How many pages are in 'annual_report.docx'?"
  - AI Infers: `GET /word/page/getPageCount?filePath=/docs/annual_report.docx`.

**Completions**: Page sizes (e.g., `A4`, `Letter`), orientations, file paths.

**When You Need to Know if It's a Leaflet or a Tome (Word Edition)**:
Before you hit print on that "brief" document, or try to impress your boss with its "conciseness," use `getPageCount`. "MCP, tell me, is 'my_quick_memo.docx' actually quick, or have I written another dissertation? Get the page count!"

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
**Description**: Converts Word to Markdown using COM Interop. Currently performs basic text extraction with comment handling (formatting, tables, and images are not preserved). **Enhancements are planned** to support rich conversion, including formatting, tables (as Markdown or HTML), image extraction to a specified directory, and optional ZIP packaging.

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

### <a name="wordapplyautotitles"></a>word/applyAutoTitles
**Description**: Unleashes the power of AI to automatically identify and apply appropriate heading styles (Heading 1, Heading 2, etc.) to an active Word document. It's like having a tiny, style-obsessed robot editor living in your computer.

**Operations**:
- **apply**: Identifies potential titles and subtitles and applies heading styles.
  - **Input**: `filePath` (string, path to the active Word document). `useComInterop` (boolean, optional, default: `true`).
  - **Output**: `{ success: boolean, changesApplied?: number, error?: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/applyAutoTitles" \
      -H "Content-Type: application/json" \
      -d '{
        "filePath": "/docs/my_rambling_manifesto.docx"
      }'
    ```
    **Response**:
    ```json
    { "success": true, "changesApplied": 12 }
    ```

**AI Request Examples**:
- User: "Can you automatically style the headings in 'my_thesis_draft.docx'?"
  - AI Infers: `POST /word/applyAutoTitles` with `filePath`.
- User: "This document is a mess. Make the titles look like titles in 'report_final.docx'."
  - AI Infers: `POST /word/applyAutoTitles` with `filePath`.

**Completions**: File paths.

**When Your Document Looks Like a Wall of Text from a Fever Dream**:
You've poured your heart and soul into writing, but it's structurally... abstract. Before your reader needs a map and compass to navigate your thoughts, let `applyAutoTitles` work its magic. "MCP, make this readable! Apply some auto-titles to 'stream_of_consciousness.docx'."

### <a name="wordsaveactivewordasmarkdown"></a>word/saveActiveWordAsMarkdown
**Description**: Converts the content of an active Word document into a beautiful, clean Markdown file. Because sometimes, `.docx` is just too mainstream.

**Operations**:
- **save**: Saves the active Word document as Markdown.
  - **Input**: `filePath` (string, path to the active Word document), `outputMarkdownPath` (string, path for the new .md file).
  - **Output**: `{ success: boolean, outputPath?: string, error?: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/saveActiveWordAsMarkdown" \
      -H "Content-Type: application/json" \
      -d '{
        "filePath": "/docs/my_blog_post.docx",
        "outputMarkdownPath": "/docs/my_blog_post.md"
      }'
    ```
    **Response**:
    ```json
    { "success": true, "outputPath": "/docs/my_blog_post.md" }
    ```

**AI Request Examples**:
- User: "Convert 'latest_article.docx' to Markdown and save it."
  - AI Infers: `POST /word/saveActiveWordAsMarkdown` with `filePath` and an inferred `outputMarkdownPath`.
- User: "I need 'meeting_notes.docx' as 'notes.md'."
  - AI Infers: `POST /word/saveActiveWordAsMarkdown` with `filePath="meeting_notes.docx"` and `outputMarkdownPath="notes.md"`.

**Completions**: File paths.

**When You're Too Cool for WYSIWYG**:
You've crafted a masterpiece in Word, but now you need to post it on your ultra-hip, text-only blog. Fear not, `saveActiveWordAsMarkdown` is here to translate your fancy formatting into pure, unadulterated Markdown. "MCP, liberate 'my_novel.docx' into 'my_novel.md'!"

### <a name="wordconcludestoryindocument"></a>word/concludeStoryInDocument
**Description**: Employs an AI to craft a fitting conclusion for a story within a specified Word document. Defaults to working its narrative magic on 'Historias de luis.docx' if you're feeling particularly unspecific or if Luis has been prolific again.

**Operations**:
- **conclude**: Generates and inserts a story conclusion.
  - **Input**: `filePath` (string, optional, path to the Word document, defaults to 'Historias de luis.docx' in a predefined user documents folder if not provided or if the provided path is not found), `storyContextPrompt` (string, optional, additional context for the AI to better understand the story, e.g., "The story is a space opera about a heroic hamster.").
  - **Output**: `{ success: boolean, message?: string, error?: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/word/concludeStoryInDocument" \
      -H "Content-Type: application/json" \
      -d '{
        "filePath": "/stories/chapter_final.docx",
        "storyContextPrompt": "It's a detective noir, and the butler definitely didn't do it."
      }'
    ```
    **Response**:
    ```json
    { "success": true, "message": "Conclusion generated and inserted successfully." }
    ```

**AI Request Examples**:
- User: "Help me finish my story in 'my_adventure_novel.docx'. It's about a pirate who finds a cursed treasure."
  - AI Infers: `POST /word/concludeStoryInDocument` with `filePath` and `storyContextPrompt`.
- User: "Just wrap up whatever Luis is writing in his story document."
  - AI Infers: `POST /word/concludeStoryInDocument` (using default filePath).

**Completions**: File paths, story context snippets.

**When Your Muse Abandons You at Chapter 27**:
You've written a masterpiece, but the ending is... elusive. Instead of staring blankly at the cursor, let the AI ghostwriter step in. "MCP, please, conclude 'TheNeverEndingTale.docx'. My brain is soup." It might not be Shakespeare, but it'll be *an* ending.

---

## Excel Tools

### excel/worksheets
**Description**: Manages worksheets and provides worksheet count information.

**Operations**:
- **add**: Adds a worksheet.
  - **Input**: `document` (string), `name` (string).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/excel/worksheets/add?document=/docs/data.xlsx&name=Data"
    ```
- **getWorksheetCount**: Retrieves the total number of worksheets in the workbook.
  - **Input**: `filePath` (string, path to the Excel workbook). `useComInterop` (boolean, optional, default: `true`).
  - **Output**: `{ success: boolean, count?: number, error?: string }`.
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/excel/worksheets/getWorksheetCount?filePath=/spreadsheets/annual_data.xlsx"
    ```
    **Response**:
    ```json
    { "success": true, "count": 12 }
    ```

**AI Request Examples**:
- User: “Add a worksheet named ‘Data’ to data.xlsx.”
  - AI Infers: `POST /excel/worksheets/add?document=/docs/data.xlsx&name=Data`.
- User: "How many sheets are in 'financial_model.xlsx'?"
  - AI Infers: `GET /excel/worksheets/getWorksheetCount?filePath=/spreadsheets/financial_model.xlsx`.

**Completions**: Worksheet names, file paths.

**When You Suspect Your Workbook is Secretly a Labyrinth (Excel Edition)**:
Is it a simple spreadsheet or a multi-dimensional data maze? Before you get lost in a sea of tabs, use `getWorksheetCount`. "MCP, how many rabbit holes... I mean, worksheets... are in 'Project_Omega_Data_v73.xlsx'?"

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

### <a name="excelformattablesinworksheet"></a>excel/formatTablesInWorksheet
**Description**: Scans a specified Excel worksheet for tables and applies AI-suggested formatting to make them look presentable. If you don't specify, it bravely tackles the second sheet of whatever Excel file is currently active. Because life's too short for ugly spreadsheets.

**Operations**:
- **format**: Identifies and formats tables in a worksheet.
  - **Input**: `filePath` (string, optional, path to the Excel workbook. If not provided, attempts to use the active document), `sheetIdentifier` (string | number, optional, name or 1-based index of the worksheet. Defaults to the second sheet if not provided).
  - **Output**: `{ success: boolean, tablesFormatted?: number, error?: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/excel/formatTablesInWorksheet" \
      -H "Content-Type: application/json" \
      -d '{
        "filePath": "/data/quarterly_results.xlsx",
        "sheetIdentifier": "Sheet1"
      }'
    ```
    **Response**:
    ```json
    { "success": true, "tablesFormatted": 3 }
    ```

**AI Request Examples**:
- User: "Make the tables in 'SalesData.xlsx' on the 'Q3_Data' sheet look nice."
  - AI Infers: `POST /excel/formatTablesInWorksheet` with `filePath` and `sheetIdentifier`.
- User: "Spruce up the tables on the second sheet of this open Excel file."
  - AI Infers: `POST /excel/formatTablesInWorksheet` (using defaults for active file and second sheet).

**Completions**: File paths, sheet names/indices.

**When Your Data Looks Like It Dressed Itself in the Dark**:
Your numbers are solid, but your tables are an eyesore. Before you blind your colleagues with unformatted data, let `formatTablesInWorksheet` give it an AI-powered makeover. "MCP, please apply some formatting wizardry to the tables in 'budget_draft_v17.xlsx', sheet 'PainfulNumbers'."

---

## PowerPoint Tools

### powerpoint/slides
**Description**: Manages slides and provides slide count information.

**Operations**:
- **add**: Adds a slide.
  - **Input**: `document` (string), `layout` (string, e.g., `TitleSlide`).
  - **Output**: `{ success: true }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/powerpoint/slides/add?document=/docs/presentation.pptx&layout=TitleSlide"
    ```
- **getSlideCount**: Retrieves the total number of slides in the presentation.
  - **Input**: `filePath` (string, path to the PowerPoint presentation). `useComInterop` (boolean, optional, default: `true`).
  - **Output**: `{ success: boolean, count?: number, error?: string }`.
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/powerpoint/slides/getSlideCount?filePath=/presentations/keynote_address.pptx"
    ```
    **Response**:
    ```json
    { "success": true, "count": 78 }
    ```

**AI Request Examples**:
- User: “Add a title slide to presentation.pptx.”
  - AI Infers: `POST /powerpoint/slides/add?document=/docs/presentation.pptx&layout=TitleSlide`.
- User: "How many slides are in 'marketing_pitch_final.pptx'?"
  - AI Infers: `GET /powerpoint/slides/getSlideCount?filePath=/presentations/marketing_pitch_final.pptx`.

**Completions**: Layout names, file paths.

**When Your "Quick Update" Presentation Has More Slides Than a Feature Film (PowerPoint Edition)**:
You promised a "short and sweet" presentation. Before you accidentally subject your audience to an epic saga, check the slide count. "MCP, how many slides did I actually make for 'FiveMinuteIntro.pptx'? Be honest."

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

### <a name="officeadaptwordtopowerpoint"></a>office/adaptWordToPowerpoint
**Description**: Magically (or through clever programming) converts an active Word document into a PowerPoint presentation. It even includes a placeholder for future AI enhancements, because the future is now...ish.

**Operations**:
- **convert**: Converts a Word document to a PowerPoint presentation.
  - **Input**: `wordFilePath` (string, path to the source Word document), `powerpointFilePath` (string, optional, path for the output .pptx file. If not provided, it's derived from the Word file name).
  - **Output**: `{ success: boolean, outputPath?: string, error?: string }`.
  - **Example**:
    ```bash
    curl -X POST "http://localhost:3000/office/adaptWordToPowerpoint" \
      -H "Content-Type: application/json" \
      -d '{
        "wordFilePath": "/reports/annual_summary.docx",
        "powerpointFilePath": "/presentations/annual_summary_slides.pptx"
      }'
    ```
    **Response**:
    ```json
    { "success": true, "outputPath": "/presentations/annual_summary_slides.pptx" }
    ```

**AI Request Examples**:
- User: "Turn 'my_detailed_report.docx' into a PowerPoint presentation."
  - AI Infers: `POST /office/adaptWordToPowerpoint` with `wordFilePath`.
- User: "I need slides from 'project_outline.docx', save it as 'project_kickoff.pptx'."
  - AI Infers: `POST /office/adaptWordToPowerpoint` with `wordFilePath` and `powerpointFilePath`.

**Completions**: File paths.

**When Your 20-Page Report Needs to Become a 5-Minute Presentation Yesterday**:
You wrote a novel, but they want a slideshow. Don't panic! `adaptWordToPowerpoint` will attempt to distill your magnum opus into digestible slides. Results may vary based on your Word document's structure and the current mood of the AI gods. "MCP, please, turn 'WarAndPeace_abridged.docx' into 'WarAndPeace_the_slideshow.pptx'. And make it snappy!"

### <a name="feature-get-document-size"></a>Feature: Get Document Size (Know Before You Commit!)
**Description**: Ever wondered if that "quick" document is actually a 500-page epic, or if your "brief" presentation has more slides than a corporate retreat? This feature lets you quickly find out the scale of your Office documents.
    *   **Word**: Get the page count using [`word/page/getPageCount`](#wordpage).
    *   **Excel**: Get the worksheet count using [`excel/worksheets/getWorksheetCount`](#excelworksheets).
    *   **PowerPoint**: Get the slide count using [`powerpoint/slides/getSlideCount`](#powerpointslides).

**When You're About to Bite Off More Than You Can Chew**:
You've been asked to "quickly review" a document. Before you dive in, get its vital statistics. "MCP, how many pages is 'WarAndPeace_Final_Final_ReallyFinal.docx'?" or "MCP, just how many slides are in 'MyLifeStory_Part1.pptx'?" This way, you can manage expectations (mostly your own).

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

## OS Tools

### <a name="osgetactiveofficedocuments"></a>os/getActiveOfficeDocuments
**Description**: Ever feel like you're juggling too many Office files and can't remember what's open? This tool is your personal digital assistant, peeking over your shoulder (with permission, of course) to list all currently open Microsoft Office documents (Word, Excel, PowerPoint) and their file paths.

**Operations**:
- **list**: Retrieves a list of active Office documents.
  - **Input**: None.
  - **Output**: `{ success: boolean, documents?: { type: string, path: string }[], error?: string }`.
    - `type`: Can be 'Word', 'Excel', or 'PowerPoint'.
    - `path`: Full file path of the open document.
  - **Example**:
    ```bash
    curl -X GET "http://localhost:3000/os/getActiveOfficeDocuments"
    ```
    **Response**:
    ```json
    {
      "success": true,
      "documents": [
        { "type": "Word", "path": "C:\\Users\\David\\Documents\\Reports\\Q1_Sales_Report.docx" },
        { "type": "Excel", "path": "C:\\Users\\David\\Spreadsheets\\Financial_Projections_v3.xlsx" }
      ]
    }
    ```

**AI Request Examples**:
- User: "What Office documents do I have open right now?"
  - AI Infers: `GET /os/getActiveOfficeDocuments`.
- User: "Can you check if I left that important PowerPoint presentation open?"
  - AI Infers: `GET /os/getActiveOfficeDocuments` (and then AI would filter/check the list).

**Completions**: None applicable for input.

**When Your Desktop is a Digital Minefield**:
You have 17 Word docs, 12 Excel sheets, and 5 PowerPoints open. One of them is critical. Which one? Instead of clicking through each window like a digital archaeologist, just ask! "MCP, what unholy mess of Office files have I created this time? `getActiveOfficeDocuments` please!"

---

This API specification provides a comprehensive guide for developers and AI agents to interact with the Office MCP Server. Each tool is detailed with clear examples and AI prompt mappings, ensuring seamless integration and automation of Office tasks.
