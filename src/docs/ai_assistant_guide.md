# Guide to AI Wizard - Office MCP Server

This guide provides AI assistants with quick-start instructions to use the Office MCP (Model Context Protocol) Server tools effectively, enabling automation of Microsoft Office tasks (Word, Excel, PowerPoint) and file/resource management. It covers available tools, how to invoke them, and references to static and dynamic resources for deeper guidance.

**Target Audience**: AI assistants (e.g., Grok) processing user requests for MCP Server operations.

**Date**: May 01, 2025

---

## Table of Contents
1. [Overview](#overview)
2. [Available Tools](#available-tools)
3. [How to Use the Tools](#how-to-use-the-tools)
4. [Resource References](#resource-references)
   - [Static Resources](#static-resources)
   - [Dynamic Resources](#dynamic-resources)

---

## Overview

The Office MCP Server provides a set of tools for automating Microsoft Office applications (Word, Excel, PowerPoint) and managing files/resources. AI assistants translate user prompts (e.g., “Merge two Word documents”) into MCP tool invocations (e.g., `word.merge`) to perform tasks like applying styles, creating tables, or exporting documents. This guide helps AI assistants identify tools, invoke them correctly, and access additional resources.

---

## Available Tools

MCP tools are organized into modules, each containing specific operations. Below is a summary of key modules and their primary tools:

### File System (`fs`)
- **Tools**:
  - `directory.list`: Lists files and subdirectories.
  - `file.read`: Reads file content.
  - `file.write`: Writes content to a file.
- **Purpose**: Manage files and directories (e.g., read `sample.docx`, create `output.md`).

### Word (`word`)
- **Tools**:
  - `styles.apply`: Applies styles (e.g., Heading1) to document ranges.
  - `merge`: Merges multiple Word documents.
  - `markdown.export`: Converts Word to Markdown.
  - `markdown.import`: Converts Markdown to Word with a template.
  - `reformat`: Reformats documents with a style set.
  - `embedded-objects.extractAll`: Extracts embedded files (e.g., PDFs).
  - `mermaid.import`: Inserts Mermaid diagrams.
  - `mermaid.export`: Exports Mermaid diagrams as images.
  - `analyze`: Analyzes content and adds comments.
  - `code-format`: Formats code with syntax highlighting.
- **Purpose**: Automate Word document tasks (e.g., style application, diagram insertion).

### Excel (`excel`)
- **Tools**:
  - `tables.insert`: Creates tables from ranges.
  - `charts.insert`: Inserts charts (e.g., column, pie).
  - `ranges.write`: Writes data to cells.
- **Purpose**: Automate Excel tasks (e.g., table creation, data visualization).

### PowerPoint (`powerpoint`)
- **Tools**:
  - `slides.add`: Adds slides with specified layouts.
  - `animations.add`: Adds animations to slides or shapes.
  - `shapes.insert`: Inserts shapes (e.g., textboxes, images).
- **Purpose**: Automate PowerPoint tasks (e.g., slide creation, animation).

### Cross-Application (`office`)
- **Tools**:
  - `pdf.export`: Exports Office documents to PDF.
  - `workflow.run`: Executes multi-step workflows across applications.
- **Purpose**: Perform tasks involving multiple Office apps or formats.

### Resource Management (`memory`)
- **Tools**:
  - `memory.read`: Reads static resources (e.g., guides).
  - `memory.write`: Updates dynamic resources (e.g., caches).
- **Purpose**: Access or manage server resources and configurations.

---

## How to Use the Tools

AI assistants process user prompts, map them to MCP tool invocations, and validate results. Follow these steps:

1. **Understand the Prompt**:
   - Identify the task (e.g., “Apply Heading1 to the first paragraph in sample.docx”).
   - Determine the module (e.g., `word`), tool (e.g., `styles`), and operation (e.g., `apply`).

2. **Query Completions**:
   - Use the `completions` operation for the tool to get valid parameters.
   - Example: Invoke `word.styles.completions` with `document=/docs/sample.docx`.
     - **Result**: `{ "style": ["Heading1", "Normal"], "range": ["paragraph:1", "selection"] }`

3. **Invoke the Tool**:
   - Select the appropriate tool and provide parameters.
   - Example:
     - **Prompt**: “Apply Heading1 to the first paragraph in sample.docx.”
     - **Tool**: `word.styles.apply`
     - **Parameters**: `{ document: "/docs/sample.docx", style: "Heading1", range: "paragraph:1" }`
   - Pass parameters as a structured object to the MCP tool.

4. **Validate the Result**:
   - Check for a successful outcome or handle errors.
   - Example Results:
     - Success: `{ success: true }`
     - Error: `{ success: false, error: { code: 400, message: "Style not found" } }`

5. **Handle Errors**:
   - If `success: false`, inspect `error.message` and query document state (e.g., invoke `fs.file.read` with `path=/docs/sample.docx`).
   - Adjust parameters or inform the user of the issue.

6. **Provide Feedback**:
   - Confirm task completion (e.g., “Style applied successfully”) or explain issues (e.g., “Style not found; available styles are Heading1, Normal”).

---

## Resource References

### Static Resources
Static resources provide detailed guidance and documentation. Access them via the `memory.read` tool.

- **`memory://ai_assistant_guide`**:
  - Comprehensive guide with detailed tool usage, prompt processing, and troubleshooting.
  - Example: Invoke `memory.read` with `path=ai_assistant_guide`.
- **`memory://office_api_doc`**:
  - Technical documentation for Office JavaScript APIs and VBA, with code examples.
  - Example: Invoke `memory.read` with `path=office_api_doc`.
- **`memory://sample_documents`**:
  - Instructions for creating sample Office files (e.g., `CV.docx`, `SalesData.xlsx`).
  - Example: Invoke `memory.read` with `path=sample_documents`.

### Dynamic Resources
Dynamic resources are generated at runtime and vary based on server state. Access them via specific tools.

- **Completions**:
  - Format: `<module>.<tool>.completions`
  - Example: Invoke `word.merge.completions` with `docs=/docs/doc1.docx`.
    - **Result**: `{ "docs": ["/docs/doc1.docx", "/docs/doc2.docx"], "output": ["/docs/merged.docx"] }`
- **Document State**:
  - Format: `<module>.<tool>.list` or `fs.file.read`
  - Example: Invoke `word.styles.list` with `document=/docs/sample.docx`.
    - **Result**: `{ "styles": ["Heading1", "Normal", "Title"] }`
- **Workflow Results**:
  - Format: `office.workflow.run`
  - Example: Invoke `office.workflow.run` with `{ steps: [{ tool: "excel.tables.insert", params: { range: "A1:D10" } }] }`.
    - **Result**: `{ success: true, output: ["/docs/output.xlsx"] }`

---

This quick-start guide equips AI assistants with the essentials to use MCP Server tools effectively. For detailed instructions, access `memory://ai_assistant_guide`. For test file setup, see `memory://sample_documents`. For Office automation details, consult `memory://office_api_doc`.