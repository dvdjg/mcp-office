# Office API Documentation for IAs running Office MCP Server

**Document Purpose**: This static resource (`memory://office_api_doc`) provides a beginner-friendly guide for an AI with limited programming knowledge to automate Microsoft Office (Word, Excel, PowerPoint) tasks using the Office MCP Server. It focuses on Office JavaScript APIs and VBA, with practical code examples for Office MCP use cases (e.g., managing styles, merging documents, formatting code). The goal is to enable an AI to understand and execute Office automation tasks effectively, even with minimal prior experience.

**Target Audience**: AI agents with basic programming skills using the Office MCP Server to automate Office tasks.

**Date**: May 01, 2025

---

## Table of Contents
1. [Introduction](#introduction)
   - [What is Office Automation?](#what-is-office-automation)
   - [Why Use Office JavaScript APIs and VBA?](#why-use-office-javascript-apis-and-vba)
   - [How Office MCP Simplifies Automation](#how-mcp-simplifies-automation)
2. [Key Concepts for Beginners](#key-concepts-for-beginners)
   - [Office JavaScript APIs](#office-javascript-apis)
   - [VBA (Visual Basic for Applications)](#vba-visual-basic-for-applications)
   - [Common Objects in Office](#common-objects-in-office)
   - [Error Handling](#error-handling)
   - [Testing Your Code](#testing-your-code)
3. [Getting Started with Office JavaScript APIs](#getting-started-with-office-javascript-apis)
   - [Setup and Environment](#setup-and-environment)
   - [Basic Structure of a Script](#basic-structure-of-a-script)
   - [Common Pitfalls](#common-pitfalls)
4. [Getting Started with VBA](#getting-started-with-vba)
   - [Accessing the VBA Editor](#accessing-the-vba-editor)
   - [Basic Structure of a Macro](#basic-structure-of-a-macro)
   - [Common Pitfalls](#common-pitfalls-1)
5. [Office MCP Use Case Examples](#mcp-use-case-examples)
   - [Use Case 1: Convert Word to Markdown with Comments](#use-case-1-convert-word-to-markdown-with-comments)
   - [Use Case 2: Merge Two Word Documents](#use-case-2-merge-two-word-documents)
   - [Use Case 3: Create Template with Placeholders](#use-case-3-create-template-with-placeholders)
   - [Use Case 4: Convert Markdown to Word with Template](#use-case-4-convert-markdown-to-word-with-template)
   - [Use Case 5: Reformat Poorly Formatted Document](#use-case-5-reformat-poorly-formatted-document)
   - [Use Case 6: Extract Embedded Documents](#use-case-6-extract-embedded-documents)
   - [Use Case 7: Import Mermaid Diagram](#use-case-7-import-mermaid-diagram)
   - [Use Case 8: Export Mermaid Diagram](#use-case-8-export-mermaid-diagram)
   - [Use Case 9: Analyze Document and Add Comments](#use-case-9-analyze-document-and-add-comments)
   - [Use Case 10: Format Code with Syntax Highlighting](#use-case-10-format-code-with-syntax-highlighting)
6. [Additional Tips for AI Agents](#additional-tips-for-ai-agents)
   - [Understanding Office MCP API Responses](#understanding-mcp-api-responses)
   - [Querying Document State](#querying-document-state)
   - [Using Completions](#using-completions)
   - [Debugging Tips](#debugging-tips)
7. [Resources and Further Learning](#resources-and-further-learning)

---

## Introduction

### What is Office Automation?
Office automation means using code to control Microsoft Office applications (Word, Excel, PowerPoint) to perform tasks like editing documents, inserting tables, or exporting files. Instead of clicking through menus, you write scripts to make Office do the work for you. For example, you can tell Word to apply a style to a paragraph or Excel to create a chart from data.

### Why Use Office JavaScript APIs and VBA?
- **Office JavaScript APIs**: These are modern, cross-platform (Windows, Mac, Office 365) tools for automating Office. They use JavaScript, run in a web-like environment, and are ideal for Office MCP's server-side tasks. They're great for consistent, scalable automation.
- **VBA (Visual Basic for Applications)**: This is Office's built-in scripting language, powerful for Windows-specific tasks like advanced document manipulation. It's older but still widely used for complex operations not fully supported by JavaScript APIs.

### How Office MCP Simplifies Automation?
The Office MCP Server acts like a friendly translator between you (the AI) and Office. Instead of writing complex JavaScript or VBA code, you send simple API requests to Office MCP (e.g., `POST /word/styles/apply`). Office MCP handles the hard parts, like interacting with Office, and returns easy-to-understand responses (e.g., `{ success: true }`). This document shows you what's happening under the hood so you can understand Office MCP's tools and troubleshoot if needed.

---

## Key Concepts for Beginners

### Office JavaScript APIs
- **What They Are**: A set of JavaScript libraries (`@microsoft/office-js`) that let you control Office apps. They work like a website interacting with a browser, but here the "browser" is Word, Excel, or PowerPoint.
- **Key Idea**: You write asynchronous code using `async/await` and a `context` object to interact with Office. You "load" properties, "sync" changes, and then apply actions.
- **Basic Example**:
  ```javascript
  async function setStyle() {
    await Word.run(async (context) => {
      const paragraph = context.document.body.paragraphs.getFirst();
      paragraph.load("style");
      await context.sync();
      paragraph.style = "Heading1";
      await context.sync();
    });
  }
  ```
  - `Word.run`: Runs code in Word.
  - `context`: A workspace to queue commands.
  - `load`: Tells Office to fetch data (e.g., paragraph style).
  - `sync`: Sends commands to Office and gets results.

### VBA (Visual Basic for Applications)
- **What It Is**: A programming language built into Office for Windows. It’s like writing instructions directly in Word or Excel.
- **Key Idea**: VBA uses a simpler, synchronous model. You directly access objects (e.g., `Document.Paragraphs`) and set properties or call methods.
- **Basic Example**:
  ```vba
  Sub SetStyle()
      Dim doc As Document
      Set doc = ActiveDocument
      doc.Paragraphs(1).Style = "Heading1"
  End Sub
  ```
  - `Sub`: A block of code (like a function).
  - `ActiveDocument`: The open Word document.
  - `Paragraphs(1)`: The first paragraph.

### Common Objects in Office
- **Word**:
  - `Document`: The entire Word file.
  - `Paragraph`: A block of text.
  - `Style`: Formatting rules (e.g., Heading1).
  - `Table`, `Chart`, `ContentControl`: For tables, charts, and placeholders.
- **Excel**:
  - `Workbook`: The Excel file.
  - `Worksheet`: A single sheet.
  - `Range`: A cell or group of cells (e.g., `A1:B2`).
  - `Chart`, `Table`: For charts and tables.
- **PowerPoint**:
  - `Presentation`: The PowerPoint file.
  - `Slide`: A single slide.
  - `Shape`: Textboxes, images, etc.

### Error Handling
- **JavaScript APIs**: Use `try/catch` to handle errors like invalid ranges or missing files.
  ```javascript
  try {
    await Word.run(async (context) => {
      const paragraph = context.document.body.paragraphs.getFirst();
      paragraph.style = "InvalidStyle"; // This will fail
      await context.sync();
    });
  } catch (error) {
    console.error("Error:", error.message); // Handle the error
  }
  ```
- **VBA**: Use `On Error` to catch errors.
  ```vba
  Sub SafeStyle()
      On Error GoTo ErrorHandler
      ActiveDocument.Paragraphs(1).Style = "InvalidStyle" ' This will fail
      Exit Sub
  ErrorHandler:
      MsgBox "Error: " & Err.Description
  End Sub
  ```

### Testing Your Code
- **JavaScript APIs**: Use a tool like Script Lab (an Office add-in) to test snippets in Office 365. Log results with `console.log`.
- **VBA**: Run macros in the VBA Editor (Alt+F11 in Office) and use `Debug.Print` to log to the Immediate Window (Ctrl+G).
- **Office MCP Testing**: Send API requests with `curl` or Postman to test Office MCP endpoints, checking responses for `{ success: true }`.

---

## Getting Started with Office JavaScript APIs

### Setup and Environment
- **Requirements**: Office 365 or Office 2016+ (Windows/Mac) with `@microsoft/office-js` installed in your Office MCP project (`npm install @microsoft/office-js`).
- **Environment**: Office MCP runs JavaScript APIs server-side, but you can test scripts in Office 365’s web client or Script Lab.
- **Loading the API**:
  ```javascript
  // Office MCP handles this, but for testing:
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  ```
- **Initialization**: Ensure Office is ready.
  ```javascript
  Office.onReady(() => {
    console.log("Office is ready!");
  });
  ```

### Basic Structure of a Script
1. **Start with `run`**: Use `Word.run`, `Excel.run`, or `PowerPoint.run` to access the app.
2. **Use `context`**: Queue commands (e.g., load properties, set styles).
3. **Sync Changes**: Call `context.sync()` to apply commands.
4. **Example** (Word, set text):
  ```javascript
  async function insertText() {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertText("Hello, World!", "End");
      await context.sync();
    });
  }
  ```

### Common Pitfalls
- **Forgetting `sync`**: Always call `await context.sync()` after loading or modifying properties.
- **Overloading `context`**: Load only necessary properties to avoid performance issues.
- **Invalid Ranges**: Ensure ranges (e.g., `paragraph:1`) exist in the document.
- **Testing Tip**: Use `console.log` to check object states before `sync`.

---

## Getting Started with VBA

### Accessing the VBA Editor
1. Open Word, Excel, or PowerPoint.
2. Press **Alt+F11** to open the VBA Editor.
3. Insert a new module: **Insert > Module**.
4. Write your code in the module.
5. Save the file as `.docm` (Word), `.xlsm` (Excel), or `.pptm` (PowerPoint) to enable macros.

### Basic Structure of a Macro
1. **Define a `Sub`**: A macro is a `Sub` (subroutine) with a name.
2. **Access Objects**: Use `ActiveDocument` (Word), `ActiveWorkbook` (Excel), or `ActivePresentation` (PowerPoint).
3. **Run the Macro**: Press **F5** in the VBA Editor or assign to a button.
4. **Example** (Word, insert text):
  ```vba
  Sub InsertText()
      ActiveDocument.Content.InsertAfter "Hello, World!"
  End Sub
  ```

### Common Pitfalls
- **Macro Security**: Enable macros in Office (File > Options > Trust Center > Macro Settings).
- **Object References**: Ensure the document is open (`ActiveDocument` fails if no file is open).
- **Case Sensitivity**: VBA is not case-sensitive, but object names (e.g., `Heading1`) must match exactly.
- **Testing Tip**: Use `Debug.Print` to log values to the Immediate Window.

---

## Office MCP Use Case Examples

Below are code examples for the 10 Word-specific Office MCP use cases, showing how MCP’s API calls map to Office JavaScript APIs and VBA. Each includes:
- **Office MCP API Call**: The request sent to Office MCP.
- **JavaScript API Code**: Equivalent code using `@microsoft/office-js`.
- **VBA Code**: Equivalent code for Windows.
- **AI Notes**: Tips for understanding and troubleshooting.

### Use Case 1: Convert Word to Markdown with Comments
**Task**: Convert a Word document (`CV.docx`) to Markdown (`CV.md`), appending comments.

**Office MCP API Call**:
```bash
curl -X POST "http://localhost:3000/word/markdown/export?document=/docs/CV.docx&output=/docs/CV.md&comments=append"
```

**JavaScript API Code**:
```javascript
async function exportToMarkdown() {
  await Word.run(async (context) => {
    // Load document content
    const body = context.document.body;
    body.load("text, comments");
    await context.sync();

    // Extract text
    let markdown = body.text;

    // Append comments
    const comments = body.comments.items.map((comment) => `> ${comment.author}: ${comment.text}`);
    markdown += "\n\n## Comments\n" + comments.join("\n");

    // Office MCP handles saving to /docs/CV.md
    console.log("Markdown:", markdown);
  });
}
```

**VBA Code**:
```vba
Sub ExportToMarkdown()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim markdown As String
    
    ' Extract text
    markdown = doc.Content.Text
    
    ' Append comments
    Dim comment As Comment
    markdown = markdown & vbCrLf & vbCrLf & "## Comments" & vbCrLf
    For Each comment In doc.Comments
        markdown = markdown & "> " & comment.Author & ": " & comment.Range.Text & vbCrLf
    Next comment
    
    ' Office MCP saves to file
    Debug.Print markdown
End Sub
```

**AI Notes**:
- **What’s Happening**: Office MCP extracts the document’s text and comments, converts to Markdown using `markdown-it`, and saves the file. Comments are appended as a `## Comments` section.
- **Troubleshooting**: Ensure `CV.docx` exists (`GET /fs/file/read?path=/docs/CV.docx`). Check if comments are loaded correctly (`body.comments` in JavaScript, `doc.Comments` in VBA).
- **Why Use Office MCP**: Office MCP handles Markdown formatting and file saving, so you don’t need to learn `markdown-it` or file I/O.

### Use Case 2: Merge Two Word Documents
**Task**: Merge `doc1.docx` and `doc2.docx` into `merged.docx` with AI conflict resolution.

**Office MCP API Call**:
```bash
curl -X POST "http://localhost:3000/word/merge?docs=/docs/doc1.docx,/docs/doc2.docx&output=/docs/merged.docx"
```

**JavaScript API Code**:
```javascript
async function mergeDocuments() {
  await Word.run(async (context) => {
    const doc = context.document;
    
    // Office MCP opens doc1.docx and doc2.docx, but here’s how to append content
    const doc1Content = await fetchDocumentContent("/docs/doc1.docx"); // Office MCP handles this
    const doc2Content = await fetchDocumentContent("/docs/doc2.docx");
    
    // Insert content from doc2 at the end of doc1
    doc.body.insertContentControl().insertText(doc2Content, "End");
    await context.sync();
    
    // Office MCP uses AI to resolve conflicts (e.g., duplicate headings)
    console.log("Merged document saved as /docs/merged.docx");
  });
}

// Mock function for fetching content (Office MCP does this)
async function fetchDocumentContent(path) {
  return "Content from " + path; // Simplified
}
```

**VBA Code**:
```vba
Sub MergeDocuments()
    Dim doc1 As Document, doc2 As Document
    Set doc1 = Documents.Open("C:\docs\doc1.docx")
    Set doc2 = Documents.Open("C:\docs\doc2.docx")
    
    ' Append doc2 content to doc1
    doc1.Content.InsertAfter doc2.Content.Text
    
    ' Office MCP handles conflict resolution
    doc1.SaveAs "C:\docs\merged.docx"
    doc2.Close
    doc1.Close
End Sub
```

**AI Notes**:
- **What’s Happening**: Office MCP opens both documents, appends content, and uses AI to resolve conflicts (e.g., merging similar paragraphs). The result is saved as `merged.docx`.
- **Troubleshooting**: Verify file paths with `GET /fs/directory/list?path=/docs`. Ensure documents are not empty.
- **Why Use Office MCP**: Office MCP simplifies multi-document operations and AI conflict resolution, which are complex in raw JavaScript/VBA.

### Use Case 3: Create Template with Placeholders
**Task**: Create a template from `budget.docx`, replacing sensitive text with placeholders (e.g., `[CompanyName]`).

**Office MCP API Call**:
```bash
curl -X POST "http://localhost:3000/word/template?document=/docs/budget.docx&output=/docs/template.docx"
```

**JavaScript API Code**:
```javascript
async function createTemplate() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();
    
    // Find sensitive text (Office MCP uses AI)
    const sensitiveText = "Acme Corp"; // Example
    const range = body.search(sensitiveText, { matchCase: true });
    range.load("text");
    await context.sync();
    
    // Replace with placeholder
    range.items.forEach((item) => {
      item.insertText("[CompanyName]", "Replace");
    });
    await context.sync();
  });
}
```

**VBA Code**:
```vba
Sub CreateTemplate()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Find and replace sensitive text
    With doc.Content.Find
        .Text = "Acme Corp"
        .Execute ReplaceWith:="[CompanyName]", Replace:=wdReplaceAll
    End With
    
    ' Save as template (Office MCP handles this)
    doc.SaveAs "C:\docs\template.docx"
End Sub
```

**AI Notes**:
- **What’s Happening**: Office MCP uses AI to identify sensitive text (e.g., company names) and replaces it with placeholders using content controls.
- **Troubleshooting**: Ensure the document has text to replace (`GET /word/text?document=/docs/budget.docx`). Check placeholder format.
- **Why Use Office MCP**: MCP’s AI detects sensitive data automatically, saving you from writing complex search logic.

### Use Case 4: Convert Markdown to Word with Template
**Task**: Convert `input.md` to a Word document using `template.docx`.

**Office MCP API Call**:
```bash
curl -X POST "http://localhost:3000/word/markdown/import?path=/docs/input.md&template=/docs/template.docx&output=/docs/output.docx"
```

**JavaScript API Code**:
```javascript
async function importMarkdown() {
  await Word.run(async (context) => {
    const doc = context.document;
    
    // Office MCP parses input.md using markdown-it
    const markdown = "# Title\nText"; // Example
    const paragraphs = markdown.split("\n").filter((line) => line);
    
    // Insert content with template styles
    paragraphs.forEach((line, index) => {
      const style = line.startsWith("#") ? "Heading1" : "Normal";
      doc.body.insertParagraph(line.replace("# ", ""), "End").style = style;
    });
    await context.sync();
    
    // Office MCP applies template.docx styles
  });
}
```

**VBA Code**:
```vba
Sub ImportMarkdown()
    Dim doc As Document
    Set doc = Documents.Open("C:\docs\template.docx")
    
    ' Office MCP parses Markdown
    Dim markdown As String
    markdown = "# Title" & vbCrLf & "Text"
    
    ' Insert content
    If InStr(markdown, "#") Then
        doc.Content.InsertAfter "Title"
        doc.Paragraphs.Last.Style = "Heading1"
    End If
    doc.Content.InsertAfter "Text"
    doc.Paragraphs.Last.Style = "Normal"
    
    doc.SaveAs "C:\docs\output.docx"
End Sub
```

**AI Notes**:
- **What’s Happening**: Office MCP uses `markdown-it` to parse Markdown, maps headers to template styles, and saves as a Word document.
- **Troubleshooting**: Verify `input.md` exists and is valid Markdown. Check template styles (`GET /word/styles/list`).
- **Why Use Office MCP**: Office MCP handles Markdown parsing and style mapping, which are tricky in raw JavaScript/VBA.

### Use Case 5: Reformat Poorly Formatted Document
**Task**: Reformat `PropuestasRandom.docx` with a professional style set.

**Office MCP API Call**:
```bash
curl -X POST "http://localhost:3000/word/reformat?document=/docs/PropuestasRandom.docx&styleSet=Professional&output=/docs/Formatted.docx"
```

**JavaScript API Code**:
```javascript
async function reformatDocument() {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("style, text");
    await context.sync();
    
    // Apply professional styles (Office MCP defines "Professional" set)
    paragraphs.items.forEach((paragraph, index) => {
      paragraph.style = index === 0 ? "Title" : "Normal";
    });
    await context.sync();
  });
}
```

**VBA Code**:
```vba
Sub ReformatDocument()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim para As Paragraph
    
    ' Apply styles
    Dim i As Integer
    i = 1
    For Each para In doc.Paragraphs
        If i = 1 Then
            para.Style = "Title"
        Else
            para.Style = "Normal"
        End If
        i = i + 1
    Next para
    
    doc.SaveAs "C:\docs\Formatted.docx"
End Sub
```

**AI Notes**:
- **What’s Happening**: Office MCP analyzes the document structure and applies a predefined style set (e.g., Title, Normal).
- **Troubleshooting**: Ensure styles exist (`GET /word/styles/list`). Check document content.
- **Why Use Office MCP**: Office MCP automates style application and structure detection, simplifying formatting logic.

### Use Case 6: Extract Embedded Documents
**Task**: Extract embedded files (e.g., PDFs) from `Composición.docx` to `/docs/extracted`.

**Office MCP API Call**:
```bash
curl -X POST "http://localhost:3000/word/embedded-objects/extractAll?document=/docs/Composición.docx&outputDir=/docs/extracted"
```

**JavaScript API Code**:
```javascript
async function extractEmbedded() {
  await Word.run(async (context) => {
    const objects = context.document.body.inlinePictures; // Simplified
    objects.load("source");
    await context.sync();
    
    // Office MCP saves objects to /docs/extracted
    objects.items.forEach((obj, index) => {
      console.log(`Saving object ${index + 1} to /docs/extracted/embedded_${index + 1}.pdf`);
    });
  });
}
```

**VBA Code**:
```vba
Sub ExtractEmbedded()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim ole As OLEFormat
    
    ' Extract OLE objects
    Dim i As Integer
    i = 1
    For Each ole In doc.InlineShapes
        If ole.Type = wdInlineShapeOLEObject Then
            ' Office MCP saves to /docs/extracted
            Debug.Print "Saving object to /docs/extracted/embedded_" & i & ".pdf"
            i = i + 1
        End If
    Next ole
End Sub
```

**AI Notes**:
- **What’s Happening**: Office MCP identifies embedded objects (e.g., PDFs) and saves them as separate files.
- **Troubleshooting**: Verify objects exist (`GET /word/embedded-objects/list`). Ensure output directory is writable.
- **Why Use Office MCP**: Office MCP handles file extraction and naming, which is complex in raw APIs.

### Use Case 7: Import Mermaid Diagram
**Task**: Insert a Mermaid diagram into `tech.docx` at paragraph 5.

**Office MCP API Call**:
```bash
curl -X POST "http://localhost:3000/word/mermaid/import?document=/docs/tech.docx&syntax=graph TD; A-->B&format=svg&position=paragraph:5"
```

**JavaScript API Code**:
```javascript
async function importMermaid() {
  await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getRange("paragraph:5");
    paragraph.load("range");
    await context.sync();
    
    // Office MCP renders Mermaid as SVG
    const svg = "<svg>...</svg>"; // Mock SVG from Mermaid
    paragraph.insertInlinePictureFromBase64(svgToBase64(svg), "After");
    await context.sync();
  });
}

function svgToBase64(svg) {
  return Buffer.from(svg).toString("base64"); // Simplified
}
```

**VBA Code**:
```vba
Sub ImportMermaid()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Insert at paragraph 5
    Dim para As Paragraph
    Set para = doc.Paragraphs(5)
    
    ' Office MCP renders SVG
    para.Range.InlineShapes.AddPicture "C:\docs\mermaid.svg"
End Sub
```

**AI Notes**:
- **What’s Happening**: Office MCP uses the `mermaid` library to render the diagram as SVG and inserts it into the document.
- **Troubleshooting**: Validate Mermaid syntax (`graph TD; A-->B`). Ensure paragraph 5 exists.
- **Why Use Office MCP**: Office MCP handles Mermaid rendering and image insertion, which are non-trivial tasks.

### Use Case 8: Export Mermaid Diagram
**Task**: Export the first Mermaid diagram in `tech.docx` as a PNG.

**Office MCP API Call**:
```bash
curl -X POST "http://localhost:3000/word/mermaid/export?document=/docs/tech.docx&diagram=1&format=png&output=/docs/diagram.png"
```

**JavaScript API Code**:
```javascript
async function exportMermaid() {
  await Word.run(async (context) => {
    const pictures = context.document.body.inlinePictures;
    pictures.load("source");
    await context.sync();
    
    // Office MCP identifies Mermaid diagram (first picture)
    const diagram = pictures.items[0];
    console.log("Exporting diagram as /docs/diagram.png");
    // Office MCP converts SVG to PNG using pdf2pic
  });
}
```

**VBA Code**:
```vba
Sub ExportMermaid()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' First inline picture (Office MCP identifies Mermaid)
    Dim shape As InlineShape
    Set shape = doc.InlineShapes(1)
    
    ' Office MCP saves as PNG
    Debug.Print "Exporting to /docs/diagram.png"
End Sub
```

**AI Notes**:
- **What’s Happening**: Office MCP extracts the diagram, converts it to PNG using `pdf2pic`, and saves it.
- **Troubleshooting**: Ensure the document has diagrams (`GET /word/mermaid/list`). Verify output path.
- **Why Use Office MCP**: Office MCP simplifies diagram identification and format conversion.

### Use Case 9: Analyze Document and Add Comments
**Task**: Analyze `PropuestaTécnica.docx` for technical clarity and add comments.

**Office MCP API Call**:
```bash
curl -X POST "http://localhost:3000/word/analyze?document=/docs/PropuestaTécnica.docx&criteria=technical&output=/docs/Commented.docx"
```

**JavaScript API Code**:
```javascript
async function analyzeDocument() {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();
    
    // Office MCP uses AI to analyze
    paragraphs.items.forEach((para, index) => {
      if (para.text.includes("jargon")) { // Simplified AI check
        para.insertComment("Simplify this jargon", "AI");
      }
    });
    await context.sync();
  });
}
```

**VBA Code**:
```vba
Sub AnalyzeDocument()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim para As Paragraph
    
    ' Office MCP uses AI
    For Each para In doc.Paragraphs
        If InStr(para.Range.Text, "jargon") Then
            doc.Comments.Add para.Range, "Simplify this jargon"
        End If
    Next para
    
    doc.SaveAs "C:\docs\Commented.docx"
End Sub
```

**AI Notes**:
- **What’s Happening**: MCP’s AI analyzes the document and adds comments on problematic sections.
- **Troubleshooting**: Verify document content (`GET /word/text`). Check comment insertion.
- **Why Use Office MCP**: MCP’s AI handles analysis, saving you from writing NLP logic.

### Use Case 10: Format Code with Syntax Highlighting
**Task**: Format code in `BuenasPrácticas.docx` with Consolas font and syntax highlighting.

**Office MCP API Call**:
```bash
curl -X POST "http://localhost:3000/word/code-format?document=/docs/BuenasPrácticas.docx&style=Código&font=Consolas&output=/docs/Formatted.docx"
```

**JavaScript API Code**:
```javascript
async function formatCode() {
  await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text, font");
    await context.sync();
    
    // Office MCP uses highlight.js for syntax highlighting
    paragraphs.items.forEach((para) => {
      if (para.text.includes("function")) { // Simplified code detection
        para.font.name = "Consolas";
        para.font.color = "blue"; // Mock highlighting
      }
    });
    await context.sync();
  });
}
```

**VBA Code**:
```vba
Sub FormatCode()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim para As Paragraph
    
    ' Office MCP uses highlight.js
    For Each para In doc.Paragraphs
        If InStr(para.Range.Text, "function") Then
            para.Range.Font.Name = "Consolas"
            para.Range.Font.Color = wdColorBlue
        End If
    Next para
    
    doc.SaveAs "C:\docs\Formatted.docx"
End Sub
```

**AI Notes**:
- **What’s Happening**: Office MCP detects code snippets, applies syntax highlighting with `highlight.js`, and formats with Consolas.
- **Troubleshooting**: Ensure code snippets exist (`GET /word/text`). Verify font availability.
- **Why Use Office MCP**: Office MCP handles code detection and highlighting, which are complex tasks.

---

## Additional Tips for AI Agents

### Understanding Office MCP API Responses
- **Success**: `{ success: true }` means the operation worked. Check `output` for file paths (e.g., `/docs/output.docx`).
- **Errors**: Look for `{ success: false, error: { message: "..." } }`. Common issues include invalid file paths or missing styles.
- **Example**:
  ```json
  { "success": false, "error": { "code": 400, "message": "Style not found" } }
  ```
  - Action: Query available styles (`GET /word/styles/list`) and retry.

### Querying Document State
- Use Office MCP endpoints to check document properties before acting:
  - `GET /word/styles/list?document=/docs/sample.docx`: Lists styles.
  - `GET /fs/file/read?path=/docs/sample.docx`: Verifies file existence.
  - `GET /word/text?document=/docs/sample.docx`: Checks content.
- Example: Before applying a style, confirm it exists to avoid errors.

### Using Completions
- Office MCP provides completions for parameters (e.g., style names, file paths). Query them to make smart choices:
  ```bash
  curl -X GET "http://localhost:3000/word/styles/completions?document=/docs/sample.docx"
  ```
  **Response**:
  ```json
  { "style": ["Heading1", "Normal"], "range": ["paragraph:1", "selection"] }
  ```
- Use completions to validate user inputs (e.g., suggest `Heading1` for a style request).

### Debugging Tips
- **Log Responses**: Check Office MCP responses for errors or unexpected outputs.
- **Test Small**: Start with simple API calls (e.g., `GET /fs/directory/list`) to build confidence.
- **Mock Data**: Use sample documents (`/docs/sample.docx`) to test scripts.
- **Ask for Help**: Query `memory://ai_assistant_guide` for usage examples:
  ```bash
  curl -X GET "http://localhost:3000/memory/ai_assistant_guide?section=tool_usage"
  ```

---

## Resources and Further Learning
- **Office JavaScript APIs**: [docs.microsoft.com/office/dev/add-ins](https://docs.microsoft.com/office/dev/add-ins)
- **VBA Documentation**: [docs.microsoft.com/office/vba](https://docs.microsoft.com/office/vba)
- **Script Lab**: Install this Office add-in to test JavaScript API code.
- **Office MCP API Docs**: `GET /memory/ai_assistant_guide` or local `docs/api`.
- **Mermaid**: [mermaid-js.github.io](https://mermaid-js.github.io/) for diagram syntax.
- **Highlight.js**: [highlightjs.org](https://highlightjs.org/) for code formatting.
- **Markdown-it**: [github.com/markdown-it/markdown-it](https://github.com/markdown-it/markdown-it) for Markdown parsing.

---

This document equips an AI with the knowledge to automate Office tasks using Office MCP, focusing on practical examples and beginner-friendly explanations. For further assistance, query `memory://ai_assistant_guide` or test Office MCP endpoints directly.