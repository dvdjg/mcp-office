# Plan: Enhance `word/markdown/export` Tool for Rich DOCX-to-Markdown Conversion

**Date:** 2025-05-05

**Author:** Roo (AI Architect)

**Status:** Approved

## 1. Goal

Transform the existing basic text extraction `word/markdown/export` tool into a comprehensive converter that accurately translates Word document structure (including formatting, tables, and images) into Markdown, with options for image handling and output packaging.

## 2. Decision: Enhance vs. New Tool

We will enhance the existing `src/tools/word/markdown.tool.ts` and its `exportToMarkdown` function. This keeps related functionality together. If the complexity becomes unmanageable, we can reconsider creating a separate tool (e.g., `word/markdown/exportRich`) later.

## 3. Core Requirements

-   **Structure Traversal:** Iterate through the Word document's primary content flow (Paragraphs, Tables, Shapes/InlineShapes) in order.
-   **Formatting Conversion:** Convert common formatting (bold, italic, strikethrough, basic lists) to Markdown syntax.
-   **Heading Conversion:** Identify headings (likely based on Word Styles like "Heading 1", "Heading 2", etc., or outline levels) and convert them to Markdown headings (`#`, `##`, etc.).
-   **Table Conversion:**
    -   Detect `Table` objects.
    -   Iterate through rows and cells.
    -   Generate standard Markdown table syntax for simple tables.
    -   Generate embedded HTML `<table>` syntax for tables with merged cells (`colspan`/`rowspan`) or complex formatting, controlled by a parameter.
-   **Image Handling:**
    -   Detect images within `InlineShapes` and potentially `Shapes`.
    *   Extract image data reliably (investigate best COM methods - `CopyAsPicture`, OLE `SaveAs`, etc.).
    *   Save extracted images as files (e.g., PNG, JPG) into a specified relative directory.
    *   Generate unique, safe filenames for images.
    *   Insert correct relative Markdown image links (`![alt text](image_dir/image_file.png)`) into the output Markdown. Use Alt Text if available in Word, otherwise use a generic placeholder or filename.
-   **Comment Handling:** Retain or enhance the existing comment extraction logic (`ignore`, `append`, `inline`).
-   **Output Packaging:** Optionally create a ZIP archive containing the generated Markdown file and the directory of extracted images.

## 4. Technical Approach (COM Interop Focus)

-   **Iteration:** Loop through `doc.StoryRanges` (main text, headers, footers, etc.) or primarily `doc.Content`. Within ranges, iterate `Range.Paragraphs`, `Range.Tables`, `Range.InlineShapes`. Need careful handling to process elements in visual order.
-   **Paragraphs:** Check `Paragraph.Style.NameLocal` for headings/styles. Check `Paragraph.Range.Font.Bold`, `.Italic`, `.StrikeThrough`, etc. Check `Paragraph.Range.ListFormat` for list type and level.
-   **Tables:** Iterate `Table.Rows`, `Table.Columns`. Access cell content via `Cell(Row, Col).Range.Text`. Check `Cell.MergeInfo` or `Cell.RowIndex`, `Cell.ColumnIndex` combined with iteration logic to detect spans for HTML output.
-   **Images:**
    *   Focus on `InlineShape` where `Type` is `wdInlineShapePicture`.
    *   Investigate `InlineShape.Select()` followed by `wordApp.Selection.CopyAsPicture()`. Paste from clipboard (requires clipboard access library or COM methods) and save.
    *   Alternatively, explore if `InlineShape.OLEFormat` is applicable and if `OLEFormat.Object.SaveAs()` can work for embedded pictures. This needs testing.
    *   Extract `InlineShape.AlternativeText` for Markdown `alt text`.
-   **File Naming:** Generate image filenames using a sequence (image1, image2) or a hash of image data to avoid collisions. Ensure filenames are filesystem-safe.
-   **Path Construction:** Use Node.js `path` module (`path.join`, `path.relative`) to correctly construct the relative image paths for Markdown links based on the `output` MD file location and the specified `imageDir`.

## 5. Proposed Tool Parameters (Schema Update)

```typescript
// In exportSchema (src/tools/word/markdown.tool.ts)
const exportSchema = z.object({
    filePath: z.string().min(1).refine(validateFilePath, { /* ... */ }),
    output: z.string().min(1).refine(/* ... */),
    // NEW Parameters:
    imageDir: z.string().default('images').describe("Relative directory to save extracted images (e.g., 'img', 'assets/images')."),
    imagePrefix: z.string().optional().describe("Optional prefix for extracted image filenames."),
    tableFormat: z.enum(['markdown', 'html']).default('html').describe("Format for tables: 'markdown' (simple) or 'html' (preserves merged cells)."),
    zipOutput: z.boolean().default(false).describe("If true, create a ZIP archive containing the Markdown file and image directory."),
    zipFileName: z.string().optional().describe("Optional name for the output ZIP file (defaults based on output MD name)."),
    // Existing Parameter:
    comments: z.enum(['ignore', 'append', 'inline']).default('ignore').describe(/* ... */),
});
```

## 6. Implementation Steps

1.  **Backup:** Create a backup or branch of `src/tools/word/markdown.tool.ts`.
2.  **Schema Update:** Modify the `exportSchema` in the file with the new parameters.
3.  **Refactor `exportToMarkdown`:**
    *   Remove the simple `doc.Content.Text` extraction.
    *   Implement the core document traversal logic (iterating paragraphs, tables, shapes).
    *   Maintain a string builder or array to accumulate the Markdown output.
4.  **Element Handlers:** Create helper functions:
    *   `handleParagraph(paragraph, imageDir, imageCounter)`: Converts paragraph text, formatting, lists, and detects/handles inline images within it. Returns Markdown string. Updates `imageCounter`.
    *   `handleTable(table, tableFormat)`: Converts a Word table to Markdown or HTML string based on `tableFormat`.
    *   `handleImage(shape, imageDir, imageCounter)`: Extracts image data from `InlineShape` or `Shape`, saves it to `imageDir`, increments `imageCounter`, and returns the Markdown link string.
5.  **Image Extraction Logic:** Implement the chosen COM method for saving image data (requires careful testing and error handling). Use `fs-extra` to write image files. Ensure the `imageDir` is created.
6.  **Path Management:** Use `path.resolve` for absolute paths during processing and `path.relative` to generate correct relative links in the Markdown output.
7.  **Zipping Logic:** If `zipOutput` is true:
    *   After the MD file and images are written, determine the final ZIP filename.
    *   Use a library like `archiver` (needs to be added as a dependency: `npm install archiver @types/archiver`) or potentially invoke the `fs/archive/create` tool's underlying logic if accessible.
    *   Add the generated MD file and the entire `imageDir` to the archive.
    *   Return the path to the ZIP file as the primary result.
8.  **Testing:** Add comprehensive unit tests for helper functions and integration tests for the overall `exportToMarkdown` process with various DOCX files (including tables, images, formatting).
9.  **Documentation:** Update the tool's description in the `McpResource` definition to reflect its new capabilities. Add comments explaining complex COM interactions.

## 7. Mermaid Diagram (Enhanced Workflow)

```mermaid
graph TD
    subgraph Initialization
        A[Start: User Request (filePath, output, imageDir, tableFormat, zipOutput, ...)] --> B{Validate Parameters};
        B --> C{Resolve Paths (Input DOCX, Output MD, Image Dir)};
        C --> D{Initialize Word App & Open DOCX via COM};
        D --> E{Create Image Directory if not exists};
        E --> F[Initialize Markdown Output String & Image Counter];
    end

    subgraph Document Processing Loop
        F --> G{Iterate Document Elements (Paragraphs, Tables, Shapes)};
        G -- Paragraph --> H[handleParagraph];
        H --> I{Append Para Markdown to Output};
        H -- Image Found --> J[handleImage];
        J --> K{Append Image Link to Output};

        G -- Table --> L[handleTable];
        L --> M{Append Table Markdown/HTML to Output};

        G -- Other Element --> N{Skip/Log Element};

        I --> O{More Elements?};
        M --> O;
        K --> O;
        N --> O;
        O -- Yes --> G;
    end

    subgraph Finalization
        O -- No --> P[Write Final Markdown Output String to MD File];
        P --> Q{zipOutput == true?};
        Q -- Yes --> R{Create ZIP Archive};
        R --> S[Add MD File to ZIP];
        S --> T[Add Image Directory to ZIP];
        T --> U[Return ZIP File Path];
        Q -- No --> V[Return MD File Path];
    end

    subgraph Cleanup
        U --> W{Close DOCX & Release COM Objects};
        V --> W;
        W --> X[End];
    end

    style J fill:#f9d,stroke:#333,stroke-width:2px
    style L fill:#f9d,stroke:#333,stroke-width:2px
```

## 8. Review

This plan outlines the necessary steps to significantly enhance the `word/markdown/export` tool. It involves complex COM Interop, especially for reliable image extraction and ordered element processing.