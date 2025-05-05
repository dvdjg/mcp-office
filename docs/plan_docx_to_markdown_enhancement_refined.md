# Plan to Enhance DOCX to Markdown Conversion

**Goal:** Improve the `word/markdown/export` tool to extract more elements from Word documents and convert them to the VSCode Markdown variant.

**Steps:**

1.  **Address Existing TODOs:**
    *   Implement robust list detection (bullet points and numbered lists) including nested lists by examining `paragraph.ListFormat` and `paragraph.ListFormat.ListLevelNumber`. Generate appropriate Markdown list syntax with correct indentation.
    *   Refine heading detection by checking `paragraph.Style.NameLocal` and `paragraph.OutlineLevel` to accurately identify heading levels and apply the corresponding Markdown prefixes (`#`, `##`, etc.).
    *   Complete inline image handling within paragraphs by checking if `inlineShape` is a picture and calling the `handleImage` function.
    *   Implement a more robust check for merged cells in tables to improve the accuracy of Markdown table conversion or determine when to fall back to HTML.
    *   Implement robust merged cell detection for HTML table conversion, adding `colspan` and `rowspan` attributes as needed.
    *   Implement the `CopyAsPicture` fallback for image extraction to handle cases where `OLEFormat.Object.SaveAs` is not available or fails.
    *   Review and confirm the COM object release strategy for image shapes to prevent memory leaks.
    *   Implement inline comment handling, potentially by inserting comments as Markdown blockquotes or similar syntax linked to the relevant text.

2.  **Implement Extraction of New Elements:**
    *   **Shapes:** Enhance the document traversal to identify and handle different types of shapes (`doc.Shapes`). For shapes that represent diagrams or drawings (including grouped shapes and shapes within canvases), the plan is to:
        *   Extract the shape data via the COM API.
        *   Utilize a suitable library (this will need to be identified during implementation) to convert the extracted shape data into SVG format.
        *   Occasionally, if configured, use a library to convert the SVG to PNG.
        *   Save the generated SVG/PNG files to the image directory.
        *   Insert a Markdown image link referencing the saved SVG/PNG file in the output.
        *   For shapes that are text boxes, extract the text content and include it in the Markdown.
    *   **Drawings/Canvases:** As mentioned in the Shapes step, shapes within canvases will be handled as part of the shape extraction process, aiming for SVG/PNG conversion.
    *   **Logos (from covers/headers/footers):** Extend the traversal to include headers and footers (`doc.Sections(i).Headers` and `doc.Sections(i).Footers`) to check for images that might be logos and extract them using the image handling logic.
    *   **Backgrounds:** Investigate if and how document backgrounds can be accessed via the COM API. It might not be possible to represent backgrounds directly in Markdown, but perhaps a note or metadata could be included.

3.  **Refine Markdown Output for VSCode Compatibility:**
    *   Ensure the generated Markdown syntax for all elements (headings, lists, tables, images, etc.) adheres to the VSCode Markdown variant. This may involve checking VSCode's specific rendering rules for certain elements.

4.  **Testing:**
    *   Create or update test cases (likely in `tests/e2e/word/markdown.test.ts`) to cover the newly implemented extraction capabilities, including documents with complex lists, nested lists, various heading levels, tables with merged cells, different types of shapes (standalone, grouped, in canvases), and documents with images in headers/footers.

## Diagram

```mermaid
graph TD
    A[Start] --> B{Read DOCX};
    B --> C[Traverse Document Elements];
    C --> D{Element Type?};
    D --> E{Paragraph};
    E --> F[Handle Paragraph Content];
    F --> G{Inline Shape?};
    G --> H[Handle Image];
    H --> I[Add Markdown Image Link];
    G --> F;
    F --> J[Add Markdown Text/Formatting];
    J --> C;
    D --> K{Table};
    K --> L[Handle Table (MD or HTML)];
    L --> C;
    D --> M{Shape};
    M --> N[Handle Shape];
    N --> O{Is Drawing/Diagram?};
    O --> P[Extract Shape Data];
    P --> Q[Convert to SVG/PNG];
    Q --> R[Save Image File];
    R --> S[Add Markdown Image Link];
    O --> T{Is Textbox?};
    T --> U[Extract Text];
    U --> V[Add Markdown Text];
    T --> N;
    N --> C;
    D --> W{Other Elements (Drawings, etc.)};
    W --> X[Investigate/Handle Other Elements];
    X --> C;
    C --> Y{End of Document?};
    Y --> Z[Handle Comments];
    Z --> AA[Write Markdown File];
    AA --> AB{Zip Output?};
    AB --> AC[Create ZIP Archive];
    AC --> AD[Save Resource];
    AB --> AD;
    AD --> AE[End];