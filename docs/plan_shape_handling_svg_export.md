# Plan for Enhanced Shape Handling and SVG Export in `word/markdown.tool.ts`

## Goal

Enhance the `word/markdown/export` tool to handle various drawing objects in a Word document and export them as SVG images.

## Steps

1.  **Refine Document Traversal and Shape Identification:** Modify the document traversal to identify all `Shape` and `InlineShape` objects and determine their specific types using `shape.Type`.
2.  **Implement Shape-Specific Extraction and SVG Generation:** Create a helper function to handle different shape types. For each type (pictures, text boxes, charts, groups, canvases, auto shapes), extract relevant properties and content. Explore the COM API for direct SVG export capabilities. If direct export is not possible, programmatically generate SVG XML based on the extracted data, mapping Word's coordinates to SVG.
3.  **Integrate SVG Export:** Call the shape handling and SVG generation logic during document traversal and save the generated SVG files.
4.  **Update Markdown Output:** Include Markdown image links pointing to the exported SVG files.
5.  **Update Zipping Logic:** Ensure the exported SVG files are included in the output ZIP archive if zipping is enabled.
6.  **Refine Error Handling and Logging:** Add specific error handling and logging for shape processing and SVG export.

## Mermaid Diagram

```mermaid
graph TD
    A[Start ExportToMarkdown] --> B(Open Word Document via COM);
    B --> C(Ensure Image Directory Exists);
    C --> D(Traverse Document Content);
    D -- For each element --> E{Element Type?};
    E -- Table --> F(Handle Table);
    E -- Paragraph --> G(Handle Paragraph);
    E -- Shape/InlineShape --> H(Handle Shape and Export SVG);
    H --> I{Shape Type?};
    I -- msoShapePicture --> J(Extract Image Data);
    J --> K{Export as SVG?};
    K -- Yes --> L(Save as SVG);
    K -- No --> M(Save as PNG/Other);
    I -- msoTextBox --> N(Extract Text and Properties);
    N --> O(Generate SVG for Text Box);
    O --> L;
    I -- msoChart --> P(Extract Chart Data and Formatting);
    P --> Q(Generate SVG for Chart);
    Q --> L;
    I -- msoGroup --> R(Iterate GroupItems);
    R --> H;
    I -- msoCanvas --> S(Iterate CanvasItems);
    S --> H;
    I -- msoAutoShape --> T(Extract Properties);
    T --> U(Generate SVG for AutoShape);
    U --> L;
    I -- Other --> V(Log Placeholder);
    L --> W(Add SVG Link to Markdown);
    M --> X(Add Image Link to Markdown);
    F --> Y(Add Table to Markdown);
    G --> Z(Add Paragraph to Markdown);
    Y --> D;
    Z --> D;
    W --> D;
    X --> D;
    V --> D;
    D --> AA{End of Document?};
    AA -- No --> D;
    AA -- Yes --> AB(Handle Comments);
    AB --> AC(Write Markdown to File);
    AC --> AD{Zip Output?};
    AD -- Yes --> AE(Create ZIP Archive);
    AD -- No --> AF(Skip Zipping);
    AE --> AG(Save Final Output as Resource);
    AF --> AG;
    AG --> AH(Release COM Objects);
    AH --> AI(End ExportToMarkdown);