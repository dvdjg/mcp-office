# Plan: Enhanced Markdown-to-Word Insertion

**Goal:** Enhance the `word/text/insert` tool to support a wide range of Markdown features, mapping them to appropriate Word formatting via COM Interop, aiming for compatibility with VSCode's Markdown feature set. Make Markdown the default input format.

**Reference:** `src/docs/VSCodeSupportedMarkdownFeatures.md`

## Phase 1: Core Formatting Fixes & Defaults
*   **Target:** Fix fundamental formatting issues and set the default.
*   **Tasks:**
    1.  **Refactor Bold/Italic/Combined:** Modify `applyMarkdownFormattingToWord` to apply `Font.Bold` and/or `Font.Italic` directly to the range of the inserted text, not `wordApp.Selection`.
    2.  **Reliable Headings:** Ensure "Heading N" styles are consistently applied to the correct paragraph range.
    3.  **Change Default Format:** In `src/tools/word/text.tool.ts`, update `insertSchema` to default `format` to `'markdown'` and adjust the description.

## Phase 2: Basic Block Elements
*   **Target:** Implement common block-level Markdown elements.
*   **Tasks:**
    1.  **Paragraphs:** Solidify logic for inserting text and ensuring correct paragraph breaks (`InsertParagraphAfter`).
    2.  **Blockquotes:** Implement by applying indentation (e.g., `ParagraphFormat.LeftIndent`) and potentially a "Block Text" or custom style to the relevant paragraph range. Handle nested blockquotes by increasing indentation.
    3.  **Horizontal Rules:** Implement by inserting a paragraph with a bottom border or inserting a specific shape/symbol.

## Phase 3: Lists (Including Nested)
*   **Target:** Implement both simple and nested ordered/unordered lists.
*   **Tasks:**
    1.  **List Detection:** Identify `bullet_list_open`, `ordered_list_open`, `list_item_open`, and closing tokens.
    2.  **List Application:** Use `Range.ListFormat.ApplyListTemplateWithLevel` or similar COM methods to apply appropriate bullet or numbering styles.
    3.  **Nesting Logic:** Track the `listLevel` accurately. Adjust indentation and potentially list templates based on the nesting depth when applying formatting to list item paragraphs. This requires careful management of the `currentRange` and potentially interacting with the `ListFormat` object of preceding paragraphs.

## Phase 4: Inline Elements
*   **Target:** Implement common inline Markdown elements.
*   **Tasks:**
    1.  **Inline Code:** Apply a monospace font (e.g., "Consolas") and potentially a subtle background shading or character style to the inserted text range.
    2.  **Strikethrough:** Apply `Font.StrikeThrough` to the inserted text range.
    3.  **Links (Standard & Autolinks):** Parse link syntax (`[text](url)`) and autodetect URLs. Use `document.Hyperlinks.Add(Anchor:=range, Address:=url, ScreenTip:=text)` to create hyperlinks on the inserted text range.
    4.  **Reference Links:** Extend link parsing to handle reference-style links (`[text][id]`, `[id]: url`). Store definitions and apply hyperlinks when references are encountered.

## Phase 5: Complex Elements & GFM
*   **Target:** Implement more complex structures like images, code blocks, tables, and footnotes.
*   **Tasks:**
    1.  **Images:** Parse `![alt](src)` syntax. Use `Range.InlineShapes.AddPicture(FileName:=src, LinkToFile:=False, SaveWithDocument:=True)`. Add `alt` text to the shape properties if possible. Handle image path resolution carefully (relative vs. absolute).
    2.  **Code Blocks:** Insert the code content within a paragraph (or multiple paragraphs). Apply a specific "Code Block" style (if defined) or manually set a monospace font (e.g., "Consolas") and potentially borders/shading to the paragraph(s). *Note: Syntax highlighting is not feasible.*
    3.  **Tables (Markdown Syntax):** Parse GFM table syntax. Use `Range.Tables.Add` to create the table structure. Populate cells with content and apply basic formatting (e.g., bold headers). Handle cell alignment based on the separator line syntax (`:---:`, `:--`, `--:`).
    4.  **Footnotes:** Parse footnote syntax (`[^id]`, `[^id]: text`). Use `document.Footnotes.Add(Range:=range, Reference:="")` to insert the footnote marker and add the corresponding text to the footnote area.

## Phase 6: Review, Refinement & Testing
*   **Target:** Ensure robustness, handle edge cases, and document.
*   **Tasks:**
    1.  **Comprehensive Testing:** Test all implemented features individually and in combination using diverse Markdown examples.
    2.  **Range Management:** Thoroughly review and test the `currentRange` manipulation logic across all implemented features to prevent errors or incorrect positioning.
    3.  **Error Handling:** Enhance error handling for both Markdown parsing edge cases and potential Word COM errors during formatting application.
    4.  **Documentation:** Update the `word/text/insert` tool's description to accurately reflect the supported Markdown features and limitations.

## Acknowledged Limitations / Out of Scope for `word/text/insert`
*   **Mermaid Diagrams:** Rendering requires external tools. Use the dedicated `word/mermaid/import` tool instead.
*   **LaTeX Math:** Rendering requires external tools/complex conversion. Not feasible for direct insertion.
*   **Task Lists:** Native interactive checkboxes within text flow are not standard Word features.
*   **Syntax Highlighting:** Word does not support syntax highlighting within standard text ranges. Code blocks will have monospace font and potentially styling, but not language-specific highlighting.
*   **HTML Embedding:** Only very basic HTML (like the existing table support) might be convertible. Complex HTML will likely be inserted as raw text or stripped.
*   **Emoji:** Relies on font support in Word. Generally works but no specific implementation needed beyond inserting the Unicode characters.
*   **Collapsible Sections, TOC Generation:** Document-level features, not applicable here.

## Diagram of Overall Process

```mermaid
graph TD
    A[Start word/text/insert] --> B{Parse Request (format='markdown' default)};
    B --> C[Get Word Application & Document];
    C --> D[Resolve Insertion Range];
    D --> E[Parse Markdown Text (markdown-it)];
    E --> F{Iterate Tokens};
    F -- Heading --> G[Apply Heading Style];
    F -- Paragraph --> H[Insert Paragraph];
    F -- List Item --> I[Apply List Formatting (Nested)];
    F -- Bold/Italic/Strike --> J[Apply Font Formatting to Range];
    F -- Link --> K[Add Hyperlink];
    F -- Image --> L[Insert Inline Picture];
    F -- Code Block --> M[Format Paragraph as Code];
    F -- Table --> N[Parse & Create Word Table];
    F -- Footnote --> O[Add Footnote];
    F -- Other Tokens --> P[Handle Others or Skip];
    G & H & I & J & K & L & M & N & O & P --> Q{Next Token?};
    Q -- Yes --> F;
    Q -- No --> R[Save Document];
    R --> S[Release COM Objects];
    S --> T[Return Success/Error];

    subgraph applyMarkdownFormattingToWord [applyMarkdownFormattingToWord Function]
        direction LR
        E --> F
    end

    style G fill:#f9f,stroke:#333,stroke-width:2px
    style H fill:#f9f,stroke:#333,stroke-width:2px
    style I fill:#f9f,stroke:#333,stroke-width:2px
    style J fill:#f9f,stroke:#333,stroke-width:2px
    style K fill:#f9f,stroke:#333,stroke-width:2px
    style L fill:#f9f,stroke:#333,stroke-width:2px
    style M fill:#f9f,stroke:#333,stroke-width:2px
    style N fill:#f9f,stroke:#333,stroke-width:2px
    style O fill:#f9f,stroke:#333,stroke-width:2px