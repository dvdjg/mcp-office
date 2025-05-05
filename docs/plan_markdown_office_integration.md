# Plan: Enhance Markdown Integration and Text Manipulation with Office Applications

**Objective:** Implement robust Markdown import/export functionality for Word, Excel, and PowerPoint, focusing initially on enhancing the `word/generate-and-insert-text` tool to correctly format LLM Markdown output, and enhance Word text manipulation tools (`word/text/insert` and `word/text/get`) to support natural language descriptions for locations and sections.

**Refined Plan (incorporating feedback):**

1.  **Develop Advanced Range Resolution Utility:**
    *   Create a new utility function or class (e.g., `src/utils/wordRangeResolver.ts`) that takes a Word Document COM object, the Word Application COM object, and a natural language description of a location or section (e.g., "after the heading 'Introduction'", "the third paragraph in the second section").
    *   This utility will need to analyze the document structure (paragraphs, headings, sections, tables, etc.) via COM and interpret the natural language description to return a specific COM `Range` object. This is a significant development effort and will require careful implementation using the COM API to traverse and identify document elements.
    *   Initially, focus on interpreting common natural language patterns related to headings, paragraphs, and sections.
2.  **Enhance `word/text/insert.tool.ts`:**
    *   Modify the `insertText` function.
    *   Update the `position` parameter handling to accept natural language descriptions in addition to the current specific formats (`start`, `end`, `paragraph:N`, `selection`).
    *   Use the new `wordRangeResolver.ts` utility to get the target insertion `Range` based on the provided position description (whether specific or natural language).
    *   Integrate the `applyMarkdownFormattingToWord` utility (from `src/utils/markdownToOffice.ts`) into `insertText` to format the inserted text if it's in Markdown. This will likely require adding an optional `format` parameter (e.g., 'markdown', 'plaintext') to the `insertSchema`.
3.  **Enhance `word/text/get.tool.ts`:**
    *   Modify the `getText` function.
    *   Update the `range` parameter handling to accept natural language descriptions in addition to the current specific formats (`document`, `paragraph:N`, `selection`).
    *   Use the new `wordRangeResolver.ts` utility to get the target `Range` for extraction based on the provided range description.
4.  **Review `word/generate-and-insert-text.tool.ts`:**
    *   Review this tool in light of the enhanced `word/text/insert`. The core functionality of generating text (via LLM) and inserting it with formatting can now be achieved by an AI orchestrating a separate LLM call and the enhanced `word/text/insert` tool.
    *   Keep `word/generate-and-insert-text` as a convenience tool for direct LLM-to-Word insertion, but ensure its internal logic utilizes the new `wordRangeResolver.ts` utility for position handling and the `applyMarkdownFormattingToWord` utility for formatting.
5.  **Enhance LLM Prompt (for `generate-and-insert-text`):** Modify the prompt sent to the LLM to request a specific format, perhaps using tags (e.g., `<summary>...</summary>`) or clearly separating the summary from the main content. This aims to make the summary identification more reliable. (This step was previously in the plan and remains relevant for the convenience tool).
6.  **Enhance `word/markdown/import.tool.ts`:**
    *   Modify the `importFromMarkdown` function.
    *   Use the core Markdown to Word formatting utility (from step 3) to parse the Markdown content from the file and apply formatting to the new Word document. (This step was previously in the plan and remains relevant).
7.  **Future Considerations (Excel and PowerPoint):** Acknowledge the complexity and future work needed for these applications.
8.  **Update Documentation:**
    *   Update `README.md` and `src/docs/office_mcp_api_doc.md` to describe the enhanced capabilities of the `word/text/insert` and `word/text/get` tools, including examples of how to use natural language descriptions for ranges and positions.
    *   Explain how AI agents can combine LLM calls with these enhanced tools for flexible text generation and insertion, highlighting the increased orchestration possibilities.
9.  **Write Notes to Basic Memory:** Document the design decisions, the implementation details of the `wordRangeResolver.ts` utility, and the complexity of interpreting natural language for COM-based range resolution.

**Process Flow for `word/generate-and-insert-text` (Mermaid Diagram):**

```mermaid
graph TD
    A[Start generateAndInsertText] --> B{Enhance Prompt & Get LLM Response};
    B --> C{Separate Initial Summary};
    C --> D[Insert Summary as Heading 1];
    C --> E[Remaining Markdown Content];
    D --> F[Determine Insertion Range for Content];
    E --> G[Call Markdown to Word Utility];
    F --> G;
    G --> H[Apply Markdown Formatting via COM];
    H --> I[Save and Close Document];
    I --> J[End];