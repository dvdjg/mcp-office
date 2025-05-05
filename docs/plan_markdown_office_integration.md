# Plan: Enhance Markdown Integration with Office Applications

**Objective:** Implement robust Markdown import/export functionality for Word, Excel, and PowerPoint, focusing initially on enhancing the `word/generate-and-insert-text` tool to correctly format LLM Markdown output.

**Refined Plan:**

1.  **Analyze LLM Response Structure:** Confirm the heuristic that the initial summary to be treated as a Heading 1 is the first paragraph of the LLM's Markdown response, before any Markdown headings or other block-level elements. (This remains a fallback/initial approach).
2.  **Enhance LLM Prompt (for `generate-and-insert-text`):** Modify the prompt sent to the LLM to request a specific format, perhaps using tags (e.g., `<summary>...</summary>`) or clearly separating the summary from the main content. This aims to make the summary identification more reliable.
3.  **Develop a Core Markdown to Office Formatting Utility:**
    *   Create a new utility function or class (e.g., `src/utils/markdownToOffice.ts`) that takes Markdown text and a reference to an Office application object (like a Word `Range`) and applies formatting based on *standard* Markdown syntax (headings, lists, bold, italic) using COM Interop.
    *   This utility will be designed to be reusable across different Office tools.
4.  **Enhance `word/generate-and-insert-text.tool.ts`:**
    *   Modify the `generateAndInsertText` function.
    *   After getting the `generatedText` from the LLM, attempt to identify the summary based on the new prompt formatting instructions (e.g., looking for specific tags). If tags are not found or the format is unexpected, fall back to the first paragraph heuristic.
    *   Insert the identified summary text into the Word document at the specified position and apply the "Heading 1" style using COM.
    *   Determine the insertion range immediately after the inserted Heading 1.
    *   Pass the remaining Markdown content (after removing the summary and any specific tags used for identification) and this new insertion range to the core Markdown to Word formatting utility developed in step 3.
    *   Ensure proper error handling and COM object release.
5.  **Enhance `word/markdown/import.tool.ts`:**
    *   Modify the `importFromMarkdown` function.
    *   Use the core Markdown to Word formatting utility (from step 3) to parse the Markdown content from the file and apply formatting to the new Word document.
6.  **Future Considerations (Excel and PowerPoint):** Acknowledge the complexity and future work needed for these applications.
7.  **Update Documentation:** Update relevant documentation files (`README.md`, `src/docs/office_mcp_api_doc.md`) to reflect the enhanced tools and the new Markdown utility.

**Process Flow for `word/generate-and-insert-text` (Mermaid Diagram):**

```mermaid
graph TD
    A[Start generateAndInsertText] --> B{Get LLM Response};
    B --> C{Separate Initial Summary};
    C --> D[Insert Summary as Heading 1];
    C --> E[Remaining Markdown Content];
    D --> F[Determine Insertion Range for Content];
    E --> G[Call Markdown to Word Utility];
    F --> G;
    G --> H[Apply Markdown Formatting via COM];
    H --> I[Save and Close Document];
    I --> J[End];