# Plan: Implement `word/tables/insertFromArray` Tool

## Objective

Enhance the `msoffice-mcp` server to allow inserting tables into Word documents directly from a 2D data array, with options for applying standard table styles and formatting. This addresses the need for a lower-level tool to programmatically create tables in Word.

## Background

Analysis revealed that while Excel table creation from arrays is possible, Word lacks a direct mechanism. The existing `word/markdown/import` tool does not support Markdown tables. The prioritized solution is to create a dedicated tool for inserting tables from arrays into Word.

## Plan Details

1.  **Prioritize New Tool:** Focus development efforts on creating the `word/tables/insertFromArray` tool.
2.  **Design `word/tables/insertFromArray`:**
    *   **Functionality:** This tool will take a 2D array of data and insert it as a new table into a specified Word document at a given position. It will also allow applying a named table style and configuring common style options like header row and first column emphasis.
    *   **Proposed Input Schema:**
        *   `filePath` (string, required): Path to the Word document.
        *   `data` (array of arrays, required): The 2D array containing the table data (e.g., `[["Header 1", "Header 2"], ["Row1Col1", "Row1Col2"]]`).
        *   `position` (string, optional, default: 'end'): Where to insert the table (e.g., 'end', 'selection', 'paragraph:N', 'bookmark:BookmarkName').
        *   `styleName` (string, optional): The name of a Word table style to apply (e.g., "Table Grid", "List Table 1 Light Accent 1"). If omitted, Word's default table style will be used.
        *   `styleOptions` (object, optional): Fine-tune the applied style's appearance.
            *   `headerRow` (boolean, default: true): Apply distinct formatting defined by the style to the first row.
            *   `firstColumn` (boolean, default: false): Apply distinct formatting defined by the style to the first column.
            *   `bandedRows` (boolean, default: false): Apply alternating row shading (banding).
            *   `bandedColumns` (boolean, default: false): Apply alternating column shading (banding).
3.  **Defer Other Enhancements:** Postpone work on adding Markdown table support to `word/markdown/import` and the `word/tables/analyze` tool for now.
4.  **Update Documentation:** Plan to add documentation for the new `word/tables/insertFromArray` tool to `README.md` and `USE_CASES.md` once implemented.
5.  **Implementation:** Proceed with implementing the new tool in `code` mode.

## Workflow Diagram

```mermaid
graph TD
    A[Start: User Request] --> B{Analyze Request & Docs};
    B --> C{Analyze markdownToOffice.ts};
    C --> D{Finding: No MD Table Support};
    B --> E{Analyze Existing Tools};
    E --> F{Finding: Excel OK from Array, Word Table Tools Limited};
    D & F --> G{Assess LLM Role};
    G --> H{Finding: LLM can generate Array (Excel) / MD (Future Word)};
    H & D & F --> I{Identify Gaps & Enhancements};
    I --> J[Initial Plan Presented];
    J --> K{User Feedback Received};
    K --> L{Priority: New Tool `word/tables/insertFromArray`};
    K --> M{Requirement: Table Styling Options};
    L & M --> N[Update Plan: Design New Tool];
    N --> O{Step 1: Prioritize New Tool};
    N --> P{Step 2: Design Tool Schema & Logic};
    N --> Q{Step 3: Defer Other Enhancements};
    N --> R{Step 4: Plan Doc Updates};
    N --> S{Step 5: Plan Implementation};
    O & P & Q & R & S --> T{Present Updated Plan to User};
    T --> U{User Feedback on Updated Plan};
    U -- Approve --> V{Offer to Save Plan};
    V -- Yes --> W[Save Plan to MD];
    V -- No / After Save --> X{Switch Mode to 'code'};
    U -- Revise --> N;

    subgraph New Tool Design (word/tables/insertFromArray)
        direction LR
        inp1[filePath: string]
        inp2[data: string[][]]
        inp3[position: string (optional)]
        inp4[styleName: string (optional)]
        inp5[styleOptions: object (optional)]
        inp5 --> opt1[headerRow: boolean]
        inp5 --> opt2[firstColumn: boolean]
        inp5 --> opt3[bandedRows: boolean]
        inp5 --> opt4[bandedColumns: boolean]
    end

    P --> New Tool Design