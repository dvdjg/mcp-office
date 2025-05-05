# Plan: Phase 2 - Enhanced Word Table Handling

## Objective

Implement enhancements to the `msoffice-mcp` server to support importing/exporting Word tables from/to Markdown (using embedded HTML for complex formatting like merged cells) and extracting data from Word tables into a structured format.

## Background

Following the implementation of `word/tables/insertFromArray`, this phase addresses the previously deferred tasks of handling Markdown tables and analyzing Word table content, incorporating user feedback regarding the need to preserve formatting (especially merged cells) during Markdown conversion.

## Detailed Plan

1.  **Enhance Markdown Import/Export (`word/markdown/import`, `word/markdown/export`)**
    *   **Goal:** Enable conversion between Word tables and Markdown, supporting merged cells via embedded HTML.
    *   **Implementation Steps:**
        *   Modify `src/utils/markdownToOffice.ts`:
            *   Configure the `markdown-it` instance to allow HTML tags (`html: true`).
            *   Add parsing logic within `applyMarkdownFormattingToWord` to specifically handle `<table>`, `<tr>`, `<th>`, `<td>` HTML tags.
            *   Interpret `colspan` and `rowspan` attributes on `<td>`/`<th>` tags.
            *   Use Word COM Interop methods (`Table.Cell(Row, Column).Merge()`) to replicate the merged cell structure during import.
        *   Modify `word/markdown/export` tool:
            *   Iterate through Word tables in the document.
            *   Detect merged cells using COM properties (e.g., checking if a cell is part of a `MergeArea`).
            *   Generate Markdown containing the equivalent HTML `<table>` structure, including `colspan` and `rowspan` attributes where necessary.
2.  **Create Table Data Extraction Tool (`word/tables/extractData`)**
    *   **Goal:** Implement a tool to extract the content of a specified Word table into a structured 2D array, correctly representing merged cells (e.g., using `null` for spanned cells).
    *   **Implementation Steps:**
        *   Define a new Zod schema (`extractTableDataSchema`) in `src/tools/word/tables.tool.ts` requiring `filePath` and `tableIndex` (1-based index of the table in the document).
        *   Create a new handler function `extractTableData`.
        *   Use Word COM Interop to access the specified table (`doc.Tables(tableIndex)`).
        *   Determine the table dimensions (rows, columns).
        *   Initialize an empty 2D array.
        *   Iterate through each cell of the table (e.g., using nested loops from `r=1` to `Rows.Count`, `c=1` to `Columns.Count`).
        *   For each `Cell(r, c)`:
            *   Check if the cell is the start of a merged area (e.g., its `RowIndex` and `ColumnIndex` match `r` and `c`).
            *   If it is the start of a merge, get its text content and place it in the corresponding array position `result[r-1][c-1]`. Mark the subsequent cells covered by its `colspan` and `rowspan` in the array as `null`.
            *   If it's *not* the start of a merge (meaning it's covered by a previous cell's span), place `null` in the array position `result[r-1][c-1]`.
        *   Return the resulting 2D array.
        *   Add the tool definition to the `wordTablesTool` export array.
3.  **Update Documentation**
    *   **Goal:** Reflect the new capabilities in user-facing documentation.
    *   **Implementation Steps:**
        *   Modify `README.md`: Update descriptions of Markdown handling and add mention of the new table extraction tool.
        *   Modify `USE_CASES.md`:
            *   Add a new use case demonstrating importing Markdown with an HTML table (including merged cells).
            *   Add a new use case demonstrating exporting a Word table with merged cells to Markdown/HTML.
            *   Add a new use case demonstrating the use of `word/tables/extractData`.
            *   Update the Table of Contents.
4.  **Future Scope Confirmation**
    *   Work on similar Markdown/HTML table import/export features for Excel and PowerPoint will be deferred and treated as a separate future enhancement.

## Workflow Diagram

```mermaid
graph TD
    A[Start: User Request for Phase 2] --> B{Refine Plan based on Feedback};
    B --> C{Decision: Use HTML for Merged Cells};
    B --> D{Decision: Focus Analysis on Data Extraction};
    C --> E{Task 1: Enhance MD Import/Export (Word)};
    E --> E1[Enable HTML in markdown-it];
    E --> E2[Handle HTML Table Tags in markdownToOffice.ts];
    E --> E3[Implement Cell Merging on Import];
    E --> E4[Generate HTML Tables on Export];
    D --> F{Task 2: Create word/tables/extractData};
    F --> F1[Define Schema (filePath, tableIndex)];
    F --> F2[Implement Handler in tables.tool.ts];
    F --> F3[Iterate Cells & Detect Merges];
    F --> F4[Build 2D Array with nulls for spanned cells];
    E & F --> G{Task 3: Update Documentation};
    G --> G1[Update README.md];
    G --> G2[Add Use Cases for MD/HTML Tables];
    G --> G3[Add Use Case for extractData];
    G --> G4[Update ToC in USE_CASES.md];
    B --> H{Task 4: Confirm Future Scope};
    H --> H1[Defer Excel/PPT MD Table Handling];
    E & F & G & H --> I{Present Final Phase 2 Plan};
    I --> J{User Feedback on Plan};
    J -- Approve --> K{Offer to Save Plan};
    K -- Yes --> L[Save Plan to MD];
    K -- No / After Save --> M{Switch Mode to 'code'};
    J -- Revise --> I;