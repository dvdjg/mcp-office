# Sample Documents for Office MCP Server API Testing

**Document Purpose**: This document provides detailed specifications for creating sample Office documents (Word, Excel, PowerPoint), Markdown, and PDF files to test the Office MCP Server API. These files serve as test inputs for validating API functionality, including Word-specific use cases (e.g., merging documents, formatting code), Excel scenarios (e.g., table creation), and PowerPoint tasks (e.g., slide addition). Instructions are provided to ensure anyone can generate the dataset consistently.

**Target Audience**: Developers and testers setting up the Office MCP Server test environment.

**Date**: May 01, 2025

---

## Table of Contents
1. [Overview](#overview)
2. [Instructions for Creating Sample Documents](#instructions-for-creating-sample-documents)
3. [Word Documents](#word-documents)
4. [Excel Documents](#excel-documents)
5. [PowerPoint Documents](#powerpoint-documents)
6. [Markdown Document](#markdown-document)
7. [PDF Document](#pdf-document)
8. [Directory Structure](#directory-structure)
9. [Notes](#notes)

---

## Overview

This document specifies 16 sample files to be used as test inputs for the Office MCP Server API:
- **Word**: 10 documents (`CV.docx`, `cuentoAladdinDraft1.docx`, etc.) for testing Markdown export, merging, template creation, and more.
- **Excel**: 2 files (`SalesData.xlsx`, `Budget.xlsx`) for table and chart operations.
- **PowerPoint**: 2 files (`Presentation.pptx`, `Tutorial.pptx`) for slide and animation tasks.
- **Markdown**: 1 file (`input.md`) for import testing.
- **PDF**: 1 file (`Offer.pdf`) for parsing and combining.

Files are stored in the `mcp-office/tests/fixtures/` directory and must be created according to the specifications below to ensure consistent test results.

---

## Instructions for Creating Sample Documents

1. **Tools**: Use Microsoft Office (Word, Excel, PowerPoint) or Office 365 to create the files. For Markdown and PDF, use a text editor and a PDF editor or Word’s PDF export feature.
2. **Location**: Save all files in the `mcp-office/tests/fixtures/` directory.
3. **Formats**:
   - Word: Save as `.docx` unless macros are required (then use `.docm`).
   - Excel: Save as `.xlsx` unless macros are required (then use `.xlsm`).
   - PowerPoint: Save as `.pptx` unless macros are required (then use `.pptm`).
   - Markdown: Save as `.md` (plain text).
   - PDF: Save as `.pdf`.
4. **Content**: Follow the specifications for each file, including text, styles, tables, comments, and other elements as described.
5. **Size**: Keep files small (1-5 pages/sheets/slides) unless specified (e.g., `LargeDoc.docx` is 100 pages).
6. **Validation**: Open each file in the respective application to confirm the content matches the description.
7. **Optional Automation**: For large files (e.g., `LargeDoc.docx`), use VBA or scripts to generate content efficiently.

---

## Word Documents

1. **CV.docx**
   - **Purpose**: Test Markdown export, style application, and metadata management.
   - **Content**:
     - Title: "John Doe Resume" (Heading1 style).
     - Section: "Experience" (Heading2 style).
     - Paragraph: "Software Engineer at Acme Corp, 2020-2025." (Normal style).
     - Comment: On the text "Acme Corp" - "Add dates" by Jane (insert via Review > New Comment).
     - Metadata: Set Author to "John Doe" and Title to "Resume" (File > Info > Properties).
   - **Size**: 1 page.
   - **Creation**:
     1. Open Word and create a new document.
     2. Apply Heading1 to the title, Heading2 to the section, and Normal to the paragraph.
     3. Add the comment via Review > New Comment.
     4. Set metadata in File > Info > Properties.
     5. Save as `CV.docx` in `tests/fixtures/`.

2. **cuentoAladdinDraft1.docx**
   - **Purpose**: Test document merging.
   - **Content**:
     - Title: "Aladdin Story Draft 1" (Heading1 style).
     - Paragraph: "Once upon a time, Aladdin found a lamp." (Normal style).
     - Paragraph: "He rubbed it and a genie appeared." (Normal style).
   - **Size**: 1 page.
   - **Creation**:
     1. Create a new Word document.
     2. Apply Heading1 to the title and Normal to the paragraphs.
     3. Save as `cuentoAladdinDraft1.docx` in `tests/fixtures/`.

3. **cuentoAladdinDraft2.docx**
   - **Purpose**: Test document merging.
   - **Content**:
     - Title: "Aladdin Story Draft 2" (Heading1 style).
     - Paragraph: "The genie granted three wishes." (Normal style).
     - Paragraph: "Aladdin wished for wisdom." (Normal style).
   - **Size**: 1 page.
   - **Creation**:
     1. Create a new Word document.
     2. Apply Heading1 to the title and Normal to the paragraphs.
     3. Save as `cuentoAladdinDraft2.docx` in `tests/fixtures/`.

4. **PresupuestosEvolutio.docx**
   - **Purpose**: Test template creation.
   - **Content**:
     - Title: "Budget Proposal" (Heading1 style).
     - Paragraph: "Client: Acme Corp" (Normal style).
     - Paragraph: "Total: $10,000" (Normal style).
     - Table: 2 columns ("Item", "Cost"), 2 rows ("Software", "$5,000"; "Services", "$5,000").
   - **Size**: 1 page.
   - **Creation**:
     1. Create a new Word document.
     2. Apply Heading1 to the title and Normal to the paragraphs.
     3. Insert a 2x2 table via Insert > Table and fill with the specified data.
     4. Save as `PresupuestosEvolutio.docx` in `tests/fixtures/`.

5. **PlantillaEvolutio.docx**
   - **Purpose**: Test Markdown import with a template.
   - **Content**:
     - Title: "[CompanyName] Proposal" (Heading1 style).
     - Paragraph: "Prepared for [ClientName]" (Normal style).
     - Content Control: Placeholder for "[Date]" (Insert > Developer > Plain Text Content Control).
   - **Size**: 1 page.
   - **Creation**:
     1. Create a new Word document.
     2. Enable the Developer tab (File > Options > Customize Ribbon > Check Developer).
     3. Apply Heading1 to the title and Normal to the paragraph.
     4. Insert a content control for "[Date]" via Developer > Controls > Plain Text Content Control.
     5. Save as `PlantillaEvolutio.docx` in `tests/fixtures/`.

6. **PropuestasRandom.docx**
   - **Purpose**: Test reformatting of poorly formatted documents.
   - **Content**:
     - Text: "Proposal" (no style, bold, 14pt Arial).
     - Text: "Details: Random text" (no style, italic, 12pt Times New Roman).
     - Text: "Conclusion" (no style, underlined, 16pt Calibri).
   - **Size**: 1 page.
   - **Creation**:
     1. Create a new Word document.
     2. Type the text and manually apply the specified formatting (bold, italic, etc.) without using styles.
     3. Save as `PropuestasRandom.docx` in `tests/fixtures/`.

7. **Composición.docx**
   - **Purpose**: Test extraction of embedded objects.
   - **Content**:
     - Title: "Composition Report" (Heading1 style).
     - Paragraph: "Includes embedded files." (Normal style).
     - Embedded Object: Embed `Offer.pdf` (Insert > Object > Create from File).
   - **Size**: 1 page.
   - **Creation**:
     1. Create a new Word document.
     2. Apply Heading1 to the title and Normal to the paragraph.
     3. Embed `Offer.pdf` via Insert > Object > Create from File > Select `Offer.pdf`.
     4. Save as `Composición.docx` in `tests/fixtures/`.

8. **PropuestaTécnica.docx**
   - **Purpose**: Test document analysis and commenting.
   - **Content**:
     - Title: "Technical Proposal" (Heading1 style).
     - Paragraph: "This proposal uses technical jargon like API and CRUD." (Normal style).
     - Paragraph: "Implementation details follow." (Normal style).
   - **Size**: 1 page.
   - **Creation**:
     1. Create a new Word document.
     2. Apply Heading1 to the title and Normal to the paragraphs.
     3. Save as `PropuestaTécnica.docx` in `tests/fixtures/`.

9. **BuenasPrácticas.docx**
   - **Purpose**: Test code formatting with syntax highlighting.
   - **Content**:
     - Title: "Best Practices" (Heading1 style).
     - Paragraph: "Code example:" (Normal style).
     - Paragraph: `function hello() { console.log("Hello, World!"); }` (Normal style, plain text).
   - **Size**: 1 page.
   - **Creation**:
     1. Create a new Word document.
     2. Apply Heading1 to the title and Normal to the paragraphs.
     3. Paste the code snippet as plain text without formatting.
     4. Save as `BuenasPrácticas.docx` in `tests/fixtures/`.

10. **LargeDoc.docx**
    - **Purpose**: Test performance for large document operations (e.g., merging, PDF export).
    - **Content**:
      - 100 pages of repeated text: "Sample paragraph" (Normal style).
      - Each page contains 5 paragraphs.
    - **Size**: 100 pages.
    - **Creation**:
      1. Create a new Word document.
      2. Manually copy-paste "Sample paragraph" 500 times (5 paragraphs × 100 pages) or use the following VBA script:
         ```vba
         Sub GenerateLargeDoc()
             Dim doc As Document
             Set doc = Documents.Add
             Dim i As Integer
             For i = 1 To 500
                 doc.Content.InsertAfter "Sample paragraph" & vbCrLf
             Next i
             doc.SaveAs "mcp-office\tests\fixtures\LargeDoc.docx"
         End Sub
         ```
      3. Run the VBA script in Word (Alt+F11 > Insert > Module > Paste and Run).
      4. Save as `LargeDoc.docx` in `tests/fixtures/`.

---

## Excel Documents

1. **SalesData.xlsx**
   - **Purpose**: Test table and chart creation.
   - **Content**:
     - Worksheet: Named "Sales".
     - Range A1:D10:
       - A1:D1: Headers "Month", "Region", "Sales", "Profit".
       - A2:D10: Sample data:
         ```
         Jan, North, 1000, 200
         Feb, South, 1200, 250
         Mar, East, 900, 180
         Apr, West, 1100, 220
         May, North, 1300, 260
         Jun, South, 1000, 200
         Jul, East, 950, 190
         Aug, West, 1150, 230
         Sep, North, 1250, 250
         ```
   - **Size**: 1 worksheet, 10 rows.
   - **Creation**:
     1. Open Excel and create a new workbook.
     2. Name the first worksheet "Sales" (double-click the sheet tab).
     3. Enter the headers and data in A1:D10.
     4. Save as `SalesData.xlsx` in `tests/fixtures/`.

2. **Budget.xlsx**
   - **Purpose**: Test data analysis and workflow automation.
   - **Content**:
     - Worksheet: Named "Budget".
     - Range A1:C5:
       - A1:C1: Headers "Category", "Amount", "Approved".
       - A2:C5: Sample data:
         ```
         Software, 5000, Yes
         Hardware, 3000, No
         Services, 4000, Yes
         Travel, 2000, No
         ```
   - **Size**: 1 worksheet, 5 rows.
   - **Creation**:
     1. Open Excel and create a new workbook.
     2. Name the first worksheet "Budget".
     3. Enter the headers and data in A1:C5.
     4. Save as `Budget.xlsx` in `tests/fixtures/`.

---

## PowerPoint Documents

1. **Presentation.pptx**
   - **Purpose**: Test slide and animation addition.
   - **Content**:
     - Slide 1: Title Slide layout, Title: "Project Overview", Subtitle: "By Team X".
     - Slide 2: Blank Slide layout, text box with "Details to be added." (centered, 12pt Calibri).
   - **Size**: 2 slides.
   - **Creation**:
     1. Open PowerPoint and create a new presentation.
     2. Use the Title Slide layout for Slide 1 and enter the title and subtitle.
     3. Add a Blank Slide for Slide 2 via Home > New Slide > Blank.
     4. Insert a text box on Slide 2 (Insert > Text Box) and type the text.
     5. Save as `Presentation.pptx` in `tests/fixtures/`.

2. **Tutorial.pptx**
   - **Purpose**: Test shape insertion and Word-to-PowerPoint conversion.
   - **Content**:
     - Slide 1: Title Slide layout, Title: "Tutorial".
     - Slide 2: Title and Content layout, Title: "Step 1", Content: "Follow instructions." (bullet point, 12pt Calibri).
   - **Size**: 2 slides.
   - **Creation**:
     1. Open PowerPoint and create a new presentation.
     2. Use the Title Slide layout for Slide 1 and enter the title.
     3. Add a Title and Content slide for Slide 2 via Home > New Slide > Title and Content.
     4. Enter the title and content as a bullet point.
     5. Save as `Tutorial.pptx` in `tests/fixtures/`.

---

## Markdown Document

1. **input.md**
   - **Purpose**: Test Markdown import to Word.
   - **Content**:
     ```markdown
     # Proposal
     This is a sample proposal.
     ## Details
     - Item 1
     - Item 2
     ```
   - **Size**: Approximately 10 lines.
   - **Creation**:
     1. Open a text editor (e.g., Notepad, VS Code).
     2. Copy the Markdown content above.
     3. Save as `input.md` in `tests/fixtures/` with UTF-8 encoding.

---

## PDF Document

1. **Offer.pdf**
   - **Purpose**: Test PDF parsing and combining with Office documents.
   - **Content**:
     - Title: "Offer Letter" (bold, 16pt).
     - Text: "We offer $10,000 for services." (12pt).
     - Image: A simple logo or graphic (e.g., a company logo or placeholder image).
   - **Size**: 1 page.
   - **Creation**:
     1. Create a Word document with the title, text, and an inserted image (Insert > Pictures).
     2. Export to PDF via File > Save As > PDF.
     3. Alternatively, use a PDF editor to create the content directly.
     4. Save as `Offer.pdf` in `tests/fixtures/`.

---

## Directory Structure

After creating the files, ensure they are organized in the following structure:
```
mcp-office/tests/fixtures/
├── CV.docx
├── cuentoAladdinDraft1.docx
├── cuentoAladdinDraft2.docx
├── PresupuestosEvolutio.docx
├── PlantillaEvolutio.docx
├── PropuestasRandom.docx
├── Composición.docx
├── PropuestaTécnica.docx
├── BuenasPrácticas.docx
├── LargeDoc.docx
├── SalesData.xlsx
├── Budget.xlsx
├── Presentation.pptx
├── Tutorial.pptx
├── input.md
└── Offer.pdf
```

---

## Notes

- **Consistency**: Ensure file names and content match the specifications exactly to avoid test failures.
- **Validation**: Open each file after creation to verify the content, styles, and structure are correct.
- **Large Files**: For `LargeDoc.docx`, consider using the provided VBA script to automate content generation, as manual creation is time-consuming.
- **Optional Text-Based Content**: If generating binary files is challenging, request text-based equivalents (e.g., Markdown for Word documents, JSON for Excel data) from a developer or tool. Examples:
  - Markdown for `CV.docx`:
    ```markdown
    # John Doe Resume
    ## Experience
    Software Engineer at Acme Corp, 2020-2025.
    <!-- Comment: Add dates -->
    ```
  - JSON for `SalesData.xlsx`:
    ```json
    {
      "Sales": [
        ["Month", "Region", "Sales", "Profit"],
        ["Jan", "North", 1000, 200],
        ["Feb", "South", 1200, 250]
      ]
    }
    ```
- **Dependencies**: Ensure `Offer.pdf` is created before `Composición.docx`, as it is embedded in the latter.
- **Testing**: After creating the files, run the Office MCP Server tests (e.g., `npm test`) to validate the dataset.

This document provides all necessary instructions to generate the sample dataset for Office MCP Server API testing. Follow the steps carefully to ensure the files meet the test requirements.