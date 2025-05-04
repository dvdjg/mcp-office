# Use Cases

Explore how the **Office MCP Server** automates Word, Excel, PowerPoint, and cross-application tasks. Each use case shows the user’s interaction with the AI, how the AI infers the solution using available tools, and any clarifications needed, with a touch of humor.

## Table of Contents
- [Word Use Cases](#word-use-cases)
  - [Markdown Magic](#markdown-magic)
  - [Merge Mayhem](#merge-mayhem)
  - [Template Trickery](#template-trickery)
  - [Markdown Import Mania](#markdown-import-mania)
  - [Reformatting Rescue](#reformatting-rescue)
  - [Embedded Object Extraction](#embedded-object-extraction)
  - [Embedded Object Insertion](#embedded-object-insertion)
  - [Embedded Object Modification](#embedded-object-modification)
  - [Embedded Object Deletion](#embedded-object-deletion)
  - [Mermaid Import Marvel](#mermaid-import-marvel)
  - [Mermaid Export Escapade](#mermaid-export-escapade)
  - [Analysis Antics](#analysis-antics)
  - [Code Formatting Fiesta](#code-formatting-fiesta)
  - [Instant Image Injection](#instant-image-injection)
  - [AI Ghostwriter](#ai-ghostwriter)
  - [Surgical Strike Styling](#surgical-strike-styling)
  - [The AI Word Analyst](#the-ai-word-analyst)
  - [The Code Stylist in Word](#the-code-stylist-in-word)
  - [The Multimodal Illustrator](#the-multimodal-illustrator)
- [Excel Use Cases](#excel-use-cases)
  - [The Excel Accountant with Judgment](#the-excel-accountant-with-judgment)
  - [The Express Excel Exporter](#the-express-excel-exporter)
- [PowerPoint Use Cases](#powerpoint-use-cases)
  - [The Christmas PowerPoint Decorator](#the-christmas-powerpoint-decorator)
- [Cross-Application Use Cases](#cross-application-use-cases)
  - [The Document Multitasker](#the-document-multitasker)
  - [The Magical Office Linker](#the-magical-office-linker)
  - [The PowerPoint Embedder](#the-powerpoint-embedder)
  - [The AI Document Creator](#the-ai-document-creator)
  - [The Universal Office Translator](#the-universal-office-translator)
  - [The Credential Guardian](#the-credential-guardian)
  - [The Compressed File Manager](#the-compressed-file-manager)

## Word Use Cases

### Markdown Magic
**Scenario**: Your boss needs a Markdown version of `CV.docx` for the wiki, comments included. Manual copying? Yawn! 😴
**Solution**:
- **User Tells AI**: “Convert CV.docx to Markdown with comments for the wiki.”
- **AI Infers**: The user wants to export a Word document to Markdown, including comments. The `word/markdown/export` tool is perfect.
- **AI Action**: Uses `word/markdown/export` with `document=/docs/CV.docx`, `output=/docs/CV.md`, and `comments=append` to include comments in a `## Comments` section. If the output path is missing, AI suggests `/docs/CV.md`.
- **Interaction**: AI confirms the document path and asks if comments should be appended or excluded.

### Merge Mayhem
**Scenario**: Conflicting drafts of `cuentoAladdin.docx`. Merging manually? Chaos! 🦁
**Solution**:
- **User Tells AI**: “Merge these two Aladdin story drafts into one.”
- **AI Infers**: The user wants to merge Word documents with conflict resolution. The `word/merge` tool with AI support is ideal.
- **AI Action**: Uses `word/merge` with `docs=/docs/cuentoAladdin.docx,/docs/Cuentos/cuentoAladdin.docx` and `output=/docs/merged.docx`. If paths are unclear, AI asks for specific file locations.
- **Interaction**: AI requests both document paths and confirms the output file name.

### Template Trickery
**Scenario**: `PresupuestosEvolutio.docx` has sensitive client names. Sharing risks leaks! 😱
**Solution**:
- **User Tells AI**: “Make a template from PresupuestosEvolutio.docx, hiding client names.”
- **AI Infers**: The user wants to replace sensitive data with placeholders. The `word/template` tool is suitable.
- **AI Action**: Uses `word/template` with `document=/docs/PresupuestosEvolutio.docx` and `output=/docs/Template.docx`. AI suggests placeholder formats if unspecified.
- **Interaction**: AI asks for the document path and confirms placeholder style (e.g., `[ClientName]`).

### Markdown Import Mania
**Scenario**: Turn `input.md` into a Word doc with `PlantillaEvolutio.docx`. Copy-pasting? So last decade!
**Solution**:
- **User Tells AI**: “Convert input.md to a Word document using my template.”
- **AI Infers**: The user wants to import Markdown into a styled Word document. The `word/markdown/import` tool fits.
- **AI Action**: Uses `word/markdown/import` with `path=/docs/input.md`, `template=/docs/PlantillaEvolutio.docx`, and `output=/docs/output.docx`. If the template is missing, AI suggests a default template.
- **Interaction**: AI confirms Markdown and template paths, asking for output path if not provided.

### Reformatting Rescue
**Scenario**: `PropuestasRandom.docx` is a formatting mess. Fixing it? Pass the aspirin!
**Solution**:
- **User Tells AI**: “Reformat PropuestasRandom.docx to look professional.”
- **AI Infers**: The user wants to apply consistent styling. The `word/reformat` tool is appropriate.
- **AI Action**: Uses `word/reformat` with `document=/docs/PropuestasRandom.docx`, `styleSet=Professional`, and `output=/docs/PropuestasFormatted.docx`. AI suggests `Professional` style if unspecified.
- **Interaction**: AI asks for the document path and preferred style set.

### Embedded Object Extraction
**Scenario**: Extract PDFs and Excel files from `Composición.docx`. Word’s UI? A dig! 🦴
**Solution**:
- **User Tells AI**: “Extract all embedded files from Composición.docx.”
- **AI Infers**: The user wants to extract embedded objects. The `word/embedded-objects/extractAll` tool is ideal.
- **AI Action**: Uses `word/embedded-objects/extractAll` with `filePath=/docs/Composición.docx` and `outputDirectory=/docs/objetos_extraidos`. If output directory is missing, AI suggests `/docs/extracted`.
- **Interaction**: AI confirms the document path and asks for the output directory.

### Embedded Object Insertion
**Solution**:
- **User Tells AI**: “Embed data.xlsx into Reporte.docx.”
- **AI Infers**: The user wants to insert an embedded object. The `word/embedded-objects/insert` tool is suitable.
- **AI Action**: Uses `word/embedded-objects/insert` with `filePath=/docs/Reporte.docx` and `objectPath=/docs/data.xlsx`. If position is unspecified, AI suggests inserting at the document’s end.
- **Interaction**: AI asks for document and object paths, confirming insertion position.

### Embedded Object Modification
**Solution**:
- **User Tells AI**: “Update the chart in Presentacion.docx with updated_data.xlsx.”
- **AI Infers**: The user wants to replace an embedded object. The `word/embedded-objects/modify` tool fits.
- **AI Action**: Uses `word/embedded-objects/modify` with `filePath=/docs/Presentacion.docx`, `objectIndex=1`, and `newObjectPath=/docs/updated_data.xlsx`. If `objectIndex` is unclear, AI asks for the object’s position.
- **Interaction**: AI confirms paths and asks which object to replace.

### Embedded Object Deletion
**Solution**:
- **User Tells AI**: “Remove an embedded file from DocumentoLargo.docx.”
- **AI Infers**: The user wants to delete an embedded object. The `word/embedded-objects/delete` tool is appropriate.
- **AI Action**: Uses `word/embedded-objects/delete` with `filePath=/docs/DocumentoLargo.docx` and `objectIndex=3`. If `objectIndex` is missing, AI suggests listing objects first.
- **Interaction**: AI asks for the document path and object index.

### Mermaid Import Marvel
**Solution**:
- **User Tells AI**: “Add a flowchart to tech.docx.”
- **AI Infers**: The user wants to insert a Mermaid diagram. The `word/mermaid/import` tool is ideal.
- **AI Action**: Uses `word/mermaid/import` with `syntax=graph TD; A-->B`, `format=svg`, `position=paragraph:5`, and `document=/docs/tech.docx`. If syntax is missing, AI suggests a default diagram.
- **Interaction**: AI asks for the document path, diagram syntax, and insertion position.

### Mermaid Export Escapade
**Solution**:
- **User Tells AI**: “Export a Mermaid diagram from tech.docx as PNG.”
- **AI Infers**: The user wants to extract a diagram. The `word/mermaid/export` tool fits.
- **AI Action**: Uses `word/mermaid/export` with `diagram=1`, `format=png`, `output=/docs/diagram.png`, and `document=/docs/tech.docx`. If diagram index is unclear, AI suggests exporting the first diagram.
- **Interaction**: AI confirms the document path and diagram number.

### Analysis Antics
**Solution**:
- **User Tells AI**: “Analyze PropuestaTécnica.docx for jargon and add comments.”
- **AI Infers**: The user wants document analysis with comments. The `word/analyze` tool is suitable.
- **AI Action**: Uses `word/analyze` with `document=/docs/PropuestaTécnica.docx`, `criteria=technical`, and `output=/docs/PropuestaTécnica_Commented.docx`. AI suggests default criteria if unspecified.
- **Interaction**: AI asks for the document path and specific analysis criteria.

### Code Formatting Fiesta
**Solution**:
- **User Tells AI**: “Format code snippets in BuenasPrácticas.docx.”
- **AI Infers**: The user wants to style code. The `word/code-format` tool is designed for this.
- **AI Action**: Uses `word/code-format` with `filePath=/docs/BuenasPrácticas.docx`, `style=Código`, and `font=Consolas`. If style/font is missing, AI suggests `Consolas` and `Código` style.
- **Interaction**: AI confirms the document path and asks for preferred style/font.

### Instant Image Injection
**Solution**:
- **User Tells AI**: “Add a cat meme to ReporteTrimestral.docx.”
- **AI Infers**: The user wants to insert an image. The `word/image/insert` tool is appropriate.
- **AI Action**: Uses `word/image/insert` with `document=/docs/ReporteTrimestral.docx`, `imagePath=/memes/cat_typing.png`, and `position=paragraph:3`. If image path is missing, AI suggests a default image.
- **Interaction**: AI asks for the document path, image location, and insertion position.

### AI Ghostwriter
**Solution**:
- **User Tells AI**: “Write a conclusion for MiNovela.docx.”
- **AI Infers**: The user wants AI-generated text. The `word/generate-and-insert-text` tool fits.
- **AI Action**: Uses `word/generate-and-insert-text` with `document=/docs/MiNovela.docx`, `prompt=write-conclusion`, and `position=end`. AI suggests a default prompt if vague.
- **Interaction**: AI confirms the document path and asks for prompt details.

### Surgical Strike Styling
**Solution**:
- **User Tells AI**: “Style paragraph 5 of DiscursoMotivador.docx with emphasis.”
- **AI Infers**: The user wants precise styling. The `word/styles/apply` tool is ideal.
- **AI Action**: Uses `word/styles/apply` with `document=office://docs/DiscursoMotivador.docx?range=paragraph:5` and `style=Emphasis`. AI confirms the style exists.
- **Interaction**: AI asks for the document path and style name.

### The AI Word Analyst
**Solution**:
- **User Tells AI**: “Analyze InformeLargo.docx for key points and errors, and add comments for the author.”
- **AI Infers**: The user wants detailed analysis with comments. The `word/analyze` tool with `analyze` and `add` operations is suitable.
- **AI Action**: Uses `word/analyze` with `filePath=/docs/InformeLargo.docx`, `operation=analyze`, and `criteria=key points,errors,sections needing detail` to analyze. Then uses `word/analyze` with `operation=add` to insert comments based on results. If criteria are vague, AI suggests defaults like `grammar,structure`.
- **Interaction**: AI confirms the document path and asks for specific criteria if unclear.

### The Code Stylist in Word
**Solution**:
- **User Tells AI**: “Format the code snippets in DocumentoConCodigo.docx.”
- **AI Infers**: The user wants to format code. The `word/code-format` tool is designed for this.
- **AI Action**: Uses `word/code-format` with `filePath=/docs/DocumentoConCodigo.docx`, `style=Código`, and `font=Consolas`. If style/font is unspecified, AI suggests `Consolas` and `Código` for readability.
- **Interaction**: AI confirms the document path and asks if the user prefers a specific code style or font.

### The Multimodal Illustrator
**Solution**:
- **User Tells AI**: “Add Mafalda-style comic strips to ReporteCreativo.docx.”
- **AI Infers**: The user wants to insert images, possibly AI-generated. The `word/image/insert` tool is suitable, but image sourcing may be needed.
- **AI Action**: Uses `word/image/insert` with `filePath=/docs/ReporteCreativo.docx`, `imageDataBase64` (if available, else `imagePath`), `position=paragraph:5`, and `altText=Comic strip`. If no Mafalda images exist, AI suggests generating them (future capability) or using a placeholder.
- **Interaction**: AI asks for the document path, position, and whether to source/generate images.

## Excel Use Cases

### The Excel Accountant with Judgment
**Solution**:
- **User Tells AI**: “In Ventas2024.xlsx, highlight or comment on all sales over €1000.”
- **AI Infers**: The user wants to filter and annotate data. The `excel/data-analysis` tool with `filter` or `apply` operations is appropriate.
- **AI Action**: Uses `excel/data-analysis` with `filePath=/docs/Ventas2024.xlsx`, `sheetName=Sheet1`, `rangeAddress=A1:Z10000`, `operation=filter`, and `filterCriteria=[{"column":"MontoVenta","criteria1":1000,"operator":"xlGreater"}]`. If sheet/range is missing, AI suggests `Sheet1` and a default range or asks for clarification.
- **Interaction**: AI confirms the file path and asks for sheet name, range, and column name for sales data.

### The Express Excel Exporter
**Solution**:
- **User Tells AI**: “Export each tab of DatosComplejos.xlsx as a separate CSV.”
- **AI Infers**: The user wants to extract sheet data and save as CSVs. A workflow using `excel/worksheets`, `excel/range`, and `fs/file` is needed.
- **AI Action**: Uses `excel/worksheets` with `operation=list` and `filePath=/docs/DatosComplejos.xlsx` to get sheet names. For each sheet, uses `excel/range` with `operation=read` to get data, then `fs/file` with `operation=write` to save as `sheet_name.csv`. If output directory is missing, AI suggests `/docs/csv`.
- **Interaction**: AI confirms the file path and asks for the output directory.

## PowerPoint Use Cases

### The Christmas PowerPoint Decorator
**Solution**:
- **User Tells AI**: “Add Christmas decorations to InformeAnual.pptx.”
- **AI Infers**: The user wants festive elements like images or shapes. The `powerpoint/shapes` and `powerpoint/animations` tools are relevant.
- **AI Action**: Uses `powerpoint/shapes` with `filePath=/docs/InformeAnual.pptx`, `slideIndex=1`, `operation=insert`, `shapeType=msoShapeRectangle`, and `imagePath=/dynamic_storage/christmas_tree.png`. Uses `powerpoint/animations` to add transitions. If images are missing, AI suggests sourcing festive images or using placeholders.
- **Interaction**: AI asks for the file path, target slides, and preferred decoration style.

## Cross-Application Use Cases

### The Document Multitasker
**Solution**:
- **User Tells AI**: “Compare data in Datos1.xlsx and Datos2.xlsx, summarize in Informe.docx, and create a chart in Presentacion.pptx.”
- **AI Infers**: The user wants cross-app operations. Uses `excel/range`, `word/text/insert`, and `powerpoint/charts`.
- **AI Action**: Uses `excel/range` with `operation=read` for both Excel files. Processes data internally. Uses `word/text/insert` to add a summary to `Informe.docx`. Uses `powerpoint/charts` with `operation=insert` and `chartType=xlColumnClustered` for `Presentacion.pptx`. If ranges are missing, AI suggests default ranges.
- **Interaction**: AI confirms file paths, ranges, and chart location.

### The Magical Office Linker
**Solution**:
- **User Tells AI**: “Link a table from ReporteFinanciero.xlsx to InformeMensual.docx so it updates automatically.”
- **AI Infers**: The user wants a dynamic link. The `office/transfer` tool with `insert` and `link=true` is suitable.
- **AI Action**: Uses `office/transfer` with `source=excel:./ReporteFinanciero.xlsx:Sheet1:A1:F20`, `target=word:./InformeMensual.docx:end`, `operation=insert`, and `link=true`. If range/position is missing, AI suggests defaults.
- **Interaction**: AI confirms the Excel range and Word insertion point.

### The PowerPoint Embedder
**Solution**:
- **User Tells AI**: “Embed InformeCompleto.docx into PresentacionFinal.pptx on slide 5.”
- **AI Infers**: The user wants to embed a document. The `office/transfer` tool with `embed` is appropriate.
- **AI Action**: Uses `office/transfer` with `source=word:./InformeCompleto.docx`, `target=powerpoint:./PresentacionFinal.pptx:slide:5`, and `operation=embed`. If slide index is missing, AI suggests slide 1.
- **Interaction**: AI confirms file paths and slide index.

### The AI Document Creator
**Solution**:
- **User Tells AI**: “Create a formal letter about product inquiries and save as CartaFormal.docx.”
- **AI Infers**: The user wants AI-generated content in a new Word file. The `word/generate-and-insert-text` tool is suitable.
- **AI Action**: Uses `word/generate-and-insert-text` with `filePath=/docs/CartaFormal.docx`, `position=start`, and `prompt=Generate a formal letter requesting information addressed to company X about product Y`. If prompt details are vague, AI suggests a standard letter format.
- **Interaction**: AI confirms the file name and asks for prompt specifics.

### The Universal Office Translator
**Solution**:
- **User Tells AI**: “Work with this Office 365 document: https://contoso.sharepoint.com/docs/ReporteCloud.docx.”
- **AI Infers**: The user wants compatibility across Office versions. No specific tool, but Office MCP handles URLs/paths.
- **AI Action**: Confirms ability to process `https://contoso.sharepoint.com/docs/ReporteCloud.docx` using standard tools like `word/styles/list`. If URL access is restricted, AI suggests local file alternatives.
- **Interaction**: AI confirms the document URL/path and access permissions.

### The Credential Guardian
**Solution**:
- **User Tells AI**: “How does Office MCP handle my Office login information?”
- **AI Infers**: The user wants security details, not a tool operation.
- **AI Action**: Explains that Office MCP securely manages credentials for local and cloud Office, avoiding repeated logins. Refers to [Technical Details](TECHNICAL_DETAILS.md) for more.
- **Interaction**: No tool use; AI provides information.

### The Compressed File Manager
**Solution**:
- **User Tells AI**: “Find report.docx in archive.zip or compress documents into archive.zip.”
- **AI Infers**: The user wants ZIP file operations, listed as a future feature.
- **AI Action**: Informs the user that ZIP tools are planned (e.g., `fs/archive/list`, `fs/archive/create`). Suggests manual extraction or compression as a workaround.
- **Interaction**: AI confirms the ZIP file path and clarifies if the user wants to search or create a ZIP.

## Next Steps
- See [Installation & API](INSTALLATION_AND_API.md) for setup and tools.
- Explore [Technical Details](TECHNICAL_DETAILS.md) for implementation and plans.
- Return to [README](README.md).