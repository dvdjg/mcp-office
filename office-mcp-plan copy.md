# Office MCP Server Implementation Plan for Microsoft Office Automation

**Document Purpose**: This plan provides a comprehensive guide for an AI coder to implement an Office MCP (Model Context Protocol) server using FastMCP in TypeScript (see file FASTMCP_README.md) to automate Microsoft Office applications (Word, Excel, PowerPoint). It includes tool specifications, dependencies, testing strategies, use cases, and enhanced requirements for resource management and documentation, ensuring a robust, testable, and maintainable solution.

**Target Audience**: AI coder tasked with implementation.

**Date**: May 01, 2025

---

## 1. Overview

### 1.1 Objectives
The Office MCP server aims to:
- Automate Word, Excel, and PowerPoint using Office JavaScript APIs, Office Scripts, VBA, COM Interop, and Power Automate.
- Provide granular control over Word documents (styles, text, tables, charts, embedded objects, search/replace).
- Support file system operations (read/write directories, files, blobs).
- Enable Markdown import/export with VS Code-compatible features and comment handling.
- Integrate Mermaid diagrams (import, render, export as SVG/PNG).
- Facilitate PDF export and parsing.
- Address specific use cases (e.g., document merging, template creation, code formatting).
- Enhance AI-driven automation via FastMCP prompts and completions.
- Implement a resource management system for static and dynamic resources.
- Provide concise, standardized documentation.

### 1.2 Scope
- **Applications**: Word (.docx), Excel (.xlsx), PowerPoint (.pptx).
- **Formats**: Markdown (.md), PDF, SVG, PNG.
- **Platforms**: Windows (VBA/COM), cross-platform (Office JavaScript APIs, Office Scripts).
- **Use Cases**: Word-specific scenarios (e.g., convert to Markdown, merge documents).
- **Resources**: Static (e.g., AI guide) and dynamic (e.g., user documents, extracted objects).

### 1.3 Key Requirements
- **Modularity**: Reusable, documented tools organized by application/functionality.
- **Security**: Access controls for file system and document operations.
- **Performance**: Optimized for complex tasks (e.g., merging, reformatting).
- **Testability**: Unit, integration, and end-to-end tests with >90% coverage.
- **AI Integration**: Clear inputs/outputs for FastMCP prompts and completions.
- **Documentation**: Concise, standardized JSDoc for tools (<200 words each).

---

## 2. Implementation Guidelines

### 2.1 General Instructions
- **Language**: TypeScript for type safety and maintainability.
- **Framework**: FastMCP for server structure, resource templates, and AI prompts.
- **Project Structure**:
  ```
  mcp-office/
  ├── src/
  │   ├── server/          # FastMCP setup
  │   ├── tools/           # Tool implementations (word/, excel/, powerpoint/, office/)
  │   ├── utils/           # Helpers
  │   ├── resources/       # Static/dynamic resources
  │   ├── types/           # TypeScript types
  │   └── docs/            # Documentation templates
  ├── tests/               # Unit, integration, end-to-end tests
  ├── package.json
  └── tsconfig.json
  ```
- **Error Handling**: Use try-catch, validate inputs with `zod`, return clear error messages.
- **Logging**: Use `winston` for debugging and monitoring.
- **Documentation**: JSDoc for functions/tools, generate API docs with `typedoc`.
- **Version Control**: Git with Conventional Commits (e.g., `feat: add word/merge tool`).

### 2.2 FastMCP Setup
- Initialize FastMCP in `src/server/index.ts`.
- Define resource templates for tools (e.g., `word/styles`, `fs/directory`).
- Implement prompt templates (e.g., “apply style {style} to {range}”) and completions (e.g., style names).
- **Example**:
  ```typescript
  import { FastMCP } from 'fastmcp';

  const mcp = new FastMCP({
    port: 3000,
    resources: [
      {
        path: 'word/styles',
        handler: async (req) => { /* Style logic */ },
        completions: async () => ({ style: ['Heading1', 'Normal'] }),
      },
    ],
  });
  mcp.start();
  ```

### 2.3 Office Automation
- **Office JavaScript APIs**: Cross-platform automation via `@microsoft/office-js`.
- **Office Scripts**: Excel automation in Office 365.
- **VBA**: Complex Word automation (e.g., search/replace) on Windows.
- **COM Interop**: Windows-specific automation via `win32com` or `node-win32ole`.
- **Power Automate**: Event-driven workflows via REST APIs.

### 2.4 AI Integration
- **Prompt Templates**: Reusable prompts (e.g., “merge {doc1} and {doc2}”).
- **Completions**: Dynamic suggestions (e.g., file paths, styles).
- **Context Awareness**: Query document states (e.g., list styles).
- **Example**:
  ```typescript
  mcp.registerPrompt({
    template: 'apply style {style} to {range}',
    handler: async ({ style, range }) => { /* Call word/styles */ },
    completions: async () => ({ style: ['Heading1', 'Normal'] }),
  });
  ```

### 2.5 Security
- Validate file paths (`path.resolve`) to prevent directory traversal.
- Restrict document access with role-based permissions.
- Use `zod` for input validation.
- **Example**:
  ```typescript
  import { z } from 'zod';
  const fileSchema = z.object({
    path: z.string().refine((p) => p.startsWith('/allowed/path'), { message: 'Invalid path' }),
  });
  ```

### 2.6 Performance
- Cache metadata with `node-cache`.
- Use batch operations (e.g., `word/batch`) to reduce API calls.
- Implement async/await for I/O operations.

---

## 3. Tool Specifications

Tools are exposed as FastMCP resource templates, organized by application/functionality. Each includes operations, APIs, and examples.

### 3.1 File System Tools
- **fs/directory**: List, create, delete directories.
  - **Operations**: `list`, `create`, `delete`.
  - **API**: `fs.readdirSync`, `fs-extra`.
  - **Example**: `GET /fs/directory/list?path=/docs&filter=*.docx`
- **fs/file**: Read/write/delete files.
  - **Operations**: `read`, `write`, `delete`, `rename`.
  - **API**: `fs.readFileSync`, `fs-extra`.
  - **Example**: `POST /fs/file/write?path=/docs/note.txt&content=Hello`
- **fs/blob**: Handle document blobs.
  - **Operations**: `save`, `read`, `validate`.
  - **API**: `fs`, `Buffer`.
  - **Example**: `POST /fs/blob/save?filename=doc1.docx`

### 3.2 Word Tools
- **word/styles**: Manage styles.
  - **Operations**: `list`, `apply`, `create`, `modify`, `delete`.
  - **API**: Office JavaScript APIs (`Word.Style`), VBA.
  - **Example**: `POST /word/styles/apply?style=Heading1&range=paragraph:1`
- **word/text**: Manipulate text/paragraphs.
  - **Operations**: `insert`, `modify`, `delete`, `get`.
  - **API**: `Word.Paragraph`, VBA.
  - **Example**: `POST /word/text/insert?text=Hello&position=paragraph:1`
- **word/search-replace**: Advanced search/replace.
  - **Operations**: `search`, `replace`, `preview`.
  - **API**: VBA (`Find/Replace`), `Word.Range.search`.
  - **Example**: `POST /word/search-replace?find=red&replace=blue&criteria=color`
- **word/page**: Configure page layout.
  - **Operations**: `set`, `modify`, `get`.
  - **API**: VBA (`PageSetup`), `Word.Section`.
  - **Example**: `POST /word/page/set?size=A4&margins=1in`
- **word/headers-footers**: Manage headers/footers.
  - **Operations**: `insert`, `modify`, `delete`, `configure`.
  - **API**: `Word.HeaderFooter`, VBA.
  - **Example**: `POST /word/headers-footers/insert?section=1&text=Confidential`
- **word/tables**: Create/manage tables.
  - **Operations**: `insert`, `modify`, `add`, `delete`, `merge`, `split`.
  - **API**: `Word.Table`, VBA.
  - **Example**: `POST /word/tables/insert?rows=3&columns=4&style=Grid`
- **word/charts**: Insert/manage charts.
  - **Operations**: `insert`, `modify`, `delete`, `reposition`.
  - **API**: `Word.Chart`, VBA.
  - **Example**: `POST /word/charts/insert?type=bar&data=[[1,2],[3,4]]`
- **word/embedded-objects**: Manage embedded objects.
  - **Operations**: `insert`, `modify`, `delete`, `extractAll`.
  - **API**: VBA (`OLEObjects`), COM Interop.
  - **Example**: `POST /word/embedded-objects/extractAll?document=/docs/Composición.docx`
- **word/metadata**: Manage properties/comments.
  - **Operations**: `set`, `get`, `add`, `remove`, `manage`.
  - **API**: `Word.DocumentProperties`, VBA.
  - **Example**: `POST /word/metadata/set?property=author&value=JohnDoe`
- **word/batch**: Execute multiple operations.
  - **Operations**: `run`, `transaction`.
  - **API**: Office JavaScript APIs, VBA.
  - **Example**: `POST /word/batch?ops=[{tool:styles/apply,...}]`
- **word/markdown/import**: Convert Markdown to Word.
  - **Operations**: `parse`, `map`, `applyTemplate`.
  - **API**: `markdown-it`, `Word.Document`.
  - **Example**: `POST /word/markdown/import?path=/docs/input.md&template=/docs/PlantillaEvolutio.docx`
- **word/markdown/export**: Convert Word to Markdown.
  - **Operations**: `extract`, `generate`, `appendComments`.
  - **API**: `markdown-it`, `Word.Comment`.
  - **Example**: `POST /word - **word/mermaid/import**: Import/render Mermaid diagrams.
  - **Operations**: `parse`, `render`, `insert`.
  - **API**: `mermaid`, `Word.ContentControl`.
  - **Example**: `POST /word/mermaid/import?syntax=graph TD; A-->B&format=svg`
- **word/mermaid/export**: Export Mermaid diagrams.
  - **Operations**: `identify`, `extract`, `save`.
  - **API**: `mermaid`, Office JavaScript APIs.
  - **Example**: `POST /word/mermaid/export?diagram=1&format=png`
- **word/merge**: Merge Word documents.
  - **Operations**: `compare`, `merge`, `resolve`.
  - **API**: `Word.Document`, VBA.
  - **Example**: `POST /word/merge?docs=[/docs/doc1.docx,/docs/doc2.docx]`
- **word/template**: Create templates with placeholders.
  - **Operations**: `analyze`, `replace`, `save`.
  - **API**: `Word.ContentControl`, VBA.
  - **Example**: `POST /word/template?document=/docs/PresupuestosEvolutio.docx`
- **word/reformat**: Reformat documents.
  - **Operations**: `analyze`, `apply`, `standardize`.
  - **API**: `Word.Style`, VBA.
  - **Example**: `POST /word/reformat?document=/docs/PropuestasRandom.docx&styleSet=Professional`
- **word/analyze**: Analyze and comment.
  - **Operations**: `analyze`, `add`, `summarize`.
  - **API**: `Word.Comment`, FastMCP prompts.
  - **Example**: `POST /word/analyze?document=/docs/PropuestaTécnica.docx&criteria=technical`
- **word/code-format**: Format code with syntax highlighting.
  - **Operations**: `identify`, `detect`, `apply`.
  - **API**: `highlight.js`, `Word.Font`.
  - **Example**: `POST /word/code-format?document=/docs/BuenasPrácticas.docx&style=Código`

### 3.3 Excel Tools
- **excel/worksheets**: Manage worksheets.
  - **Operations**: `add`, `delete`, `rename`, `set`.
  - **API**: `ExcelScript.Workbook`, VBA.
  - **Example**: `POST /excel/worksheets/add?name=Data`
- **excel/range**: Manipulate ranges.
  - **Operations**: `read`, `write`, `format`, `apply`.
  - **API**: `ExcelScript.Range`.
  - **Example**: `POST /excel/range/set?range=A1:B2&values=[[1,2],[3,4]]`
- **excel/tables**: Create/manage tables.
  - **Operations**: `insert`, `modify`, `add`, `delete`.
  - **API**: `ExcelScript.Table`.
  - **Example**: `POST /excel/tables/insert?range=A1:D10`
- **excel/charts**: Insert/manage charts.
  - **Operations**: `insert`, `modify`, `delete`, `reposition`.
  - **API**: `ExcelScript.Chart`.
  - **Example**: `POST /excel/charts/insert?type=column`
- **excel/data-analysis**: Process data.
  - **Operations**: `sort`, `filter`, `pivot`, `calculate`.
  - **API**: Office Scripts, VBA.
  - **Example**: `POST /excel/data-analysis/sort?range=A1:D100`

### 3.4 PowerPoint Tools
- **powerpoint/slides**: Manage slides.
  - **Operations**: `add`, `delete`, `set`.
  - **API**: `PowerPoint.Slide`, VBA.
  - **Example**: `POST /powerpoint/slides/add?layout=TitleSlide`
- **powerpoint/shapes**: Manipulate shapes/text.
  - **Operations**: `insert`, `modify`, `format`.
  - **API**: `PowerPoint.Shape`, VBA.
  - **Example**: `POST /powerpoint/shapes/insert?type=textbox`
- **powerpoint/properties**: Manage settings.
  - **Operations**: `set`, `configure`, `add`.
  - **API**: `PowerPoint.Presentation`, VBA.
  - **Example**: `POST /powerpoint/properties/set?size=16:9`
- **powerpoint/animations**: Manage animations/transitions.
  - **Operations**: `add`, `configure`.
  - **API**: VBA (`SlideShowTransition`).
  - **Example**: `POST /powerpoint/animations/add?slide=1&type=fade`

### 3.5 Cross-Application Tools
- **office/transfer**: Move data between apps.
  - **Operations**: `embed`, `insert`, `copy`.
  - **API**: COM Interop, VBA.
  - **Example**: `POST /office/transfer?source=excel:range:A1:B2&target=word:paragraph:1`
- **office/workflow**: Multi-step automation.
  - **Operations**: `run`, `trigger`.
  - **API**: Power Automate, Office Scripts.
  - **Example**: `POST /office/workflow?steps=[{tool:excel/tables/insert,...}]`
- **office/ai-suggest**: AI-driven recommendations.
  - **Operations**: `format`, `search`, `chart`.
  - **API**: FastMCP prompts.
  - **Example**: `POST /office/ai-suggest?context=word:document&task=format`
- **office/pdf/export**: Export to PDF.
  - **Operations**: `convert`, `configure`, `save`.
  - **API**: `Document.saveAs`, `pdfkit`.
  - **Example**: `POST /office/pdf/export?document=/docs/doc1.docx`
- **office/pdf/parse**: Extract PDF content.
  - **Operations**: `parse`, `convert`.
  - **API**: `pdf-parse`, `pdf2pic`.
  - **Example**: `POST /office/pdf/parse?file=/docs/offer.pdf`
- **office/combine**: Combine files into Word.
  - **Operations**: `read`, `normalize`, `insert`.
  - **API**: `pdf-parse`, VBA.
  - **Example**: `POST /office/combine?directory=/docs/Oferta`
- **office/word-to-powerpoint**: Convert Word to PowerPoint.
  - **Operations**: `analyze`, `create`, `transfer`.
  - **API**: `Word.Document`, `PowerPoint.Slide`.
  - **Example**: `POST /office/word-to-powerpoint?document=/docs/PasoAPaso.docx`

### 3.6 Resource Management
- **memory/ai_assistant_guide**: Static guide for API usage.
  - **Operations**: `read`, `list`.
  - **Content**: JSON with sections (introduction, tool_usage, use_cases, best_practices).
  - **API**: FastMCP resource.
  - **Example**: `GET /memory/ai_assistant_guide?section=tool_usage`
  - **Implementation**: Store in `src/resources/static/ai_assistant_guide.ts`.
- **dynamic/resources**: Manage user-generated content.
  - **Operations**: `list`, `read`, `write`, `delete`, `metadata`, `search`.
  - **Resource Types**: Documents (.docx, .xlsx, .pptx), images (PNG, SVG), embedded objects, sub-documents.
  - **API**: FastMCP resource, `fs-extra`, `node-cache`.
  - **Example**: `GET /dynamic/resources/list?type=docx`
  - **Implementation**: Store in `/dynamic_storage`, cache metadata.

---

## 4. Dependencies
Install via `npm`:
- **Core**: `fastmcp`, `typescript`, `node` (v18+), `@microsoft/office-js`, `win32com`/`node-win32ole`, `zod`, `winston`, `node-cache`.
- **File System**: `fs-extra`.
- **Markdown**: `markdown-it`.
- **Mermaid**: `mermaid`.
- **PDF**: `pdf-parse`, `pdf2pic`, `pdfkit`.
- **Syntax Highlighting**: `highlight.js`.
- **Testing**: `jest`, `mocha`.
- **Optional**: `guesslang`, `puppeteer`.
```bash
npm install fastmcp typescript @microsoft/office-js zod winston node-cache fs-extra markdown-it mermaid pdf-parse pdf2pic pdfkit highlight.js jest
```

---

## 5. Testing Strategy

### 5.1 Test Types
- **Unit Tests**: Test tool operations (e.g., `word/styles/apply`) with Jest.
- **Integration Tests**: Test tool interactions (e.g., `word/merge` + `word/styles`).
- **End-to-End Tests**: Test workflows (e.g., convert Word to Markdown) with Mocha.
- **Performance Tests**: Measure complex operations (e.g., merge 100-page docs <10s).
- **Security Tests**: Validate inputs and access controls.
- **Documentation Tests**: Verify JSDoc coverage and length (<200 words/tool).

### 5.2 Test Cases for Use Cases
1. **Convert Word to Markdown with Comments**:
   - Tool: `word/markdown/export`.
   - Test: Convert `CV.docx` to `CV.md`, verify comments appended.
2. **Merge Two Word Documents**:
   - Tool: `word/merge`.
   - Test: Merge `cuentoAladdin.docx` drafts, verify content preserved.
3. **Create Template with Placeholders**:
   - Tool: `word/template`.
   - Test: Generate template from `PresupuestosEvolutio.docx`, verify placeholders.
4. **Convert Markdown to Word with Template**:
   - Tool: `word/markdown/import`.
   - Test: Convert `input.md` with `PlantillaEvolutio.docx`, verify styles.
5. **Reformat Poorly Formatted Document**:
   - Tool: `word/reformat`.
   - Test: Reformat `PropuestasRandom.docx`, verify professional formatting.
6. **Extract Embedded Documents**:
   - Tool: `word/embedded-objects`.
   - Test: Extract files from `Composición.docx`, verify saved.
7. **Combine Multiple Documents**:
   - Tool: `office/combine`.
   - Test: Combine files in `Oferta`, verify Word output.
8. **Convert Word to PowerPoint**:
   - Tool: `office/word-to-powerpoint`.
   - Test: Convert `PasoAPaso.docx` to `.pptx`, verify slides.
9. **Analyze Document and Add Comments**:
   - Tool: `word/analyze`.
   - Test: Analyze `PropuestaTécnica.docx`, verify comments.
10. **Format Code and Metadata**:
    - Tool: `word/code-format`.
    - Test: Format code in `BuenasPrácticas.docx`, verify highlighting.

### 5.3 Test Implementation
- **Setup**: Use `tests/` with `unit/`, `integration/`, `e2e/`, and `fixtures/`.
- **Coverage**: >90% with `jest --coverage`.
- **Automation**: Run via `npm test` in CI/CD.
- **Data**: Use sample docs (e.g., `CV.docx`) and malformed inputs.
```json
{
  "scripts": {
    "test": "jest",
    "test:coverage": "jest --coverage",
    "test:e2e": "mocha e2e",
    "test:docs": "jest tests/docs"
  }
}
```

---

## 6. Enhanced Documentation Guidelines

### 6.1 Objectives
- Provide concise, standardized JSDoc (<200 words/tool).
- Support automated generation with `typedoc`.
- Integrate with `memory://ai_assistant_guide`.

### 6.2 Standards
- **Format**: JSDoc with description, path, operations, parameters, returns, example, errors.
- **Example**:
  ```typescript
  /**
   * Manages styles in Word documents.
   * @path word/styles
   * @operations
   *   - list: Retrieves styles.
   *   - apply: Applies a style.
   * @parameters
   *   | Name  | Type   | Description          | Optional |
   *   |-------|--------|----------------------|----------|
   *   | style | string | Style name           | No       |
   *   | range | string | Range (e.g., paragraph:1) | Yes |
   * @returns JSON (e.g., { success: true })
   * @example POST /word/styles/apply?style=Heading1
   * @errors
   *   - 400: Invalid style
   */
  ```
- **Length**: Description <50 words, operations <20 words each, parameters <30 words each.
- **Integration**: Summarize in `memory://ai_assistant_guide`.
- **Generation**: Use `typedoc` with:
  ```json
  {
    "entryPoints": ["src/tools"],
    "out": "docs/api",
    "excludePrivate": true
  }
  ```

### 6.3 Implementation
- Create JSDoc template in `src/docs/templates/tool.jsdoc`.
- Validate length in CI/CD.
- Test coverage and accessibility in `memory://ai_assistant_guide`.

---

## 7. Implementation Steps
1. **Setup Project**:
   - Initialize Node.js/TypeScript (`npm init`, `tsc --init`).
   - Install dependencies.
   - Configure `tsconfig.json`:
     ```json
     {
       "compilerOptions": {
         "target": "ES2020",
         "module": "commonjs",
         "strict": true,
         "outDir": "./dist",
         "rootDir": "./src"
       }
     }
     ```
2. **Implement FastMCP Server**:
   - Create `src/server/index.ts`.
   - Define resource templates and prompts.
3. **Implement Tools**:
   - Organize in `src/tools/` (e.g., `word/`, `excel/`).
   - Example (`word/styles`):
     ```typescript
     import { Word } from '@microsoft/office-js';
     import { z } from 'zod';

     const schema = z.object({
       style: z.string(),
       range: z.string(),
     });

     export const wordStyles = {
       path: 'word/styles',
       handler: async (req) => {
         const { style, range } = schema.parse(req.params);
         await Word.run(async (context) => {
           const paragraph = context.document.body.paragraphs.getRange(range);
           paragraph.load('style');
           await context.sync();
           paragraph.style = style;
           await context.sync();
         });
         return { success: true };
       },
       completions: async () => ({
         style: ['Heading1', 'Normal'],
         range: ['paragraph:1', 'paragraph:2'],
       }),
     };
     ```
4. **Implement Resources**:
   - Static: `src/resources/static/ai_assistant_guide.ts` with JSON content.
   - Dynamic: `src/tools/dynamic/resources.ts` with `list`, `read`, etc.
     ```typescript
     import { z } from 'zod';
     import fs from 'fs-extra';
     import { nanoid } from 'nanoid';

     const resourceSchema = z.object({
       id: z.string().optional(),
       path: z.string().optional(),
       type: z.enum(['docx', 'xlsx', 'pptx', 'png', 'svg', 'pdf', 'other']).optional(),
     });

     mcp.registerResource({
       path: 'dynamic/resources',
       handler: async (req) => {
         const { operation, id, path, type } = req.params;
         switch (operation) {
           case 'list':
             const files = await fs.readdir('/dynamic_storage');
             return files.map((file) => ({
               id: nanoid(),
               path: `/dynamic_storage/${file}`,
               type: file.split('.').pop(),
             }));
           case 'read':
             const data = await fs.readFile(path);
             return { path, data };
         }
       },
     });
     ```
5. **Integrate Dependencies**:
   - Use `markdown-it`, `mermaid`, `pdf-parse`, `highlight.js`.
   - Example (`word/code-format`):
     ```typescript
     import hljs from 'highlight.js';
     import { Word } from '@microsoft/office-js';

     export const wordCodeFormat = {
       path: 'word/code-format',
       handler: async (req) => {
         const { document, style, font } = req.params;
         await Word.run(async (context) => {
           const paragraphs = context.document.body.paragraphs;
           paragraphs.load('text');
           await context.sync();
           paragraphs.items.forEach((p) => {
             const code = hljs.highlightAuto(p.text).value;
             // Apply style and colors
           });
         });
         return { success: true };
       },
     };
     ```
6. **Implement Tests**:
   - Write unit, integration, end-to-end, and documentation tests.
   - Example:
     ```typescript
     import { wordStyles } from '../src/tools/word/styles';
     import { mockWord } from './mocks';

     describe('word/styles', () => {
       it('applies style', async () => {
         mockWord.setup();
         const result = await wordStyles.handler({ style: 'Heading1', range: 'paragraph:1' });
         expect(result.success).toBe(true);
       });
     });
     ```
7. **Document Code**:
   - Use JSDoc for all tools.
   - Generate API docs with `typedoc`.
8. **Optimize and Deploy**:
   - Profile with `performance.now()`.
   - Optimize caching and streaming.
   - Deploy with `pm2` or Docker.

---

## 8. Use Case Implementation Notes
1. **Convert Word to Markdown with Comments**:
   - Use `word/markdown/export`, append comments with `markdown-it`.
   - Test: Verify comments in `## Comments` section.
2. **Merge Two Word Documents**:
   - Use `word/merge` with AI conflict resolution.
   - Test: Verify unique content preserved.
3. **Create Template with Placeholders**:
   - Use `word/template`, AI to detect sensitive text.
   - Test: Verify placeholders replace data.
4. **Convert Markdown to Word with Template**:
   - Use `word/markdown/import`, map headers to styles.
   - Test: Verify styles and content accuracy.
5. **Reformat Poorly Formatted Document**:
   - Use `word/reformat`, AI to detect sections.
   - Test: Verify consistent formatting.
6. **Extract Embedded Documents**:
   - Use `word/embedded-objects/extractAll`, save with unique names.
   - Test: Verify all files extracted.
7. **Combine Multiple Documents**:
   - Use `office/combine`, integrate with `pdf-parse`.
   - Test: Verify consistent formatting.
8. **Convert Word to PowerPoint**:
   - Use `office/word-to-powerpoint`, preserve formatting.
   - Test: Verify slides match structure.
9. **Analyze Document and Add Comments**:
   - Use `word/analyze`, AI for issue detection.
   - Test: Verify relevant comments.
10. **Format Code and Metadata**:
    - Use `word/code-format`, `highlight.js` for highlighting.
    - Test: Verify Consolas and colors.

---

## 9. Deliverables
- **Source Code**: `mcp-office/` repository.
- **Tests**: Unit, integration, end-to-end, documentation tests (>90% coverage).
- **Documentation**: API docs (`docs/api`), JSDoc template, user guide.
- **Resources**: `memory://ai_assistant_guide`, `dynamic/resources`.
- **Sample Files**: Test fixtures for use cases.
- **Deployment Script**: `npm start` or Docker config.

---

## 10. Notes for AI Coder
- **Testability**: Write testable code with mocked dependencies.
- **TypeScript**: Use strict typing and interfaces.
- **Use Cases**: Ensure all 10 are supported with end-to-end tests.
- **Edge Cases**: Handle empty docs, invalid inputs, unsupported formats.
- **AI Optimization**: Design tools for clear prompts/completions.
- **Resources**: Implement `memory://ai_assistant_guide` early, use streaming for `dynamic/resources`.
- **Documentation**: Enforce brevity, reuse in `memory://ai_assistant_guide`.

---

## 11. Approval and Next Steps
- **Approval**: Confirm or provide feedback (e.g., prioritize `word/code-format`).
- **Steps**:
  1. Initialize project and install dependencies.
  2. Implement FastMCP server and core tools.
  3. Develop/test tools incrementally, starting with Word.
  4. Validate use cases with end-to-end tests.
  5. Submit for review.

**Feedback Request**: Approve or suggest modifications, including prioritization or additional scenarios for `memory://ai_assistant_guide`.

---

This plan provides a clear, actionable roadmap for implementing the Office MCP server, integrating all requirements and ensuring a robust, AI-driven solution for Office automation. Contact for code samples or clarification.