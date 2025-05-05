# Markdown Features Supported by Visual Studio Code

Visual Studio Code (VSCode) supports a robust set of Markdown features, enhanced by its built-in renderer and popular extensions like **Markdown All in One**, **Markdown Preview Enhanced**, and **Markdown+Math**. Below is a comprehensive overview of the Markdown features supported in VSCode, including standard Markdown syntax, GitHub Flavored Markdown (GFM), and additional capabilities like strikethrough, Mermaid diagrams, and LaTeX-style math formulas.

## 1. Basic Markdown Syntax
VSCode supports standard Markdown syntax based on John Gruber’s original specification, with enhancements from Pandoc and GFM.

- **Headings**: Create headings using `#` (H1 to H6).
  ```markdown
  # Heading 1
  ## Heading 2
  ### Heading 3
  ```
- **Bold and Italics**:
  ```markdown
  **Bold text** or __Bold text__
  *Italic text* or _Italic text_
  ```
- **Combined Emphasis**:
  ```markdown
  **_Bold and italic_**
  ```
- **Lists**:
  - Unordered lists with `-`, `*`, or `+`:
    ```markdown
    - Item 1
    - Item 2
      - Subitem
    ```
  - Ordered lists:
    ```markdown
    1. First item
    2. Second item
    ```
- **Links**:
  ```markdown
  [Link text](https://example.com)
  ```
- **Images**:
  ```markdown
  ![Alt text](image.jpg)
  ```
- **Blockquotes**:
  ```markdown
  > This is a blockquote.
  > > Nested blockquote.
  ```
- **Code**:
  - Inline code: `` `code` ``
  - Code blocks:
    ```markdown
    ```javascript
    console.log("Hello, world!");
    ```
    ```
- **Horizontal Rules**:
  ```markdown
  ---
  ```

## 2. GitHub Flavored Markdown (GFM) Features
VSCode’s Markdown renderer supports GFM, which adds features commonly used on GitHub.

- **Strikethrough**:
  ```markdown
  ~~Strikethrough text~~
  ```
  Renders as ~~Strikethrough text~~. Enabled via the `markdown-it-strikethrough-alt` plugin in settings.[](https://www.markdowntoolbox.com/blog/markdown-in-visual-studio-code-shortcuts-and-commands/)

- **Task Lists**:
  ```markdown
  - [x] Completed task
  - [ ] Incomplete task
  ```
  Renders as checkboxes in the preview.

- **Tables**:
  ```markdown
  | Header 1 | Header 2 |
  |----------|----------|
  | Cell 1   | Cell 2   |
  ```
  Align text using `:` in the header row (e.g., `:--` for left, `--:` for right, `:--:` for center).

- **Autolinks**:
  ```markdown
  https://example.com
  ```
  Automatically converts to clickable links.

- **Emoji**:
  ```markdown
  :smile: :rocket:
  ```
  Requires the `markdown-it-emoji` plugin in settings.[](https://www.markdowntoolbox.com/blog/markdown-in-visual-studio-code-shortcuts-and-commands/)

## 3. Mermaid Diagrams
VSCode supports Mermaid diagrams for creating flowcharts, sequence diagrams, Gantt charts, and more using text-based syntax. This is enabled through extensions like **Markdown Preview Mermaid Support** or **Mermaid Chart**.[](https://marketplace.visualstudio.com/items?itemName=bierner.markdown-mermaid)[](https://docs.mermaidchart.com/blog/posts/mermaid-chart-vs-code-plugin-create-and-edit-mermaid-js-diagrams-in-visual-studio-code)

- **Usage**:
  ```markdown
  ```mermaid
  graph TD;
      A-->B;
      A-->C;
      B-->D;
      C-->D;
  ```
  ```
  Renders as a flowchart in the preview pane. Supported diagram types include flowcharts, sequence diagrams, class diagrams, Gantt charts, and entity-relationship diagrams.[](https://www.freecodecamp.org/news/diagrams-as-code-with-mermaid-github-and-vs-code/)[](https://docs.mermaidchart.com/blog/posts/mermaid-chart-vs-code-plugin-create-and-edit-mermaid-js-diagrams-in-visual-studio-code)

- **Integration**:
  - The **Mermaid Chart VS Code Plugin** allows real-time editing and previewing of Mermaid diagrams, with features like pan/zoom and cloud sync for collaboration.[](https://docs.mermaidchart.com/blog/posts/mermaid-chart-vs-code-plugin-create-and-edit-mermaid-js-diagrams-in-visual-studio-code)
  - Diagrams in `.md` files are auto-detected and rendered.[](https://docs.mermaidchart.com/blog/posts/mermaid-chart-vs-code-plugin-create-and-edit-mermaid-js-diagrams-in-visual-studio-code)

## 4. LaTeX-Style Math Formulas
VSCode supports rendering mathematical expressions using LaTeX syntax, powered by KaTeX, through extensions like **Markdown+Math** or **Markdown Preview Enhanced**.[](https://marketplace.visualstudio.com/items?itemName=goessner.mdmath)[](https://www.markdowntoolbox.com/blog/markdown-in-visual-studio-code-shortcuts-and-commands/)

- **Inline Math**:
  ```markdown
  Inline math: $E = mc^2$
  ```
  Renders as \( E = mc^2 \).

- **Math Blocks**:
  ```markdown
  $$
  \sum_{k=1}^n a_k b_k \leq \sqrt{\sum_{k=1}^n a_k^2} \sqrt{\sum_{k=1}^n b_k^2}
  $$
  ```
  Renders as a centered equation.

- **Supported Commands**:
  - Standard LaTeX commands like `\sum`, `\frac`, `\sqrt`, `\ket`, `\bra`, and more.
  - Limited support for advanced commands (e.g., `\ket{\psi}` works in newer VSCode versions).[](https://stackoverflow.com/questions/67562772/vscode-markdown-preview-not-supporting-some-math-commands)
  - Enable/disable via `"markdown.math.enabled": true/false` in settings.[](https://stackoverflow.com/questions/67562772/vscode-markdown-preview-not-supporting-some-math-commands)

- **Limitations**:
  - KaTeX supports a subset of LaTeX. Some commands (e.g., equation numbering) may require specific extension configurations.[](https://stackoverflow.com/questions/67562772/vscode-markdown-preview-not-supporting-some-math-commands)
  - Accessibility support for math is limited.[](https://stackoverflow.com/questions/39207954/how-can-i-insert-mathematical-formulas-into-a-markdown-file-with-visual-studio-c)

## 5. Additional Features via Extensions
Extensions enhance VSCode’s Markdown capabilities beyond the built-in renderer.

- **Markdown Preview Enhanced**:
  - Supports GitHub Flavored Markdown, Mermaid, PlantUML, and KaTeX.
  - Allows embedding media (images, videos) and exporting to HTML/PDF.
  - Customizable via settings for markdown-it plugins.[](https://www.markdowntoolbox.com/blog/markdown-in-visual-studio-code-shortcuts-and-commands/)[](https://www.markdowntoolbox.com/es/blog/atajos-y-comandos-de-markdown-en-visual-studio-code/)

- **Markdown All in One**:
  - Adds keyboard shortcuts, auto-completion, and table of contents generation.
  - Supports Pandoc-style LaTeX math and syntax highlighting for code blocks.
  - Compatible with other Markdown extensions.[](https://markdown-all-in-one.github.io/docs/guide/)[](https://marketplace.visualstudio.com/items?itemName=yzhang.markdown-all-in-one)

- **Markdown+Math**:
  - Enhances LaTeX math rendering with macro support.
  - Improves preview for complex equations.[](https://marketplace.visualstudio.com/items?itemName=goessner.mdmath)

- **Mermaid Markdown Syntax Highlighting**:
  - Provides syntax highlighting for Mermaid code blocks in the editor.[](https://marketplace.visualstudio.com/items?itemName=bpruitt-goddard.mermaid-markdown-syntax-highlighting)

## 6. Code Blocks and Syntax Highlighting
VSCode supports fenced code blocks with syntax highlighting for over 100 programming languages.

- **Example**:
  ```markdown
  ```python
  def hello_world():
      print("Hello, world!")
  ```
  ```
- **Custom Language**:
  Specify the language after the opening ``` (e.g., `python`, `javascript`, `mermaid`).

## 7. Other Features
- **Reference Links**:
  ```markdown
  [Link text][id]
  [id]: https://example.com
  ```
- **Footnotes** (via extensions like Markdown Preview Enhanced):
  ```markdown
  Here is a footnote reference[^1].
  [^1]: This is the footnote.
  ```
- **HTML Embedding**:
  ```markdown
  <div style="color: red;">Red text</div>
  ```
  Supported for cases where Markdown lacks features (e.g., text highlighting).[](https://joplinapp.org/help/apps/markdown/)

- **Collapsible Sections** (GFM):
  ```markdown
  <details>
  <summary>Click to expand</summary>
  Hidden content here.
  </details>
  ```

- **Table of Contents**:
  Auto-generated by **Markdown All in One** using:
  ```markdown
  <!-- toc -->
  ```
  Updates automatically on save.[](https://markdown-all-in-one.github.io/docs/guide/)

## 8. Customization
- **Settings**:
  Customize Markdown rendering via `settings.json`:
  ```json
  {
    "markdown.markdownItPlugins": [
      "markdown-it-emoji",
      "markdown-it-task-lists",
      {"markdown-it-strikethrough-alt": {}}
    ],
    "markdown.math.enabled": true
  }
  ```
  Enables strikethrough, task lists, emojis, and math rendering.[](https://www.markdowntoolbox.com/blog/markdown-in-visual-studio-code-shortcuts-and-commands/)

- **Preview Customization**:
  Adjust preview styles via CSS or extension settings in **Markdown Preview Enhanced**.[](https://www.markdowntoolbox.com/blog/markdown-in-visual-studio-code-shortcuts-and-commands/)

## 9. Limitations
- **Math Accessibility**: KaTeX lacks built-in accessibility support, which may affect screen reader compatibility.[](https://stackoverflow.com/questions/39207954/how-can-i-insert-mathematical-formulas-into-a-markdown-file-with-visual-studio-c)
- **Mermaid Diagram Themes**: Diagrams render on a white background, regardless of VSCode’s theme, to ensure color compatibility.[](https://joplinapp.org/help/apps/markdown/)
- **Advanced LaTeX**: Some LaTeX commands (e.g., custom macros or equation numbering) may require additional configuration or extensions.[](https://stackoverflow.com/questions/67562772/vscode-markdown-preview-not-supporting-some-math-commands)

## 10. Recommended Extensions
To fully leverage VSCode’s Markdown capabilities, install:
- **Markdown All in One**: For shortcuts, TOC, and general enhancements.
- **Markdown Preview Enhanced**: For Mermaid, PlantUML, and advanced math.
- **Markdown+Math**: For robust LaTeX math support.
- **Mermaid Chart** or **Markdown Preview Mermaid Support**: For Mermaid diagram rendering.
- **Mermaid Markdown Syntax Highlighting**: For better Mermaid code editing.

## Conclusion
VSCode’s Markdown support, combined with extensions, provides a powerful environment for creating rich documents with strikethrough, Mermaid diagrams, LaTeX-style math, and more. By leveraging GFM and extensions like Markdown Preview Enhanced, users can create dynamic, visually appealing documentation directly within VSCode.

For more details, explore the [VSCode Markdown documentation](https://code.visualstudio.com/docs/languages/markdown) or the extension pages on the [Visual Studio Marketplace](https://marketplace.visualstudio.com/vscode).