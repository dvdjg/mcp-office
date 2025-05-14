## Detailed Test Plan and Documentation Outline

**Goal:** Create comprehensive tests for all Office MCP tools and documentation on test execution and interpretation.

**Phase 1: Test Creation**

This phase focuses on writing new test cases for tools that currently have limited or no dedicated tests, as well as adding tests for parameter variations for existing tools. Tests will be categorized by tool module and progress from simple to more complex scenarios.

**1. File System (fs) Tools:**

*   **fs/directory:**
    *   **Simple:**
        *   `list`: Test listing files in a known directory (`tests/fixtures/`).
        *   `create`: Test creating a new directory.
        *   `delete`: Test deleting an empty directory.
    *   **Complex/Variations:**
        *   `list`: Test with `filter` parameter (e.g., `*.docx`, `*.ts`). Test recursive listing. Test listing in a non-existent directory (error handling).
        *   `create`: Test creating nested directories. Test creating a directory that already exists (error handling).
        *   `delete`: Test deleting a non-empty directory (should fail or require force, depending on tool implementation - need to check API spec). Test deleting a non-existent directory (error handling).
*   **fs/file:**
    *   **Simple:**
        *   `read`: Test reading a small text file (`README.md`).
        *   `write`: Test writing a new text file.
        *   `delete`: Test deleting a newly created file.
    *   **Complex/Variations:**
        *   `read`: Test reading a binary file (e.g., `.docx`). Test reading a non-existent file (error handling).
        *   `write`: Test overwriting an existing file. Test writing to a path with non-existent parent directories (should be created by the tool).
        *   `delete`: Test deleting a non-existent file (error handling).
*   **fs/blob:**
    *   **Simple:**
        *   `save`: Test saving a small blob (e.g., base64 encoded text).
        *   `read`: Test reading a previously saved blob.
    *   **Complex/Variations:**
        *   `save`: Test saving a larger blob (e.g., base64 encoded image).
        *   `read`: Test reading a non-existent blob (error handling).
*   **fs/archive:**
    *   **Simple:**
        *   `list`: Test listing contents of a known `.zip` file.
        *   `extract`: Test extracting a single file from a `.zip` archive.
        *   `create`: Test creating a simple `.zip` archive with one file.
    *   **Complex/Variations:**
        *   `list`: Test listing contents of a `.7z` archive. Test listing contents of a non-existent archive (error handling). Test listing contents of an archive with nested directories.
        *   `extract`: Test extracting multiple files. Test extracting to a specified directory. Test extracting from a `.7z` archive. Test extracting a non-existent file (error handling).
        *   `create`: Test creating a `.7z` archive with multiple files and nested directories. Test creating an archive with a non-existent source file (error handling).

**2. Excel Tools:**

*   **excel/range:**
    *   **Simple:**
        *   `read`: Test reading a simple range (e.g., A1:B2) from a known `.xlsx` file.
        *   `write`: Test writing simple data to a range in a new `.xlsx` file.
    *   **Complex/Variations:**
        *   `read`: Test reading from different sheets. Test reading a named range. Test reading a large range. Test reading from a non-existent file/sheet/range (error handling).
        *   `write`: Test writing to an existing file/sheet/range. Test writing different data types (numbers, strings, formulas). Test writing to a non-existent sheet (should be created).
        *   `format`: Test applying basic formatting (e.g., bold, currency) to a range.
        *   `apply`: Test applying a style to a range.
*   **excel/data-analysis:**
    *   **Simple:**
        *   Test filtering data based on a simple criterion (e.g., value > X).
    *   **Complex/Variations:**
        *   Test filtering with multiple criteria. Test applying analysis operations (need to confirm available operations from API spec). Test handling errors with invalid criteria or non-existent ranges.

**3. PowerPoint Tools:**

*   **powerpoint/slides:**
    *   **Simple:**
        *   `add`: Test adding a new slide to a presentation.
        *   `delete`: Test deleting a slide.
    *   **Complex/Variations:**
        *   `add`: Test adding slides with different layouts. Test adding slides at specific positions.
        *   `delete`: Test deleting multiple slides. Test deleting a non-existent slide (error handling).
        *   `set`: Test setting properties of a slide (e.g., background).
*   **powerpoint/shapes:**
    *   **Simple:**
        *   `add`: Test adding a simple shape (e.g., rectangle) to a slide.
    *   **Complex/Variations:**
        *   `add`: Test adding different shape types. Test adding shapes with text. Test adding images as shapes. Test adding shapes at specific coordinates.
*   **powerpoint/charts:**
    *   **Simple:**
        *   `insert`: Test inserting a basic chart type (e.g., column chart) with sample data.
    *   **Complex/Variations:**
        *   `insert`: Test inserting different chart types. Test inserting charts linked to Excel data (if supported). Test customizing chart properties.

**4. Word Tools:**

*   **Review and Enhance Existing Tests:**
    *   Review existing unit, integration, and e2e tests for Word tools (`merge`, `markdown/export`, etc.).
    *   Add tests for parameter variations not currently covered (e.g., different merge options, export formats, analysis criteria).
    *   Improve content verification in e2e tests where marked with `TODO`.
*   **Add Tests for Other Word Tools:**
    *   Create unit, integration, and e2e tests for Word tools that have limited or no coverage based on the file listing (e.g., `styles`, `markdown/import`, `embedded-objects`, `image`, `generate-and-insert-text`, `page`, `analyze`, `code-format`).
    *   Cover simple cases first, then add complex scenarios and parameter variations.

**5. Cross-Application (office) Tools:**

*   **office/pdf/export:**
    *   **Simple:** Test exporting a simple Word document to PDF.
    *   **Complex/Variations:** Test exporting Excel and PowerPoint files. Test exporting with different PDF settings (if available in API).
*   **office/combine:**
    *   **Simple:** Test combining two simple documents (e.g., Word + Word).
    *   **Complex/Variations:** Test combining different file types (e.g., Word + PDF, Excel + Word - need to confirm supported combinations from API spec). Test combining multiple documents.
*   **office/transfer:**
    *   **Simple:** Test transferring a simple range from Excel to Word.
    *   **Complex/Variations:** Test transferring data with linking (`link=true`). Test transferring data between other application combinations (e.g., Word to PowerPoint). Test embedding content (`operation=embed`). Test transferring non-existent sources/targets (error handling).

**6. Static and Dynamic Resources:**

*   **memory/ai_assistant_guide:**
    *   **Simple:** Test reading the resource.
*   **dynamic/resources:**
    *   **Simple:**
        *   `list`: Test listing available dynamic resources.
        *   `read`: Test reading a known dynamic resource.
    *   **Complex/Variations:**
        *   `read`: Test reading a non-existent resource (error handling).
        *   `write`, `delete`, `metadata`, `search`: Create tests for these operations based on the API specification.

**Phase 2: Documentation Creation**

This phase focuses on creating a markdown document explaining how to request and interpret the tests.

**Document Title:** `TESTING_GUIDE.md`

**Structure:**

1.  **Introduction:**
    *   Purpose of the document.
    *   Overview of the test suite (unit, integration, e2e).
    *   **Running Tests:**
        *   **Prerequisites:** Before running any tests, ensure all project dependencies are installed. Open your terminal in the project root directory and run:
            ```bash
            npm install
            ```
        *   **Running All Tests:** To execute the complete test suite, which includes unit, integration, and end-to-end tests, use the following command:
            ```bash
            npm test
            ```
            This command is defined in the `scripts` section of your [`package.json`](../../package.json:11) file.
        *   **Running Specific Test Suites:**
            *   To run only unit tests:
                ```bash
                npm run test:unit
                ```
            *   To run only end-to-end (e2e) tests:
                ```bash
                npm run test:e2e
                ```
        *   **Running Individual Test Files or Patterns (using Jest):**
            Jest is the testing framework used in this project. You can run specific tests directly using `npx jest` followed by the path to the test file or a pattern.
            *   To run tests for a single specific file:
                ```bash
                npx jest path/to/your/test-file.test.ts
                ```
                For example:
                ```bash
                npx jest tests/unit/fs/fileContent.test.ts
                ```
            *   To run tests matching a name pattern (e.g., all tests related to 'fileContent'):
                ```bash
                npx jest fileContent
                ```
            *   To run all tests within a specific directory (e.g., all tests in `tests/unit/fs/`):
                ```bash
                npx jest tests/unit/fs/
                ```
            *   Alternatively, the project provides a script to run a single test file:
                ```bash
                npm run test:single path/to/your/test-file.test.ts
                ```
2.  **Requesting Tests via AI:**
    *   Explain how a user would ask the AI to run specific tests or test suites.
    *   Provide example AI prompts for simple and complex test requests (e.g., "Run the basic file read test", "Run all end-to-end tests for Word merge with different options").
    *   Explain how the AI interprets the request and selects the appropriate test(s).
3.  **Requesting Tests via cURL:**
    *   Explain how to trigger specific tests or tool operations directly via cURL commands to the MCP server's API.
    *   Provide example cURL commands for various tools and operations, including different parameters.
    *   Reference the `INSTALLATION_AND_API.md` for full API details.
4.  **Interpreting Test Results:**
    *   Explain the structure of the test output (console output from `npm test`).
    *   How to identify successful tests, failed tests, and errors.
    *   How to interpret the output of individual tool calls (JSON response structure - `success`, `data`, `error`).
    *   Explain common error codes and messages.
5.  **Test Case Details (Optional but Recommended):**
    *   Briefly describe each major test case or category.
    *   Mention the tool(s) and operations being tested.
    *   Describe the scenario (e.g., "Testing Word merge with two documents").
    *   (Could link to the actual test files for detailed implementation).
6.  **Contributing to Tests:**
    *   Brief guidelines on how to add new tests to the suite.

**Leveraging Existing Tests:**

*   The structure and patterns used in the existing `tests/` directory (unit, integration, e2e) will be followed for new tests.
*   Existing test files will be used as templates for creating new tests for similar tools.
*   Parameter variations and more complex scenarios will be added to existing test files where appropriate.

**Mermaid Diagrams:**

Mermaid diagrams can be included in the `TESTING_GUIDE.md` to illustrate:
*   The overall test architecture (unit -> integration -> e2e).
*   The flow of a test request (AI/curl -> Server -> Tool -> Office App).
*   Examples of tool workflows (similar to those in `INSTALLATION_AND_API.md` but focused on the testing perspective).