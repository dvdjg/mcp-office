# Office MCP Server API Test Plan

**Document Purpose**: This test plan outlines a comprehensive, automated testing strategy for the Office MCP Server API, built with FastMCP in TypeScript for automating Microsoft Office (Word, Excel, PowerPoint). It includes unit, integration, end-to-end, performance, security, and documentation tests, ensuring >90% coverage for 10 Word-specific use cases and additional Excel/PowerPoint scenarios. The plan is designed for automation with tools like Jest and Postman, while also providing detailed instructions for manual reproduction by a human tester using an AI (e.g., Claude). Each test case is documented with API calls, expected outcomes, and steps for manual execution.

**Target Audience**: Developers automating tests and human testers manually verifying API functionality with AI assistance.

**Date**: May 01, 2025

---

## Table of Contents
1. [Overview](#overview)
   - [Objectives](#objectives)
   - [Scope](#scope)
   - [Test Types](#test-types)
2. [Test Environment](#test-environment)
   - [Setup](#setup)
   - [Tools](#tools)
   - [Sample Documents](#sample-documents)
3. [Test Strategy](#test-strategy)
   - [Automation Approach](#automation-approach)
   - [Manual Reproduction](#manual-reproduction)
4. [Test Cases](#test-cases)
   - [Unit Tests](#unit-tests)
   - [Integration Tests](#integration-tests)
   - [End-to-End Tests](#end-to-end-tests)
   - [Performance Tests](#performance-tests)
   - [Security Tests](#security-tests)
   - [Documentation Tests](#documentation-tests)
5. [Execution and Reporting](#execution-and-reporting)
   - [Automated Execution](#automated-execution)
   - [Manual Execution](#manual-execution)
   - [Reporting](#reporting)
6. [Risks and Mitigation](#risks-and-mitigation)
7. [Appendices](#appendices)
   - [Sample API Calls](#sample-api-calls)
   - [Manual Test Script Template](#manual-test-script-template)

---

## Overview

### Objectives
- Ensure the Office MCP Server API functions correctly for all tools (Word, Excel, PowerPoint, file system, cross-application, resource management).
- Achieve >90% code coverage for unit and integration tests.
- Validate all 10 Word-specific use cases and additional Excel/PowerPoint scenarios.
- Confirm performance (e.g., merge 100-page documents in <10s).
- Verify security (e.g., input validation, access controls).
- Ensure documentation meets standards (<200 words per tool).
- Enable manual reproduction by human testers with AI assistance.

### Scope
- **In Scope**:
  - All API endpoints (`/word/*`, `/excel/*`, `/powerpoint/*`, `/office/*`, `/fs/*`, `/memory/*`, `/dynamic/*`).
  - 10 Word use cases (e.g., convert to Markdown, merge documents).
  - Excel scenarios (e.g., table creation, chart insertion).
  - PowerPoint scenarios (e.g., slide addition, animation).
  - Automated tests with Jest, Mocha, and Postman.
  - Manual test instructions for human testers.
- **Out of Scope**:
  - UI testing (Office MCP is API-only).
  - Third-party service integrations (e.g., Power Automate).

### Test Types
- **Unit Tests**: Test individual tool operations (e.g., `word/styles/apply`).
- **Integration Tests**: Test tool interactions (e.g., `word/merge` + `word/styles`).
- **End-to-End Tests**: Test full workflows (e.g., convert Word to Markdown).
- **Performance Tests**: Measure operation speed (e.g., merge large documents).
- **Security Tests**: Validate input sanitization and access controls.
- **Documentation Tests**: Verify JSDoc coverage and length.

---

## Test Environment

### Setup
- **Hardware**: Standard laptop/server (8GB RAM, 4-core CPU).
- **OS**: Windows 10/11 (for VBA/COM Interop) or cross-platform (Office 365).
- **Software**:
  - Node.js v18+.
  - Microsoft Office 365 or 2016+ (Word, Excel, PowerPoint).
  - Office MCP Server running on `localhost:3000`.
- **Dependencies**: Install via `npm install` (see `package.json`).
  ```bash
  npm install jest mocha postman-newman fastmcp @microsoft/office-js zod winston node-cache fs-extra markdown-it mermaid pdf-parse pdf2pic pdfkit highlight.js
  ```
- **Directory Structure**:
  ```
  mcp-office/
  ├── src/               # Source code
  ├── tests/             # Test scripts
  │   ├── unit/         # Unit tests
  │   ├── integration/  # Integration tests
  │   ├── e2e/          # End-to-end tests
  │   ├── performance/  # Performance tests
  │   ├── security/     # Security tests
  │   ├── docs/         # Documentation tests
  │   └── fixtures/     # Sample documents
  ├── package.json
  └── tsconfig.json
  ```

### Tools
- **Jest**: Unit and integration testing (`npm test`).
- **Mocha**: End-to-end testing (`npm run test:e2e`).
- **Postman/Newman**: API testing and automation.
- **Winston**: Logging for debugging.
- **Node-Cache**: Mocking resource states.
- **Curl**: Manual API testing.
- **AI (e.g., Claude)**: Assists human testers with API calls and validation.

### Sample Documents
Sample Office files are stored in `tests/fixtures/` (see [Sample Office Documents](#sample-office-documents) below for details):
- **Word**: `CV.docx`, `cuentoAladdinDraft1.docx`, `cuentoAladdinDraft2.docx`, `PresupuestosEvolutio.docx`, `PlantillaEvolutio.docx`, `PropuestasRandom.docx`, `Composición.docx`, `PropuestaTécnica.docx`, `BuenasPrácticas.docx`, `LargeDoc.docx`.
- **Excel**: `SalesData.xlsx`, `Budget.xlsx`.
- **PowerPoint**: `Presentation.pptx`, `Tutorial.pptx`.
- **Markdown**: `input.md`.
- **PDF**: `Offer.pdf`.

---

## Test Strategy

### Automation Approach
- **Unit Tests**: Use Jest to mock dependencies (e.g., `@microsoft/office-js`, `fs-extra`) and test tool operations in isolation. Example: Test `word/styles/apply` with mocked `Word.run`.
- **Integration Tests**: Use Jest to test tool combinations (e.g., `word/merge` + `word/styles`). Mock Office APIs minimally to simulate real interactions.
- **End-to-End Tests**: Use Mocha to test full workflows with real Office files. Run Office MCP Server and Office locally.
- **Performance Tests**: Use Postman/Newman to measure operation times (e.g., merge large documents).
- **Security Tests**: Use Postman to test invalid inputs and unauthorized access.
- **Documentation Tests**: Use Jest to parse JSDoc and verify coverage/length.
- **CI/CD**: Run tests on GitHub Actions or Jenkins with `npm test`.

**Automation Scripts**:
```json
{
  "scripts": {
    "test": "jest",
    "test:coverage": "jest --coverage",
    "test:e2e": "mocha tests/e2e",
    "test:performance": "newman run tests/performance/collection.json",
    "test:security": "newman run tests/security/collection.json",
    "test:docs": "jest tests/docs"
  }
}
```

### Manual Reproduction
- **Human Tester Role**: Execute API calls using `curl`, Postman, or an AI (e.g., Claude) to validate responses.
- **AI Role**: Assist with generating API calls, interpreting responses, and suggesting fixes based on `memory://ai_assistant_guide`.
- **Steps**:
  1. **Setup**: Ensure Office MCP Server is running (`npm start`) and Office is installed.
  2. **Prepare Files**: Place sample documents in `tests/fixtures/`.
  3. **Execute Tests**: Follow test case instructions (e.g., send `curl` commands).
  4. **Validate Results**: Check API responses and output files (e.g., `merged.docx`).
  5. **Log Issues**: Note failures in a test log (see [Manual Test Script Template](#manual-test-script-template)).
- **AI Interaction**: Phrase requests clearly (e.g., “Send a POST request to merge doc1.docx and doc2.docx”). Use completions (`GET /word/merge/completions`) for valid parameters.

---

## Test Cases

### Unit Tests
**Purpose**: Test individual tool operations in isolation.

**Examples**:
1. **word/styles/apply**:
   - **Test**: Apply `Heading1` to `paragraph:1` in `CV.docx`.
   - **API Call**: `POST /word/styles/apply?document=/tests/fixtures/CV.docx&style=Heading1&range=paragraph:1`
   - **Expected**: `{ success: true }`
   - **Manual Steps**:
     1. Open `CV.docx` in Word.
     2. Run: `curl -X POST "http://localhost:3000/word/styles/apply?document=/tests/fixtures/CV.docx&style=Heading1&range=paragraph:1"`
     3. Verify first paragraph has `Heading1` style.
     4. Ask AI: “Check if the first paragraph in CV.docx has Heading1 style.”
   - **Automation**: `tests/unit/word/styles.test.ts`
     ```typescript
     import { wordStyles } from '../../src/tools/word/styles';
     jest.mock('@microsoft/office-js');
     test('apply style', async () => {
       const result = await wordStyles.handler({ document: 'CV.docx', style: 'Heading1', range: 'paragraph:1' });
       expect(result.success).toBe(true);
     });
     ```
2. **fs/file/read**:
   - **Test**: Read `input.md`.
   - **API Call**: `GET /fs/file/read?path=/tests/fixtures/input.md&format=text`
   - **Expected**: `{ content: "# Sample\nText" }`
   - **Manual Steps**:
     1. Run: `curl -X GET "http://localhost:3000/fs/file/read?path=/tests/fixtures/input.md&format=text"`
     2. Verify response matches file content.
     3. Ask AI: “Read input.md and confirm it starts with a heading.”

**Total**: ~100 tests (5-10 per tool).

### Integration Tests
**Purpose**: Test tool interactions.

**Examples**:
1. **word/merge + word/styles**:
   - **Test**: Merge `cuentoAladdinDraft1.docx` and `cuentoAladdinDraft2.docx`, then apply `Normal` style.
   - **API Calls**:
     - `POST /word/merge?docs=/tests/fixtures/cuentoAladdinDraft1.docx,/tests/fixtures/cuentoAladdinDraft2.docx&output=/tests/fixtures/merged.docx`
     - `POST /word/styles/apply?document=/tests/fixtures/merged.docx&style=Normal&range=paragraph:1`
   - **Expected**: `{ success: true }` for both; `merged.docx` has combined content with `Normal` style.
   - **Manual Steps**:
     1. Run merge: `curl -X POST "http://localhost:3000/word/merge?docs=/tests/fixtures/cuentoAladdinDraft1.docx,/tests/fixtures/cuentoAladdinDraft2.docx&output=/tests/fixtures/merged.docx"`
     2. Run style: `curl -X POST "http://localhost:3000/word/styles/apply?document=/tests/fixtures/merged.docx&style=Normal&range=paragraph:1"`
     3. Open `merged.docx` and verify content and style.
     4. Ask AI: “Merge two drafts and apply Normal style to the first paragraph.”
   - **Automation**: `tests/integration/word/merge-styles.test.ts`
2. **excel/tables + excel/charts**:
   - **Test**: Create a table in `SalesData.xlsx` and insert a chart.
   - **API Calls**:
     - `POST /excel/tables/insert?document=/tests/fixtures/SalesData.xlsx&range=A1:D10&name=Sales`
     - `POST /excel/charts/insert?document=/tests/fixtures/SalesData.xlsx&type=column&range=A1:D10`
   - **Expected**: `{ success: true }`; table and chart in `SalesData.xlsx`.

**Total**: ~50 tests (2-5 per tool combination).

### End-to-End Tests
**Purpose**: Test full workflows for Office MCP use cases.

**Examples** (covering 10 Word use cases + 2 Excel + 2 PowerPoint):
1. **Use Case 1: Convert Word to Markdown with Comments**:
   - **Test**: Convert `CV.docx` to `CV.md` with comments appended.
   - **API Call**: `POST /word/markdown/export?document=/tests/fixtures/CV.docx&output=/tests/fixtures/CV.md&comments=append`
   - **Expected**: `{ success: true, output: "/tests/fixtures/CV.md" }`; `CV.md` has comments in `## Comments`.
   - **Manual Steps**:
     1. Run: `curl -X POST "http://localhost:3000/word/markdown/export?document=/tests/fixtures/CV.docx&output=/tests/fixtures/CV.md&comments=append"`
     2. Open `CV.md` and verify comments section.
     3. Ask AI: “Convert CV.docx to Markdown and check for comments.”
   - **Automation**: `tests/e2e/word/markdown-export.test.js`
2. **Excel: Create Table and Chart**:
   - **Test**: Insert a table and chart in `Budget.xlsx`.
   - **API Call**: `POST /office/workflow/run` with `steps=[{tool:"excel/tables/insert",params:{range:"A1:D10",name:"Budget"}},{tool:"excel/charts/insert",params:{type:"column",range:"A1:D10"}}]`
   - **Expected**: `{ success: true }`; table and chart in `Budget.xlsx`.
   - **Manual Steps**:
     1. Run: `curl -X POST "http://localhost:3000/office/workflow/run" -d '{"steps":[{"tool":"excel/tables/insert","params":{"document":"/tests/fixtures/Budget.xlsx","range":"A1:D10","name":"Budget"}},{"tool":"excel/charts/insert","params":{"document":"/tests/fixtures/Budget.xlsx","type":"column","range":"A1:D10"}}]}'`
     2. Open `Budget.xlsx` and verify table and chart.
     3. Ask AI: “Create a table and column chart in Budget.xlsx.”
3. **PowerPoint: Add Slide and Animation**:
   - **Test**: Add a title slide with a fade transition to `Presentation.pptx`.
   - **API Call**: `POST /office/workflow/run` with `steps=[{tool:"powerpoint/slides/add",params:{layout:"TitleSlide"}},{tool:"powerpoint/animations/add",params:{slide:1,type:"fade"}}]`
   - **Expected**: `{ success: true }`; new slide with fade transition.

**Total**: 14 tests (10 Word + 2 Excel + 2 PowerPoint).

### Performance Tests
**Purpose**: Measure operation speed for large inputs.

**Examples**:
1. **Merge Large Documents**:
   - **Test**: Merge `LargeDoc.docx` (100 pages) with itself in <10s.
   - **API Call**: `POST /word/merge?docs=/tests/fixtures/LargeDoc.docx,/tests/fixtures/LargeDoc.docx&output=/tests/fixtures/merged_large.docx`
   - **Expected**: `{ success: true }` in <10s.
   - **Manual Steps**:
     1. Time the request: `time curl -X POST "http://localhost:3000/word/merge?docs=/tests/fixtures/LargeDoc.docx,/tests/fixtures/LargeDoc.docx&output=/tests/fixtures/merged_large.docx"`
     2. Verify time <10s and output file exists.
     3. Ask AI: “Merge two large documents and check if it takes less than 10 seconds.”
   - **Automation**: `tests/performance/merge.json` (Postman collection).
2. **Export Large Document to PDF**:
   - **Test**: Export `LargeDoc.docx` to PDF in <5s.
   - **API Call**: `POST /office/pdf/export?document=/tests/fixtures/LargeDoc.docx&output=/tests/fixtures/large.pdf`

**Total**: 5 tests.

### Security Tests
**Purpose**: Validate input sanitization and access controls.

**Examples**:
1. **Invalid File Path**:
   - **Test**: Attempt directory traversal (`../../etc/passwd`).
   - **API Call**: `GET /fs/file/read?path=../../etc/passwd`
   - **Expected**: `{ success: false, error: { code: 400, message: "Invalid path" } }`
   - **Manual Steps**:
     1. Run: `curl -X GET "http://localhost:3000/fs/file/read?path=../../etc/passwd"`
     2. Verify error response.
     3. Ask AI: “Try reading an invalid file path and check for an error.”
   - **Automation**: `tests/security/fs.json`
2. **Invalid Style**:
   - **Test**: Apply non-existent style.
   - **API Call**: `POST /word/styles/apply?document=/tests/fixtures/CV.docx&style=InvalidStyle&range=paragraph:1`
   - **Expected**: `{ success: false, error: { code: 400, message: "Style not found" } }`

**Total**: 10 tests.

### Documentation Tests
**Purpose**: Verify JSDoc coverage and length.

**Examples**:
1. **JSDoc Coverage**:
   - **Test**: Ensure all tools have JSDoc.
   - **Expected**: >90% coverage.
   - **Manual Steps**:
     1. Run: `npm run test:docs`
     2. Check coverage report (`coverage/docs`).
     3. Ask AI: “Verify all tools have documentation.”
   - **Automation**: `tests/docs/coverage.test.ts`
2. **JSDoc Length**:
   - **Test**: Ensure JSDoc <200 words per tool.
   - **Expected**: No violations.
   - **Automation**: `tests/docs/length.test.ts`

**Total**: 5 tests.

---

## Execution and Reporting

### Automated Execution
- **Run All Tests**: `npm test`
- **Specific Tests**:
  - Unit/Integration: `npm test`
  - End-to-End: `npm run test:e2e`
  - Performance/Security: `npm run test:performance`, `npm run test:security`
  - Documentation: `npm run test:docs`
- **CI/CD**: Configure GitHub Actions:
  ```yaml
  name: CI
  on: [push]
  jobs:
    test:
      runs-on: ubuntu-latest
      steps:
        - uses: actions/checkout@v3
        - uses: actions/setup-node@v3
          with: { node-version: '18' }
        - run: npm install
        - run: npm test
  ```

### Manual Execution
- **Prerequisites**: Office MCP Server running, Office installed, sample documents in `tests/fixtures/`.
- **Steps**:
  1. Open terminal or Postman.
  2. For each test case, copy the `curl` command from the test case description.
  3. Run the command and capture the response.
  4. Validate the response against the expected outcome (e.g., `{ success: true }`).
  5. For file outputs, open the file (e.g., `merged.docx`) in Office and verify content.
  6. Log results in a test script (see [Manual Test Script Template](#manual-test-script-template)).
  7. Use AI to assist: “Run this curl command and tell me if it worked: [command].”
- **AI Assistance**:
  - Query completions: `GET /word/styles/completions` for valid parameters.
  - Check responses: “Is this response correct? `{ success: true }`”
  - Debug failures: “Why did this API call fail? `{ success: false, error: ... }`”

### Reporting
- **Automated**:
  - Jest/Mocha: Generate coverage reports (`coverage/`).
  - Newman: Export performance/security results (`tests/results/`).
  - Example: `jest --coverage && newman run tests/performance/collection.json -r cli,html`
- **Manual**:
  - Use a test log spreadsheet or document.
  - Columns: Test ID, Description, API Call, Expected Result, Actual Result, Pass/Fail, Notes.
  - Example:
    ```
    | ID | Description | API Call | Expected | Actual | Pass/Fail | Notes |
    |----|-------------|----------|----------|--------|-----------|-------|
    | 1  | Convert to Markdown | POST /word/markdown/export | { success: true } | { success: true } | Pass | CV.md created |
    ```

---

## Risks and Mitigation
- **Risk**: Sample documents missing or incorrect.
  - **Mitigation**: Provide detailed specs (see [Sample Office Documents](#sample-office-documents)). Validate files before testing.
- **Risk**: Office not installed or configured.
  - **Mitigation**: Include setup instructions. Use Office 365 for cross-platform testing.
- **Risk**: AI misinterprets manual test instructions.
  - **Mitigation**: Use clear, action-oriented phrases (e.g., “Run this command”). Provide completions.
- **Risk**: Performance tests vary by hardware.
  - **Mitigation**: Run on standardized hardware. Average multiple runs.

---

## Appendices

### Sample API Calls
- **List Styles**: `curl -X GET "http://localhost:3000/word/styles/list?document=/tests/fixtures/CV.docx"`
- **Merge Documents**: `curl -X POST "http://localhost:3000/word/merge?docs=/tests/fixtures/cuentoAladdinDraft1.docx,/tests/fixtures/cuentoAladdinDraft2.docx&output=/tests/fixtures/merged.docx"`
- **Create Table**: `curl -X POST "http://localhost:3000/excel/tables/insert?document=/tests/fixtures/SalesData.xlsx&range=A1:D10&name=Sales"`

### Manual Test Script Template
```
Test ID: [e.g., E2E-01]
Description: [e.g., Convert Word to Markdown with Comments]
API Call: [e.g., POST /word/markdown/export?document=/tests/fixtures/CV.docx&output=/tests/fixtures/CV.md&comments=append]
Expected Result: [e.g., { success: true, output: "/tests/fixtures/CV.md" }; CV.md has comments]
Steps:
1. Run: [curl command]
2. Check response: [expected JSON]
3. Open output file: [e.g., CV.md]
4. Verify: [e.g., comments in ## Comments section]
5. Ask AI: [e.g., "Convert CV.docx to Markdown and check for comments"]
Actual Result: [e.g., { success: true }; comments present]
Pass/Fail: [Pass/Fail]
Notes: [e.g., File created successfully]
```

---

This test plan ensures robust, automated testing of the Office MCP Server API while enabling manual reproduction by human testers with AI assistance. Execute tests as described and report results for review.