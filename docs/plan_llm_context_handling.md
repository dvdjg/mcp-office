# Plan: Enhance LLM Context Handling and Document API Key Configuration

**Objective:** Modify Office MCP tools that rely on LLM sampling to gracefully handle cases where a dedicated LLM API key is not configured for the server, and document the API key configuration process.

**Steps:**

1.  **Identify LLM-Dependent Tools:** Confirm which tools use `context.session.requestSampling`. (Based on previous search, this includes `generateAndInsertText.tool.ts` and `aiSuggest.tool.ts`).
2.  **Analyze Current Error Handling:** Review how these tools currently handle the absence of `context.session` or `requestSampling`. (They currently throw an error).
3.  **Design Fallback/Error Strategy:** Determine the desired behavior when `context.session.requestSampling` is unavailable (due to missing API key configuration).
    *   **Option A (Preferred):** Implement a mechanism where the tool attempts to use the context of the AI running the tool as a fallback for sampling. This would require investigating if the FastMCP framework supports such a fallback or if modifications are needed.
    *   **Option B:** If Option A is not immediately feasible, improve the error handling to provide a clear, user-friendly message indicating that an LLM API key is required for this tool and directing the user to the relevant documentation for configuration.
4.  **Outline Code Modifications:** Detail the specific code changes needed in the identified tool files (`generateAndInsertText.tool.ts`, `aiSuggest.tool.ts`) to implement the chosen strategy (Option A or B). This would involve adding checks and alternative logic.
5.  **Update Documentation:**
    *   Add a clear section in `INSTALLATION_AND_API.md` explaining how to configure the LLM API key for the Office MCP server.
    *   Update the documentation for the LLM-dependent tools (e.g., in `USE_CASES.md` or potentially a new `AI_INTEGRATION.md`) to explain their reliance on LLM context and how the fallback/error handling works (based on the chosen strategy).
6.  **Present Plan to User:** Share this plan for your review and approval.
7.  **Implement Changes (Requires Mode Switch):** Once the plan is approved, switch to 'code' mode to perform the code modifications and documentation updates.