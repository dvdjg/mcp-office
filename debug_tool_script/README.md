# Menu-Driven Debugging Script

This script provides a menu-driven interface to help select and prepare commands for debugging a suite of TypeScript-based tools. It guides the user through selecting a tool category, then a specific tool, fulfilling any preconditions, and providing necessary parameters. It supports suggesting fixture files for relevant parameters. Finally, it displays the `ts-node` command that would be used to execute the tool and offers an option to run that command directly.

## Features

*   **Interactive Menu:** Easy-to-use command-line interface powered by `inquirer`.
*   **Tool Categorization:** Tools are organized by category for easier navigation.
*   **Precondition Handling:** Displays any preconditions for a selected tool and requires user acknowledgment.
*   **Parameter Input:** Prompts for all necessary parameters for the selected tool.
    *   **Fixture Support:** For parameters representing file paths that are relevant to fixtures, the script can list files from the `tests/fixtures/` directory (relative to the main project root) for easy selection.
*   **Command Display:** Shows the fully constructed `ts-node` command.
*   **Optional Command Execution:** After displaying the command, the script asks the user if they want to execute it. If confirmed, the script runs the command and streams its `stdout` and `stderr`.

## File Structure

Located within the `debug_tool_script/` directory:

*   `debug-tool-menu.ts`: The main entry point for the script.
*   `toolRegistry.ts`: Contains the definitions of all available tools, their parameters, and preconditions.
*   `uiHandler.ts`: Manages all user interface interactions (prompts, confirmations).
*   `commandBuilder.ts`: Responsible for constructing the `ts-node` command string.
*   `fileSystemHelper.ts`: Utility functions, primarily for listing fixture files.
*   `package.json`: Defines script dependencies (e.g., `inquirer`) and a start script.
*   `README.md`: This documentation file.

## Prerequisites

*   Node.js and npm (or a compatible package manager like yarn).
*   The main project's tools (expected to be `.tool.ts` files) should be accessible via the paths defined in `toolRegistry.ts`.

## Setup

1.  Navigate to this script's directory:
    ```bash
    cd path/to/your/project/debug_tool_script
    ```
2.  Install dependencies:
    ```bash
    npm install
    ```

## Running the Script

From within the `debug_tool_script/` directory:

```bash
npm start
```

Alternatively, you can run it directly with `ts-node`:

```bash
npx ts-node debug-tool-menu.ts
```

The script will then guide you through the menus.

## Tool Definitions (`toolRegistry.ts`)

The `toolRegistry.ts` file exports an array named `TOOLS_DATA`. Each object in this array represents a `ToolDefinition` and includes:
*   `id`: A unique identifier (e.g., "word.export-word-to-markdown").
*   `name`: A user-friendly name (e.g., "Export Word to Markdown").
*   `category`: The category the tool belongs to (e.g., "Word", "Excel", "FileSystem").
*   `scriptPath`: The relative path from the main project root to the tool's `.tool.ts` file (e.g., `"src/tools/word/markdown.tool.ts"`).
*   `preconditions`: An optional array of strings, where each string is a precondition message to be displayed to the user.
*   `parameters`: An array of `ToolParameter` objects, defining each parameter the tool accepts. Each parameter includes:
    *   `name`: The command-line argument name (e.g., "filePath").
    *   `type`: The expected data type (`'string'`, `'boolean'`, `'filePath'`, `'number'`, `'enum'`).
    *   `description`: A user-friendly description.
    *   `required`: A boolean indicating if the parameter is mandatory.
    *   `defaultValue`: An optional default value.
    *   `isFixtureRelevant`: A boolean indicating if the parameter (typically of type `'filePath'`) can use files from the `tests/fixtures/` directory.
    *   `enumValues`: An optional array of strings for parameters of type `'enum'`.

The tool data is currently hardcoded but derived from an analysis of available `.tool.ts` files.