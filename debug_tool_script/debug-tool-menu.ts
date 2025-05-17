import * as toolRegistry from './toolRegistry';
import * as uiHandler from './uiHandler';
import * as commandBuilder from './commandBuilder';
import { ToolDefinition } from './toolRegistry';
import { exec } from 'child_process';
import * as path from 'path'; // Import path

async function main() {
  console.log("Starting Debug Tool Menu Script...");
  const toolsData = toolRegistry.loadTools();
  const projectRoot = path.resolve(__dirname, '..'); // Determine project root

  let keepRunning = true;
  while (keepRunning) {
    const categories = toolRegistry.getUniqueCategories(toolsData);
    if (categories.length === 0) {
        console.log("No tools configured. Exiting.");
        keepRunning = false;
        break;
    }

    const selectedCategory = await uiHandler.promptSelectCategory(categories);
    if (selectedCategory === null) { // User chose to exit
      keepRunning = false;
      break;
    }

    const toolsInCategory = toolRegistry.filterToolsByCategory(toolsData, selectedCategory);
    if (toolsInCategory.length === 0) {
        console.log(`No tools found in category: ${selectedCategory}`);
        continue;
    }

    const selectedTool: ToolDefinition | null = await uiHandler.promptSelectTool(toolsInCategory);
    if (selectedTool === null) { // User chose to go back or cancelled
      continue;
    }

    // Handle Preconditions
    if (selectedTool.preconditions && selectedTool.preconditions.length > 0) {
      console.log(`\n--- Preconditions for ${selectedTool.name} ---`);
      for (const precondition of selectedTool.preconditions) {
        await uiHandler.displayPrecondition(precondition);
      }
      console.log("--- End of Preconditions ---");
    }

    // Prompt for Parameters
    const userParameters = await uiHandler.promptForParameters(selectedTool);

    // Construct Command
    const commandToRun = commandBuilder.buildCommand(selectedTool, userParameters);
    uiHandler.displayCommand(commandToRun);

    // Ask if user wants to execute the command
    const execute = await uiHandler.confirm("Do you want to execute this command?");
    if (execute) {
      console.log(`\nAttempting to execute tool from project root: ${projectRoot}`);
      console.log(`Executing command: ${commandToRun}`);
      
      const child = exec(commandToRun, { cwd: projectRoot }); // Set CWD to project root

      child.stdout?.on('data', (data) => {
        process.stdout.write(data);
      });

      child.stderr?.on('data', (data) => {
        process.stderr.write(data);
      });

      // It's important to handle the 'exit' event as well for a complete picture
      child.on('exit', (code, signal) => {
        if (code !== null) {
          console.log(`\nTool execution process exited with code ${code}.`);
        } else if (signal !== null) {
          console.log(`\nTool execution process was killed with signal ${signal}.`);
        }
        // The 'close' event might still be useful for when stdio streams are closed.
      });
      
      child.on('close', (code) => {
        // This message might be redundant if 'exit' is handled well, but can be kept for clarity.
        // console.log(`\nTool execution finished (stdio streams closed) with code ${code}. Check terminal for output/errors.`);
        askToContinue();
      });

      child.on('error', (err) => {
        console.error(`\nFailed to start or run tool: ${err.message}`);
        askToContinue();
      });
    } else {
      askToContinue();
    }

    async function askToContinue() {
        keepRunning = await uiHandler.confirm("Do you want to select another tool?");
        if (!keepRunning) {
            console.log("Exiting Debug Tool Menu Script.");
        }
    }
  }
}

main().catch(error => {
  console.error("An unexpected error occurred in main:", error);
  if (process.exitCode === undefined) {
    process.exitCode = 1;
  }
});