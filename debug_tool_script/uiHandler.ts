import inquirer from 'inquirer';
import { ToolDefinition, ToolParameter } from './toolRegistry';
import * as fileSystemHelper from './fileSystemHelper';
import * as pathModule from 'path'; // Import path module

const FIXTURE_PATH = 'tests/fixtures/';

export async function promptSelectCategory(categories: string[]): Promise<string | null> {
  if (categories.length === 0) {
    console.log("No tool categories found.");
    return null;
  }
  const answers = await inquirer.prompt([
    {
      type: 'list',
      name: 'category',
      message: 'Select a tool category:',
      choices: [...categories, new inquirer.Separator(), { name: 'Exit', value: null }],
    },
  ]);
  return answers.category;
}

export async function promptSelectTool(tools: ToolDefinition[]): Promise<ToolDefinition | null> {
  if (tools.length === 0) {
    console.log("No tools found in this category.");
    return null;
  }
  const choices = tools.map(tool => ({ name: `${tool.name} (${tool.id})`, value: tool.id }));
  const answers = await inquirer.prompt([
    {
      type: 'list',
      name: 'toolId',
      message: 'Select a tool:',
      choices: [...choices, new inquirer.Separator(), { name: 'Back to categories', value: null }],
    },
  ]);
  if (!answers.toolId) return null;
  return tools.find(tool => tool.id === answers.toolId) || null;
}

export async function displayPrecondition(message: string): Promise<void> {
  console.log(`\n⚠️ PRECONDITION: ${message}`);
  await inquirer.prompt([{ type: 'input', name: 'ack', message: "Press Enter to acknowledge and continue..." }]);
}

async function promptForFixtureRelevantPath(paramDef: ToolParameter): Promise<string | undefined> {
  // Initialize as a mutable array
  const choices: Array<inquirer.DistinctChoice<inquirer.Answers, inquirer.ListChoiceMap<inquirer.Answers>>> = [
    { name: `List relevant files from ${FIXTURE_PATH}`, value: 'list_fixtures' },
    { name: 'Provide a custom path', value: 'custom_path' },
  ];

  if (paramDef.defaultValue) {
    choices.push({ name: `Use default: ${paramDef.defaultValue}`, value: 'default_value' });
  }
  if (!paramDef.required) {
    choices.push({ name: 'Skip this parameter', value: 'skip_parameter' });
  }

  const { action } = await inquirer.prompt<{ action: string }>([{
    type: 'list',
    name: 'action',
    message: `For parameter "${paramDef.name}" (${paramDef.description}):`,
    choices: choices,
  }]);

  if (action === 'list_fixtures') {
    const fixtureFiles = await fileSystemHelper.listFixtureFiles(FIXTURE_PATH);
    if (fixtureFiles.length === 0) {
      console.log(`No files found in ${FIXTURE_PATH}. Please provide a custom path or skip if optional.`);
      const nestedChoicesList: Array<inquirer.DistinctChoice<inquirer.Answers, inquirer.ListChoiceMap<inquirer.Answers>>> = [{ name: 'Provide a custom path', value: 'custom_path_nested' }];
      if (!paramDef.required) {
        nestedChoicesList.push({ name: 'Skip this parameter', value: 'skip_parameter_nested' });
      }
      const { nestedAction } = await inquirer.prompt<{ nestedAction: string }>([{
        type: 'list',
        name: 'nestedAction',
        message: 'No fixtures found. Choose an option:',
        choices: nestedChoicesList
      }]);

      if (nestedAction === 'custom_path_nested') {
        const { customPath } = await inquirer.prompt<{ customPath: string }>([{ type: 'input', name: 'customPath', message: `Enter custom path for ${paramDef.name}:`, validate: (input) => (paramDef.required && !input.trim() ? 'This field is required.' : true) }]);
        return customPath;
      }
      return undefined; // Skip if chosen
    }

    const { selectedPath } = await inquirer.prompt<{ selectedPath: string | undefined }>([{
      type: 'list',
      name: 'selectedPath',
      message: 'Select a fixture file (or skip if optional):',
      choices: [...fixtureFiles.map(f => ({ name: f, value: pathModule.join(FIXTURE_PATH, f) })), ...(paramDef.required ? [] : [{ name: '(Skip)', value: undefined }])],
    }]);
    return selectedPath;
  } else if (action === 'default_value') {
    return paramDef.defaultValue;
  } else if (action === 'skip_parameter') {
    return undefined;
  } else { // custom_path
    const { customPath } = await inquirer.prompt<{ customPath: string }>([{
      type: 'input',
      name: 'customPath',
      message: `Enter custom path for ${paramDef.name}:`,
      validate: (input) => {
        if (paramDef.required && (input === null || String(input).trim() === '')) {
          return 'This field is required.';
        }
        return true;
      }
    }]);
    return customPath;
  }
}

export async function promptForParameters(tool: ToolDefinition): Promise<Record<string, any>> {
  const params: Record<string, any> = {};
  console.log(`\nConfiguring parameters for: ${tool.name}`);

  if (!tool.parameters || tool.parameters.length === 0) {
    console.log("This tool has no configurable parameters.");
    return params;
  }

  for (const paramDef of tool.parameters) {
    let value: any;
    if (paramDef.type === 'filePath' && paramDef.isFixtureRelevant) {
      value = await promptForFixtureRelevantPath(paramDef);
    } else if (paramDef.type === 'boolean') {
      const { answer } = await inquirer.prompt<{ answer: boolean }>([{
        type: 'confirm',
        name: 'answer',
        message: `${paramDef.description} (${paramDef.name}):`,
        default: paramDef.defaultValue !== undefined ? paramDef.defaultValue : false,
      }]);
      value = answer;
    } else if (paramDef.type === 'enum' && paramDef.enumValues && paramDef.enumValues.length > 0) {
      const { answer } = await inquirer.prompt<{ answer: string }> ([{
        type: 'list',
        name: 'answer',
        message: `${paramDef.description} (${paramDef.name}):`,
        choices: paramDef.enumValues,
        default: paramDef.defaultValue,
      }]);
      value = answer;
    } else { // string, number, or non-fixture filePath
      const { answer } = await inquirer.prompt<{ answer: string | number }> ([{
        type: paramDef.type === 'number' ? 'number' : 'input',
        name: 'answer',
        message: `${paramDef.description} (${paramDef.name}):`,
        default: paramDef.defaultValue,
        validate: (input: string | number) => {
          if (paramDef.required && (input === null || input === undefined || String(input).trim() === '')) {
            return 'This field is required.';
          }
          if (paramDef.type === 'number' && typeof input === 'string' && isNaN(parseFloat(input))) {
            return 'Please enter a valid number.';
          }
          return true;
        }
      }]);
      value = answer;
    }
    if (value !== undefined) { 
        params[paramDef.name] = value;
    } else if (paramDef.required) {
        console.warn(`Warning: Required parameter "${paramDef.name}" was not provided and has no default.`);
    }
  }
  return params;
}


export function displayCommand(command: string): void {
  console.log("\n--------------------------------------------------");
  console.log("🚀 Command to execute (copy and paste to run):");
  console.log("--------------------------------------------------");
  console.log(command);
  console.log("--------------------------------------------------\n");
}

export async function confirm(message: string): Promise<boolean> {
  const { confirmed } = await inquirer.prompt<{ confirmed: boolean }>([
    {
      type: 'confirm',
      name: 'confirmed',
      message: message,
      default: true,
    },
  ]);
  return confirmed;
}