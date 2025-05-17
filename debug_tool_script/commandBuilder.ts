import { ToolDefinition, ToolParameter } from './toolRegistry';
import * as path from 'path';

export function buildCommand(tool: ToolDefinition, params: Record<string, any>): string {
  const projectRoot = path.resolve(__dirname, '..'); // __dirname is debug_tool_script/
  const absoluteScriptPath = path.resolve(projectRoot, tool.scriptPath); // tool.scriptPath is like src/tools/...

  const tsconfigPath = path.join(projectRoot, 'tsconfig.json');
  let command = `npx cross-env TS_NODE_PROJECT="${tsconfigPath}" ts-node -r tsconfig-paths/register "${absoluteScriptPath}"`;

  // Parameter handling logic remains the same
  for (const paramName in params) {
      const value = params[paramName];
      const paramDef = tool.parameters.find(p => p.name === paramName);

      if (value !== undefined && value !== null && value !== '') {
          if (paramDef && paramDef.type === 'boolean') {
              if (value === true) {
                  command += ` --${paramName}`;
              }
          } else {
              const formattedValue = typeof value === 'string' && value.includes(' ') ? `"${value}"` : value;
              command += ` --${paramName} ${formattedValue}`;
          }
      }
  }
  return command;
}