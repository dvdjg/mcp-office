// /src/tools/index.ts
// =============================================================================
/**
 * @file Aggregates all tool definitions for registration with the MCP server.
 */
import { McpResource } from '@/types/common.types';
import { fsDirectoryTool } from './fs/directory.tool';
import { fsFileTool } from './fs/file.tool';
import { wordStylesTool } from './word/styles.tool';
import { wordMarkdownTool } from './word/markdown.tool';
import { wordMergeTool } from './word/merge.tool';
import { wordTemplateTool } from './word/template.tool'; // Importar la nueva herramienta
import { wordTextTool } from './word/text.tool'; // Importar la herramienta de texto
import { wordSearchReplaceTool } from './word/searchReplace.tool'; // Importar la nueva herramienta
// Import static resources
import aiAssistantGuideResource from '../resources/static/ai_assistant_guide.resource'; // Added import
// Import other tools (Excel, PowerPoint, Office, etc.) here
// import { excelRangeTool } from './excel/range.tool';
// import { powerpointSlidesTool } from './powerpoint/slides.tool';
// import { officePdfTool } from './office/pdf.tool';

const allTools: McpResource[] = [
    ...fsDirectoryTool,
    ...fsFileTool,
    ...wordStylesTool,
    ...wordMarkdownTool,
    ...wordMergeTool,
    ...wordTemplateTool, // Añadir la nueva herramienta
    ...wordTextTool, // Añadir la herramienta de texto
    ...wordSearchReplaceTool, // Añadir la nueva herramienta
    ...aiAssistantGuideResource, // Added the new resource
    // ...excelRangeTool,
    // ...powerpointSlidesTool,
    // ...officePdfTool,
    // Add other imported tools here
];

// Add placeholder entries for tools defined in the plan but not implemented yet
const placeholderTools: McpResource[] = [
    // Word
    // { path: 'word/search-replace', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/search-replace not implemented.' } }), description: 'Perform advanced search and replace (Not Implemented - Requires VBA/COM)' }, // Placeholder removed
    { path: 'word/page', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/page not implemented.' } }), description: 'Configure page layout (Not Implemented)' },
    { path: 'word/headers-footers', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/headers-footers not implemented.' } }), description: 'Manage headers and footers (Not Implemented)' },
    { path: 'word/tables', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/tables not implemented.' } }), description: 'Create and manage tables (Not Implemented)' },
    { path: 'word/charts', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/charts not implemented.' } }), description: 'Insert and manage charts (Not Implemented)' },
    { path: 'word/embedded-objects', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/embedded-objects not implemented.' } }), description: 'Manage embedded objects (Not Implemented - Requires VBA/COM)' },
    { path: 'word/metadata', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/metadata not implemented.' } }), description: 'Manage document properties and comments (Not Implemented)' },
    { path: 'word/batch', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/batch not implemented.' } }), description: 'Execute multiple operations (Not Implemented)' },
    { path: 'word/mermaid/import', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/mermaid/import not implemented.' } }), description: 'Import and render Mermaid diagrams (Not Implemented)' },
    { path: 'word/mermaid/export', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/mermaid/export not implemented.' } }), description: 'Export Mermaid diagrams (Not Implemented)' },
    // { path: 'word/merge', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/merge not implemented.' } }), description: 'Merge multiple Word documents (Not Implemented - Requires VBA/COM)' }, // Eliminado placeholder
    // { path: 'word/template', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/template not implemented.' } }), description: 'Create a template with placeholders (Not Implemented)' }, // Eliminado placeholder
    { path: 'word/reformat', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/reformat not implemented.' } }), description: 'Reformat a document professionally (Not Implemented)' },
    { path: 'word/analyze', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/analyze not implemented.' } }), description: 'Analyze content and add comments (Not Implemented - Requires AI integration)' },
    { path: 'word/code-format', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool word/code-format not implemented.' } }), description: 'Format code and metadata with syntax highlighting (Not Implemented)' },
    // Excel (Placeholders)
    { path: 'excel/worksheets', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool excel/worksheets not implemented.' } }), description: 'Manage Excel worksheets (Not Implemented - Requires Office JS/Scripts)' },
    { path: 'excel/range', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool excel/range not implemented.' } }), description: 'Manipulate cell ranges (Not Implemented - Requires Office JS/Scripts)' },
    // PowerPoint (Placeholders)
    { path: 'powerpoint/slides', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool powerpoint/slides not implemented.' } }), description: 'Manage PowerPoint slides (Not Implemented - Requires Office JS/VBA)' },
    // Office (Placeholders)
    { path: 'office/transfer', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool office/transfer not implemented.' } }), description: 'Move data between applications (Not Implemented - Requires COM/VBA)' },
    { path: 'office/pdf/export', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool office/pdf/export not implemented.' } }), description: 'Export documents to PDF (Not Implemented - Requires Office JS/VBA)' },
    { path: 'office/combine', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool office/combine not implemented.' } }), description: 'Combine multiple files into a Word document (Not Implemented)' },
    { path: 'office/word-to-powerpoint', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool office/word-to-powerpoint not implemented.' } }), description: 'Convert Word to PowerPoint (Not Implemented)' },
    { path: 'office/ai-suggest', handler: async () => ({ success: false, error: { code: 'NOT_IMPLEMENTED', message: 'Tool office/ai-suggest not implemented.' } }), description: 'Provide AI-driven recommendations (Not Implemented)' },
];


export default [...allTools, ...placeholderTools];