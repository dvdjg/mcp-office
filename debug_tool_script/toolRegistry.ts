// Suggested interface for tool parameters
export interface ToolParameter {
  name: string; // e.g., "filePath", "output", "useComInterop"
  type: 'string' | 'boolean' | 'filePath' | 'number' | 'enum'; // 'enum' for predefined choices
  description: string;
  required: boolean;
  defaultValue?: any;
  isFixtureRelevant?: boolean; // True if this 'filePath' parameter can use fixtures
  enumValues?: string[]; // For type 'enum'
}

// Suggested interface for tool definitions
export interface ToolDefinition {
  id: string; // Unique identifier, e.g., "word/markdown.export"
  name: string; // User-friendly name, e.g., "Export Word to Markdown"
  category: string; // e.g., "Word", "Excel", "FS", "Office"
  scriptPath: string; // Relative path to the .tool.ts file, e.g., "src/tools/word/markdown.tool.ts"
  preconditions?: string[]; // Array of precondition messages
  parameters: ToolParameter[];
}

function parsePreconditions(preconditionString: string): string[] {
  const lowerPreconditionString = preconditionString.toLowerCase();
  if (!preconditionString || lowerPreconditionString === 'none obvious beyond system access' || lowerPreconditionString === 'none obvious' || lowerPreconditionString === 'directory path' || lowerPreconditionString === 'file path' || lowerPreconditionString === 'text file path') {
    // For very generic preconditions that are more like parameter types,
    // or "none obvious", treat as no specific user-facing precondition message.
    // The actual path requirement will be handled by parameters.
    if (lowerPreconditionString === 'directory path' || lowerPreconditionString === 'file path' || lowerPreconditionString === 'text file path') {
        return []; // These are handled by parameters
    }
    return [];
  }
  // Treat the entire string as a single precondition message
  return [preconditionString.trim()].filter(p => p.length > 0);
}

function getCategoryFromPath(filePath: string): string {
  const parts = filePath.split('/');
  if (parts.length > 1) {
    // Use the first part of the path (e.g., "dynamic", "excel", "fs") as the category
    const categoryKey = parts[0].toLowerCase();
    if (categoryKey === 'dynamic') return 'Dynamic';
    if (categoryKey === 'fs') return 'FileSystem';
    if (categoryKey === 'os') return 'OS';
    // Capitalize the first letter for other categories
    return categoryKey.charAt(0).toUpperCase() + categoryKey.slice(1);
  }
  return 'General'; // Default category if path structure is unexpected
}

function generateId(filePath: string, functionName: string): string {
    // Derive category from the first directory in the filePath
    const categoryPart = filePath.split('/')[0].toLowerCase();
    // Sanitize function name for ID: lowercase, replace spaces and special chars with hyphens
    const namePart = functionName
        .toLowerCase()
        .replace(/\s*\([^)]*\)\s*/g, '') // Remove content in parentheses
        .replace(/\s*\[[^\]]*\]\s*/g, '') // Remove content in brackets
        .replace(/[^a-z0-9\s-]/g, '')    // Remove special characters except hyphens and spaces
        .replace(/\s+/g, '-')           // Replace spaces with hyphens
        .replace(/-+/g, '-');           // Replace multiple hyphens with single
    return `${categoryPart}.${namePart}`;
}


export const TOOLS_DATA: ToolDefinition[] = [
  {
    id: generateId('dynamic/resources.tool.ts', "Manages dynamic resources (CRUD, search)"),
    name: "Manages dynamic resources (CRUD, search)",
    category: getCategoryFromPath('dynamic/resources.tool.ts'),
    scriptPath: "src/tools/dynamic/resources.tool.ts",
    preconditions: parsePreconditions("None obvious beyond system access"),
    parameters: [
         { name: "operation", type: "enum", description: "Operation to perform (list, read, write, delete, metadata, search)", required: true, enumValues: ["list", "read", "write", "delete", "metadata", "search"] },
         { name: "resourceId", type: "string", description: "ID of the resource (for read, delete, metadata)", required: false },
         { name: "filePath", type: "filePath", description: "File path for resource (for write)", required: false, isFixtureRelevant: true }, // Fixture relevance "Maybe"
         { name: "content", type: "string", description: "Content for write operation", required: false },
         { name: "type", type: "string", description: "Filter by resource type (for list, search)", required: false },
         { name: "query", type: "string", description: "Search query (for search)", required: false },
    ]
  },
  {
    id: generateId('excel/charts.tool.ts', "Manages charts in Excel files"),
    name: "Manages charts in Excel files",
    category: getCategoryFromPath('excel/charts.tool.ts'),
    scriptPath: "src/tools/excel/charts.tool.ts",
    preconditions: parsePreconditions("Active Excel document or Excel file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Excel file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('excel/dataAnalysis.tool.ts', "Performs data analysis in Excel (sort, filter, pivot)"),
    name: "Performs data analysis in Excel (sort, filter, pivot)",
    category: getCategoryFromPath('excel/dataAnalysis.tool.ts'),
    scriptPath: "src/tools/excel/dataAnalysis.tool.ts",
    preconditions: parsePreconditions("Active Excel document or Excel file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Excel file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('excel/formatTablesInWorksheet.tool.ts', "Formats tables within an Excel worksheet"),
    name: "Formats tables within an Excel worksheet",
    category: getCategoryFromPath('excel/formatTablesInWorksheet.tool.ts'),
    scriptPath: "src/tools/excel/formatTablesInWorksheet.tool.ts",
    preconditions: parsePreconditions("Active Excel document or Excel file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Excel file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('excel/range.tool.ts', "Manipulates cell ranges in Excel"),
    name: "Manipulates cell ranges in Excel",
    category: getCategoryFromPath('excel/range.tool.ts'),
    scriptPath: "src/tools/excel/range.tool.ts",
    preconditions: parsePreconditions("Active Excel document or Excel file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Excel file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('excel/tables.tool.ts', "Manages tables in Excel files"),
    name: "Manages tables in Excel files",
    category: getCategoryFromPath('excel/tables.tool.ts'),
    scriptPath: "src/tools/excel/tables.tool.ts",
    preconditions: parsePreconditions("Active Excel document or Excel file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Excel file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('excel/worksheets.tool.ts', "Manages worksheets in Excel files"),
    name: "Manages worksheets in Excel files",
    category: getCategoryFromPath('excel/worksheets.tool.ts'),
    scriptPath: "src/tools/excel/worksheets.tool.ts",
    preconditions: parsePreconditions("Active Excel document or Excel file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Excel file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('fs/archive.tool.ts', "Manages archive files (ZIP, 7z; list, extract, create)"),
    name: "Manages archive files (ZIP, 7z; list, extract, create)",
    category: getCategoryFromPath('fs/archive.tool.ts'),
    scriptPath: "src/tools/fs/archive.tool.ts",
    preconditions: parsePreconditions("Input archive file path (for list/extract), source paths (for create)"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input archive file path (for list/extract)", required: false, isFixtureRelevant: true },
      { name: "sourcePaths", type: "string", description: "Comma-separated source paths (for create)", required: false, isFixtureRelevant: true }, // Assuming source paths can be fixtures
      { name: "outputDirectory", type: "string", description: "Output directory for extraction", required: false }
    ]
  },
  {
    id: generateId('fs/directory.tool.ts', "Basic directory operations (list, create, delete)"),
    name: "Basic directory operations (list, create, delete)",
    category: getCategoryFromPath('fs/directory.tool.ts'),
    scriptPath: "src/tools/fs/directory.tool.ts",
    preconditions: parsePreconditions("Directory path"),
    parameters: [
      { name: "path", type: "string", description: "Directory path", required: true, isFixtureRelevant: false }
    ]
  },
  {
    id: generateId('fs/directoryOperations.tool.ts', "Advanced directory operations (find, tree)"),
    name: "Advanced directory operations (find, tree)",
    category: getCategoryFromPath('fs/directoryOperations.tool.ts'),
    scriptPath: "src/tools/fs/directoryOperations.tool.ts",
    preconditions: parsePreconditions("Directory path"),
    parameters: [
      { name: "path", type: "string", description: "Directory path", required: true, isFixtureRelevant: false }
    ]
  },
  {
    id: generateId('fs/file.tool.ts', "Basic file operations (read, write, delete, rename)"),
    name: "Basic file operations (read, write, delete, rename)",
    category: getCategoryFromPath('fs/file.tool.ts'),
    scriptPath: "src/tools/fs/file.tool.ts",
    preconditions: parsePreconditions("File path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Target file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('fs/fileContent.tool.ts', "Text file content manipulation"),
    name: "Text file content manipulation (insert/delete/replace lines, search/replace text)",
    category: getCategoryFromPath('fs/fileContent.tool.ts'),
    scriptPath: "src/tools/fs/fileContent.tool.ts",
    preconditions: parsePreconditions("Text file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Target text file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('fs/structuredData.tool.ts', "Reads/writes structured data files (CSV, JSON, XML)"),
    name: "Reads/writes structured data files (CSV, JSON, XML)",
    category: getCategoryFromPath('fs/structuredData.tool.ts'),
    scriptPath: "src/tools/fs/structuredData.tool.ts",
    preconditions: parsePreconditions("File path (CSV, JSON, XML)"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Target structured data file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('office/adaptWordToPowerpoint.tool.ts', "Adapts Word documents to PowerPoint presentations"),
    name: "Adapts Word documents to PowerPoint presentations",
    category: getCategoryFromPath('office/adaptWordToPowerpoint.tool.ts'),
    scriptPath: "src/tools/office/adaptWordToPowerpoint.tool.ts",
    preconditions: parsePreconditions("Input Word file path, (optional) active Word document"),
    parameters: [
      { name: "inputWordPath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('office/aiSuggest.tool.ts', "Provides AI-based suggestions for Office tasks"),
    name: "Provides AI-based suggestions for Office tasks",
    category: getCategoryFromPath('office/aiSuggest.tool.ts'),
    scriptPath: "src/tools/office/aiSuggest.tool.ts",
    preconditions: parsePreconditions("Potentially active Office document, context dependent"),
    parameters: [
        { name: "context", type: "string", description: "Context for AI suggestion", required: true, isFixtureRelevant: false } // Fixture relevance "Maybe"
    ]
  },
  {
    id: generateId('office/combine.tool.ts', "Combines multiple files into a single Word document"),
    name: "Combines multiple files into a single Word document",
    category: getCategoryFromPath('office/combine.tool.ts'),
    scriptPath: "src/tools/office/combine.tool.ts",
    preconditions: parsePreconditions("Directory path with source files, output Word file path"),
    parameters: [
      { name: "directoryPath", type: "filePath", description: "Directory path with source files", required: true, isFixtureRelevant: true },
      { name: "outputWordPath", type: "filePath", description: "Output Word file path", required: true, isFixtureRelevant: false }
    ]
  },
  {
    id: generateId('office/generalParseText.tool.ts', "Extracts text from various Office document types"),
    name: "Extracts text from various Office document types",
    category: getCategoryFromPath('office/generalParseText.tool.ts'),
    scriptPath: "src/tools/office/generalParseText.tool.ts",
    preconditions: parsePreconditions("Office file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Office file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('office/imageAnalysisGeneration.tool.ts', "Analyzes images and generates/inserts images in Office docs"),
    name: "Analyzes images and generates/inserts images in Office docs",
    category: getCategoryFromPath('office/imageAnalysisGeneration.tool.ts'),
    scriptPath: "src/tools/office/imageAnalysisGeneration.tool.ts",
    preconditions: parsePreconditions("Image source (file, URL, doc), target Office doc path"),
    parameters: [
      { name: "imageSourcePath", type: "filePath", description: "Image source file path or URL", required: true, isFixtureRelevant: true },
      { name: "targetDocPath", type: "filePath", description: "Target Office document path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('office/pdfExport.tool.ts', "Exports Office documents to PDF"),
    name: "Exports Office documents to PDF",
    category: getCategoryFromPath('office/pdfExport.tool.ts'),
    scriptPath: "src/tools/office/pdfExport.tool.ts",
    preconditions: parsePreconditions("Office file path (Word, Excel, PowerPoint)"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Office file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('office/pdfParse.tool.ts', "Parses and converts PDF files"),
    name: "Parses and converts PDF files",
    category: getCategoryFromPath('office/pdfParse.tool.ts'),
    scriptPath: "src/tools/office/pdfParse.tool.ts",
    preconditions: parsePreconditions("PDF file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "PDF file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('office/transfer.tool.ts', "Transfers data between Office applications"),
    name: "Transfers data between Office applications",
    category: getCategoryFromPath('office/transfer.tool.ts'),
    scriptPath: "src/tools/office/transfer.tool.ts",
    preconditions: parsePreconditions("Source/Target Office file paths/elements"),
    parameters: [
      { name: "sourcePath", type: "filePath", description: "Source Office file path/element specifier", required: true, isFixtureRelevant: true },
      { name: "targetPath", type: "filePath", description: "Target Office file path/element specifier", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('office/wordToPowerpoint.tool.ts', "Converts Word documents to PowerPoint"),
    name: "Converts Word documents to PowerPoint (likely core of adapt)",
    category: getCategoryFromPath('office/wordToPowerpoint.tool.ts'),
    scriptPath: "src/tools/office/wordToPowerpoint.tool.ts",
    preconditions: parsePreconditions("Input Word file path"),
    parameters: [
      { name: "inputWordPath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('office/workflow.tool.ts', "Executes multi-step Office tool workflows"),
    name: "Executes multi-step Office tool workflows",
    category: getCategoryFromPath('office/workflow.tool.ts'),
    scriptPath: "src/tools/office/workflow.tool.ts",
    preconditions: parsePreconditions("Configuration of workflow steps"),
    parameters: [
        { name: "workflowConfigPath", type: "filePath", description: "Path to workflow configuration file (e.g., JSON)", required: true, isFixtureRelevant: true } // Fixture relevance "Maybe"
    ]
  },
  {
    id: generateId('os/getActiveOfficeDocuments.tool.ts', "Lists currently open Microsoft Office documents"),
    name: "Lists currently open Microsoft Office documents",
    category: getCategoryFromPath('os/getActiveOfficeDocuments.tool.ts'),
    scriptPath: "src/tools/os/getActiveOfficeDocuments.tool.ts",
    preconditions: parsePreconditions("Running Office applications (Word, Excel, PowerPoint)"),
    parameters: []
  },
  {
    id: generateId('powerpoint/animations.tool.ts', "Manages animations and transitions in PowerPoint"),
    name: "Manages animations and transitions in PowerPoint",
    category: getCategoryFromPath('powerpoint/animations.tool.ts'),
    scriptPath: "src/tools/powerpoint/animations.tool.ts",
    preconditions: parsePreconditions("Active PowerPoint document or PowerPoint file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input PowerPoint file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('powerpoint/properties.tool.ts', "Manages PowerPoint presentation properties"),
    name: "Manages PowerPoint presentation properties",
    category: getCategoryFromPath('powerpoint/properties.tool.ts'),
    scriptPath: "src/tools/powerpoint/properties.tool.ts",
    preconditions: parsePreconditions("Active PowerPoint document or PowerPoint file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input PowerPoint file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('powerpoint/shapes.tool.ts', "Manages shapes in PowerPoint slides"),
    name: "Manages shapes in PowerPoint slides",
    category: getCategoryFromPath('powerpoint/shapes.tool.ts'),
    scriptPath: "src/tools/powerpoint/shapes.tool.ts",
    preconditions: parsePreconditions("Active PowerPoint document or PowerPoint file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input PowerPoint file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('powerpoint/slides.tool.ts', "Manages slides in PowerPoint presentations"),
    name: "Manages slides in PowerPoint presentations",
    category: getCategoryFromPath('powerpoint/slides.tool.ts'),
    scriptPath: "src/tools/powerpoint/slides.tool.ts",
    preconditions: parsePreconditions("Active PowerPoint document or PowerPoint file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input PowerPoint file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/analyze.tool.ts', "Analyzes Word documents, adds comments, summarizes"),
    name: "Analyzes Word documents, adds comments, summarizes",
    category: getCategoryFromPath('word/analyze.tool.ts'),
    scriptPath: "src/tools/word/analyze.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/applyAutoTitles.tool.ts', "Applies heading styles automatically in Word"),
    name: "Applies heading styles automatically in Word",
    category: getCategoryFromPath('word/applyAutoTitles.tool.ts'),
    scriptPath: "src/tools/word/applyAutoTitles.tool.ts",
    preconditions: parsePreconditions("Active Word document or Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: false, isFixtureRelevant: true } // Optional if active doc is used
    ]
  },
  {
    id: generateId('word/batch.tool.ts', "Executes multiple Word tool operations in batch"),
    name: "Executes multiple Word tool operations in batch",
    category: getCategoryFromPath('word/batch.tool.ts'),
    scriptPath: "src/tools/word/batch.tool.ts",
    preconditions: parsePreconditions("Word file path, list of operations"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true },
      { name: "operationsConfigPath", type: "filePath", description: "Path to operations configuration file (e.g., JSON)", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/charts.tool.ts', "Manages charts in Word documents"),
    name: "Manages charts in Word documents",
    category: getCategoryFromPath('word/charts.tool.ts'),
    scriptPath: "src/tools/word/charts.tool.ts",
    preconditions: parsePreconditions("Active Word document or Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/codeFormat.tool.ts', "Formats code blocks and metadata within Word docs"),
    name: "Formats code blocks and metadata within Word docs",
    category: getCategoryFromPath('word/codeFormat.tool.ts'),
    scriptPath: "src/tools/word/codeFormat.tool.ts",
    preconditions: parsePreconditions("Word file path with code blocks"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path with code blocks", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/concludeStoryInDocument.tool.ts', "Generates and inserts a story conclusion in Word"),
    name: "Generates and inserts a story conclusion in Word",
    category: getCategoryFromPath('word/concludeStoryInDocument.tool.ts'),
    scriptPath: "src/tools/word/concludeStoryInDocument.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/embeddedObjects.tool.ts', "Manages embedded OLE objects in Word documents"),
    name: "Manages embedded OLE objects in Word documents",
    category: getCategoryFromPath('word/embeddedObjects.tool.ts'),
    scriptPath: "src/tools/word/embeddedObjects.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/generateAndInsertText.tool.ts', "Generates text via LLM and inserts into Word"),
    name: "Generates text via LLM and inserts into Word",
    category: getCategoryFromPath('word/generateAndInsertText.tool.ts'),
    scriptPath: "src/tools/word/generateAndInsertText.tool.ts",
    preconditions: parsePreconditions("Word file path, prompt"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true },
      { name: "prompt", type: "string", description: "Prompt for text generation", required: true }
    ]
  },
  {
    id: generateId('word/headersFooters.tool.ts', "Manages headers and footers in Word documents"),
    name: "Manages headers and footers in Word documents",
    category: getCategoryFromPath('word/headersFooters.tool.ts'),
    scriptPath: "src/tools/word/headersFooters.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/image.tool.ts', "Manages images in Word documents (extract, insert)"),
    name: "Manages images in Word documents (extract, insert)",
    category: getCategoryFromPath('word/image.tool.ts'),
    scriptPath: "src/tools/word/image.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/markdown.tool.ts', "Handles Markdown to Word and Word to Markdown conversion"),
    name: "Handles Markdown to Word and Word to Markdown conversion",
    category: getCategoryFromPath('word/markdown.tool.ts'),
    scriptPath: "src/tools/word/markdown.tool.ts",
    preconditions: parsePreconditions("Word file path or Markdown file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word or Markdown file path", required: true, isFixtureRelevant: true },
      { name: "output", type: "filePath", description: "Output file path", required: true, isFixtureRelevant: false },
      { name: "useComInterop", type: "boolean", description: "Use COM Interop", required: false, defaultValue: false },
      { name: "imageDir", type: "string", description: "Directory for images (Word to MD)", required: false, defaultValue: "images" }
    ]
  },
  {
    id: generateId('word/merge.tool.ts', "Merges multiple Word documents"),
    name: "Merges multiple Word documents",
    category: getCategoryFromPath('word/merge.tool.ts'),
    scriptPath: "src/tools/word/merge.tool.ts",
    preconditions: parsePreconditions("List of Word file paths, output Word file path"),
    parameters: [
      { name: "docs", type: "string", description: "Comma-separated list of Word file paths to merge", required: true, isFixtureRelevant: true },
      { name: "output", type: "filePath", description: "Output Word file path", required: true, isFixtureRelevant: false }
    ]
  },
  {
    id: generateId('word/mermaidExport.tool.ts', "Exports Mermaid diagrams from Word documents"),
    name: "Exports Mermaid diagrams from Word documents",
    category: getCategoryFromPath('word/mermaidExport.tool.ts'),
    scriptPath: "src/tools/word/mermaidExport.tool.ts",
    preconditions: parsePreconditions("Word file path with Mermaid diagrams"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path with Mermaid diagrams", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/mermaidImport.tool.ts', "Imports Mermaid diagrams into Word documents"),
    name: "Imports Mermaid diagrams into Word documents",
    category: getCategoryFromPath('word/mermaidImport.tool.ts'),
    scriptPath: "src/tools/word/mermaidImport.tool.ts",
    preconditions: parsePreconditions("Word file path, Mermaid syntax"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true },
      { name: "mermaidSyntax", type: "string", description: "Mermaid diagram syntax", required: true }
    ]
  },
  {
    id: generateId('word/metadata.tool.ts', "Manages metadata of Word documents"),
    name: "Manages metadata of Word documents",
    category: getCategoryFromPath('word/metadata.tool.ts'),
    scriptPath: "src/tools/word/metadata.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/page.tool.ts', "Configures page layout settings in Word"),
    name: "Configures page layout settings in Word",
    category: getCategoryFromPath('word/page.tool.ts'),
    scriptPath: "src/tools/word/page.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/reformat.tool.ts', "Reformats Word documents (analyze, apply styles)"),
    name: "Reformats Word documents (analyze, apply styles)",
    category: getCategoryFromPath('word/reformat.tool.ts'),
    scriptPath: "src/tools/word/reformat.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/saveActiveWordAsMarkdown.tool.ts', "Saves the active Word document as Markdown"),
    name: "Saves the active Word document as Markdown",
    category: getCategoryFromPath('word/saveActiveWordAsMarkdown.tool.ts'),
    scriptPath: "src/tools/word/saveActiveWordAsMarkdown.tool.ts",
    preconditions: parsePreconditions("Active Word document"),
    parameters: [
        { name: "output", type: "filePath", description: "Output Markdown file path", required: true, isFixtureRelevant: false } // Even if active, it needs an output path
    ]
  },
  {
    id: generateId('word/searchReplace.tool.ts', "Searches and replaces text in Word documents"),
    name: "Searches and replaces text in Word documents",
    category: getCategoryFromPath('word/searchReplace.tool.ts'),
    scriptPath: "src/tools/word/searchReplace.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/styles.tool.ts', "Manages styles in Word documents (apply, list)"),
    name: "Manages styles in Word documents (apply, list)",
    category: getCategoryFromPath('word/styles.tool.ts'),
    scriptPath: "src/tools/word/styles.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/tables.tool.ts', "Manages tables in Word documents"),
    name: "Manages tables in Word documents",
    category: getCategoryFromPath('word/tables.tool.ts'),
    scriptPath: "src/tools/word/tables.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/template.tool.ts', "Manages Word templates (analyze, replace placeholders)"),
    name: "Manages Word templates (analyze, replace placeholders)",
    category: getCategoryFromPath('word/template.tool.ts'),
    scriptPath: "src/tools/word/template.tool.ts",
    preconditions: parsePreconditions("Word template file path (.dotx) or .docx with controls"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word template or document file path", required: true, isFixtureRelevant: true }
    ]
  },
  {
    id: generateId('word/text.tool.ts', "Manages text in Word documents (get, insert, modify, delete)"),
    name: "Manages text in Word documents (get, insert, modify, delete)",
    category: getCategoryFromPath('word/text.tool.ts'),
    scriptPath: "src/tools/word/text.tool.ts",
    preconditions: parsePreconditions("Word file path"),
    parameters: [
      { name: "filePath", type: "filePath", description: "Input Word file path", required: true, isFixtureRelevant: true }
    ]
  }
];

export function loadTools(): ToolDefinition[] {
  return TOOLS_DATA;
}

export function getUniqueCategories(tools: ToolDefinition[]): string[] {
  const categories = new Set<string>();
  tools.forEach(tool => categories.add(tool.category));
  return Array.from(categories).sort();
}

export function filterToolsByCategory(tools: ToolDefinition[], category: string): ToolDefinition[] {
  return tools.filter(tool => tool.category === category);
}