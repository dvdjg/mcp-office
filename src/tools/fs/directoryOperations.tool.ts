import * as fs from 'fs/promises';
import * as path from 'path';
import { z } from 'zod';

// --- Helper function for glob to regex (basic implementation) ---
function globToRegex(glob: string, caseSensitive: boolean = true): RegExp {
  const regexString = glob
    .replace(/[.+^${}()|[\]\\]/g, '\\$&') // Escape special regex characters
    .replace(/\*/g, '.*') // Convert * to .*
    .replace(/\?/g, '.')  // Convert ? to .
  return new RegExp(`^${regexString}$`, caseSensitive ? '' : 'i');
}

// --- list_directory_contents ---

export const listDirectoryContentsParams = z.object({
  path: z.string().min(1, "Path cannot be empty."),
  options: z.object({
    recursive: z.boolean().optional().default(false),
    includeSize: z.boolean().optional().default(false),
    maxDepth: z.number().int().min(0).optional(),
  }).optional(),
});

export type ListDirectoryContentsParams = z.infer<typeof listDirectoryContentsParams>;
export type ListDirectoryItem = {
  name: string;
  type: 'file' | 'directory' | 'other';
  path: string;
  size?: number;
  depth: number;
};

async function listDirectoryContentsRecursive(
  currentPath: string,
  basePath: string,
  options: NonNullable<ListDirectoryContentsParams['options']>,
  currentDepth: number,
  maxDepthEffective: number,
): Promise<ListDirectoryItem[]> {
  const results: ListDirectoryItem[] = [];
  if (currentDepth > maxDepthEffective) {
    return results;
  }

  let entries;
  try {
    entries = await fs.readdir(currentPath, { withFileTypes: true });
  } catch (error: any) {
    // Log error or handle specific errors like EACCES or ENOENT
    // console.error(`Error reading directory ${currentPath}: ${error.message}`); // Removed for cleaner test output, error is re-thrown
    throw new Error(`Failed to read directory ${currentPath}: ${error.message}`);
  }

  for (const entry of entries) {
    const entryPath = path.resolve(currentPath, entry.name);
    const item: Partial<ListDirectoryItem> = {
      name: entry.name,
      path: entryPath,
      depth: currentDepth,
    };

    if (entry.isFile()) {
      item.type = 'file';
      if (options.includeSize) {
        try {
          const stats = await fs.stat(entryPath);
          item.size = stats.size;
        } catch (statError: any) {
          console.warn(`Could not get stats for file ${entryPath}: ${statError.message}`);
          item.size = 0; // Or undefined, based on preference
        }
      }
    } else if (entry.isDirectory()) {
      item.type = 'directory';
      if (options.includeSize) {
        item.size = 0; // Reporting 0 for directories as per simplified approach
      }
    } else {
      item.type = 'other';
      if (options.includeSize) {
        item.size = 0; // Reporting 0 for 'other' types
      }
    }
    results.push(item as ListDirectoryItem);

    if (item.type === 'directory' && options.recursive && currentDepth < maxDepthEffective) {
      results.push(
        ...(await listDirectoryContentsRecursive(
          entryPath,
          basePath,
          options,
          currentDepth + 1,
          maxDepthEffective,
        )),
      );
    }
  }
  return results;
}

export async function listDirectoryContents(
  params: ListDirectoryContentsParams,
): Promise<ListDirectoryItem[]> {
  const resolvedPath = path.resolve(params.path);
  const normalizedPath = path.normalize(resolvedPath);

  const effectiveOptions = {
    recursive: params.options?.recursive ?? false,
    includeSize: params.options?.includeSize ?? false,
    maxDepth: params.options?.maxDepth, // Keep undefined if not set
  };

  let maxDepthEffective: number;
  if (effectiveOptions.recursive) {
    maxDepthEffective = effectiveOptions.maxDepth ?? Infinity;
  } else {
    maxDepthEffective = 0; // Only top-level if not recursive
  }
  // If maxDepth is explicitly set, it overrides the recursive default.
  if (typeof effectiveOptions.maxDepth === 'number') {
      maxDepthEffective = effectiveOptions.maxDepth;
  }


  return listDirectoryContentsRecursive(normalizedPath, normalizedPath, effectiveOptions, 0, maxDepthEffective);
}


// --- find_files ---

export const findFilesParams = z.object({
  directoryPath: z.string().min(1, "Directory path cannot be empty."),
  namePattern: z.string().min(1, "Name pattern cannot be empty."),
  options: z.object({
    isRegexName: z.boolean().optional().default(false),
    recursive: z.boolean().optional().default(true),
    contentPattern: z.string().optional(),
    contentIsRegex: z.boolean().optional().default(false),
    maxDepth: z.number().int().min(0).optional(),
    encoding: z.string().optional().default('utf8') as z.ZodType<BufferEncoding>, // BufferEncoding is a wide type
  }).optional(),
});

export type FindFilesParams = z.infer<typeof findFilesParams>;
export type FindFilesResultItem = {
  path: string;
  matchesContent?: boolean;
};

async function findFilesRecursive(
  currentSearchPath: string,
  nameMatcher: RegExp,
  options: NonNullable<FindFilesParams['options']>,
  currentDepth: number,
  maxDepthEffective: number,
): Promise<FindFilesResultItem[]> {
  const results: FindFilesResultItem[] = [];
  if (currentDepth > maxDepthEffective) {
    return results;
  }

  let entries;
  try {
    entries = await fs.readdir(currentSearchPath, { withFileTypes: true });
  } catch (error: any) {
    console.error(`Error reading directory ${currentSearchPath} in findFiles: ${error.message}`);
    // Silently skip unreadable directories or rethrow based on desired strictness
    return results;
  }

  for (const entry of entries) {
    const entryPath = path.resolve(currentSearchPath, entry.name);
    if (entry.isFile()) {
      if (nameMatcher.test(entry.name)) {
        let matchesContent: boolean | undefined = undefined;
        if (options.contentPattern) {
          try {
            const fileContent = await fs.readFile(entryPath, { encoding: options.encoding });
            const contentRegex = options.contentIsRegex
              ? new RegExp(options.contentPattern)
              : new RegExp(options.contentPattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'); // Basic literal string to regex
            matchesContent = contentRegex.test(fileContent);
          } catch (readError: any) {
            console.warn(`Could not read file ${entryPath} for content search: ${readError.message}`);
            matchesContent = false; // Or handle as error
          }
        }
        // If name matches, the file is a candidate.
        // The `matchesContent` variable will be:
        // - true: if contentPattern was provided and matched.
        // - false: if contentPattern was provided and did NOT match.
        // - undefined: if contentPattern was NOT provided.
        // We always push the file if the name matches, and `matchesContent` carries the content match status.
        results.push({ path: entryPath, matchesContent });

      }
    } else if (entry.isDirectory() && options.recursive && currentDepth < maxDepthEffective) {
      results.push(
        ...(await findFilesRecursive(
          entryPath,
          nameMatcher,
          options,
          currentDepth + 1,
          maxDepthEffective,
        )),
      );
    }
  }
  return results;
}


export async function findFiles(params: FindFilesParams): Promise<FindFilesResultItem[]> {
  const resolvedPath = path.resolve(params.directoryPath);
  const normalizedPath = path.normalize(resolvedPath);

  try {
      const stats = await fs.stat(normalizedPath);
      if (!stats.isDirectory()) {
          throw new Error(`Path is not a directory: ${normalizedPath}`);
      }
  } catch (error: any) {
      throw new Error(`Failed to access directory path ${normalizedPath}: ${error.message}`);
  }

  const effectiveOptions = {
    isRegexName: params.options?.isRegexName ?? false,
    recursive: params.options?.recursive ?? true,
    contentPattern: params.options?.contentPattern,
    contentIsRegex: params.options?.contentIsRegex ?? false,
    maxDepth: params.options?.maxDepth,
    encoding: params.options?.encoding ?? 'utf8',
  };

  const nameMatcher = effectiveOptions.isRegexName
    ? new RegExp(params.namePattern)
    : globToRegex(params.namePattern);

  let maxDepthEffective: number;
   if (effectiveOptions.recursive) {
    maxDepthEffective = effectiveOptions.maxDepth ?? Infinity;
  } else {
    maxDepthEffective = 0; // Only top-level if not recursive
  }
  // If maxDepth is explicitly set, it overrides the recursive default.
  if (typeof effectiveOptions.maxDepth === 'number') {
      maxDepthEffective = effectiveOptions.maxDepth;
  }

  return findFilesRecursive(normalizedPath, nameMatcher, effectiveOptions, 0, maxDepthEffective);
}


// --- get_directory_tree ---

export const getDirectoryTreeParams = z.object({
  path: z.string().min(1, "Path cannot be empty."),
  options: z.object({
    maxDepth: z.number().int().min(0).optional().default(3),
    includeFiles: z.boolean().optional().default(true),
    includeSize: z.boolean().optional().default(false),
  }).optional(),
});

export type GetDirectoryTreeParams = z.infer<typeof getDirectoryTreeParams>;

async function getDirectoryTreeRecursive(
  currentPath: string,
  options: NonNullable<GetDirectoryTreeParams['options']>,
  currentDepth: number,
  prefix: string,
): Promise<string> {
  let treeString = '';
  if (currentDepth > options.maxDepth) {
    return treeString;
  }

  let entries;
  try {
    entries = await fs.readdir(currentPath, { withFileTypes: true });
    // Sort entries: directories first, then files, then alphabetically
    entries.sort((a, b) => {
        if (a.isDirectory() && !b.isDirectory()) return -1;
        if (!a.isDirectory() && b.isDirectory()) return 1;
        return a.name.localeCompare(b.name);
    });
  } catch (error: any) {
    console.error(`Error reading directory ${currentPath} for tree: ${error.message}`);
    return `${prefix}└── [Error reading directory: ${error.message}]\n`;
  }

  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    const isLast = i === entries.length - 1;
    const connector = isLast ? '└── ' : '├── ';
    const entryPath = path.join(currentPath, entry.name); // path.join for relative display

    if (entry.isDirectory()) {
      treeString += `${prefix}${connector}${entry.name}/\n`;
      const newPrefix = prefix + (isLast ? '    ' : '│   ');
      treeString += await getDirectoryTreeRecursive(
        path.resolve(currentPath, entry.name), // Use resolved path for fs operations
        options,
        currentDepth + 1,
        newPrefix,
      );
    } else if (entry.isFile() && options.includeFiles) {
      let sizeInfo = '';
      if (options.includeSize) {
        try {
          const stats = await fs.stat(path.resolve(currentPath, entry.name));
          sizeInfo = ` (${stats.size} bytes)`;
        } catch (statError: any) {
          sizeInfo = ' (size unavailable)';
        }
      }
      treeString += `${prefix}${connector}${entry.name}${sizeInfo}\n`;
    } else if (options.includeFiles) { // For 'other' types if includeFiles is true
        treeString += `${prefix}${connector}${entry.name} [other]\n`;
    }
  }
  return treeString;
}

export async function getDirectoryTree(params: GetDirectoryTreeParams): Promise<string> {
  const resolvedPath = path.resolve(params.path);
  const normalizedPath = path.normalize(resolvedPath);
  
  const effectiveOptions = {
    maxDepth: params.options?.maxDepth ?? 3,
    includeFiles: params.options?.includeFiles ?? true,
    includeSize: params.options?.includeSize ?? false,
  };

  try {
      const stats = await fs.stat(normalizedPath);
      if (!stats.isDirectory()) {
          throw new Error(`Path is not a directory: ${normalizedPath}`);
      }
  } catch (error: any) {
      throw new Error(`Failed to access path ${normalizedPath}: ${error.message}`);
  }
  
  let rootName = path.basename(normalizedPath);
  if (normalizedPath === path.resolve(normalizedPath, '..')) { // Check if it's a root like C:\
    rootName = normalizedPath;
  }


  const treeHeader = `${rootName}/\n`;
  const treeBody = await getDirectoryTreeRecursive(normalizedPath, effectiveOptions, 0, '');
  return treeHeader + treeBody;
}


// --- Tool Registration Object ---
export const directoryOperationsTools = {
  list_directory_contents: {
    name: 'list_directory_contents',
    description: 'Lists the contents of a specified directory, optionally including file sizes and recursion.',
    schema: listDirectoryContentsParams,
    execute: listDirectoryContents,
    inputSchema: listDirectoryContentsParams.shape.options.unwrap().shape, // For better display if needed
    outputSchema: z.array(z.object({
        name: z.string(),
        type: z.enum(['file', 'directory', 'other']),
        path: z.string(),
        size: z.number().optional(),
        depth: z.number(),
    })),
  },
  find_files: {
    name: 'find_files',
    description: 'Searches for files within a directory (and its subdirectories) based on a name pattern (glob or regex) and optionally content.',
    schema: findFilesParams,
    execute: findFiles,
    inputSchema: findFilesParams.shape.options.unwrap().shape,
    outputSchema: z.array(z.object({
        path: z.string(),
        matchesContent: z.boolean().optional(),
    })),
  },
  get_directory_tree: {
    name: 'get_directory_tree',
    description: 'Generates a string representation of a directory tree structure.',
    schema: getDirectoryTreeParams,
    execute: getDirectoryTree,
    inputSchema: getDirectoryTreeParams.shape.options.unwrap().shape,
    outputSchema: z.string(),
  },
};