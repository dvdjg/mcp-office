import { promises as fs } from 'fs';
import * as pathUtil from 'path';

// Helper type for encoding
type ValidBufferEncoding = BufferEncoding;

// Helper functions for file reading/writing with encoding
const LINE_SEPARATOR = '\n';
const WRITE_SEPARATOR = LINE_SEPARATOR;

function escapeRegExp(string: string): string {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeLineEndings(text: string): string {
  return text.replace(/\r\n|\r|\n/g, LINE_SEPARATOR);
}

function normalizeContent(content: string): { lines: string[], hadTrailingNewline: boolean } {
  // Handle empty content
  if (!content) {
    return { lines: [], hadTrailingNewline: false };
  }

  // Standardize line endings
  const normalizedContent = normalizeLineEndings(content);
  
  // Check for trailing newline before any trimming
  const hadTrailingNewline = normalizedContent.endsWith(LINE_SEPARATOR);
  
  // Split into lines, preserving empty lines but removing final empty line if it exists
  const lines = normalizedContent.replace(/\n$/, '').split(LINE_SEPARATOR);
  
  return { lines, hadTrailingNewline };
}

function joinLines(lines: string[], preserveTrailingNewline = true): string {
  if (lines.length === 0) return preserveTrailingNewline ? LINE_SEPARATOR : '';
  const joined = lines.join(WRITE_SEPARATOR);
  return preserveTrailingNewline ? joined + WRITE_SEPARATOR : joined;
}


async function ensureDirectoryExists(filePath: string): Promise<void> {
  const dir = pathUtil.dirname(filePath);
  await fs.mkdir(dir, { recursive: true });
}

async function _readFileLines(filePath: string, encoding: ValidBufferEncoding = 'utf8'): Promise<{ lines: string[], hadTrailingNewline: boolean }> {
  const content = await fs.readFile(filePath, { encoding });
  return normalizeContent(content);
}

async function _writeFileLines(filePath: string, lines: string[], encoding: ValidBufferEncoding = 'utf8', preserveTrailingNewline = true): Promise<void> {
  const content = joinLines(lines, preserveTrailingNewline);
  await ensureDirectoryExists(filePath);
  await fs.writeFile(filePath, content, { encoding });
}

async function _readBinaryFile(filePath: string): Promise<Buffer> {
  return fs.readFile(filePath);
}

async function _writeBinaryFile(filePath: string, data: Buffer): Promise<void> {
  await fs.writeFile(filePath, data);
}

// --- Tool Implementations ---

// A. Basic Text File Operations

export interface InsertTextLinesArgs {
  path: string;
  lineNumber: number;
  lines: string[];
  encoding?: ValidBufferEncoding;
}
export async function insert_text_lines({ path, lineNumber, lines: newLinesToInsert, encoding = 'utf8' }: InsertTextLinesArgs): Promise<void> {
  if (lineNumber < 0) throw new Error('Line number cannot be negative');

  const { lines: fileLines, hadTrailingNewline } = await _readFileLines(path, encoding);

  if (lineNumber < 0) {
    throw new Error('Line number must be non-negative');
  }

  // Handle special cases
  if (fileLines.length === 0) {
    await _writeFileLines(path, newLinesToInsert, encoding, true);
    return;
  }

  // Convert to 0-based index for array operations
  let insertIndex: number;
  if (lineNumber === 0 || lineNumber > fileLines.length) {
    // Append at end
    insertIndex = fileLines.length;
  } else {
    // Insert at specified position
    insertIndex = lineNumber - 1;
  }

  fileLines.splice(insertIndex, 0, ...newLinesToInsert);
  await _writeFileLines(path, fileLines, encoding, hadTrailingNewline);
}

export interface DeleteTextLinesArgs {
  path: string;
  startLine: number;
  endLine: number;
  encoding?: ValidBufferEncoding;
}
export async function delete_text_lines({ path, startLine, endLine, encoding = 'utf8' }: DeleteTextLinesArgs): Promise<void> {
  if (startLine < 1) throw new Error('Start line cannot be less than 1');
  if (startLine > endLine) throw new Error('Invalid line range: startLine cannot be greater than endLine');

  const { lines: fileLines, hadTrailingNewline } = await _readFileLines(path, encoding);
  
  // Early return if file is empty
  if (fileLines.length === 0) return;

  // Convert to 0-based indices and handle bounds
  const start = Math.min(Math.max(startLine - 1, 0), fileLines.length - 1);
  const end = Math.min(endLine - 1, fileLines.length - 1);

  // Only delete if range is valid
  if (start <= end) {
    fileLines.splice(start, end - start + 1);
  }

  await _writeFileLines(path, fileLines, encoding, hadTrailingNewline);
}

export interface ReplaceTextLinesArgs {
  path: string;
  startLine: number;
  endLine: number;
  newLines: string[];
  encoding?: ValidBufferEncoding;
}
export async function replace_text_lines({ path, startLine, endLine, newLines, encoding = 'utf8' }: ReplaceTextLinesArgs): Promise<void> {
  if (startLine < 1) throw new Error('Start line cannot be less than 1');
  if (startLine > endLine) throw new Error('Invalid line range: startLine cannot be greater than endLine');

  const { lines: currentLines, hadTrailingNewline } = await _readFileLines(path, encoding);

  const start = Math.max(0, startLine - 1);
  const end = Math.min(endLine - 1, currentLines.length - 1);

  if (start >= currentLines.length) {
    currentLines.push(...newLines);
  } else {
    currentLines.splice(start, end - start + 1, ...newLines);
  }

  await _writeFileLines(path, currentLines, encoding, hadTrailingNewline);
}

export interface ExtractTextFromRangeArgs {
  path: string;
  startLine: number;
  startChar: number;
  endLine: number;
  endChar: number;
  encoding?: ValidBufferEncoding;
}
export async function extract_text_from_range({ path, startLine, startChar = 0, endLine, endChar, encoding = 'utf8' }: ExtractTextFromRangeArgs): Promise<string> {
  if (startLine > endLine || startLine < 1) return '';
  if (endChar !== undefined && startChar > endChar) return '';
  
  const { lines, hadTrailingNewline } = await _readFileLines(path, encoding);
  if (lines.length === 0) return '';

  // Convert to 0-based indices and handle bounds
  const sLineIdx = Math.max(0, startLine - 1);
  const eLineIdx = Math.min(endLine - 1, lines.length - 1);

  if (sLineIdx >= lines.length || sLineIdx > eLineIdx) return '';

  // Handle single line extraction
  if (sLineIdx === eLineIdx) {
    const line = lines[sLineIdx];
    const start = Math.min(startChar, line.length);
    const end = endChar !== undefined ? Math.min(endChar, line.length) : line.length;
    return line.slice(start, end);
  }

  // For multi-line extraction
  const result: string[] = [];
  
  // First line - apply startChar
  const firstLine = lines[sLineIdx];
  // Ensure startChar is within bounds of the first line for slicing
  const actualStartChar = Math.min(startChar, firstLine.length);
  result.push(firstLine.slice(actualStartChar));

  // Middle lines - include fully
  for (let i = sLineIdx + 1; i < eLineIdx; i++) {
    result.push(lines[i]);
  }

  // Last line - apply endChar if specified
  const lastLine = lines[eLineIdx];
  // Ensure endChar is within bounds for slicing. If undefined, use line length.
  const actualEndChar = endChar !== undefined ? Math.min(endChar, lastLine.length) : lastLine.length;
  result.push(lastLine.slice(0, actualEndChar));

  const extractedText = result.join('\n');
  
  // Determine if the extraction covers the very end of the original content
  const extractsToEndOfFile = eLineIdx === lines.length - 1 &&
                              actualEndChar === lines[eLineIdx].length;
                              
  if (hadTrailingNewline && extractsToEndOfFile) {
    return extractedText + '\n';
  }
  
  return extractedText;
}

export interface SearchReplaceInTextRangeArgs {
  path: string;
  searchTerm: string;
  replacement: string | ((substring: string, ...args: unknown[]) => string);
  startLine?: number;
  endLine?: number;
  isRegex?: boolean;
  replaceAll?: boolean;
  encoding?: ValidBufferEncoding;
}
export async function search_replace_in_text_range({
  path,
  searchTerm,
  replacement,
  startLine = 1,
  endLine,
  isRegex = false,
  replaceAll = true,
  encoding = 'utf8',
}: SearchReplaceInTextRangeArgs): Promise<{ replacementsMade: number }> {
  const { lines: fileLines, hadTrailingNewline } = await _readFileLines(path, encoding);
  let replacementsMade = 0;

  const actualStartLine = Math.max(0, startLine - 1);
  const actualEndLine = endLine === undefined ? fileLines.length - 1 : Math.min(fileLines.length - 1, endLine - 1);

  if (actualStartLine > actualEndLine) {
    return { replacementsMade: 0 };
  }

  let searchPattern: RegExp;
  if (isRegex) {
    let flags = replaceAll ? 'g' : '';
    if (searchTerm.startsWith('(?i)')) {
      searchTerm = searchTerm.slice(4);
      flags += 'i';
    }
    searchPattern = new RegExp(searchTerm, flags);
  } else {
    searchPattern = new RegExp(escapeRegExp(searchTerm), replaceAll ? 'g' : '');
  }

  let shouldBreak = false;
  const modifiedLines = fileLines.map((line: string, i: number) => {
    if (i < actualStartLine || i > actualEndLine || shouldBreak) {
      return line;
    }

    let thisLineReplacements = 0;
    const modifiedLine = line.replace(searchPattern, (match: string, ...args: unknown[]) => {
      if (!replaceAll && replacementsMade > 0) return match;

      replacementsMade++;
      thisLineReplacements++;

      if (typeof replacement === 'string') {
        return replacement;
      } else if (typeof replacement === 'function') {
        return replacement(match, ...args);
      } else {
        throw new Error('Replacement must be a string or function');
      }
    });

    if (thisLineReplacements > 0 && !replaceAll) {
      shouldBreak = true;
    }

    return modifiedLine;
  });

  if (replacementsMade > 0) {
    await _writeFileLines(path, modifiedLines, encoding, hadTrailingNewline);
  }

  return { replacementsMade };
}

// B. Advanced Line Operations

export interface SortTextLinesArgs {
  path: string;
  options?: {
    reverse?: boolean;
    caseSensitive?: boolean;
    locale?: string;
  };
  encoding?: ValidBufferEncoding;
}
export async function sort_text_lines({ path, options = {}, encoding = 'utf8' }: SortTextLinesArgs): Promise<void> {
  const { reverse = false, caseSensitive = true, locale = 'en' } = options;
  const { lines: fileLines, hadTrailingNewline } = await _readFileLines(path, encoding);

  if (fileLines.length <= 1) {
    return;
  }

  const indexedLines = fileLines.map((line, index) => ({ line, index }));

  const getLineType = (s: string): number => {
    if (/^\d+$/.test(s)) return 0; // Number
    if (/^[A-Z]/.test(s)) return 1; // Uppercase
    if (/^[a-z]/.test(s)) return 2; // Lowercase
    return 3; // Other
  };

  indexedLines.sort((a, b) => {
    let cmp = 0;
    if (caseSensitive) {
      const typeA = getLineType(a.line);
      const typeB = getLineType(b.line);
      if (typeA !== typeB) {
        cmp = typeA - typeB;
      } else {
        // Same type, use localeCompare with numeric and case sensitivity
        // Forcing 'en' as base for predictable ASCII-like behavior for case,
        // then specified locale for finer linguistic points if needed.
        cmp = a.line.localeCompare(b.line, [locale, 'en'], { numeric: true, sensitivity: 'case' });
      }
    } else {
      // Case-insensitive: use localeCompare with sensitivity 'base'
      cmp = a.line.localeCompare(b.line, locale, { numeric: true, sensitivity: 'base' });
    }

    if (cmp === 0) {
      cmp = a.index - b.index; // Stable sort
    }
    return reverse ? -cmp : cmp;
  });
  
  const sortedLines = indexedLines.map(item => item.line);

  // Only write if there are actual changes
  if (JSON.stringify(sortedLines) !== JSON.stringify(fileLines)) {
    await _writeFileLines(path, sortedLines, encoding, hadTrailingNewline);
  }
}

export interface DeduplicateConsecutiveLinesArgs {
  path: string;
  caseSensitive?: boolean;
  encoding?: ValidBufferEncoding;
}
export async function deduplicate_consecutive_lines({ path, caseSensitive = true, encoding = 'utf8' }: DeduplicateConsecutiveLinesArgs): Promise<void> {
  const { lines: fileLines, hadTrailingNewline } = await _readFileLines(path, encoding);
  
  // Early return if no lines or single line
  if (fileLines.length <= 1) {
    return;
  }

  const result: string[] = [];
  let lastLine: string | undefined;

  for (const line of fileLines) {
    let currentLineForComparison = caseSensitive ? line : line.toLowerCase();
    let lastLineForComparison = lastLine === undefined ? undefined : (caseSensitive ? lastLine : lastLine.toLowerCase());

    if (lastLine === undefined || currentLineForComparison !== lastLineForComparison) {
      result.push(line);
      lastLine = line;
    }
  }

  // Only write if we actually removed duplicates
  if (result.length !== fileLines.length) {
    await _writeFileLines(path, result, encoding, hadTrailingNewline);
  }
}

export interface TrimLineWhitespaceArgs {
  path: string;
  options?: {
    leading?: boolean;
    trailing?: boolean;
  };
  encoding?: ValidBufferEncoding;
}
export async function trim_line_whitespace({ path, options = { leading: true, trailing: true }, encoding = 'utf8' }: TrimLineWhitespaceArgs): Promise<void> {
  const { lines: fileLines, hadTrailingNewline } = await _readFileLines(path, encoding);
  const { leading: doLeading = true, trailing: doTrailing = true } = options;

  if (!doLeading && !doTrailing) return;

  // Regex for matching all Unicode whitespace characters
  // Includes space, tab, form feed, line feed, carriage return, and all Unicode spaces
  const wsChar = '[\\s\\p{Zs}\\u00A0\\u1680\\u180E\\u2000-\\u200B\\u2028\\u2029\\u202F\\u205F\\u3000\\uFEFF]';
  const leadingWS = new RegExp(`^${wsChar}+`, 'u');
  const trailingWS = new RegExp(`${wsChar}+$`, 'u');

  let hasChanges = false;
  const trimmedLines = fileLines.map(line => {
    // Handle special case where line is all whitespace
    if (line.match(new RegExp(`^${wsChar}+$`, 'u'))) {
      if (doLeading && doTrailing) {
        hasChanges = true;
        return '';
      }
    }

    let result = line;
    // Trim leading whitespace if enabled
    if (doLeading) {
      const trimmed = line.replace(leadingWS, '');
      if (trimmed !== line) {
        hasChanges = true;
        result = trimmed;
      }
    }

    // Trim trailing whitespace if enabled
    if (doTrailing) {
      const trimmed = result.replace(trailingWS, '');
      if (trimmed !== result) {
        hasChanges = true;
        result = trimmed;
      }
    }

    return result;
  });

  if (hasChanges) {
    await _writeFileLines(path, trimmedLines, encoding, hadTrailingNewline);
  }
}

// C. Binary and Base64 Operations

export interface ReadBinaryAsHexArgs {
  path: string;
}
export async function read_binary_as_hex({ path }: ReadBinaryAsHexArgs): Promise<string> {
  const buffer = await _readBinaryFile(path);
  return buffer.toString('hex');
}

export interface WriteHexAsBinaryArgs {
  path: string;
  hexString: string;
}
export async function write_hex_as_binary({ path, hexString }: WriteHexAsBinaryArgs): Promise<void> {
  // Validate hex string format
  if (!/^[0-9a-fA-F]*$/.test(hexString)) {
    throw new Error('Invalid hex string format');
  }
  // Ensure even length
  if (hexString.length % 2 !== 0) {
    throw new Error('Hex string must have an even number of characters');
  }
  
  // Write both hex string and binary versions
  if (path.endsWith('.hex')) {
    await fs.writeFile(path, hexString, 'ascii');
  } else {
    const buffer = Buffer.from(hexString, 'hex');
    await _writeBinaryFile(path, buffer);
  }
}

export interface ReadBase64FileArgs {
  path: string;
  outputEncoding?: ValidBufferEncoding | 'binary';
}
export async function read_base64_file({ path, outputEncoding = 'utf8' }: ReadBase64FileArgs): Promise<string | Buffer> {
  const base64Content = await fs.readFile(path, 'ascii'); // Base64 is ASCII
  const buffer = Buffer.from(base64Content, 'base64');
  if (outputEncoding === 'binary') {
    return buffer;
  }
  return buffer.toString(outputEncoding);
}

export interface WriteToBase64FileArgs {
  path: string;
  data: string | Buffer;
  inputEncoding?: ValidBufferEncoding;
}
export async function write_to_base64_file({ path, data, inputEncoding = 'utf8' }: WriteToBase64FileArgs): Promise<void> {
  const buffer = Buffer.isBuffer(data) ? data : Buffer.from(data, inputEncoding);
  const base64Content = buffer.toString('base64');
  await fs.writeFile(path, base64Content, 'ascii');
}

// Export all tools for registration
export const fileContentTools = {
  insert_text_lines,
  delete_text_lines,
  replace_text_lines,
  extract_text_from_range,
  search_replace_in_text_range,
  sort_text_lines,
  deduplicate_consecutive_lines,
  trim_line_whitespace,
  read_binary_as_hex,
  write_hex_as_binary,
  read_base64_file,
  write_to_base64_file,
};