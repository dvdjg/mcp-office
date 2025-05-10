const fs = require('fs-extra');
import * as path from 'path';
import {
    insert_text_lines as insertTextLines,
    delete_text_lines as deleteTextLines,
    replace_text_lines as replaceTextLines,
    extract_text_from_range as extractTextFromRange,
    search_replace_in_text_range as searchReplaceInTextRange,
    sort_text_lines as sortTextLines,
    deduplicate_consecutive_lines as deduplicateConsecutiveLines,
    trim_line_whitespace as trimLineWhitespace,
    read_binary_as_hex as readBinaryAsHex,
    write_hex_as_binary as writeHexAsBinary,
    read_base64_file as readBase64File,
    write_to_base64_file as writeToBase64File,
} from '../../../src/tools/fs/fileContent.tool'; // Adjust path as necessary

const fixturesDir = path.join(process.cwd(), 'tests', 'fixtures');
const tempTestDir = path.join(process.cwd(), 'tests', 'temp_file_content_tests');

describe('File Content Tools', () => {
    beforeAll(async () => {
        await fs.ensureDir(tempTestDir);
    });

    afterAll(async () => {
        await fs.remove(tempTestDir);
    });

    beforeEach(async () => {
        // Clean up temp dir before each test, or copy fresh fixtures
        await fs.emptyDir(tempTestDir);
    });

    describe('insertTextLines', () => {
        const sampleFilePath = path.join(fixturesDir, 'sample.txt');
        let testFilePath: string;

        beforeEach(async () => {
            testFilePath = path.join(tempTestDir, 'test_insert.txt');
            await fs.copyFile(sampleFilePath, testFilePath);
        });

        it('should insert lines at the beginning of the file', async () => {
            const linesToInsert = ['Line A', 'Line B'];
            await insertTextLines({ path: testFilePath, lineNumber: 1, lines: linesToInsert });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.split('\n');
            expect(lines[0]).toBe('Line A');
            expect(lines[1]).toBe('Line B');
            expect(lines[2]).toBe('This is the first line.');
        });

        it('should insert lines in the middle of the file', async () => {
            const linesToInsert = ['--Inserted Here--'];
            await insertTextLines({ path: testFilePath, lineNumber: 3, lines: linesToInsert });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.split('\n');
            expect(lines[1]).toBe('This is the second line, with some repetition: line line.');
            expect(lines[2]).toBe('--Inserted Here--');
            expect(lines[3]).toBe('And a third line.');
        });

        it('should append lines to the end of the file (lineNumber 0 or > line count)', async () => {
            const linesToInsert = ['End Line 1', 'End Line 2'];
            await insertTextLines({ path: testFilePath, lineNumber: 0, lines: linesToInsert });
            let content = await fs.readFile(testFilePath, 'utf-8');
            let lines = content.split('\n');
            // Adjusting for potential trailing newline from file read
            const relevantLines = lines.filter(line => line.length > 0);
            expect(relevantLines[relevantLines.length - 2]).toBe('End Line 1');
            expect(relevantLines[relevantLines.length - 1]).toBe('End Line 2');

            // Test with lineNumber > line count
            await fs.copyFile(sampleFilePath, testFilePath); // Reset file
            const initialFileContent = await fs.readFile(sampleFilePath, 'utf-8');
            const initialLinesCount = initialFileContent.split('\n').filter(l => l.length > 0).length;
            await insertTextLines({ path: testFilePath, lineNumber: initialLinesCount + 5, lines: linesToInsert });
            content = await fs.readFile(testFilePath, 'utf-8');
            lines = content.split('\n').filter(line => line.length > 0);
            expect(lines[lines.length - 2]).toBe('End Line 1');
            expect(lines[lines.length - 1]).toBe('End Line 2');
        });

        it('should handle inserting into an empty file', async () => {
            const emptyFilePath = path.join(tempTestDir, 'empty_insert.txt');
            await fs.writeFile(emptyFilePath, '');
            const linesToInsert = ['First line in empty', 'Second line'];
            await insertTextLines({ path: emptyFilePath, lineNumber: 1, lines: linesToInsert });
            const content = await fs.readFile(emptyFilePath, 'utf-8');
            const lines = content.split('\n').filter(l => l.length > 0);
            expect(lines[0]).toBe('First line in empty');
            expect(lines[1]).toBe('Second line');
            expect(lines.length).toBe(2);
        });

        it('should use specified encoding for insertion', async () => {
            const linesToInsert = ['你好世界']; // Chinese characters
            await insertTextLines({ path: testFilePath, lineNumber: 1, lines: linesToInsert, encoding: 'utf16le' });
            const stats = await fs.stat(testFilePath);
            expect(stats.size).toBeGreaterThan(0); // Basic check
        });

        it('should throw an error if file does not exist', async () => {
            await expect(insertTextLines({
                path: path.join(tempTestDir, 'nonexistent.txt'),
                lineNumber: 1,
                lines: ['test']
            })).rejects.toThrow();
        });
    });

    describe('deleteTextLines', () => {
        const sampleFilePath = path.join(fixturesDir, 'sample.txt');
        let testFilePath: string;

        beforeEach(async () => {
            testFilePath = path.join(tempTestDir, 'test_delete.txt');
            await fs.copyFile(sampleFilePath, testFilePath);
        });

        it('should delete lines from the beginning', async () => {
            await deleteTextLines({ path: testFilePath, startLine: 1, endLine: 2 });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.split('\n');
            expect(lines[0]).toBe('And a third line.');
        });

        it('should delete lines from the middle', async () => {
            await deleteTextLines({ path: testFilePath, startLine: 2, endLine: 3 });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.split('\n');
            expect(lines[0]).toBe('This is the first line.');
            expect(lines[1]).toBe('alpha'); // Original 4th line
        });

        it('should delete lines to the end', async () => {
            const originalFileContent = (await fs.readFile(testFilePath, 'utf-8')).replace(/\r\n/g, '\n');
            // Consistent way to get lines, removing a single trailing newline if present, then splitting.
            // This ensures that if the file was "a\n", lines is ["a"], not ["a", ""].
            const getLinesArray = (text: string) => text.replace(/\n$/, '').split('\n').map(line => line.replace(/\r$/, ''));
            
            const originalLines = getLinesArray(originalFileContent);
            const count = originalLines.length;

            const startDelLine = Math.max(1, count - 2); // Delete last 3 lines, or fewer if not enough lines
            const endDelLine = count;

            await deleteTextLines({ path: testFilePath, startLine: startDelLine, endLine: endDelLine });
            
            const finalFileContent = await fs.readFile(testFilePath, 'utf-8');
            const finalLines = getLinesArray(finalFileContent);

            // Calculate expected remaining lines
            // Number of lines actually deleted. Can't be more than 'count'.
            // And can't be more than (endDelLine - startDelLine + 1)
            const numEffectivelyDeleted = Math.min(count, endDelLine) - startDelLine + 1;
            const expectedRemainingCount = Math.max(0, count - numEffectivelyDeleted);
            
            // If the file content becomes empty string after deletion, split('\n') results in ['']
            // So, if expectedRemainingCount is 0, finalLines should be [''] if not empty, or [] if tool makes it empty.
            // The tool _writeFileLines with an empty array and preserveTrailingNewline=true writes "\n".
            // If preserveTrailingNewline=false, it writes "".
            // Our _readFileLines -> normalizeContent for "\n" results in lines: [''], hadTrailingNewline: true.
            // So, if the file is just "\n", getLinesArray("\n") = [''].
            // If the file is "", getLinesArray("") = [''].
            // This means if expectedRemainingCount is 0, finalLines should be [''] if the file isn't truly empty due to a single newline.
            // Or, if the tool makes it truly empty, finalLines would be [''].
            // This logic is tricky. Let's simplify: if expected 0 lines, content should be "" or "\n".
            if (expectedRemainingCount === 0) {
                expect(finalFileContent === '' || finalFileContent === '\n').toBe(true);
            } else {
                expect(finalLines.length).toBe(expectedRemainingCount);
                // If lines remain, check the last one.
                // The line before the deleted block was originalLines[startDelLine - 2]
                if (startDelLine > 1 && expectedRemainingCount > 0) {
                     // originalLines[startDelLine - 2] is the line just before the deleted block
                    expect(finalLines[finalLines.length - 1]).toBe(originalLines[startDelLine - 2]);
                } else if (startDelLine === 1 && expectedRemainingCount > 0 && originalLines.length > numEffectivelyDeleted) {
                    // If we deleted from the start, the new first line should be originalLines[numEffectivelyDeleted]
                     expect(finalLines[0]).toBe(originalLines[numEffectivelyDeleted]);
                }
            }
        });

        it('should handle deleting more lines than exist (delete all)', async () => {
            await deleteTextLines({ path: testFilePath, startLine: 1, endLine: 100 });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content.trim()).toBe(''); // File might have a trailing newline
        });

        it('should handle deleting a single line', async () => {
            await deleteTextLines({ path: testFilePath, startLine: 2, endLine: 2 });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.split('\n');
            expect(lines[0]).toBe('This is the first line.');
            expect(lines[1]).toBe('And a third line.');
        });

        it('should throw error for invalid range (startLine > endLine)', async () => {
            await expect(deleteTextLines({ path: testFilePath, startLine: 3, endLine: 1 })).rejects.toThrow(/Invalid line range/);
        });

        it('should throw error if file does not exist', async () => {
            await expect(deleteTextLines({
                path: path.join(tempTestDir, 'nonexistent.txt'),
                startLine: 1,
                endLine: 1
            })).rejects.toThrow();
        });
    });

    describe('replaceTextLines', () => {
        const sampleFilePath = path.join(fixturesDir, 'sample.txt');
        let testFilePath: string;

        beforeEach(async () => {
            testFilePath = path.join(tempTestDir, 'test_replace.txt');
            await fs.copyFile(sampleFilePath, testFilePath);
        });

        it('should replace lines with the same number of new lines', async () => {
            const replacement = ['Replaced Line 1', 'Replaced Line 2'];
            await replaceTextLines({ path: testFilePath, startLine: 1, endLine: 2, newLines: replacement });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.split('\n');
            expect(lines[0]).toBe('Replaced Line 1');
            expect(lines[1]).toBe('Replaced Line 2');
            expect(lines[2]).toBe('And a third line.');
        });

        it('should replace lines with fewer new lines', async () => {
            const replacement = ['Single Replacement'];
            await replaceTextLines({ path: testFilePath, startLine: 1, endLine: 3, newLines: replacement });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.split('\n');
            expect(lines[0]).toBe('Single Replacement');
            expect(lines[1]).toBe('alpha'); // Original 4th line
        });

        it('should replace lines with more new lines', async () => {
            const replacement = ['New Line A', 'New Line B', 'New Line C'];
            await replaceTextLines({ path: testFilePath, startLine: 2, endLine: 2, newLines: replacement });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.split('\n');
            expect(lines[0]).toBe('This is the first line.');
            expect(lines[1]).toBe('New Line A');
            expect(lines[2]).toBe('New Line B');
            expect(lines[3]).toBe('New Line C');
            expect(lines[4]).toBe('And a third line.'); // Original 3rd line pushed down
        });

        it('should replace all lines if range covers entire file', async () => {
            const replacement = ['File completely replaced.'];
            await replaceTextLines({ path: testFilePath, startLine: 1, endLine: 20, newLines: replacement }); // Assuming sample.txt has < 20 lines
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.split('\n').filter(l => l.length > 0);
            expect(lines[0]).toBe('File completely replaced.');
            expect(lines.length).toBe(1);
        });

        it('should handle replacing in an empty file (effectively inserts)', async () => {
            const emptyFilePath = path.join(tempTestDir, 'empty_replace.txt');
            await fs.writeFile(emptyFilePath, '');
            const replacement = ['Content for empty file'];
            await replaceTextLines({ path: emptyFilePath, startLine: 1, endLine: 1, newLines: replacement });
            const content = await fs.readFile(emptyFilePath, 'utf-8');
            expect(content.replace(/\n$/, '')).toBe('Content for empty file');
        });

        it('should throw error for invalid range (startLine > endLine)', async () => {
            await expect(replaceTextLines({ path: testFilePath, startLine: 3, endLine: 1, newLines: ['test'] })).rejects.toThrow(/Invalid line range/);
        });

        it('should throw error if file does not exist', async () => {
            await expect(replaceTextLines({
                path: path.join(tempTestDir, 'nonexistent.txt'),
                startLine: 1,
                endLine: 1,
                newLines: ['test']
            })).rejects.toThrow();
        });
    });

    describe('extractTextFromRange', () => {
        const sampleFilePath = path.join(fixturesDir, 'sample.txt');
        let testFilePath: string;

        beforeEach(async () => {
            testFilePath = path.join(tempTestDir, 'test_extract.txt');
            await fs.copyFile(sampleFilePath, testFilePath);
        });

        it('should extract a single line', async () => {
            const result = await extractTextFromRange({ path: testFilePath, startLine: 1, endLine: 1, startChar: 0, endChar: Infinity });
            expect(result).toBe('This is the first line.');
        });

        it('should extract multiple lines', async () => {
            const result = await extractTextFromRange({ path: testFilePath, startLine: 1, endLine: 3, startChar: 0, endChar: Infinity });
            expect(result).toBe('This is the first line.\nThis is the second line, with some repetition: line line.\nAnd a third line.');
        });

        it('should extract with character precision (startChar)', async () => {
            const result = await extractTextFromRange({ path: testFilePath, startLine: 1, endLine: 1, startChar: 8, endChar: Infinity }); // "the first line."
            expect(result).toBe('the first line.');
        });

        it('should extract with character precision (endChar)', async () => {
            const result = await extractTextFromRange({ path: testFilePath, startLine: 1, endLine: 1, startChar: 0, endChar: 7 }); // "This is"
            expect(result).toBe('This is');
        });

        it('should extract with character precision (startChar and endChar)', async () => {
            const result = await extractTextFromRange({ path: testFilePath, startLine: 1, endLine: 1, startChar: 5, endChar: 12 }); // "is the "
            expect(result).toBe('is the ');
        });

        it('should handle range exceeding file length (extracts to end)', async () => {
            const originalFileContent = (await fs.readFile(testFilePath, 'utf-8')).replace(/\r\n/g, '\n');
            const result = await extractTextFromRange({ path: testFilePath, startLine: 1, endLine: 100, startChar: 0, endChar: Infinity });
            
            // If the tool is correct, the extracted result for the whole file should be identical to the original read (after normalizing original to \n).
            expect(result).toBe(originalFileContent);
        });

        it('should handle startChar/endChar on multi-line extraction (applied to first/last line of range)', async () => {
            const result = await extractTextFromRange({ path: testFilePath, startLine: 1, endLine: 2, startChar: 5, endChar: 10 });
            expect(result).toBe('is the first line.\nThis is th');
        });

        it('should return empty string for invalid range (startLine > endLine)', async () => {
            const result = await extractTextFromRange({ path: testFilePath, startLine: 3, endLine: 1, startChar: 0, endChar: 0 });
            expect(result).toBe('');
        });
        
        it('should return empty string for startChar > endChar on single line', async () => {
            const result = await extractTextFromRange({ path: testFilePath, startLine: 1, endLine: 1, startChar: 10, endChar: 5 });
            expect(result).toBe('');
        });

        it('should throw error if file does not exist', async () => {
            await expect(extractTextFromRange({
                path: path.join(tempTestDir, 'nonexistent.txt'),
                startLine: 1,
                endLine: 1,
                startChar: 0,
                endChar: 0
            })).rejects.toThrow();
        });
    });

    describe('searchReplaceInTextRange', () => {
        const sampleFilePath = path.join(fixturesDir, 'sample.txt');
        let testFilePath: string;

        beforeEach(async () => {
            testFilePath = path.join(tempTestDir, 'test_search_replace.txt');
            await fs.copyFile(sampleFilePath, testFilePath);
        });

        it('should replace string occurrences globally in the file', async () => {
            await searchReplaceInTextRange({ path: testFilePath, searchTerm: 'line', replacement: 'LINE', replaceAll: true });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content).toContain('first LINE.');
            expect(content).toContain('second LINE, with some repetition: LINE LINE.');
            expect(content).toContain('third LINE.');
            expect(content).toContain('Another LINE for testing.');
        });

        it('should replace regex occurrences globally', async () => {
            await searchReplaceInTextRange({ path: testFilePath, searchTerm: 'l[io]ne', replacement: 'L***E', isRegex: true, replaceAll: true });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content).toContain('first L***E.');
            expect(content).toContain('second L***E, with some repetition: L***E L***E.');
            expect(content).toContain('third L***E.');
        });

        it('should replace only the first occurrence if replaceAll is false', async () => {
            await searchReplaceInTextRange({ path: testFilePath, searchTerm: 'line', replacement: 'LINE', replaceAll: false });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content).toContain('first LINE.'); // Replaced
            expect(content).toContain('second line, with some repetition: line line.'); // Not replaced
        });

        it('should perform case-insensitive search (Note: tool itself needs to implement ignoreCase for regex)', async () => {
            await fs.copyFile(path.join(fixturesDir, 'sample.txt'), testFilePath); // reset
            await searchReplaceInTextRange({ path: testFilePath, searchTerm: '(?i)case test', replacement: 'REPLACED', isRegex: true, replaceAll: true });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content).toContain('REPLACED'); // For 'CASE TEST'
            expect(content).toContain('REPLACED'); // For 'case test'
        });
        
        it('should replace within a specified line range', async () => {
            await searchReplaceInTextRange({ path: testFilePath, searchTerm: 'line', replacement: 'LINE', startLine: 2, endLine: 2, replaceAll: true });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.split('\n');
            expect(lines[0]).toBe('This is the first line.'); // Unchanged
            expect(lines[1]).toBe('This is the second LINE, with some repetition: LINE LINE.'); // Changed
            expect(lines[2]).toBe('And a third line.'); // Unchanged
        });

        it('should handle no matches found', async () => {
            const originalContent = await fs.readFile(testFilePath, 'utf-8');
            await searchReplaceInTextRange({ path: testFilePath, searchTerm: 'nonexistent_term', replacement: 'REPLACED' });
            const newContent = await fs.readFile(testFilePath, 'utf-8');
            expect(newContent).toBe(originalContent);
        });

        it('should throw error if file does not exist', async () => {
            await expect(searchReplaceInTextRange({
                path: path.join(tempTestDir, 'nonexistent.txt'),
                searchTerm: 'a',
                replacement: 'b'
            })).rejects.toThrow();
        });
    });

    describe('sortTextLines', () => {
        const sampleFilePath = path.join(fixturesDir, 'sample.txt');
        let testFilePath: string;

        beforeEach(async () => {
            testFilePath = path.join(tempTestDir, 'test_sort.txt');
            const contentToSort = "gamma\nalpha\nbeta\n123\nXYZ\nabc";
            await fs.writeFile(testFilePath, contentToSort);
        });

        it('should sort lines in ascending order by default', async () => {
            await sortTextLines({ path: testFilePath });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.replace(/\n$/, '').split('\n');
            expect(lines).toEqual(['123', 'XYZ', 'abc', 'alpha', 'beta', 'gamma']);
        });

        it('should sort lines in descending order', async () => {
            await sortTextLines({ path: testFilePath, options: { reverse: true } });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const lines = content.replace(/\n$/, '').split('\n');
            expect(lines).toEqual(['gamma', 'beta', 'alpha', 'abc', 'XYZ', '123']);
        });

        it('should sort case-sensitively by default (or if specified)', async () => {
            const caseSensitiveContent = "apple\nBanana\nApple";
            await fs.writeFile(testFilePath, caseSensitiveContent);
            await sortTextLines({ path: testFilePath, options: { caseSensitive: true } });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content.replace(/\n$/, '').split('\n')).toEqual(['Apple', 'Banana', 'apple']);
        });

        it('should sort case-insensitively', async () => {
            const caseSensitiveContent = "apple\nBanana\nApple";
            await fs.writeFile(testFilePath, caseSensitiveContent);
            await sortTextLines({ path: testFilePath, options: { caseSensitive: false } });
            const content = await fs.readFile(testFilePath, 'utf-8');
            const sortedLower = content.replace(/\n$/, '').split('\n').map(s => s.toLowerCase());
            expect(sortedLower).toEqual(['apple', 'apple', 'banana']);
        });
        
        it('should throw error if file does not exist', async () => {
            await expect(sortTextLines({ path: path.join(tempTestDir, 'nonexistent.txt') })).rejects.toThrow();
        });
    });

    describe('deduplicateConsecutiveLines', () => {
        let testFilePath: string;

        beforeEach(() => {
            testFilePath = path.join(tempTestDir, 'test_dedupe.txt');
        });

        it('should remove consecutive duplicate lines (case-sensitive by default)', async () => {
            const contentWithDupes = "a\nb\nb\nc\nc\nc\na\na";
            await fs.writeFile(testFilePath, contentWithDupes);
            await deduplicateConsecutiveLines({ path: testFilePath });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content.replace(/\n$/, '')).toBe("a\nb\nc\na");
        });

        it('should remove consecutive duplicate lines (case-insensitive)', async () => {
            const contentWithDupes = "a\nA\nB\nb\nc\nC\nC";
            await fs.writeFile(testFilePath, contentWithDupes);
            await deduplicateConsecutiveLines({ path: testFilePath, caseSensitive: false });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content.replace(/\n$/, '')).toBe("a\nB\nc");
        });

        it('should not change file if no consecutive duplicates', async () => {
            const contentNoDupes = "a\nb\nc\na\nb";
            await fs.writeFile(testFilePath, contentNoDupes);
            await deduplicateConsecutiveLines({ path: testFilePath });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content.replace(/\n$/, '')).toBe(contentNoDupes);
        });

        it('should handle empty file', async () => {
            await fs.writeFile(testFilePath, "");
            await deduplicateConsecutiveLines({ path: testFilePath });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content).toBe("");
        });

        it('should throw error if file does not exist', async () => {
            await expect(deduplicateConsecutiveLines({ path: path.join(tempTestDir, 'nonexistent.txt') })).rejects.toThrow();
        });
    });

    describe('trimLineWhitespace', () => {
        let testFilePath: string;

        beforeEach(() => {
            testFilePath = path.join(tempTestDir, 'test_trim.txt');
        });

        it('should trim leading and trailing whitespace from all lines by default', async () => {
            const contentToTrim = "  line 1  \n\tline 2\nline 3    ";
            await fs.writeFile(testFilePath, contentToTrim);
            await trimLineWhitespace({ path: testFilePath });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content.replace(/\n$/, '')).toBe("line 1\nline 2\nline 3");
        });

        it('should trim only leading whitespace', async () => {
            const contentToTrim = "  line 1  \n\tline 2\n  line 3    ";
            await fs.writeFile(testFilePath, contentToTrim);
            await trimLineWhitespace({ path: testFilePath, options: { leading: true, trailing: false } });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content.replace(/\n$/, '')).toBe("line 1  \nline 2\nline 3    ");
        });

        it('should trim only trailing whitespace', async () => {
            const contentToTrim = "  line 1  \n\tline 2\n  line 3    ";
            await fs.writeFile(testFilePath, contentToTrim);
            await trimLineWhitespace({ path: testFilePath, options: { leading: false, trailing: true } });
            const content = await fs.readFile(testFilePath, 'utf-8');
            expect(content.replace(/\n$/, '')).toBe("  line 1\n\tline 2\n  line 3");
        });

        it('should not change file if no whitespace to trim or options are false', async () => {
            const contentClean = "line 1\nline 2";
            await fs.writeFile(testFilePath, contentClean);
            await trimLineWhitespace({ path: testFilePath }); // Default trims both
            let content = await fs.readFile(testFilePath, 'utf-8');
            expect(content.replace(/\n$/, '')).toBe(contentClean);

            const contentToTrim = "  line 1  ";
            await fs.writeFile(testFilePath, contentToTrim);
            await trimLineWhitespace({ path: testFilePath, options: { leading: false, trailing: false } });
            content = await fs.readFile(testFilePath, 'utf-8');
            expect(content.replace(/\n$/, '')).toBe(contentToTrim);
        });

        it('should throw error if file does not exist', async () => {
            await expect(trimLineWhitespace({ path: path.join(tempTestDir, 'nonexistent.txt') })).rejects.toThrow();
        });
    });

    describe('Binary Hex Operations', () => {
        const binaryFilePath = path.join(fixturesDir, 'binary_data.bin'); // Contains "test"
        const testHexFilePath = path.join(tempTestDir, 'test.hex');
        const testBinFilePath = path.join(tempTestDir, 'test_from_hex.bin');
        const expectedHex = '74657374'; // "test" in hex

        it('readBinaryAsHex should convert binary file to hex string', async () => {
            const hexString = await readBinaryAsHex({ path: binaryFilePath });
            expect(hexString).toBe(expectedHex);
        });

        it('writeHexAsBinary should write hex string to binary file', async () => {
            await writeHexAsBinary({ path: testHexFilePath, hexString: expectedHex });
            const fileContent = await fs.readFile(testHexFilePath, 'utf-8'); 
            expect(fileContent).toBe(expectedHex); 
            
            await writeHexAsBinary({ path: testBinFilePath, hexString: expectedHex });
            const binaryContent = await fs.readFile(testBinFilePath);
            expect(binaryContent.toString('utf-8')).toBe('test');
        });

        it('readBinaryAsHex and writeHexAsBinary round trip', async () => {
            const originalContent = await fs.readFile(binaryFilePath);
            const hex = await readBinaryAsHex({ path: binaryFilePath });
            await writeHexAsBinary({ path: testBinFilePath, hexString: hex });
            const newContent = await fs.readFile(testBinFilePath);
            expect(newContent).toEqual(originalContent);
        });

        it('readBinaryAsHex should throw if file not found', async () => {
            await expect(readBinaryAsHex({ path: 'nonexistent.bin' })).rejects.toThrow();
        });

        it('writeHexAsBinary should throw for invalid hex string', async () => {
            await expect(writeHexAsBinary({ path: testBinFilePath, hexString: 'invalidhex' })).rejects.toThrow();
        });
    });

    describe('Base64 File Operations', () => {
        const textFilePath = path.join(fixturesDir, 'sample.txt'); 
        const testBase64FilePath = path.join(tempTestDir, 'test.b64');
        const testRestoredFilePath = path.join(tempTestDir, 'test_from_b64.txt');

        it('readBase64File should read and decode a base64 encoded file', async () => {
            const originalContent = await fs.readFile(textFilePath, 'utf-8');
            const base64Content = Buffer.from(originalContent, 'utf-8').toString('base64');
            await fs.writeFile(testBase64FilePath, base64Content, 'ascii'); 

            const decodedContent = await readBase64File({ path: testBase64FilePath, outputEncoding: 'utf-8' });
            expect(decodedContent).toBe(originalContent);
        });

        it('writeToBase64File should encode content and write to a file', async () => {
            const originalContent = "Hello Base64 World!";
            await writeToBase64File({ path: testBase64FilePath, data: originalContent, inputEncoding: 'utf-8' });
            
            const writtenBase64 = await fs.readFile(testBase64FilePath, 'ascii');
            const expectedBase64 = Buffer.from(originalContent, 'utf-8').toString('base64');
            expect(writtenBase64).toBe(expectedBase64);

            const decodedWrittenContent = Buffer.from(writtenBase64, 'base64').toString('utf-8');
            expect(decodedWrittenContent).toBe(originalContent);
        });

        it('readBase64File and writeToBase64File round trip', async () => {
            const originalContent = await fs.readFile(textFilePath, 'utf-8');
            
            await writeToBase64File({ path: testBase64FilePath, data: originalContent, inputEncoding: 'utf-8' });
            
            const restoredContent = await readBase64File({ path: testBase64FilePath, outputEncoding: 'utf-8' });
            
            expect(restoredContent).toBe(originalContent);
        });
        
        it('readBase64File should handle different target encodings', async () => {
            const originalContent = "你好世界"; // Chinese characters
            const base64Content = Buffer.from(originalContent, 'utf16le').toString('base64');
            await fs.writeFile(testBase64FilePath, base64Content, 'ascii');

            const decodedContent = await readBase64File({ path: testBase64FilePath, outputEncoding: 'utf16le' });
            expect(decodedContent).toBe(originalContent);
        });

        it('writeToBase64File should handle different content encodings', async () => {
            const originalContent = "你好世界"; // Chinese characters
            await writeToBase64File({ path: testBase64FilePath, data: originalContent, inputEncoding: 'utf16le' });
            
            const writtenBase64 = await fs.readFile(testBase64FilePath, 'ascii');
            const expectedBase64 = Buffer.from(originalContent, 'utf16le').toString('base64');
            expect(writtenBase64).toBe(expectedBase64);
        });

        it('readBase64File should throw if file not found', async () => {
            await expect(readBase64File({ path: 'nonexistent.b64', outputEncoding: 'utf-8' })).rejects.toThrow();
        });
    });
});