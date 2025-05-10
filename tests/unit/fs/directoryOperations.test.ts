const fs = require('fs-extra');
import * as path from 'path';
import {
    listDirectoryContents as list_directory_contents,
    findFiles as find_files,
    getDirectoryTree as get_directory_tree,
} from '../../../src/tools/fs/directoryOperations.tool'; // Adjust path as necessary

const fixturesDir = path.join(process.cwd(), 'tests', 'fixtures');
const tempTestDir = path.join(process.cwd(), 'tests', 'temp_dir_ops_tests');

describe('Directory Operations Tools', () => {
    beforeAll(async () => {
        await fs.ensureDir(tempTestDir);
    });

    afterAll(async () => {
        await fs.remove(tempTestDir);
    });

    beforeEach(async () => {
        await fs.emptyDir(tempTestDir);
        // Setup a predictable directory structure for testing
        await fs.ensureDir(path.join(tempTestDir, 'dir1'));
        await fs.ensureDir(path.join(tempTestDir, 'dir1', 'subdir1_1'));
        await fs.ensureDir(path.join(tempTestDir, 'dir2'));
        await fs.writeFile(path.join(tempTestDir, 'file1.txt'), 'content1');
        await fs.writeFile(path.join(tempTestDir, 'dir1', 'file1_1.txt'), 'content_file_1_1');
        await fs.writeFile(path.join(tempTestDir, 'dir1', 'subdir1_1', 'file1_1_1.log'), 'log content');
        await fs.writeFile(path.join(tempTestDir, 'dir2', 'file2_1.md'), 'markdown content');
        await fs.ensureDir(path.join(tempTestDir, 'empty_dir'));
    });

    describe('list_directory_contents', () => {
        it('should list top-level contents of a directory', async () => {
            const result = await list_directory_contents({ path: tempTestDir });
            // Expected: dir1, dir2, empty_dir, file1.txt
            expect(result.length).toBe(4);
            expect(result).toEqual(expect.arrayContaining([
                expect.objectContaining({ name: 'dir1', type: 'directory' }),
                expect.objectContaining({ name: 'dir2', type: 'directory' }),
                expect.objectContaining({ name: 'empty_dir', type: 'directory' }),
                expect.objectContaining({ name: 'file1.txt', type: 'file' }),
            ]));
        });

        it('should list contents recursively', async () => {
            const result = await list_directory_contents({ path: tempTestDir, options: { recursive: true, includeSize: false } });
            // Expected: dir1, dir1/subdir1_1, dir1/subdir1_1/file1_1_1.log, dir1/file1_1.txt,
            // dir2, dir2/file2_1.md, empty_dir, file1.txt
            expect(result.length).toBe(8); // 4 dirs, 4 files
            expect(result).toEqual(expect.arrayContaining([
                // Files
                expect.objectContaining({ name: 'file1.txt', type: 'file', depth: 0, path: path.join(tempTestDir, 'file1.txt') }),
                expect.objectContaining({ name: 'file1_1.txt', type: 'file', depth: 1, path: path.join(tempTestDir, 'dir1', 'file1_1.txt') }),
                expect.objectContaining({ name: 'file1_1_1.log', type: 'file', depth: 2, path: path.join(tempTestDir, 'dir1', 'subdir1_1', 'file1_1_1.log') }),
                expect.objectContaining({ name: 'file2_1.md', type: 'file', depth: 1, path: path.join(tempTestDir, 'dir2', 'file2_1.md') }),
                // Directories
                expect.objectContaining({ name: 'dir1', type: 'directory', depth: 0, path: path.join(tempTestDir, 'dir1') }),
                expect.objectContaining({ name: 'subdir1_1', type: 'directory', depth: 1, path: path.join(tempTestDir, 'dir1', 'subdir1_1') }),
                expect.objectContaining({ name: 'dir2', type: 'directory', depth: 0, path: path.join(tempTestDir, 'dir2') }),
                expect.objectContaining({ name: 'empty_dir', type: 'directory', depth: 0, path: path.join(tempTestDir, 'empty_dir') }),
            ]));
        });

        it('should include size if requested', async () => {
            const result = await list_directory_contents({ path: tempTestDir, options: { includeSize: true, recursive: false } });
            const file1 = result.find(f => f.name === 'file1.txt');
            expect(file1).toBeDefined();
            expect(file1?.size).toBeGreaterThan(0); // "content1"
        });

        it('should respect maxDepth for recursion', async () => {
            // Test intent: maxDepth 0 (code interpretation) should yield top-level items only.
            // Test comment "maxDepth 1 ... (same as non-recursive)" implies the test's "maxDepth 1" means code's "maxDepth 0".
            let result = await list_directory_contents({ path: tempTestDir, options: { recursive: true, maxDepth: 0, includeSize: false } });
            expect(result.length).toBe(4); // dir1, dir2, empty_dir, file1.txt

            // Test intent: maxDepth 1 (code interpretation) should yield top-level and one level down.
            // Test comment "maxDepth 2: dir1, dir1/subdir1_1, dir1/file1_1.txt, dir2, dir2/file2_1.md, empty_dir, file1.txt"
            // These are items at depth 0 and 1. This means the test's "maxDepth 2" means code's "maxDepth 1".
            result = await list_directory_contents({ path: tempTestDir, options: { recursive: true, maxDepth: 1, includeSize: false } });
            expect(result.length).toBe(7); // 4 items at depth 0, 3 items at depth 1
             expect(result).toEqual(expect.arrayContaining([
                expect.objectContaining({ name: 'subdir1_1', path: expect.stringContaining('dir1'), type: 'directory' }),
                expect.objectContaining({ name: 'file1_1.txt', path: expect.stringContaining('dir1'), type: 'file' }),
                expect.objectContaining({ name: 'file2_1.md', path: expect.stringContaining('dir2'), type: 'file' }),
            ]));
            // file1_1_1.log should NOT be present as it's at depth 3
            expect(result.find(f => f.name === 'file1_1_1.log')).toBeUndefined();
        });

        it('should return empty array for an empty directory', async () => {
            const result = await list_directory_contents({ path: path.join(tempTestDir, 'empty_dir') });
            expect(result).toEqual([]);
        });

        it('should throw error for non-existent path', async () => {
            await expect(list_directory_contents({ path: path.join(tempTestDir, 'non_existent_dir') })).rejects.toThrow();
        });
    });

    describe('find_files', () => {
        it('should find files by glob name pattern (non-recursive by default if recursive not specified in options)', async () => {
            const result = await find_files({ directoryPath: tempTestDir, namePattern: '*.txt', options: { recursive: false, isRegexName: false, contentIsRegex: false, encoding: 'utf8' } });
            expect(result.length).toBe(1);
            expect(result[0].path).toEqual(expect.stringMatching(/file1\.txt$/));
        });

        it('should find files by glob name pattern (recursive)', async () => {
            const result = await find_files({ directoryPath: tempTestDir, namePattern: '*.txt', options: { recursive: true, isRegexName: false, contentIsRegex: false, encoding: 'utf8' } });
            expect(result.length).toBe(2);
            expect(result.map(r => r.path)).toEqual(expect.arrayContaining([
                expect.stringMatching(/file1\.txt$/),
                expect.stringMatching(/file1_1\.txt$/),
            ]));
        });

        it('should find files by regex name pattern', async () => {
            const result = await find_files({ directoryPath: tempTestDir, namePattern: 'file\\d_\\d.*', options: { isRegexName: true, recursive: true, contentIsRegex: false, encoding: 'utf8' } });
            expect(result.length).toBe(3); // file1_1.txt, file1_1_1.log, file2_1.md
            expect(result.map(r => r.path)).toEqual(expect.arrayContaining([
                expect.stringMatching(/file1_1\.txt$/),
                expect.stringMatching(/file2_1\.md$/),
                expect.stringMatching(/file1_1_1\.log$/),
            ]));
        });

        it('should find files by string content search (recursive)', async () => {
            const result = await find_files({ directoryPath: tempTestDir, namePattern: '*', options: { contentPattern: 'content1', recursive: true, isRegexName: false, contentIsRegex: false, encoding: 'utf8' } });
            // Tool returns all files matching namePattern, matchesContent indicates content match.
            // All 4 files in the fixture match namePattern '*'.
            expect(result.length).toBe(4);
            const matchedFile = result.find(r => r.matchesContent === true);
            expect(matchedFile).toBeDefined();
            expect(matchedFile?.path).toEqual(expect.stringMatching(/file1\.txt$/));
            // Verify others didn't match content
            const otherFiles = result.filter(r => r.path !== matchedFile?.path);
            otherFiles.forEach(f => expect(f.matchesContent).toBe(false));
        });

        it('should find files by regex content search (recursive)', async () => {
            const result = await find_files({ directoryPath: tempTestDir, namePattern: '*', options: { contentPattern: 'log\\scontent', contentIsRegex: true, recursive: true, isRegexName: false, encoding: 'utf8' } });
            expect(result.length).toBe(4); // All 4 files match namePattern '*'
            const matchedFile = result.find(r => r.matchesContent === true);
            expect(matchedFile).toBeDefined();
            expect(matchedFile?.path).toEqual(expect.stringMatching(/file1_1_1\.log$/));
            const otherFiles = result.filter(r => r.path !== matchedFile?.path);
            otherFiles.forEach(f => expect(f.matchesContent).toBe(false));
        });

        it('should combine name pattern and content search', async () => {
            const result = await find_files({
                directoryPath: tempTestDir,
                namePattern: '*.txt',
                options: {
                    contentPattern: 'content_file',
                    recursive: true,
                    isRegexName: false,
                    contentIsRegex: false,
                    encoding: 'utf8'
                }
            });
            // Files matching namePattern '*.txt': file1.txt, file1_1.txt
            expect(result.length).toBe(2);
            const matchedFile = result.find(r => r.matchesContent === true);
            expect(matchedFile).toBeDefined();
            expect(matchedFile?.path).toEqual(expect.stringMatching(/file1_1\.txt$/));

            const nonMatchedFile = result.find(r => r.matchesContent === false);
            expect(nonMatchedFile).toBeDefined();
            expect(nonMatchedFile?.path).toEqual(expect.stringMatching(/file1\.txt$/));
        });

        it('should respect maxDepth for recursion', async () => {
            // Test wants items at depth 0 and 1, excluding depth 2. This is code's maxDepth: 1.
            const result = await find_files({ directoryPath: tempTestDir, namePattern: '*.*', options: { recursive: true, maxDepth: 1, isRegexName: false, contentIsRegex: false, encoding: 'utf8' } });
            // Expected: file1.txt (depth 0), file1_1.txt (depth 1), file2_1.md (depth 1)
            expect(result.length).toBe(3);
            // file1_1_1.log is at depth 2, so it should not be found with maxDepth: 1
            expect(result.find(f => f.path.endsWith('file1_1_1.log'))).toBeUndefined();
        });
        
        it('should return empty array if no files match', async () => {
            const result = await find_files({ directoryPath: tempTestDir, namePattern: '*.nonexistent' });
            expect(result).toEqual([]);
        });

        it('should throw error for non-existent directory path', async () => {
            await expect(find_files({ directoryPath: path.join(tempTestDir, 'non_existent_dir'), namePattern: '*' })).rejects.toThrow();
        });
    });

    describe('get_directory_tree', () => {
        it('should get basic directory tree string', async () => {
            const tree = await get_directory_tree({ path: tempTestDir });
            expect(tree).toContain(path.basename(tempTestDir) + '/');
            expect(tree).toContain('dir1/');
            expect(tree).toContain('file1.txt');
            // Check structure for a couple of items
            expect(tree).toMatch(/├── dir1\//);
            expect(tree).toMatch(/└── file1\.txt/); // Assuming file1.txt is last or near last at top level based on typical sort
        });

        it('should respect maxDepth', async () => {
            // Test's "maxDepth: 1" (show 1 level) corresponds to code's maxDepth: 0 (show items at depth 0)
            const treeDepth0 = await get_directory_tree({ path: tempTestDir, options: { maxDepth: 0, includeFiles: true, includeSize: false } });
            expect(treeDepth0).toContain('dir1/'); // dir1/ is at depth 0
            expect(treeDepth0).not.toContain('subdir1_1/'); // subdir1_1/ is at depth 1
            expect(treeDepth0).not.toContain('file1_1.txt'); // file1_1.txt is at depth 1

            // Test's "maxDepth: 2" (show 2 levels) corresponds to code's maxDepth: 1 (show items at depth 0 and 1)
            const treeDepth1 = await get_directory_tree({ path: tempTestDir, options: { maxDepth: 1, includeFiles: true, includeSize: false } });
            expect(treeDepth1).toContain('subdir1_1/'); // subdir1_1/ is at depth 1
            expect(treeDepth1).toContain('file1_1.txt'); // file1_1.txt is at depth 1
            expect(treeDepth1).not.toContain('file1_1_1.log'); // file1_1_1.log is at depth 2
        });

        it('should exclude files if specified', async () => {
            const tree = await get_directory_tree({ path: tempTestDir, options: { includeFiles: false, maxDepth: 3, includeSize: false } });
            expect(tree).not.toContain('file1.txt');
            expect(tree).toContain('dir1/');
            expect(tree).toContain('dir2/');
            expect(tree).toContain('empty_dir/');
        });

        it('should include size if requested', async () => {
            const tree = await get_directory_tree({ path: tempTestDir, options: { includeSize: true, includeFiles: true, maxDepth: 3 } });
            expect(tree).toMatch(/file1\.txt \((\d+) bytes\)/);
            // Directory sizes are not part of the string output per current tool implementation, only file sizes
        });

        it('should handle empty directory', async () => {
            const tree = await get_directory_tree({ path: path.join(tempTestDir, 'empty_dir') });
            expect(tree).toBe('empty_dir/\n'); // Root dir name + newline
        });

        it('should throw error for non-existent path', async () => {
            await expect(get_directory_tree({ path: path.join(tempTestDir, 'non_existent_dir') })).rejects.toThrow();
        });
    });
});