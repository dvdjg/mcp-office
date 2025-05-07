// =============================================================================
/**
 * @file Unit tests for security utility functions.
 */
import path from 'path';
import { validateFilePath } from '../../../src/utils/security.js'; // Use relative path and .js extension

// Mock ALLOWED_BASE_PATHS for testing if needed, or configure via environment variables for tests
const testAllowedBase = path.resolve('./data/test_files');

describe('Security Utils', () => {
    describe('validateFilePath', () => {
        const allowedPaths = [testAllowedBase]; // Use a dedicated test path

        it('should allow paths within the allowed base directory', () => {
            const validPath = path.join(testAllowedBase, 'document.docx');
            // Pass allowedPaths as the second argument
            expect(validateFilePath(validPath, allowedPaths)).toBe(path.resolve(validPath));
        });

        it('should allow paths in subdirectories of the allowed base', () => {
            const validPath = path.join(testAllowedBase, 'subdir', 'image.png');
             // We might need to create the subdir for the test if existence check is enabled
            // Pass allowedPaths as the second argument
            expect(validateFilePath(validPath, allowedPaths)).toBe(path.resolve(validPath));
        });

        it('should disallow paths outside the allowed base directory', () => {
            const invalidPath = path.resolve('/etc/passwd');
            // Pass allowedPaths as the second argument
            expect(() => validateFilePath(invalidPath, allowedPaths)).toThrow('Access denied');
        });

        it('should disallow directory traversal attempts (../)', () => {
            const traversalPath = path.join(testAllowedBase, '..', 'some_other_dir', 'file.txt');
            // Pass allowedPaths as the second argument
            expect(() => validateFilePath(traversalPath, allowedPaths)).toThrow('Access denied');
        });

         it('should disallow directory traversal attempts (absolute path with ..)', () => {
            const traversalPath = path.resolve(testAllowedBase, '../secret_file');
            // Pass allowedPaths as the second argument
            expect(() => validateFilePath(traversalPath, allowedPaths)).toThrow('Access denied');
        });

        it('should handle relative paths correctly', () => {
            // Assuming the test runs from the project root where './data/test_files' resolves correctly
            const relativeValidPath = './data/test_files/relative.txt';
            // Pass allowedPaths as the second argument
            expect(validateFilePath(relativeValidPath, allowedPaths)).toBe(path.resolve(relativeValidPath));

            const relativeInvalidPath = './src/server/index.ts'; // Assuming src is outside allowed test path
             // Pass allowedPaths as the second argument
             expect(() => validateFilePath(relativeInvalidPath, allowedPaths)).toThrow('Access denied');
        });
    });
});

// Add more tests for other utilities and tools...