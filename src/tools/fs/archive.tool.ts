/**
 * @file Implements tools for handling archive files (ZIP, 7z, etc.) using Node.js libraries and child processes.
 * @author David Jurado
 * @date 2025-05-04
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import { z } from 'zod';
import AdmZip from 'adm-zip';
import * as path from 'path';
import * as fs from 'fs-extra';
// Assuming list and extractFull are available exports from node-7z
import { extractFull, list } from 'node-7z';
import { spawn } from 'child_process'; // Use Node.js child_process for command execution

// Helper function to determine archive type based on file extension
function getArchiveType(filePath: string): 'zip' | '7z' | 'unknown' {
    const ext = path.extname(filePath).toLowerCase();
    if (ext === '.zip') {
        return 'zip';
    }
    if (ext === '.7z') {
        return '7z';
    }
    // Add more formats here if needed for read-only support
    return 'unknown';
}

/** Schema for the input parameters of the 'fs/archive/list' tool. */
const archiveListInputSchema = z.object({
    /** The path to the archive file. */
    filePath: z.string().describe('The path to the archive file.'),
});

type ArchiveListInput = z.infer<typeof archiveListInputSchema>;

/**
 * Lists the contents of an archive file.
 * @param input - The input parameters.
 * @returns A promise resolving to an object indicating success and the list of contents.
 */
const listArchiveContents = async (input: ArchiveListInput) => {
    const { filePath } = input;
    const archiveType = getArchiveType(filePath);

    try {
        let contents: string[] = [];
        if (archiveType === 'zip') {
            const zip = new AdmZip(filePath);
            contents = zip.getEntries().map(entry => entry.entryName);
        } else if (archiveType === '7z') {
            // Use node-7z list command
            // This might require the 7z executable in the system's PATH
            // The list method typically returns a stream or a promise
            const stream = list(filePath);

            contents = await new Promise<string[]>((resolve, reject) => {
                const items: string[] = [];
                stream.on('data', (data) => {
                    // Check if data is a string or an object with a 'name' property
                    if (typeof data === 'string') {
                         // Type assertion needed here as TypeScript's control flow
                         // doesn't fully narrow down 'data' to string in this branch.
                         items.push((data as string).trim());
                    } else if (data && typeof data === 'object' && 'name' in data) {
                        // Assuming data objects have a 'name' property
                        items.push((data as any).name); // Use 'any' to access name property
                    }
                });
                stream.on('end', () => resolve(items));
                stream.on('error', (err) => reject(err));
            });

        } else {
            throw new Error(`Unsupported archive format for listing: ${archiveType}`);
        }

        return { success: true, contents };
    } catch (error: any) {
        return { success: false, error: error.message };
    }
};

/** Schema for the input parameters of the 'fs/archive/extract' tool. */
const archiveExtractInputSchema = z.object({
    /** The path to the archive file. */
    filePath: z.string().describe('The path to the archive file.'),
    /** The target file or directory within the archive to extract (optional). */
    target: z.string().optional().describe('The target file or directory within the archive to extract.'),
    /** The directory to extract the contents to. */
    outputDirectory: z.string().describe('The directory to extract the contents to.'),
});

type ArchiveExtractInput = z.infer<typeof archiveExtractInputSchema>;

/**
 * Extracts contents from an archive file.
 * @param input - The input parameters.
 * @returns A promise resolving to an object indicating success.
 */
const extractArchive = async (input: ArchiveExtractInput) => {
    const { filePath, target, outputDirectory } = input;
    const archiveType = getArchiveType(filePath);

    try {
        await fs.ensureDir(outputDirectory); // Ensure output directory exists

        if (archiveType === 'zip') {
            const zip = new AdmZip(filePath);
            if (target) {
                // AdmZip does not have a direct method to extract a single file/directory easily by path
                // Need to find the entry and extract it. This is a simplified approach.
                const entry = zip.getEntry(target);
                if (entry) {
                     zip.extractEntryTo(entry, outputDirectory, false, true);
                } else {
                     throw new Error(`Target "${target}" not found in the ZIP archive.`);
                }
            } else {
                zip.extractAllTo(outputDirectory, true);
            }
        } else if (archiveType === '7z') {
            // node-7z extraction
            // This might require the 7z executable in the system's PATH
            const stream = extractFull(filePath, outputDirectory, {
                 // Options might be needed here, check node-7z documentation
                 // include: target ? [target] : undefined, // Include specific target if provided
            });

            await new Promise((resolve, reject) => {
                stream.on('end', () => resolve({}));
                stream.on('error', (err) => reject(err));
            });

        } else {
            throw new Error(`Unsupported archive format for extraction: ${archiveType}`);
        }

        return { success: true };
    } catch (error: any) {
        return { success: false, error: error.message };
    }
};

/** Schema for the input parameters of the 'fs/archive/create' tool. */
const archiveCreateInputSchema = z.object({
    /** The path for the output archive file. */
    outputFilePath: z.string().describe('The path for the output archive file.'),
    /** A list of file or directory paths to include in the archive. */
    sourcePaths: z.array(z.string()).describe('A list of file or directory paths to include in the archive.'),
    /** The format of the archive to create ('zip' or '7z'). Defaults to 'zip'. */
    format: z.enum(['zip', '7z']).default('zip').describe('The format of the archive to create (zip or 7z).'),
});

type ArchiveCreateInput = z.infer<typeof archiveCreateInputSchema>;

/**
 * Creates an archive file from a list of source paths.
 * @param input - The input parameters.
 * @returns A promise resolving to an object indicating success.
 */
const createArchive = async (input: ArchiveCreateInput) => {
    const { outputFilePath, sourcePaths, format } = input;

    try {
        if (format === 'zip') {
            const zip = new AdmZip();
            sourcePaths.forEach(sourcePath => {
                if (fs.statSync(sourcePath).isDirectory()) {
                    zip.addLocalFolder(sourcePath, path.basename(sourcePath));
                } else {
                    zip.addLocalFile(sourcePath);
                }
            });
            zip.writeZip(outputFilePath);
        } else if (format === '7z') {
            // Create 7z archive using the 7z command-line tool via child_process
            // This requires the 7z executable in the system's PATH
            // The syntax for adding multiple files/directories is '7z a <archive_name> <file1> <dir1> ...'
            const args = ['a', outputFilePath, ...sourcePaths];
            console.warn(`Executing 7z command: 7z ${args.join(' ')}`);

            const child = spawn('7z', args);

            let stdout = '';
            let stderr = '';

            child.stdout.on('data', (data) => {
                stdout += data;
            });

            child.stderr.on('data', (data) => {
                stderr += data;
            });

            await new Promise<void>((resolve, reject) => {
                child.on('close', (code) => {
                    if (code !== 0) {
                        reject(new Error(`7z command failed with code ${code}. Stderr: ${stderr}`));
                    } else {
                        console.warn(`7z command stdout: ${stdout}`);
                        resolve();
                    }
                });
                child.on('error', (err) => {
                    reject(new Error(`Failed to start 7z process: ${err.message}`));
                });
            });

            console.warn("Creating 7z archives using node-7z via command execution.");

        } else {
            throw new Error(`Unsupported archive format for creation: ${format}`);
        }

        return { success: true };
    } catch (error: any) {
        return { success: false, error: error.message };
    }
};


/**
 * Function to register the archive handling tools with the FastMCP server.
 * @param server - The FastMCP server instance.
 */
export const registerArchiveTools = (server: any) => {
    server.addTool({
        name: 'fs/archive/list',
        description: 'Lists the contents of an archive file (ZIP, 7z, etc.).',
        parameters: archiveListInputSchema,
        execute: listArchiveContents,
    });

    server.addTool({
        name: 'fs/archive/extract',
        description: 'Extracts contents from an archive file (ZIP, 7z, etc.).',
        parameters: archiveExtractInputSchema,
        execute: extractArchive,
    });

    server.addTool({
        name: 'fs/archive/create',
        description: 'Creates an archive file (ZIP or 7z) from a list of source paths.',
        parameters: archiveCreateInputSchema,
        execute: createArchive,
    });
};