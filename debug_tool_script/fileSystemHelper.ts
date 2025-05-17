import * as fs from 'fs/promises';
import * as path from 'path';

export async function listFixtureFiles(basePath: string = 'tests/fixtures/'): Promise<string[]> {
  try {
    const absoluteBasePath = path.resolve(basePath); // Ensure correct path resolution
    const entries = await fs.readdir(absoluteBasePath, { withFileTypes: true });
    const files = entries
      .filter(dirent => dirent.isFile())
      .map(dirent => dirent.name);
    return files;
  } catch (error: any) {
    // Check if the error is due to the directory not existing
    if (error.code === 'ENOENT') {
      console.warn(`Warning: Fixture directory not found at ${basePath}.`);
    } else {
      console.warn(`Warning: Could not read fixture directory at ${basePath}. Error: ${error.message}`);
    }
    return []; // Return empty array if directory doesn't exist or other error
  }
}