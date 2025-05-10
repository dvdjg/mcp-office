import * as fs from 'fs-extra';
import * as pathUtil from 'path';
import * as Papa from 'papaparse';
import { JSONPath } from 'jsonpath-plus';
import * as cheerio from 'cheerio';
import { z } from 'zod';

// Helper type for encoding
type ValidBufferEncoding = BufferEncoding;

// Helper function for robust file reading
async function readFileContent(filePath: string, encoding: ValidBufferEncoding = 'utf8'): Promise<string> {
    if (!await fs.pathExists(filePath)) {
        throw new Error(`File not found at path: ${filePath}`);
    }
    return fs.readFile(filePath, encoding);
}

// Helper function for robust file writing
async function writeFileContent(filePath: string, data: string, encoding: ValidBufferEncoding = 'utf8'): Promise<void> {
    await fs.ensureDir(pathUtil.dirname(filePath));
    return fs.writeFile(filePath, data, { encoding });
}

// --- CSV Operations ---

export const readCsvDataOptionsSchema = z.object({
    delimiter: z.string().optional(),
    columns: z.union([z.array(z.string()), z.array(z.number())]).optional(),
    startRow: z.number().int().positive().optional().default(1),
    endRow: z.number().int().positive().optional(),
    hasHeaders: z.boolean().optional().default(true),
    encoding: z.string().optional().default('utf8') as z.ZodType<ValidBufferEncoding>,
}).optional();

export const readCsvDataInputSchema = z.object({
    path: z.string(),
    options: readCsvDataOptionsSchema,
});

export const readCsvDataOutputSchema = z.array(z.union([z.record(z.string(), z.any()), z.array(z.any())]));

export type ReadCsvDataOptions = z.infer<typeof readCsvDataOptionsSchema>;
export type ReadCsvDataInput = z.infer<typeof readCsvDataInputSchema>;

export async function read_csv_data(
    { path, options }: ReadCsvDataInput
): Promise<z.infer<typeof readCsvDataOutputSchema>> {
    const {
        delimiter,
        columns,
        startRow = 1,
        endRow,
        hasHeaders = true,
        encoding = 'utf8',
    } = options || {};

    const fileContent = await readFileContent(path, encoding);

    return new Promise((resolve, reject) => {
        Papa.parse(fileContent, {
            delimiter: delimiter || undefined,
            header: hasHeaders && !(Array.isArray(columns) && typeof columns?.[0] === 'number'),
            dynamicTyping: true,
            skipEmptyLines: true,
            complete: (results: Papa.ParseResult<any>) => {
                let data = (hasHeaders && !(Array.isArray(columns) && typeof columns?.[0] === 'number'))
                    ? results.data as Record<string, any>[]
                    : results.data as any[][];

                const actualStartRow = hasHeaders ? Math.max(0, startRow - 1) : Math.max(0, startRow - 1);

                if (endRow) {
                    const actualEndRow = hasHeaders ? endRow : endRow;
                    data = data.slice(actualStartRow, actualEndRow);
                } else {
                    data = data.slice(actualStartRow);
                }

                if (columns && columns.length > 0) {
                    if (Array.isArray(data) && data.length > 0 && typeof data[0] === 'object' && data[0] !== null && !Array.isArray(data[0]) && typeof columns[0] === 'string') {
                        const selectedColumns = columns as string[];
                        data = (data as Record<string, any>[]).map(row => {
                            const newRow: Record<string, any> = {};
                            selectedColumns.forEach(colName => {
                                if (row.hasOwnProperty(colName)) {
                                    newRow[colName] = row[colName];
                                }
                            });
                            return newRow;
                        }).filter(row => Object.keys(row).length > 0);
                    } else if (!hasHeaders && Array.isArray(data) && data.length > 0 && Array.isArray(data[0]) && typeof columns[0] === 'number') {
                        // This case is for when hasHeaders is explicitly false, and we're selecting by numeric index from raw rows
                        const selectedIndices = columns as number[];
                        data = (data as any[][]).map(row =>
                            selectedIndices.map(index => row[index]).filter(value => value !== undefined)
                        ).filter(row => row.length > 0);
                    } else if (hasHeaders && Array.isArray(columns) && columns.length > 0 && typeof columns[0] === 'number' && Array.isArray(results.data) && results.data.length > 0) {
                        // This case: options.hasHeaders is true, but user wants to select columns by original index
                        // PapaParse would have run with header:false due to `typeof columns[0] === 'number'` in line 63
                        // So, results.data includes the header row. We need to skip it for data processing.
                        let rawData = results.data as any[][];
                        const headerOffset = 1; // Account for header row
                        const dataStartRow = actualStartRow + headerOffset;
                        const dataEndRow = endRow ? endRow + headerOffset : undefined;

                        if (dataEndRow) {
                            rawData = rawData.slice(dataStartRow, dataEndRow);
                        } else {
                            rawData = rawData.slice(dataStartRow);
                        }
                        const selectedIndices = columns as number[];
                        data = rawData.map(row =>
                            selectedIndices.map(index => row[index]).filter(value => value !== undefined)
                        ).filter(row => row.length > 0);
                    }
                }
                resolve(data);
            },
            error: (error: Error) => {
                reject(new Error(`CSV parsing error for ${path}: ${error.message}`));
            },
        });
    });
}


export const writeCsvDataOptionsSchema = z.object({
    delimiter: z.string().optional().default(','),
    headers: z.array(z.string()).optional(),
    includeHeaders: z.boolean().optional(),
    encoding: z.string().optional().default('utf8') as z.ZodType<ValidBufferEncoding>,
    mode: z.enum(['overwrite', 'append']).optional().default('overwrite'),
}).optional();

export const writeCsvDataInputSchema = z.object({
    path: z.string(),
    data: z.array(z.union([z.record(z.string(), z.any()), z.array(z.any())])),
    options: writeCsvDataOptionsSchema,
});
export const writeCsvDataOutputSchema = z.void();

export type WriteCsvDataOptions = z.infer<typeof writeCsvDataOptionsSchema>;
export type WriteCsvDataInput = z.infer<typeof writeCsvDataInputSchema>;


export async function write_csv_data(
    { path, data, options }: WriteCsvDataInput
): Promise<void> {
    const {
        delimiter = ',',
        headers,
        encoding = 'utf8',
        mode = 'overwrite',
    } = options || {};
    let { includeHeaders } = options || {};

    if (data.length === 0 && !headers && mode === 'overwrite') {
        await writeFileContent(path, '', encoding);
        return;
    }

    if (includeHeaders === undefined) {
        if (headers && headers.length > 0) {
            includeHeaders = true;
        } else if (data.length > 0 && typeof data[0] === 'object' && !Array.isArray(data[0]) && data[0] !== null) {
            includeHeaders = true;
        } else {
            includeHeaders = false;
        }
    }

    const config: Papa.UnparseConfig = {
        delimiter,
        header: includeHeaders,
        skipEmptyLines: true,
    };

    if (headers && headers.length > 0) {
        config.columns = headers;
    }

    let csvString = Papa.unparse(data, config);

    if (mode === 'append') {
        let existingContent = '';
        let needsNewline = false;
        if (await fs.pathExists(path)) {
            existingContent = await readFileContent(path, encoding);
            if (existingContent.length > 0 && !existingContent.endsWith('\n')) {
                needsNewline = true;
            }
            if (includeHeaders && existingContent.length > 0) { // If appending and existing file might have headers
                const lines = csvString.split('\n');
                if (lines.length > 1) {
                    csvString = lines.slice(1).join('\n');
                } else if (data.length === 0) {
                    csvString = "";
                }
            }
        }
        const contentToAppend = (needsNewline ? '\n' : '') + csvString;
        if (contentToAppend.trim().length > 0 || (existingContent.length === 0 && contentToAppend.length > 0) ){ // only append if there's actual content
             await fs.appendFile(path, contentToAppend, { encoding });
        }
    } else {
        await writeFileContent(path, csvString, encoding);
    }
}


// --- JSON Operations ---

export const readJsonPathOptionsSchema = z.object({
    encoding: z.string().optional().default('utf8') as z.ZodType<ValidBufferEncoding>,
}).optional();
export const readJsonPathInputSchema = z.object({
    path: z.string(),
    jsonPath: z.string(),
    options: readJsonPathOptionsSchema,
});
export const readJsonPathOutputSchema = z.any().optional();

export type ReadJsonPathOptions = z.infer<typeof readJsonPathOptionsSchema>;
export type ReadJsonPathInput = z.infer<typeof readJsonPathInputSchema>;

export async function read_json_path(
    { path, jsonPath, options }: ReadJsonPathInput
): Promise<any | undefined> {
    const { encoding = 'utf8' } = options || {};
    const fileContent = await readFileContent(path, encoding);
    try {
        const jsonData = JSON.parse(fileContent);
        const result = JSONPath({ path: jsonPath, json: jsonData, wrap: false }); // wrap: false to get value directly
        return result; // JSONPath returns the value or undefined if not found with wrap:false
    } catch (error: any) {
        if (error instanceof SyntaxError) {
            throw new Error(`Invalid JSON in file ${path}: ${error.message}`);
        }
        throw new Error(`Error processing JSONPath for ${path}: ${error.message}`);
    }
}

export const writeJsonPathOptionsSchema = z.object({
    createMissing: z.boolean().optional().default(true),
    encoding: z.string().optional().default('utf8') as z.ZodType<ValidBufferEncoding>,
}).optional();

export const writeJsonPathInputSchema = z.object({
    path: z.string(),
    jsonPath: z.string(),
    value: z.any(),
    options: writeJsonPathOptionsSchema,
});
export const writeJsonPathOutputSchema = z.void();

export type WriteJsonPathOptions = z.infer<typeof writeJsonPathOptionsSchema>;
export type WriteJsonPathInput = z.infer<typeof writeJsonPathInputSchema>;

export async function write_json_path(
    { path, jsonPath, value, options }: WriteJsonPathInput
): Promise<void> {
    const { createMissing = true, encoding = 'utf8' } = options || {};
    let jsonData: any;

    if (await fs.pathExists(path)) {
        const fileContent = await readFileContent(path, encoding);
        try {
            jsonData = JSON.parse(fileContent);
        } catch (error: any) {
            if (error instanceof SyntaxError) {
                throw new Error(`Invalid JSON in file ${path}: ${error.message}`);
            }
            throw error;
        }
    } else if (createMissing) {
        jsonData = {};
    } else {
        throw new Error(`File not found at ${path} and createMissing is false.`);
    }

    // Simplified and corrected manual path traversal and setting
    const pathSegments = JSONPath.toPathArray(jsonPath); // Use library's robust path parser
    if (pathSegments[0] === '$') {
        pathSegments.shift(); // Remove root '$'
    }

    let currentContext = jsonData;
    for (let i = 0; i < pathSegments.length - 1; i++) {
        const segment = pathSegments[i];
        const nextSegment = pathSegments[i + 1];
        const isNextSegmentNumericArrayIndex = typeof nextSegment === 'number' || /^\d+$/.test(String(nextSegment));

        if (!currentContext.hasOwnProperty(segment) || typeof currentContext[segment] !== 'object' || currentContext[segment] === null) {
            if (createMissing) {
                currentContext[segment] = isNextSegmentNumericArrayIndex ? [] : {};
            } else {
                throw new Error(`Path segment "${segment}" in "${jsonPath}" does not exist or is not an object/array, and createMissing is false.`);
            }
        }
        currentContext = currentContext[segment];
         // If we expect an array for the next segment but current is not an array (e.g. it's an object), and createMissing is true
        if (isNextSegmentNumericArrayIndex && !Array.isArray(currentContext) && createMissing) {
             // This case is tricky: if currentContext[segment] was an object, and next is an index.
             // This implies the path is trying to treat an object as an array.
             // For robust creation, if we hit this, we might need to overwrite currentContext[segment] with an array.
             // However, this is destructive. A safer approach is to ensure the structure matches.
             // For now, let's assume the structure will be built correctly if createMissing.
             // If currentContext is not an object (e.g. primitive), throw.
             if (typeof currentContext !== 'object' || currentContext === null) {
                throw new Error(`Path segment "${segment}" leads to a non-object/array, cannot create nested structure for "${jsonPath}".`);
             }
        }
    }

    const finalSegment = pathSegments[pathSegments.length - 1];
    if (typeof currentContext !== 'object' || currentContext === null) {
        if (createMissing && pathSegments.length === 1) { // Special case: root is not an object, but we are setting a top-level key
             // This shouldn't happen if jsonData starts as {} for new files.
             // If jsonData was loaded and is a primitive, this is an issue.
        } else {
            throw new Error(`Cannot set property on a non-object at path ending before "${finalSegment}" in "${jsonPath}".`);
        }
    }
    
    // If the final segment is a number, ensure the current context is an array
    if ((typeof finalSegment === 'number' || /^\d+$/.test(String(finalSegment))) && !Array.isArray(currentContext) && createMissing) {
        // This implies we are trying to set an array index on an object.
        // This part of the logic is complex if we need to convert an object to an array.
        // For now, this specific scenario might still be problematic if the path implies an array but an object exists.
        // However, if `createMissing` built the path, it should be an array.
    }
    
    currentContext[finalSegment] = value;

    await writeFileContent(path, JSON.stringify(jsonData, null, 2), encoding);
}


// --- HTML/XML Operations ---

export const extractFromMarkupOptionsSchema = z.object({
    extract: z.union([
        z.literal('text'),
        z.literal('html'),
        z.object({ attribute: z.string() })
    ]).optional().default('text'),
    encoding: z.string().optional().default('utf8') as z.ZodType<ValidBufferEncoding>,
    isXml: z.boolean().optional().default(false),
}).optional();

export const extractFromMarkupInputSchema = z.object({
    path: z.string(),
    selector: z.string(),
    options: extractFromMarkupOptionsSchema,
});
export const extractFromMarkupOutputSchema = z.union([z.string(), z.array(z.string()), z.null()]);

export type ExtractFromMarkupOptions = z.infer<typeof extractFromMarkupOptionsSchema>;
export type ExtractFromMarkupInput = z.infer<typeof extractFromMarkupInputSchema>;

export async function extract_from_markup(
    { path, selector, options }: ExtractFromMarkupInput
): Promise<string | string[] | null> {
    const { extract = 'text', encoding = 'utf8', isXml = false } = options || {};
    const fileContent = await readFileContent(path, encoding);

    const cheerioLoadOptions: { decodeEntities?: boolean, xml?: { xmlMode?: boolean, lowerCaseAttributeNames?: boolean } } = { decodeEntities: true };
    if (isXml) {
        cheerioLoadOptions.xml = {
            xmlMode: true,
            lowerCaseAttributeNames: false,
        };
    }
    const $ = cheerio.load(fileContent, cheerioLoadOptions);
    const elements = $(selector);

    if (elements.length === 0) {
        return null;
    }

    const results: string[] = [];
    elements.each((i: number, el: cheerio.Element) => {
        const element = $(el);
        if (extract === 'text') {
            results.push(element.text());
        } else if (extract === 'html') {
            results.push($.html(element) || ''); // $.html(element) for outer HTML
        } else if (typeof extract === 'object' && extract.attribute) {
            const attrValue = element.attr(extract.attribute);
            if (attrValue !== undefined) {
                 results.push(attrValue);
            }
        }
    });

    return results.length === 1 ? results[0] : (results.length > 0 ? results : null) ;
}

export const updateMarkupContentOptionsSchema = z.object({
    updateType: z.union([
        z.literal('text'),
        z.literal('html'),
        z.object({ attribute: z.string() })
    ]).optional().default('text'),
    encoding: z.string().optional().default('utf8') as z.ZodType<ValidBufferEncoding>,
    isXml: z.boolean().optional().default(false),
}).optional();

export const updateMarkupContentInputSchema = z.object({
    path: z.string(),
    selector: z.string(),
    newContent: z.string(),
    options: updateMarkupContentOptionsSchema,
});
export const updateMarkupContentOutputSchema = z.object({ modifiedCount: z.number() });

export type UpdateMarkupContentOptions = z.infer<typeof updateMarkupContentOptionsSchema>;
export type UpdateMarkupContentInput = z.infer<typeof updateMarkupContentInputSchema>;

export async function update_markup_content(
    { path, selector, newContent, options }: UpdateMarkupContentInput
): Promise<{ modifiedCount: number }> {
    const { updateType = 'text', encoding = 'utf8', isXml = false } = options || {};
    const fileContent = await readFileContent(path, encoding);

    const cheerioLoadOptionsUpdate: { decodeEntities?: boolean, xml?: { xmlMode?: boolean, lowerCaseAttributeNames?: boolean } } = { decodeEntities: true };
    if (isXml) {
        cheerioLoadOptionsUpdate.xml = {
            xmlMode: true,
            lowerCaseAttributeNames: false,
        };
    }
    const $ = cheerio.load(fileContent, cheerioLoadOptionsUpdate);
    const elements = $(selector);

    if (elements.length === 0) {
        return { modifiedCount: 0 };
    }

    elements.each((i: number, el: cheerio.Element) => {
        const element = $(el);
        if (updateType === 'text') {
            element.text(newContent);
        } else if (updateType === 'html') {
            element.html(newContent); // Sets inner HTML
        } else if (typeof updateType === 'object' && updateType.attribute) {
            element.attr(updateType.attribute, newContent);
        }
    });

    await writeFileContent(path, isXml ? $.xml() : $.html(), encoding);
    return { modifiedCount: elements.length };
}

export const structuredDataTools = {
    read_csv_data,
    write_csv_data,
    read_json_path,
    write_json_path,
    extract_from_markup,
    update_markup_content,
};

export const structuredDataSchemas = {
    readCsvDataInputSchema,
    readCsvDataOutputSchema,
    writeCsvDataInputSchema,
    writeCsvDataOutputSchema,
    readJsonPathInputSchema,
    readJsonPathOutputSchema,
    writeJsonPathInputSchema,
    writeJsonPathOutputSchema,
    extractFromMarkupInputSchema,
    extractFromMarkupOutputSchema,
    updateMarkupContentInputSchema,
    updateMarkupContentOutputSchema,
};