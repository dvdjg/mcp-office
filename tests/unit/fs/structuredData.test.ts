import fs from 'fs-extra';
import * as path from 'path';
import {
    read_csv_data,
    write_csv_data,
    read_json_path,
    write_json_path,
    extract_from_markup,
    update_markup_content,
} from '../../../src/tools/fs/structuredData.tool'; // Adjust path as necessary

const fixturesDir = path.join(process.cwd(), 'tests', 'fixtures');
const tempTestDir = path.join(process.cwd(), 'tests', 'temp_structured_data_tests');

describe('Structured Data Tools', () => {
    beforeAll(async () => {
        await fs.ensureDir(tempTestDir);
    });

    afterAll(async () => {
        await fs.remove(tempTestDir);
    });

    beforeEach(async () => {
        await fs.emptyDir(tempTestDir);
    });

    const normalizeNewlines = (str: string) => str.replace(/\r\n/g, '\n');

    describe('CSV Operations', () => {
        const sampleCsvPath = path.join(fixturesDir, 'sample.csv');
        const tempCsvPath = path.join(tempTestDir, 'test.csv');

        // Content of sample.csv:
        // Name,Age,City
        // Alice,30,New York
        // Bob,24,San Francisco
        // Charlie,35,London
        // David,22,"Paris;France"

        it('read_csv_data should read all data with headers', async () => {
            const result = await read_csv_data({ path: sampleCsvPath, options: { hasHeaders: true, startRow: 1, encoding: 'utf8' } });
            expect(result.length).toBe(4);
            expect(result[0]).toEqual({ Name: 'Alice', Age: 30, City: 'New York' });
            expect(result[3]).toEqual({ Name: 'David', Age: 22, City: 'Paris;France' });
        });

        it('read_csv_data should read without headers', async () => {
            const result = await read_csv_data({ path: sampleCsvPath, options: { hasHeaders: false, startRow: 1, encoding: 'utf8' } });
            expect(result.length).toBe(5); // Includes header row as data
            expect(result[0]).toEqual(['Name', 'Age', 'City']);
            expect(result[1]).toEqual(['Alice', 30, 'New York']);
        });

        it('read_csv_data should handle specified delimiter', async () => {
            const csvWithSemicolon = "Name;Age;City\nEve;40;Berlin";
            await fs.writeFile(tempCsvPath, csvWithSemicolon);
            const result = await read_csv_data({ path: tempCsvPath, options: { delimiter: ';', hasHeaders: true, startRow: 1, encoding: 'utf8' } });
            expect(result.length).toBe(1);
            expect(result[0]).toEqual({ Name: 'Eve', Age: 40, City: 'Berlin' });
        });

        it('read_csv_data should select specific columns by name', async () => {
            const result = await read_csv_data({ path: sampleCsvPath, options: { columns: ['Name', 'City'], hasHeaders: true, startRow: 1, encoding: 'utf8' } });
            expect(result.length).toBe(4);
            expect(result[0]).toEqual({ Name: 'Alice', City: 'New York' });
            expect(Object.keys(result[0] as object).length).toBe(2);
        });

        it('read_csv_data should select specific columns by index (no headers)', async () => {
            const result = await read_csv_data({ path: sampleCsvPath, options: { columns: [0, 2], hasHeaders: false, startRow: 1, encoding: 'utf8' } });
            expect(result.length).toBe(5);
            expect(result[0]).toEqual(['Name', 'City']); // Header row processed
            expect(result[1]).toEqual(['Alice', 'New York']);
        });
        
        it('read_csv_data should select specific columns by index (with headers)', async () => {
            // This scenario implies we want specific columns from the raw data, ignoring header mapping for selection
            // but still treating the first row as headers for the output structure if not overridden.
            // The tool's behavior for `columns: number[]` with `hasHeaders: true` might be to select based on original column index.
            const result = await read_csv_data({ path: sampleCsvPath, options: { columns: [0, 2], hasHeaders: true, startRow: 1, encoding: 'utf8' } });
            // Expecting it to pick 0th and 2nd column data, using original headers for keys if possible
            // PapaParse with `header:true` returns objects. If columns are numbers, it's tricky.
            // The tool implementation seems to handle this by re-processing if columns are numbers with hasHeaders:true
            expect(result.length).toBe(4);
            expect(result[0]).toEqual(['Alice', 'New York']); // Data from columns 0 and 2
        });


        it('read_csv_data should handle startRow and endRow', async () => {
            const result = await read_csv_data({ path: sampleCsvPath, options: { startRow: 2, endRow: 3, hasHeaders: true, encoding: 'utf8' } }); // Bob and Charlie
            expect(result.length).toBe(2);
            expect(result[0]).toEqual({ Name: 'Bob', Age: 24, City: 'San Francisco' });
            expect(result[1]).toEqual({ Name: 'Charlie', Age: 35, City: 'London' });
        });

        it('write_csv_data should write data with headers', async () => {
            const dataToWrite = [{ Name: 'Test', Value: 100 }, { Name: 'Another', Value: 200 }];
            await write_csv_data({ path: tempCsvPath, data: dataToWrite, options: { includeHeaders: true, delimiter: ',', mode: 'overwrite', encoding: 'utf8' } });
            const content = await fs.readFile(tempCsvPath, 'utf-8');
            const lines = normalizeNewlines(content.trim()).split('\n');
            expect(lines[0]).toBe('Name,Value');
            expect(lines[1]).toBe('Test,100');
        });

        it('write_csv_data should write data without headers', async () => {
            const dataToWrite = [['Test', 100], ['Another', 200]];
            await write_csv_data({ path: tempCsvPath, data: dataToWrite, options: { includeHeaders: false, delimiter: ',', mode: 'overwrite', encoding: 'utf8' } });
            const content = await fs.readFile(tempCsvPath, 'utf-8');
            const lines = normalizeNewlines(content.trim()).split('\n');
            expect(lines[0]).toBe('Test,100');
        });

        it('write_csv_data should use specified delimiter', async () => {
            const dataToWrite = [{ Name: 'Test', Value: 100 }];
            await write_csv_data({ path: tempCsvPath, data: dataToWrite, options: { delimiter: ';', includeHeaders: true, mode: 'overwrite', encoding: 'utf8' } });
            const content = await fs.readFile(tempCsvPath, 'utf-8');
            expect(normalizeNewlines(content.trim())).toBe('Name;Value\nTest;100');
        });

        it('write_csv_data should append data', async () => {
            const initialData = [{ ID: 1, Status: 'A' }];
            await write_csv_data({ path: tempCsvPath, data: initialData, options: { includeHeaders: true, delimiter: ',', mode: 'overwrite', encoding: 'utf8' } });
            const dataToAppend = [{ ID: 2, Status: 'B' }];
            await write_csv_data({ path: tempCsvPath, data: dataToAppend, options: { mode: 'append', includeHeaders: true, delimiter: ',', encoding: 'utf8' } }); // includeHeaders true for append means don't re-write header
            
            const content = await fs.readFile(tempCsvPath, 'utf-8');
            const lines = normalizeNewlines(content.trim()).split('\n');
            expect(lines.length).toBe(3);
            expect(lines[0]).toBe('ID,Status');
            expect(lines[1]).toBe('1,A');
            expect(lines[2]).toBe('2,B');
        });
    });

    describe('JSON Operations', () => {
        const sampleJsonPath = path.join(fixturesDir, 'sample.json');
        const tempJsonPath = path.join(tempTestDir, 'test.json');
        // sample.json content: { "name": "John Doe", "age": 30, "address": { "street": "123 Main St" }, "courses": [...] }

        it('read_json_path should extract a simple value', async () => {
            const name = await read_json_path({ path: sampleJsonPath, jsonPath: '$.name' });
            expect(name).toBe('John Doe');
        });

        it('read_json_path should extract a nested value', async () => {
            const street = await read_json_path({ path: sampleJsonPath, jsonPath: '$.address.street' });
            expect(street).toBe('123 Main St');
        });

        it('read_json_path should extract an array element', async () => {
            const courseTitle = await read_json_path({ path: sampleJsonPath, jsonPath: '$.courses[0].title' });
            expect(courseTitle).toBe('History');
        });

        it('read_json_path should return undefined for non-existent path', async () => {
            const result = await read_json_path({ path: sampleJsonPath, jsonPath: '$.nonexistent.path' });
            expect(result).toBeUndefined();
        });

        it('write_json_path should write a new value to an existing file', async () => {
            await fs.copyFile(sampleJsonPath, tempJsonPath);
            await write_json_path({ path: tempJsonPath, jsonPath: '$.age', value: 31 });
            const newAge = await read_json_path({ path: tempJsonPath, jsonPath: '$.age' });
            expect(newAge).toBe(31);
        });

        it('write_json_path should create a new file and path if createMissing is true', async () => {
            await write_json_path({ path: tempJsonPath, jsonPath: '$.newly.created.path', value: 'testValue', options: { createMissing: true, encoding: 'utf8' } });
            const result = await read_json_path({ path: tempJsonPath, jsonPath: '$.newly.created.path' });
            expect(result).toBe('testValue');
            const content = JSON.parse(await fs.readFile(tempJsonPath, 'utf-8'));
            expect(content.newly.created.path).toBe('testValue');
        });

        it('write_json_path should update an existing nested value', async () => {
            await fs.copyFile(sampleJsonPath, tempJsonPath);
            await write_json_path({ path: tempJsonPath, jsonPath: '$.address.city', value: 'New City' });
            const newCity = await read_json_path({ path: tempJsonPath, jsonPath: '$.address.city' });
            expect(newCity).toBe('New City'); // Original sample.json has Anytown
        });
        
        it('write_json_path should add a new property to an existing object', async () => {
            await fs.copyFile(sampleJsonPath, tempJsonPath);
            await write_json_path({ path: tempJsonPath, jsonPath: '$.address.zip', value: '90210', options: { createMissing: true, encoding: 'utf8' } });
            const zip = await read_json_path({ path: tempJsonPath, jsonPath: '$.address.zip' });
            expect(zip).toBe('90210');
        });

        it('write_json_path should throw if path does not exist and createMissing is false', async () => {
            await expect(write_json_path({ path: tempJsonPath, jsonPath: '$.some.path', value: 'test', options: { createMissing: false, encoding: 'utf8' } })).rejects.toThrow();
        });
    });

    describe('Markup Operations (HTML/XML)', () => {
        const sampleHtmlPath = path.join(fixturesDir, 'sample.html');
        const tempHtmlPath = path.join(tempTestDir, 'test.html');
        // sample.html: <h1 id="main-heading">Hello World!</h1> <p class="content">This is a sample paragraph.</p> ...

        it('extract_from_markup should extract text from HTML element', async () => {
            const text = await extract_from_markup({ path: sampleHtmlPath, selector: '#main-heading' });
            expect(text).toBe('Hello World!');
        });

        it('extract_from_markup should extract HTML content of an element', async () => {
            const html = await extract_from_markup({ path: sampleHtmlPath, selector: 'div[data-testid="test-div"]', options: { extract: 'html', isXml: false, encoding: 'utf8' } });
            expect(html).toContain('<span class="nested">Nested span 1</span>');
        });

        it('extract_from_markup should extract an attribute value', async () => {
            const href = await extract_from_markup({ path: sampleHtmlPath, selector: '#link', options: { extract: { attribute: 'href' }, isXml: false, encoding: 'utf8' } });
            expect(href).toBe('https://example.com');
        });
        
        it('extract_from_markup should return array for multiple matches', async () => {
            const items = await extract_from_markup({ path: sampleHtmlPath, selector: 'ul li', options: { extract: 'text', isXml: false, encoding: 'utf8' } });
            expect(items).toEqual(['Item 1', 'Item 2', 'Item 3']);
        });

        it('extract_from_markup should return null if selector does not match', async () => {
            const result = await extract_from_markup({ path: sampleHtmlPath, selector: '#nonexistent' });
            expect(result).toBeNull();
        });

        it('update_markup_content should update text of an HTML element', async () => {
            await fs.copyFile(sampleHtmlPath, tempHtmlPath);
            const { modifiedCount } = await update_markup_content({ path: tempHtmlPath, selector: '#main-heading', newContent: 'Updated Heading' });
            expect(modifiedCount).toBe(1);
            const updatedText = await extract_from_markup({ path: tempHtmlPath, selector: '#main-heading' });
            expect(updatedText).toBe('Updated Heading');
        });

        it('update_markup_content should update HTML content of an element', async () => {
            await fs.copyFile(sampleHtmlPath, tempHtmlPath);
            const newHtml = '<p>Brand new content</p>';
            await update_markup_content({ path: tempHtmlPath, selector: 'div[data-testid="test-div"]', newContent: newHtml, options: { updateType: 'html', isXml: false, encoding: 'utf8' } });
            const updatedHtml = await extract_from_markup({ path: tempHtmlPath, selector: 'div[data-testid="test-div"]', options: { extract: 'html', isXml: false, encoding: 'utf8' } });
            expect(updatedHtml).toContain(newHtml);
        });

        it('update_markup_content should update an attribute value', async () => {
            await fs.copyFile(sampleHtmlPath, tempHtmlPath);
            await update_markup_content({ path: tempHtmlPath, selector: '#link', newContent: 'https://newexample.com', options: { updateType: { attribute: 'href' }, isXml: false, encoding: 'utf8' } });
            const updatedHref = await extract_from_markup({ path: tempHtmlPath, selector: '#link', options: { extract: { attribute: 'href' }, isXml: false, encoding: 'utf8' } });
            expect(updatedHref).toBe('https://newexample.com');
        });

        // XML specific tests
        const sampleXmlPath = path.join(fixturesDir, 'sample.xml');
        const tempXmlPath = path.join(tempTestDir, 'test.xml');
        // sample.xml: <root><item id="1"><name>Item A</name>...</item>...</root>

        it('extract_from_markup should extract text from XML element', async () => {
            const text = await extract_from_markup({ path: sampleXmlPath, selector: 'item[id="1"] name', options: { isXml: true, extract: 'text', encoding: 'utf8' } });
            expect(text).toBe('Item A');
        });

        it('extract_from_markup should extract attribute from XML element', async () => {
            // Corrected selector to find an item that has a child 'name' tag containing "Item B"
            const id = await extract_from_markup({ path: sampleXmlPath, selector: 'item:has(name:contains("Item B"))', options: { extract: { attribute: 'id' }, isXml: true, encoding: 'utf8' } });
            expect(id).toBe('2');
        });

        it('update_markup_content should update text of an XML element', async () => {
            await fs.copyFile(sampleXmlPath, tempXmlPath);
            await update_markup_content({ path: tempXmlPath, selector: 'item[id="1"] name', newContent: 'Updated Item A', options: { isXml: true, updateType: 'text', encoding: 'utf8' } });
            const updatedText = await extract_from_markup({ path: tempXmlPath, selector: 'item[id="1"] name', options: { isXml: true, extract: 'text', encoding: 'utf8' } });
            expect(updatedText).toBe('Updated Item A');
        });
    });
});
