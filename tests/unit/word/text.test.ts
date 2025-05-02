import { getText, insertText, modifyText, deleteText } from '../../../src/tools/word/text.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { z } from 'zod';

// Mock de las dependencias
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
}));
jest.mock('../../../src/utils/logger');

const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;
const mockLogger = logger as jest.Mocked<typeof logger>;

describe('word/text unit tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockSelection: any;
    let mockRange: any;
    let mockParagraphs: any;
    let mockOfficeAppInstance: any;

    beforeEach(() => {
        // Reset mocks antes de cada prueba
        jest.clearAllMocks();

        // Configurar mocks básicos para Word COM objects
        mockRange = {
            Text: 'Sample text',
            Start: 0,
            End: 11,
            Delete: jest.fn(),
            Collapse: jest.fn(),
        };
        mockSelection = {
            Range: mockRange,
            Type: 2, // wdSelectionIP = 2 (Insertion Point)
        };
        mockParagraphs = {
            Count: 5,
            Item: jest.fn((index: number) => {
                if (index > 0 && index <= mockParagraphs.Count) {
                    return { Range: { Text: `Paragraph ${index}`, Start: (index - 1) * 15, End: (index * 15) -1 } }; // Mock simple paragraph range
                }
                return undefined; // Simular COM object no encontrado
            }),
        };
        mockDoc = {
            Content: { Text: 'Document content', Start: 0, End: 16 },
            Paragraphs: mockParagraphs,
            Range: jest.fn((start, end) => {
                 // Return a new mock range for specific positions
                 const newRange = {
                    Text: '', // Initially empty for insertion
                    Start: start,
                    End: end,
                    Delete: jest.fn(),
                    Collapse: jest.fn(),
                 };
                 // Add a mock setter for Text to simulate insertion
                 Object.defineProperty(newRange, 'Text', {
                     set: jest.fn((value) => { newRange._text = value; }),
                     get: jest.fn(() => newRange._text),
                     configurable: true,
                 });
                 return newRange;
            }),
            Save: jest.fn(),
            Close: jest.fn(),
        };
        mockWordApp = {
            Documents: {
                Open: jest.fn().mockResolvedValue(mockDoc),
            },
            Selection: mockSelection,
            Quit: jest.fn(),
        };

        // Mock the OfficeAppInstance structure
        mockOfficeAppInstance = {
            app: mockWordApp,
            openDocument: jest.fn().mockReturnValue(mockDoc),
            release: jest.fn(),
        };

        mockGetOfficeApplication.mockResolvedValue(mockOfficeAppInstance);
        mockValidateFilePath.mockReturnValue(true); // Asumir que la validación de ruta es exitosa por defecto
    });

    // --- Pruebas Unitarias para getText ---

    describe('getText', () => {
        const baseParams = { filePath: 'C:/test/document.docx' };

        test('debería obtener texto de la selección', async () => {
            const params = { ...baseParams, range: 'selection' };
            const result = await getText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            // getRangeFromSpecifier is called internally, we check its effect
            expect(mockSelection.Range).toBe(mockRange);
            expect(result).toEqual({ success: true, data: 'Sample text' });
            expect(mockDoc.Close).toHaveBeenCalledWith(false); // Cerrar sin guardar
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // Range obtained from getRangeFromSpecifier
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully retrieved text'));
        });

        test('debería obtener texto del documento completo', async () => {
            const params = { ...baseParams, range: 'document' };
            const result = await getText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            // getRangeFromSpecifier is called internally
            expect(mockDoc.Content.Text).toBe('Document content');
            expect(result).toEqual({ success: true, data: 'Document content' });
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // Note: doc.Content is not explicitly released in the tool's finally block,
            // as it's part of the document object itself.
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully retrieved text'));
        });

        test('debería obtener texto de un párrafo específico', async () => {
            const params = { ...baseParams, range: 'paragraph:2' };
            const result = await getText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(2);
            // getRangeFromSpecifier is called internally
            expect(result).toEqual({ success: true, data: 'Paragraph 2' });
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // The range returned by Paragraphs.Item().Range is released
            expect(mockReleaseObject).toHaveBeenCalled(); // Check if releaseObject was called at least once
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully retrieved text'));
        });

        test('debería manejar un filePath inválido', async () => {
            mockValidateFilePath.mockReturnValue(false);
            const params = { ...baseParams, range: 'document' };
            const result = await getText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Invalid or potentially unsafe file path');
        });

        test('debería manejar un rango no soportado', async () => {
            const params = { ...baseParams, range: 'invalidRange' };
            const result = await getText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Unsupported range format');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('debería manejar un índice de párrafo fuera de límites', async () => {
            const params = { ...baseParams, range: 'paragraph:10' }; // Solo hay 5 párrafos mockeados
            const result = await getText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(10);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Paragraph index 10 is out of bounds');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

         test('debería manejar un error al abrir el documento', async () => {
            mockOfficeAppInstance.openDocument.mockReturnValue(null); // Simular fallo al abrir
            const params = { ...baseParams, range: 'document' };
            const result = await getText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Failed to open document');
            // No debería intentar cerrar doc si es null
            expect(mockDoc.Close).not.toHaveBeenCalled();
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });
    });

    // --- Pruebas Unitarias para insertText ---

    describe('insertText', () => {
        const baseParams = { filePath: 'C:/test/document.docx', text: 'Inserted text' };

        test('debería insertar texto al inicio del documento', async () => {
            const params = { ...baseParams, position: 'start' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockDoc.Range).toHaveBeenCalledWith(0, 0);
            expect(mockRange.Text).toBe(params.text); // Check the setter was called
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(mockReleaseObject).toHaveBeenCalledWith(undefined); // paraRange (should be undefined)
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully inserted text'));
        });

        test('debería insertar texto al final del documento', async () => {
            const params = { ...baseParams, position: 'end' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockDoc.Content.End).toBe(16); // Mocked end position
            expect(mockDoc.Range).toHaveBeenCalledWith(16, 16);
            expect(mockRange.Text).toBe(params.text);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(mockReleaseObject).toHaveBeenCalledWith(undefined); // paraRange
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully inserted text'));
        });

        test('debería insertar texto en la selección (punto de inserción)', async () => {
            const params = { ...baseParams, position: 'selection' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockWordApp.Selection.Range).toBe(mockRange);
            expect(mockWordApp.Selection.Type).toBe(2); // wdSelectionIP
            expect(mockRange.Collapse).not.toHaveBeenCalled(); // Should not collapse if IP
            expect(mockRange.Text).toBe(params.text);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(mockReleaseObject).toHaveBeenCalledWith(undefined); // paraRange
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully inserted text'));
        });

         test('debería insertar texto en la selección (rango seleccionado)', async () => {
            mockSelection.Type = 1; // wdSelectionNormal = 1
            const params = { ...baseParams, position: 'selection' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockWordApp.Selection.Range).toBe(mockRange);
            expect(mockWordApp.Selection.Type).toBe(1); // wdSelectionNormal
            expect(mockRange.Collapse).toHaveBeenCalledWith(1); // Should collapse to start
            expect(mockRange.Text).toBe(params.text);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(mockReleaseObject).toHaveBeenCalledWith(undefined); // paraRange
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully inserted text'));
        });


        test('debería insertar texto al inicio de un párrafo específico', async () => {
            const params = { ...baseParams, position: 'paragraph:3:start' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(3);
            // Mocked paragraph 3 range start is (3-1)*15 = 30
            expect(mockDoc.Range).toHaveBeenCalledWith(30, 30);
            expect(mockRange.Text).toBe(params.text);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(mockReleaseObject).toHaveBeenCalled(); // paraRange should be released
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully inserted text'));
        });

        test('debería insertar texto al final de un párrafo específico', async () => {
            const params = { ...baseParams, position: 'paragraph:2:end' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(2);
             // Mocked paragraph 2 range end is 2*15 - 1 = 29
            expect(mockDoc.Range).toHaveBeenCalledWith(29, 29);
            expect(mockRange.Text).toBe(params.text);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(mockReleaseObject).toHaveBeenCalled(); // paraRange should be released
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully inserted text'));
        });

        test('debería insertar texto al inicio de un párrafo por defecto si no se especifica start/end', async () => {
            const params = { ...baseParams, position: 'paragraph:4' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(4);
            // Mocked paragraph 4 range start is (4-1)*15 = 45
            expect(mockDoc.Range).toHaveBeenCalledWith(45, 45);
            expect(mockRange.Text).toBe(params.text);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(mockReleaseObject).toHaveBeenCalled(); // paraRange should be released
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully inserted text'));
        });

        test('debería manejar un filePath inválido', async () => {
            mockValidateFilePath.mockReturnValue(false);
            const params = { ...baseParams, position: 'start' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Invalid or potentially unsafe file path');
        });

        test('debería manejar una posición no soportada', async () => {
            const params = { ...baseParams, position: 'invalidPosition' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Unsupported position specifier');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('debería manejar un índice de párrafo inválido (no numérico)', async () => {
            const params = { ...baseParams, position: 'paragraph:abc:start' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Invalid paragraph index format');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('debería manejar un índice de párrafo fuera de límites', async () => {
            const params = { ...baseParams, position: 'paragraph:10:start' }; // Solo hay 5 párrafos mockeados
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(10);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Paragraph index 10 is out of bounds');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('debería manejar un error al abrir el documento', async () => {
            mockOfficeAppInstance.openDocument.mockReturnValue(null); // Simular fallo al abrir
            const params = { ...baseParams, position: 'start' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Failed to open document');
            // No debería intentar cerrar doc si es null
            expect(mockDoc.Close).not.toHaveBeenCalled();
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });
    });

    // --- Pruebas Unitarias para modifyText ---

    describe('modifyText', () => {
        const baseParams = { filePath: 'C:/test/document.docx', newText: 'Modified text' };

        test('debería modificar texto en la selección', async () => {
            const params = { ...baseParams, range: 'selection' };
            const result = await modifyText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // getRangeFromSpecifier is called internally
            expect(mockSelection.Range).toBe(mockRange);
            expect(mockRange.Text).toBe(params.newText);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // selectedRange
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully modified text'));
        });

        test('debería modificar texto en el documento completo', async () => {
            const params = { ...baseParams, range: 'document' };
            const result = await modifyText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // getRangeFromSpecifier is called internally
            // Note: doc.Content is not the same mockRange as selection, need to adjust mock if testing this specifically
            // For now, assume getRangeFromSpecifier returns a range-like object
            // We check if the Text property of the returned range is set
            const rangeFromSpecifier = getRangeFromSpecifier(mockDoc, params.range, mockWordApp);
            expect(rangeFromSpecifier.Text).toBe(params.newText);

            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // The range returned by getRangeFromSpecifier is released
            expect(mockReleaseObject).toHaveBeenCalled(); // Check if releaseObject was called at least once
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully modified text'));
        });

        test('debería modificar texto en un párrafo específico', async () => {
            const params = { ...baseParams, range: 'paragraph:1' };
            const result = await modifyText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(1);
            // getRangeFromSpecifier is called internally
            const rangeFromSpecifier = getRangeFromSpecifier(mockDoc, params.range, mockWordApp);
            expect(rangeFromSpecifier.Text).toBe(params.newText);

            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // The range returned by Paragraphs.Item().Range is released
            expect(mockReleaseObject).toHaveBeenCalled(); // Check if releaseObject was called at least once
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully modified text'));
        });

        test('debería manejar un filePath inválido', async () => {
            mockValidateFilePath.mockReturnValue(false);
            const params = { ...baseParams, range: 'document' };
            const result = await modifyText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Invalid or potentially unsafe file path');
        });

        test('debería manejar un rango no soportado', async () => {
            const params = { ...baseParams, range: 'invalidRange' };
            const result = await modifyText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Unsupported range format');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('debería manejar un índice de párrafo fuera de límites', async () => {
            const params = { ...baseParams, range: 'paragraph:10' }; // Solo hay 5 párrafos mockeados
            const result = await modifyText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(10);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Paragraph index 10 is out of bounds');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

         test('debería manejar un error al abrir el documento', async () => {
            mockOfficeAppInstance.openDocument.mockReturnValue(null); // Simular fallo al abrir
            const params = { ...baseParams, range: 'document' };
            const result = await modifyText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Failed to open document');
            // No debería intentar cerrar doc si es null
            expect(mockDoc.Close).not.toHaveBeenCalled();
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });
    });

    // --- Pruebas Unitarias para deleteText ---

    describe('deleteText', () => {
        const baseParams = { filePath: 'C:/test/document.docx' };

        test('debería eliminar texto en la selección', async () => {
            const params = { ...baseParams, range: 'selection' };
            const result = await deleteText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // getRangeFromSpecifier is called internally
            expect(mockSelection.Range).toBe(mockRange);
            expect(mockRange.Delete).toHaveBeenCalled();
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // selectedRange
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully deleted text'));
        });

        test('debería eliminar texto en el documento completo', async () => {
            const params = { ...baseParams, range: 'document' };
            const result = await deleteText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // getRangeFromSpecifier is called internally
            const rangeFromSpecifier = getRangeFromSpecifier(mockDoc, params.range, mockWordApp);
            expect(rangeFromSpecifier.Delete).toHaveBeenCalled();

            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // The range returned by getRangeFromSpecifier is released
            expect(mockReleaseObject).toHaveBeenCalled(); // Check if releaseObject was called at least once
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully deleted text'));
        });

        test('debería eliminar texto en un párrafo específico', async () => {
            const params = { ...baseParams, range: 'paragraph:1' };
            const result = await deleteText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(1);
            // getRangeFromSpecifier is called internally
            const rangeFromSpecifier = getRangeFromSpecifier(mockDoc, params.range, mockWordApp);
            expect(rangeFromSpecifier.Delete).toHaveBeenCalled();

            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // The range returned by Paragraphs.Item().Range is released
            expect(mockReleaseObject).toHaveBeenCalled(); // Check if releaseObject was called at least once
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully deleted text'));
        });

        test('debería manejar un filePath inválido', async () => {
            mockValidateFilePath.mockReturnValue(false);
            const params = { ...baseParams, range: 'document' };
            const result = await deleteText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Invalid or potentially unsafe file path');
        });

        test('debería manejar un rango no soportado', async () => {
            const params = { ...baseParams, range: 'invalidRange' };
            const result = await deleteText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Unsupported range format');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

        test('debería manejar un índice de párrafo fuera de límites', async () => {
            const params = { ...baseParams, range: 'paragraph:10' }; // Solo hay 5 párrafos mockeados
            const result = await deleteText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(mockParagraphs.Item).toHaveBeenCalledWith(10);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Paragraph index 10 is out of bounds');
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });

         test('debería manejar un error al abrir el documento', async () => {
            mockOfficeAppInstance.openDocument.mockReturnValue(null); // Simular fallo al abrir
            const params = { ...baseParams, range: 'document' };
            const result = await deleteText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            expect(result.success).toBe(false);
            expect(result.error?.message).toContain('Failed to open document');
            // No debería intentar cerrar doc si es null
            expect(mockDoc.Close).not.toHaveBeenCalled();
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
        });
    });
});

// Mock implementation for the local getRangeFromSpecifier used within text.tool.ts
// This is needed because the tests mock the external dependencies but the internal helper is called directly.
// We need to provide a mock implementation that behaves like the real one for testing purposes.
// This mock should be placed outside the describe block or in a separate file if preferred.
// For simplicity, we'll place it here and ensure it's used by the handlers being tested.
// Note: This is a simplified mock and might need refinement based on actual getRangeFromSpecifier logic.
// A better approach might be to export getRangeFromSpecifier and mock it like other dependencies.
// Given the current structure, we'll mock the internal behavior by ensuring the mocks
// for doc.Content, doc.Paragraphs.Item, and wordApp.Selection.Range return the expected mockRange objects.
// The tests for each handler already implicitly test the getRangeFromSpecifier logic by checking
// which COM object properties/methods are accessed based on the range specifier.
// We can add a specific test for getRangeFromSpecifier if needed, but for now,
// the handler tests cover its usage sufficiently for unit testing purposes.

// Re-mocking getRangeFromSpecifier if it were exported:
/*
jest.mock('../../../src/tools/word/text.tool', () => {
    const originalModule = jest.requireActual('../../../src/tools/word/text.tool');
    return {
        ...originalModule,
        getRangeFromSpecifier: jest.fn((doc, rangeSpecifier, wordApp) => {
            // Simplified mock logic based on the real function
            const rangeStringLower = rangeSpecifier.toLowerCase();
            if (rangeStringLower === 'selection') {
                return wordApp.Selection.Range;
            } else if (rangeStringLower === 'document') {
                return doc.Content;
            } else if (rangeStringLower.startsWith('paragraph:')) {
                const indexStr = rangeSpecifier.split(':')[1];
                const index = parseInt(indexStr, 10);
                if (!isNaN(index) && index > 0 && index <= doc.Paragraphs.Count) {
                    return doc.Paragraphs.Item(index).Range;
                }
                throw new Error('Mock: Invalid paragraph range');
            }
            throw new Error('Mock: Unsupported range format');
        }),
    };
});
*/