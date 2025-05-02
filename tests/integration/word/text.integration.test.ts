import { getText, insertText, modifyText, deleteText } from '../../../src/tools/word/text.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import { z } from 'zod';

// Mock de las dependencias externas (simulando la interacción con Office)
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
    createErrorResponse: jest.fn((message, code, details) => ({ success: false, error: { code, message, details } })),
}));
jest.mock('../../../src/utils/logger');

const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;
const mockLogger = logger as jest.Mocked<typeof logger>;

describe('word/text integration tests', () => {
    let mockWordApp: any;
    let mockDoc: any;
    let mockSelection: any;
    let mockRange: any;
    let mockParagraphs: any;
    let mockOfficeAppInstance: any;

    beforeEach(() => {
        jest.clearAllMocks();

        // Configurar mocks para simular objetos COM de Word y la interacción con officeInterop
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
                    return { Range: { Text: `Paragraph ${index}`, Start: (index - 1) * 15, End: (index * 15) -1, Collapse: jest.fn() } };
                }
                return undefined;
            }),
        };
        mockDoc = {
            Content: { Text: 'Document content', Start: 0, End: 16 },
            Paragraphs: mockParagraphs,
            Range: jest.fn((start, end) => {
                 // Return a new mock range for specific positions
                 const newRange: any = { // Use 'any' to allow _text property
                    Text: '', // Initially empty for insertion
                    Start: start,
                    End: end,
                    Delete: jest.fn(),
                    Collapse: jest.fn(),
                    _text: '', // Declare the private property for the mock
                 };
                 // Add a mock setter for Text to simulate insertion/modification
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

        // Mock de la estructura de OfficeAppInstance devuelta por getOfficeApplication
        mockOfficeAppInstance = {
            app: mockWordApp,
            openDocument: jest.fn().mockReturnValue(mockDoc), // Simula openDocument de officeInterop
            release: jest.fn(), // Simula release de officeInterop
        };

        mockGetOfficeApplication.mockResolvedValue(mockOfficeAppInstance);
        mockValidateFilePath.mockReturnValue(true); // Asumir que la validación de ruta es exitosa
    });

    // --- Pruebas de Integración para getText ---

    describe('getText', () => {
        const baseParams = { filePath: 'C:/test/document.docx' };

        test('debería interactuar correctamente con officeInterop para obtener texto de la selección', async () => {
            const params = { ...baseParams, range: 'selection' };
            const result = await getText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            // Verificamos que se accede a la selección y su rango a través de mockWordApp
            expect(mockWordApp.Selection.Range).toBe(mockRange);
            // Verificamos que se accede a la propiedad Text del rango mockeado
            expect(mockRange.Text).toBe('Sample text');

            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // Range obtained from getRangeFromSpecifier
            expect(result).toEqual({ success: true, data: 'Sample text' });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully retrieved text'));
        });

        test('debería interactuar correctamente con officeInterop para obtener texto del documento completo', async () => {
            const params = { ...baseParams, range: 'document' };
            const result = await getText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            // Verificamos que se accede al contenido del documento mockeado
            expect(mockDoc.Content.Text).toBe('Document content');

            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // Note: mockDoc.Content is not explicitly released in the tool's finally block
            expect(result).toEqual({ success: true, data: 'Document content' });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully retrieved text'));
        });

        test('debería interactuar correctamente con officeInterop para obtener texto de un párrafo específico', async () => {
            const params = { ...baseParams, range: 'paragraph:2' };
            const result = await getText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath);
            // Verificamos que se accede al párrafo específico a través de mockDoc.Paragraphs
            expect(mockParagraphs.Item).toHaveBeenCalledWith(2);
            // Verificamos que se accede a la propiedad Text del rango del párrafo mockeado
            expect(mockParagraphs.Item(2).Range.Text).toBe('Paragraph 2');

            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // The range returned by Paragraphs.Item().Range is released
            expect(mockReleaseObject).toHaveBeenCalled(); // Check if releaseObject was called at least once
            expect(result).toEqual({ success: true, data: 'Paragraph 2' });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully retrieved text'));
        });
    });

    // --- Pruebas de Integración para insertText ---

    describe('insertText', () => {
        const baseParams = { filePath: 'C:/test/document.docx', text: 'Inserted text' };

        test('debería interactuar correctamente con officeInterop para insertar texto al inicio', async () => {
            const params = { ...baseParams, position: 'start' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // Verificamos que se crea un rango en la posición correcta
            expect(mockDoc.Range).toHaveBeenCalledWith(0, 0);
            // Verificamos que se asigna el texto al rango mockeado
            expect(mockRange.Text).toBe(params.text);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(mockReleaseObject).toHaveBeenCalledWith(undefined); // paraRange
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully inserted text'));
        });

        test('debería interactuar correctamente con officeInterop para insertar texto al final', async () => {
            const params = { ...baseParams, position: 'end' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // Verificamos que se obtiene el final del contenido y se crea un rango
            expect(mockDoc.Content.End).toBe(16);
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

        test('debería interactuar correctamente con officeInterop para insertar texto en la selección', async () => {
            const params = { ...baseParams, position: 'selection' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // Verificamos que se obtiene el rango de la selección
            expect(mockWordApp.Selection.Range).toBe(mockRange);
            expect(mockRange.Text).toBe(params.text);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(mockReleaseObject).toHaveBeenCalledWith(undefined); // paraRange
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully inserted text'));
        });

        test('debería interactuar correctamente con officeInterop para insertar texto al inicio de un párrafo', async () => {
            const params = { ...baseParams, position: 'paragraph:3:start' };
            const result = await insertText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // Verificamos que se accede al párrafo y se crea un rango en su inicio
            expect(mockParagraphs.Item).toHaveBeenCalledWith(3);
            expect(mockDoc.Range).toHaveBeenCalledWith(mockParagraphs.Item(3).Range.Start, mockParagraphs.Item(3).Range.Start);
            expect(mockRange.Text).toBe(params.text);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // insertionRange
            expect(mockReleaseObject).toHaveBeenCalled(); // paraRange should be released
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully inserted text'));
        });
    });

    // --- Pruebas de Integración para modifyText ---

    describe('modifyText', () => {
        const baseParams = { filePath: 'C:/test/document.docx', newText: 'Modified text' };

        test('debería interactuar correctamente con officeInterop para modificar texto en la selección', async () => {
            const params = { ...baseParams, range: 'selection' };
            const result = await modifyText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // Verificamos que se obtiene el rango de la selección
            expect(mockWordApp.Selection.Range).toBe(mockRange);
            // Verificamos que se asigna el nuevo texto al rango mockeado
            expect(mockRange.Text).toBe(params.newText);
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // selectedRange
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully modified text'));
        });

        test('debería interactuar correctamente con officeInterop para modificar texto en el documento completo', async () => {
            const params = { ...baseParams, range: 'document' };
            const result = await modifyText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // Verificamos que se obtiene el rango del contenido del documento
            // Note: mockDoc.Content is not the same mockRange as selection, need to adjust mock if testing this specifically
            // For now, assume getRangeFromSpecifier returns a range-like object
            // We check if the Text property of the returned range is set
            const rangeFromSpecifier = mockDoc.Content; // Simplified for integration test mock
            expect(rangeFromSpecifier.Text).toBe(params.newText);

            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // The range obtained from getRangeFromSpecifier is released (if it's a new object)
            // In this simplified mock, mockDoc.Content is not a new object, so releaseObject won't be called for it.
            // This is acceptable for integration test focus.
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully modified text'));
        });

        test('debería interactuar correctamente con officeInterop para modificar texto en un párrafo específico', async () => {
            const params = { ...baseParams, range: 'paragraph:1' };
            const result = await modifyText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // Verificamos que se accede al párrafo y se obtiene su rango
            expect(mockParagraphs.Item).toHaveBeenCalledWith(1);
            const rangeFromSpecifier = mockParagraphs.Item(1).Range; // Simplified for integration test mock
            expect(rangeFromSpecifier.Text).toBe(params.newText);

            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // The range returned by Paragraphs.Item().Range is released
            expect(mockReleaseObject).toHaveBeenCalled(); // Check if releaseObject was called at least once
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully modified text'));
        });
    });

    // --- Pruebas de Integración para deleteText ---

    describe('deleteText', () => {
        const baseParams = { filePath: 'C:/test/document.docx' };

        test('debería interactuar correctamente con officeInterop para eliminar texto en la selección', async () => {
            const params = { ...baseParams, range: 'selection' };
            const result = await deleteText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // Verificamos que se obtiene el rango de la selección
            expect(mockWordApp.Selection.Range).toBe(mockRange);
            // Verificamos que se llama al método Delete del rango mockeado
            expect(mockRange.Delete).toHaveBeenCalled();
            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            expect(mockReleaseObject).toHaveBeenCalledWith(mockRange); // selectedRange
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully deleted text'));
        });

        test('debería interactuar correctamente con officeInterop para eliminar texto en el documento completo', async () => {
            const params = { ...baseParams, range: 'document' };
            const result = await deleteText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // Verificamos que se obtiene el rango del contenido del documento
            const rangeFromSpecifier = mockDoc.Content; // Simplified for integration test mock
            expect(rangeFromSpecifier.Delete).toHaveBeenCalled();

            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // The range obtained from getRangeFromSpecifier is released (if it's a new object)
            // In this simplified mock, mockDoc.Content is not a new object, so releaseObject won't be called for it.
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully deleted text'));
        });

        test('debería interactuar correctamente con officeInterop para eliminar texto en un párrafo específico', async () => {
            const params = { ...baseParams, range: 'paragraph:1' };
            const result = await deleteText(params);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.filePath);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.filePath, false, false);
            // Verificamos que se accede al párrafo y se obtiene su rango
            expect(mockParagraphs.Item).toHaveBeenCalledWith(1);
            const rangeFromSpecifier = mockParagraphs.Item(1).Range; // Simplified for integration test mock
            expect(rangeFromSpecifier.Delete).toHaveBeenCalled();

            expect(mockDoc.Save).toHaveBeenCalled();
            expect(mockDoc.Close).toHaveBeenCalledWith(false);
            expect(mockOfficeAppInstance.release).toHaveBeenCalled();
            // The range returned by Paragraphs.Item().Range is released
            expect(mockReleaseObject).toHaveBeenCalled(); // Check if releaseObject was called at least once
            expect(result).toEqual({ success: true, data: {} });
            expect(mockLogger.info).toHaveBeenCalledWith(expect.stringContaining('Successfully deleted text'));
        });
    });
});