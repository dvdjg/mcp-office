import { mergeDocuments } from '../../../src/tools/word/merge.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import path from 'path';
import { z } from 'zod';
import { FastMCPContext } from '../../../src/types/common.types'; // Import FastMCPContext

// Mock de las dependencias
jest.mock('../../../src/utils/officeInterop');
jest.mock('../../../src/utils/security');
jest.mock('../../../src/utils/errorHandler', () => ({
    handleToolError: jest.fn((error, code) => ({ success: false, error: { code, message: error.message } })),
    createErrorResponse: jest.fn((message, code, details) => ({ success: false, error: { code, message, details } })),
}));
jest.mock('../../../src/utils/logger');
jest.mock('path'); // Mock the path module

const mockGetOfficeApplication = getOfficeApplication as jest.Mock;
const mockReleaseObject = releaseObject as jest.Mock;
const mockValidateFilePath = validateFilePath as jest.Mock;
const mockLogger = logger as jest.Mocked<typeof logger>;
const mockPath = path as jest.Mocked<typeof path>;

describe('word/merge unit tests', () => {
    let mockWordApp: any;
    let mockTargetDoc: any;
    let mockSourceDoc1: any;
    let mockSourceDoc2: any;
    let mockTargetDocContentRange: any;
    let mockSourceDocContentRange: any;
    let mockOfficeAppInstance: any;
    let mockContext: FastMCPContext<undefined>;

    beforeEach(() => {
        // Reset mocks antes de cada prueba
        jest.clearAllMocks();

        // Configurar mocks básicos para Word COM objects
        mockTargetDocContentRange = {
            End: 100,
            Collapse: jest.fn(),
            Paste: jest.fn(),
            InsertBreak: jest.fn(),
        };
        mockSourceDocContentRange = {
            Copy: jest.fn(),
        };
        mockTargetDoc = {
            Activate: jest.fn(),
            Content: mockTargetDocContentRange,
            SaveAs2: jest.fn(),
            Close: jest.fn(),
        };
        mockSourceDoc1 = {
            Content: mockSourceDocContentRange,
            Close: jest.fn(),
        };
        mockSourceDoc2 = {
            Content: mockSourceDocContentRange,
            Close: jest.fn(),
        };
        mockWordApp = {
            Documents: {
                Add: jest.fn().mockReturnValue(mockTargetDoc),
                Open: jest.fn(),
            },
            Visible: false,
            DisplayAlerts: 0,
        };

        // Mock the OfficeAppInstance structure
        mockOfficeAppInstance = {
            app: mockWordApp,
            openDocument: jest.fn(), // Will be overridden in tests as needed
            release: jest.fn(),
        };

        // Default mock implementations
        mockGetOfficeApplication.mockResolvedValue(mockOfficeAppInstance);
        mockValidateFilePath.mockReturnValue(true); // Assume valid paths by default
        mockPath.dirname.mockImplementation((filePath) => `C:/test/dir/${path.basename(filePath)}/..`); // Simple mock for dirname

        // Mock openDocument to return specific source docs based on path
        mockOfficeAppInstance.openDocument.mockImplementation((filePath: string) => {
            if (filePath.includes('doc1.docx')) return mockSourceDoc1;
            if (filePath.includes('doc2.docx')) return mockSourceDoc2;
            // For the target doc created by Add(), we don't need openDocument to return it
            return null; // Default for other paths if needed
        });

        // Mock context with log and reportProgress
        mockContext = {
            log: {
                info: jest.fn(),
                warn: jest.fn(),
                error: jest.fn(),
                debug: jest.fn(),
            },
            reportProgress: jest.fn(),
            userId: 'testUser', // Example userId
            // Add other context properties if needed by the tool
        } as unknown as FastMCPContext<undefined>; // Cast to match expected type
    });

    // --- Pruebas Unitarias para mergeDocuments ---

    describe('mergeDocuments', () => {
        const baseParams = {
            docs: ['C:/test/doc1.docx', 'C:/test/doc2.docx'],
            output: 'C:/test/output.docx',
        };

        test('debería fusionar documentos exitosamente', async () => {
            const params = { ...baseParams };
            const result = await mergeDocuments(params, mockContext);

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.output, expect.any(Array)); // Check output path validation
            expect(mockPath.dirname).toHaveBeenCalledWith(params.output); // Check dirname call
            expect(mockValidateFilePath).toHaveBeenCalledWith('C:/test/dir/output.docx/..', expect.any(Array)); // Check output dir validation
            expect(mockValidateFilePath).toHaveBeenCalledWith(params.docs[0], expect.any(Array)); // Check source path validation 1
            expect(mockValidateFilePath).toHaveBeenCalledWith(params.docs[1], expect.any(Array)); // Check source path validation 2

            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Add).toHaveBeenCalled(); // Target doc created
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.docs[0], false, true); // Source doc 1 opened
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.docs[1], false, true); // Source doc 2 opened

            expect(mockSourceDocContentRange.Copy).toHaveBeenCalledTimes(2); // Content copied from both source docs
            expect(mockTargetDoc.Activate).toHaveBeenCalledTimes(2); // Target doc activated before each paste
            expect(mockTargetDocContentRange.Collapse).toHaveBeenCalledWith(0); // Collapsed to end before paste
            expect(mockTargetDocContentRange.Paste).toHaveBeenCalledTimes(2); // Content pasted twice

            expect(mockTargetDocContentRange.InsertBreak).toHaveBeenCalledTimes(1); // Page break after first doc

            expect(mockTargetDoc.SaveAs2).toHaveBeenCalledWith(params.output); // Target doc saved
            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false); // Target doc closed

            expect(mockSourceDoc1.Close).toHaveBeenCalledWith(false); // Source doc 1 closed
            expect(mockSourceDoc2.Close).toHaveBeenCalledWith(false); // Source doc 2 closed

            expect(mockReleaseObject).toHaveBeenCalledWith(mockTargetDoc); // Target doc released
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc1); // Source doc 1 released
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc2); // Source doc 2 released
            expect(mockReleaseObject).toHaveBeenCalledWith(mockOfficeAppInstance); // Office app instance released

            expect(mockContext.reportProgress).toHaveBeenCalledWith({ progress: 0, total: 2 }); // Initial progress
            expect(mockContext.reportProgress).toHaveBeenCalledWith({ progress: 1, total: 2 }); // Progress after doc 1
            expect(mockContext.reportProgress).toHaveBeenCalledWith({ progress: 2, total: 2 }); // Progress after doc 2 and saving

            expect(result).toEqual({ success: true, data: { outputPath: params.output }, message: 'Successfully merged 2 documents into C:/test/output.docx.' });
            expect(mockContext.log.info).toHaveBeenCalledWith(expect.stringContaining('Starting merge process for 2 documents...'));
            expect(mockContext.log.info).toHaveBeenCalledWith(expect.stringContaining('Merged document saved successfully.'));
        });

        test('debería manejar un array de docs vacío', async () => {
            const params = { ...baseParams, docs: [] };
            // Zod validation should catch this before the handler logic
            try {
                mergeDocuments(params, mockContext);
            } catch (error: any) {
                expect(error).toBeInstanceOf(z.ZodError);
                expect(error.errors[0].message).toContain('At least two documents are required');
            }

            // No debería llamar a ninguna función COM si la validación falla
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
            expect(mockOfficeAppInstance.openDocument).not.toHaveBeenCalled();
            expect(mockWordApp.Documents.Add).not.toHaveBeenCalled();
            expect(mockContext.log.error).toHaveBeenCalled(); // Zod error should be logged by the tool wrapper
        });

        test('debería manejar un filePath de salida inválido', async () => {
            mockValidateFilePath.mockImplementation((filePath) => filePath !== baseParams.output); // Simulate output path validation failure
            const params = { ...baseParams };
            try {
                await mergeDocuments(params, mockContext);
            } catch (error: any) {
                 expect(error.message).toContain(`Output path validation failed for: ${params.output}`);
            }

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.output, expect.any(Array));
            // No debería proceder si la validación falla
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error in mergeDocuments'));
        });

         test('debería manejar un filePath de fuente inválido', async () => {
            mockValidateFilePath.mockImplementation((filePath) => filePath !== baseParams.docs[0]); // Simulate first source path validation failure
            const params = { ...baseParams };
            try {
                await mergeDocuments(params, mockContext);
            } catch (error: any) {
                 expect(error.message).toContain(`Source path validation failed for: ${params.docs[0]}`);
            }

            expect(mockValidateFilePath).toHaveBeenCalledWith(params.output, expect.any(Array));
            expect(mockValidateFilePath).toHaveBeenCalledWith(params.docs[0], expect.any(Array));
            // No debería proceder si la validación falla
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error in mergeDocuments'));
        });


        test('debería manejar un error al obtener la aplicación Word', async () => {
            mockGetOfficeApplication.mockRejectedValue(new Error('Failed to get Word app'));
            const params = { ...baseParams };
            try {
                await mergeDocuments(params, mockContext);
            } catch (error: any) {
                expect(error.message).toContain('Failed to get Word app');
            }

            expect(mockValidateFilePath).toHaveBeenCalledTimes(3); // Output and two source docs
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            // No debería intentar abrir documentos ni crear el destino
            expect(mockWordApp.Documents.Add).not.toHaveBeenCalled();
            expect(mockOfficeAppInstance.openDocument).not.toHaveBeenCalled();
            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error in mergeDocuments'));
        });

        test('debería manejar un error al crear el documento destino', async () => {
            mockWordApp.Documents.Add.mockReturnValue(null); // Simular fallo al crear
            const params = { ...baseParams };
            try {
                await mergeDocuments(params, mockContext);
            } catch (error: any) {
                expect(error.message).toContain('Failed to create target document.');
            }

            expect(mockValidateFilePath).toHaveBeenCalledTimes(3);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Add).toHaveBeenCalled();
            // No debería intentar abrir documentos fuente
            expect(mockOfficeAppInstance.openDocument).not.toHaveBeenCalled();
            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error in mergeDocuments'));
            expect(mockReleaseObject).toHaveBeenCalledWith(mockOfficeAppInstance); // Word app should be released
        });

        test('debería manejar un error al abrir un documento fuente', async () => {
            // Simular fallo al abrir el primer documento fuente
            mockOfficeAppInstance.openDocument.mockImplementation((filePath: string) => {
                if (filePath.includes('doc1.docx')) throw new Error('Failed to open doc1');
                if (filePath.includes('doc2.docx')) return mockSourceDoc2;
                return null;
            });

            const params = { ...baseParams };
            try {
                await mergeDocuments(params, mockContext);
            } catch (error: any) {
                expect(error.message).toContain('Failed to open doc1');
            }

            expect(mockValidateFilePath).toHaveBeenCalledTimes(3);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Add).toHaveBeenCalled(); // Target doc created
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.docs[0], false, true); // Attempted to open doc1
            expect(mockOfficeAppInstance.openDocument).not.toHaveBeenCalledWith(params.docs[1], false, true); // Did not attempt to open doc2

            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false); // Target doc should be closed
            expect(mockReleaseObject).toHaveBeenCalledWith(mockTargetDoc); // Target doc released
            expect(mockReleaseObject).toHaveBeenCalledWith(mockOfficeAppInstance); // Office app instance released
            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error processing source document'));
            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error in mergeDocuments'));
        });

        test('debería manejar un error al guardar el documento destino', async () => {
            mockTargetDoc.SaveAs2.mockImplementation(() => { throw new Error('Failed to save'); });
            const params = { ...baseParams };
            try {
                await mergeDocuments(params, mockContext);
            } catch (error: any) {
                expect(error.message).toContain('Failed to save');
            }

            expect(mockValidateFilePath).toHaveBeenCalledTimes(3);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Add).toHaveBeenCalled();
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledTimes(2); // Both source docs opened
            expect(mockSourceDocContentRange.Copy).toHaveBeenCalledTimes(2);
            expect(mockTargetDocContentRange.Paste).toHaveBeenCalledTimes(2);
            expect(mockTargetDocContentRange.InsertBreak).toHaveBeenCalledTimes(1);

            expect(mockTargetDoc.SaveAs2).toHaveBeenCalledWith(params.output); // Attempted to save
            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false); // Target doc should be closed
            expect(mockReleaseObject).toHaveBeenCalledWith(mockTargetDoc); // Target doc released
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc1); // Source doc 1 released
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc2); // Source doc 2 released
            expect(mockReleaseObject).toHaveBeenCalledWith(mockOfficeAppInstance); // Office app instance released
            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error in mergeDocuments'));
        });

        test('debería liberar documentos fuente si ocurre un error durante el procesamiento', async () => {
             // Simular fallo al abrir el segundo documento fuente
            mockOfficeAppInstance.openDocument.mockImplementation((filePath: string) => {
                if (filePath.includes('doc1.docx')) return mockSourceDoc1;
                if (filePath.includes('doc2.docx')) throw new Error('Failed to open doc2');
                return null;
            });

            const params = { ...baseParams };
            try {
                await mergeDocuments(params, mockContext);
            } catch (error: any) {
                expect(error.message).toContain('Failed to open doc2');
            }

            expect(mockValidateFilePath).toHaveBeenCalledTimes(3);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Add).toHaveBeenCalled(); // Target doc created
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.docs[0], false, true); // doc1 opened successfully
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledWith(params.docs[1], false, true); // Attempted to open doc2

            expect(mockSourceDoc1.Close).toHaveBeenCalledWith(false); // doc1 should be closed
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc1); // doc1 should be released

            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false); // Target doc should be closed
            expect(mockReleaseObject).toHaveBeenCalledWith(mockTargetDoc); // Target doc released
            expect(mockReleaseObject).toHaveBeenCalledWith(mockOfficeAppInstance); // Office app instance released

            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error processing source document'));
            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error in mergeDocuments'));
        });

        test('debería usar el logger del contexto si está disponible', async () => {
            const params = { ...baseParams };
            await mergeDocuments(params, mockContext);

            expect(mockContext.log.info).toHaveBeenCalled();
            expect(mockLogger.info).not.toHaveBeenCalled(); // Global logger should not be used
        });

         test('debería usar el logger global si el contexto no está disponible', async () => {
            const params = { ...baseParams };
            // Call without context
            const result = await mergeDocuments(params);

            expect(mockLogger.info).toHaveBeenCalled(); // Global logger should be used
            expect(mockContext.log.info).not.toHaveBeenCalled(); // Context logger should not be used

            expect(mockValidateFilePath).toHaveBeenCalledTimes(3);
            expect(mockGetOfficeApplication).toHaveBeenCalledWith('Word.Application');
            expect(mockWordApp.Documents.Add).toHaveBeenCalled();
            expect(mockOfficeAppInstance.openDocument).toHaveBeenCalledTimes(2);
            expect(mockTargetDoc.SaveAs2).toHaveBeenCalledWith(params.output);
            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false);
            expect(mockSourceDoc1.Close).toHaveBeenCalledWith(false);
            expect(mockSourceDoc2.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledTimes(4); // Target, Source1, Source2, OfficeAppInstance
            expect(result).toEqual({ success: true, data: { outputPath: params.output }, message: 'Successfully merged 2 documents into C:/test/output.docx.' });
        });

        test('debería lanzar un error si el contexto está incompleto', async () => {
            const params = { ...baseParams };
            const incompleteContext = { log: { info: jest.fn() } } as unknown as FastMCPContext<undefined>; // Missing reportProgress

            try {
                await mergeDocuments(params, incompleteContext);
            } catch (error: any) {
                expect(error.message).toContain('Tool context is missing required properties');
            }

            expect(incompleteContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Merge tool received context but it is missing required properties'));
            // No debería llamar a ninguna función COM
            expect(mockGetOfficeApplication).not.toHaveBeenCalled();
        });
    });
});