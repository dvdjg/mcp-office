import { mergeDocuments } from '../../../src/tools/word/merge.tool';
import { getOfficeApplication, releaseObject } from '../../../src/utils/officeInterop';
import { validateFilePath } from '../../../src/utils/security';
import logger from '../../../src/utils/logger';
import path from 'path';
import { z } from 'zod';
import { FastMCPContext } from '../../../src/types/common.types'; // Import FastMCPContext

// Mock de las dependencias externas (simulando la interacción con Office y el sistema de archivos/seguridad)
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

describe('word/merge integration tests', () => {
    let mockWordApp: any;
    let mockTargetDoc: any;
    let mockSourceDoc1: any;
    let mockSourceDoc2: any;
    let mockTargetDocContentRange: any;
    let mockSourceDocContentRange: any;
    let mockOfficeAppInstance: any;
    let mockContext: FastMCPContext<undefined>;

    beforeEach(() => {
        jest.clearAllMocks();

        // Configurar mocks para simular objetos COM de Word y la interacción con officeInterop
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

        // Mock de la estructura de OfficeAppInstance devuelta por getOfficeApplication
        mockOfficeAppInstance = {
            app: mockWordApp,
            openDocument: jest.fn(), // Will be overridden in tests as needed
            release: jest.fn(), // Simula release de officeInterop
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

    // --- Pruebas de Integración para mergeDocuments ---

    describe('mergeDocuments', () => {
        const baseParams = {
            docs: ['C:/test/doc1.docx', 'C:/test/doc2.docx'],
            output: 'C:/test/output.docx',
        };

        test('debería interactuar correctamente con officeInterop para fusionar documentos', async () => {
            const params = { ...baseParams };
            const result = await mergeDocuments(params, mockContext);

            // Verificar interacciones con officeInterop y mocks de COM
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

            // Verificar el resultado de la herramienta
            expect(result).toEqual({ success: true, data: { outputPath: params.output }, message: 'Successfully merged 2 documents into C:/test/output.docx.' });
            expect(mockContext.log.info).toHaveBeenCalledWith(expect.stringContaining('Starting merge process for 2 documents...'));
            expect(mockContext.log.info).toHaveBeenCalledWith(expect.stringContaining('Merged document saved successfully.'));
        });

        test('debería manejar un error al abrir un documento fuente durante la interacción', async () => {
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

            // Debería intentar cerrar el documento destino y liberar objetos
            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockTargetDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockOfficeAppInstance);

            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error processing source document'));
            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error in mergeDocuments'));
        });

        test('debería manejar un error al guardar el documento destino durante la interacción', async () => {
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
            // Debería intentar cerrar el documento destino y liberar objetos
            expect(mockTargetDoc.Close).toHaveBeenCalledWith(false);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockTargetDoc);
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc1); // Source doc 1 released
            expect(mockReleaseObject).toHaveBeenCalledWith(mockSourceDoc2); // Source doc 2 released
            expect(mockReleaseObject).toHaveBeenCalledWith(mockOfficeAppInstance);

            expect(mockContext.log.error).toHaveBeenCalledWith(expect.stringContaining('Error in mergeDocuments'));
        });
    });
});