import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Asumiendo que el servidor se ejecuta en localhost:3000
const FIXTURES_DIR = path.join(__dirname, '../fixtures');
const inputFile1 = path.join(FIXTURES_DIR, 'cuentoAladdin_draft1.docx');
const inputFile2 = path.join(FIXTURES_DIR, 'cuentoAladdin_draft2.docx');
const outputFilePath = path.join(FIXTURES_DIR, 'cuentoAladdin_merged.docx');
const nonExistentFile = path.join(FIXTURES_DIR, 'non_existent_file.docx');


// Define a basic type for the expected successful response
interface SuccessResponse {
  success: true;
  data: any; // Use a more specific type if the data structure is known
  message?: string;
}

// Define a basic type for the expected error response
interface ErrorResponse {
  success: false;
  error: {
    code: string;
    message: string;
    details?: any;
  };
}

type ToolResponse = SuccessResponse | ErrorResponse;


describe('Caso de Uso 2: Fusionar Dos Documentos Word', () => {

  beforeAll(async () => {
    // Configuración: Asegurarse de que los archivos de entrada existen.
    // Estos archivos deberían ser parte de los fixtures del proyecto.
    // Check if fixture files exist
    const fixtures = [inputFile1, inputFile2];
    for (const fixturePath of fixtures) {
      try {
        await fs.stat(fixturePath);
      } catch (error: any) {
        if (error.code === 'ENOENT') {
           console.error(`Fixture file ${fixturePath} not found. Please ensure it exists for merge tests.`);
           // Depending on test setup, might throw or skip tests
        } else {
          console.error(`Error checking fixture file ${fixturePath}:`, error);
        }
      }
    }
  });

  afterAll(async () => {
    // Limpieza: Eliminar el archivo de salida si se creó.
    try {
      await fs.unlink(outputFilePath);
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        console.error(`Error al eliminar el archivo de salida ${outputFilePath}:`, error);
      }
    }
  });

  test('Debería fusionar dos borradores de cuentoAladdin.docx y verificar el contenido', async () => {
    // Ejecutar la herramienta MCP para fusionar documentos
    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/merge',
        arguments: {
          input_paths: [inputFile1, inputFile2],
          output_path: outputFilePath,
        },
      }),
    });

    // Verificar que la solicitud fue exitosa
    expect(response.ok).toBe(true);

    const result = await response.json() as ToolResponse; // Type assertion here

    // Verificar que la herramienta se ejecutó correctamente
    expect(result.success).toBe(true); // Now TypeScript knows 'result' has 'success'

    if (result.success) {
      expect(result.data).toHaveProperty('outputPath', outputFilePath);

      // Verificar que el archivo de salida fue creado
      await expect(fs.access(outputFilePath)).resolves.toBeUndefined();

      // Para verificar el contenido fusionado, necesitaríamos una forma de leer el texto de un archivo .docx
      // o comparar hashes si el contenido esperado es fijo.
      // Por ahora, solo verificaremos la existencia del archivo.
      // TODO: Implementar verificación de contenido más robusta si es posible.
    } else {
      fail('Expected merge to succeed but it failed.');
    }
  });

  test('Debería manejar un error al intentar fusionar archivos no existentes', async () => {
    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/merge',
        arguments: {
          input_paths: [inputFile1, nonExistentFile], // One existing, one non-existent
          output_path: outputFilePath,
        },
      }),
    });

    // Expect a non-OK response or a success: false in the body depending on server error handling
    // For this test, we expect the tool to report failure.
    expect(response.ok).toBe(true); // Assuming server returns 200 even on tool error

    const result = await response.json() as ToolResponse;

    expect(result.success).toBe(false);

    if (!result.success) {
      expect(result.error).toHaveProperty('code');
      expect(result.error).toHaveProperty('message');
      // TODO: Check for a specific error code related to file not found if the API defines one
      expect(result.error.message).toContain('Error processing source document'); // Basic check based on integration test error message
    } else {
      fail('Expected merge to fail but it succeeded.');
    }
  });

  // TODO: Add tests for other word/merge variations if applicable (e.g., different numbers of files)
});