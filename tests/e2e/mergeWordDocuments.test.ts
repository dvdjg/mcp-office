import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Asumiendo que el servidor se ejecuta en localhost:3000

describe('Caso de Uso 2: Fusionar Dos Documentos Word', () => {
  const inputFile1 = path.join(__dirname, '../fixtures/cuentoAladdin_draft1.docx');
  const inputFile2 = path.join(__dirname, '../fixtures/cuentoAladdin_draft2.docx');
  const outputFilePath = path.join(__dirname, '../fixtures/cuentoAladdin_merged.docx');

  beforeAll(async () => {
    // Configuración: Asegurarse de que los archivos de entrada existen.
    // Estos archivos deberían ser parte de los fixtures del proyecto.
    // Podríamos añadir código aquí para crear archivos .docx de prueba si fuera necesario.
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

    const result = await response.json();

    // Verificar que la herramienta se ejecutó correctamente
    expect(result).toHaveProperty('success', true); // Ajustar según la estructura de respuesta real

    // Verificar que el archivo de salida fue creado
    await expect(fs.access(outputFilePath)).resolves.toBeUndefined();

    // Para verificar el contenido fusionado, necesitaríamos una forma de leer el texto de un archivo .docx
    // o comparar hashes si el contenido esperado es fijo.
    // Por ahora, solo verificaremos la existencia del archivo.
    // TODO: Implementar verificación de contenido más robusta si es posible.
  });
});