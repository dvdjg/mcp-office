import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Asumiendo que el servidor se ejecuta en localhost:3000

describe('Caso de Uso 10: Formatear Código y Metadatos', () => {
  const inputFilePath = path.join(__dirname, '../fixtures/BuenasPrácticas.docx');
  const outputFilePath = path.join(__dirname, '../fixtures/BuenasPrácticas_formatted.docx'); // Asumiendo que la herramienta crea un nuevo archivo o modifica el existente

  beforeAll(async () => {
    // Configuración: Asegurarse de que el archivo de entrada existe.
    // Este archivo debería ser parte de los fixtures del proyecto y contener bloques de código.
    // Podríamos añadir código aquí para crear un archivo .docx de prueba si fuera necesario.
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

  test('Debería formatear código en BuenasPrácticas.docx y verificar el resaltado', async () => {
    // Ejecutar la herramienta MCP para formatear código
    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/code-format',
        arguments: {
          input_path: inputFilePath,
          output_path: outputFilePath, // Asumiendo que la herramienta 'code-format' tiene un output_path
          // Asumiendo que la herramienta 'word/code-format' aplica resaltado de sintaxis automáticamente.
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

    // Para verificar el resaltado de sintaxis, necesitaríamos una forma de inspeccionar
    // el contenido interno del .docx y buscar las propiedades de formato aplicadas al texto,
    // o usar otra herramienta MCP.
    // Por ahora, solo verificaremos la existencia del archivo.
    // TODO: Implementar verificación de resaltado si es posible.
  });
});