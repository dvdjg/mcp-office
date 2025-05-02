import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Asumiendo que el servidor se ejecuta en localhost:3000

describe('Caso de Uso 9: Analizar Documento y Añadir Comentarios', () => {
  const inputFilePath = path.join(__dirname, '../fixtures/PropuestaTécnica.docx');
  const outputFilePath = path.join(__dirname, '../fixtures/PropuestaTécnica_analyzed.docx'); // Asumiendo que la herramienta crea un nuevo archivo o modifica el existente

  beforeAll(async () => {
    // Configuración: Asegurarse de que el archivo de entrada existe.
    // Este archivo debería ser parte de los fixtures del proyecto.
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

  test('Debería analizar PropuestaTécnica.docx y verificar la adición de comentarios', async () => {
    // Ejecutar la herramienta MCP para analizar el documento y añadir comentarios
    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/analyze',
        arguments: {
          input_path: inputFilePath,
          output_path: outputFilePath, // Asumiendo que la herramienta 'analyze' tiene un output_path
          // Asumiendo que la herramienta 'word/analyze' añade comentarios automáticamente
          // o que se puede especificar qué tipo de análisis realizar.
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

    // Para verificar la adición de comentarios, necesitaríamos una forma de inspeccionar
    // el contenido interno del .docx y buscar comentarios, o usar otra herramienta MCP.
    // Por ahora, solo verificaremos la existencia del archivo.
    // TODO: Implementar verificación de comentarios si es posible.
  });
});