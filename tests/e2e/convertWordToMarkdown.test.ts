import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch'; // O 'fetch' si está disponible globalmente

const MCP_SERVER_URL = 'http://localhost:3000'; // Asumiendo que el servidor se ejecuta en localhost:3000

describe('Caso de Uso 1: Convertir Word a Markdown con Comentarios', () => {
  const inputFilePath = path.join(__dirname, '../fixtures/CV.docx');
  const outputFilePath = path.join(__dirname, '../fixtures/CV.md');

  beforeAll(async () => {
    // Configuración: Asegurarse de que el archivo de entrada existe.
    // En un escenario real, este archivo debería ser parte de los fixtures del proyecto.
    // Por ahora, asumiremos que CV.docx existe en tests/fixtures.
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

  test('Debería convertir CV.docx a Markdown y verificar los comentarios', async () => {
    // Ejecutar la herramienta MCP haciendo una solicitud HTTP al servidor
    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/markdown/export',
        arguments: {
          input_path: inputFilePath,
          output_path: outputFilePath,
          include_comments: true,
        },
      }),
    });

    // Verificar que la solicitud fue exitosa
    expect(response.ok).toBe(true);

    const result = await response.json();

    // Verificar que la herramienta se ejecutó correctamente (esto puede variar dependiendo de la respuesta de la herramienta)
    // Asumiendo que la respuesta JSON incluye un campo 'success' o similar.
    expect(result).toHaveProperty('success', true); // Ajustar según la estructura de respuesta real del servidor MCP

    // Verificar que el archivo de salida fue creado
    await expect(fs.access(outputFilePath)).resolves.toBeUndefined();

    // Leer el contenido del archivo de salida
    const outputContent = await fs.readFile(outputFilePath, 'utf-8');

    // Verificar que el contenido incluye comentarios (esto dependerá del contenido de CV.docx y cómo la herramienta exporta comentarios)
    // Esta es una verificación básica. Se necesitaría un análisis más detallado del formato de salida esperado.
    expect(outputContent).toContain('<!--'); // Ejemplo: buscar el inicio de un comentario HTML/Markdown
    expect(outputContent).toContain('-->'); // Ejemplo: buscar el final de un comentario HTML/Markdown

    // Aquí se podrían añadir verificaciones más específicas sobre el contenido y los comentarios esperados.
  });
});