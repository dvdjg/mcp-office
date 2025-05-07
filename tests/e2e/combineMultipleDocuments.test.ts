import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Asumiendo que el servidor se ejecuta en localhost:3000

describe('Caso de Uso 7: Combinar Múltiples Documentos', () => {
  const inputFilesDir = path.join(__dirname, '../fixtures/Oferta');
  const outputFilePath = path.join(__dirname, '../fixtures/Oferta_combined.docx');

  beforeAll(async () => {
    // Configuración: Asegurarse de que el directorio de entrada existe y contiene archivos.
    // Este directorio y sus archivos deberían ser parte de los fixtures del proyecto.
    // Podríamos añadir código aquí para crear archivos de prueba si fuera necesario.
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

  test('Debería combinar archivos en el directorio Oferta y verificar la salida Word', async () => {
    // Ejecutar la herramienta MCP para combinar documentos
    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'office/combine',
        arguments: {
          input_dir: inputFilesDir,
          output_path: outputFilePath,
          output_format: 'docx', // Asumiendo que la herramienta permite especificar el formato de salida
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

    // Para verificar el contenido combinado, necesitaríamos una forma de leer el texto de un archivo .docx
    // o comparar hashes si el contenido esperado es fijo.
    // Por ahora, solo verificaremos la existencia del archivo.
    // TODO: Implementar verificación de contenido más robusta si es posible.
  });
});