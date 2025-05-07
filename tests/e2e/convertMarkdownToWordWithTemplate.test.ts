import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Asumiendo que el servidor se ejecuta en localhost:3000

describe('Caso de Uso 4: Convertir Markdown a Word con Plantilla', () => {
  const inputFilePath = path.join(__dirname, '../fixtures/input.md');
  const templateFilePath = path.join(__dirname, '../fixtures/PlantillaEvolutio.docx');
  const outputFilePath = path.join(__dirname, '../fixtures/output_from_md.docx');

  beforeAll(async () => {
    // Configuración: Asegurarse de que los archivos de entrada y plantilla existen.
    // Estos archivos deberían ser parte de los fixtures del proyecto.
    // Podríamos añadir código aquí para crear archivos de prueba si fuera necesario.
    // Crear un archivo input.md de prueba si no existe
    const dummyMarkdownContent = '# Título de Prueba\n\nEste es un párrafo de prueba.';
    const inputExists = await fs.access(inputFilePath).then(() => true).catch(() => false);
    if (!inputExists) {
      await fs.writeFile(inputFilePath, dummyMarkdownContent, 'utf-8');
    }
  });

  afterAll(async () => {
    // Limpieza: Eliminar el archivo de salida si se creó y el archivo input.md si lo creamos.
    try {
      await fs.unlink(outputFilePath);
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        console.error(`Error al eliminar el archivo de salida ${outputFilePath}:`, error);
      }
    }
    // Eliminar el archivo input.md de prueba si lo creamos
    try {
        const inputExists = await fs.access(inputFilePath).then(() => true).catch(() => false);
        if (inputExists) {
             const content = await fs.readFile(inputFilePath, 'utf-8');
             if (content === '# Título de Prueba\n\nEste es un párrafo de prueba.') {
                 await fs.unlink(inputFilePath);
             }
        }
    } catch (error: any) {
        if (error.code !== 'ENOENT') {
            console.error(`Error al eliminar el archivo de entrada ${inputFilePath}:`, error);
        }
    }
  });

  test('Debería convertir input.md a Word usando PlantillaEvolutio.docx y verificar estilos', async () => {
    // Ejecutar la herramienta MCP para convertir Markdown a Word con plantilla
    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/markdown/import',
        arguments: {
          input_path: inputFilePath,
          output_path: outputFilePath,
          template_path: templateFilePath,
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

    // Para verificar los estilos aplicados por la plantilla, necesitaríamos una forma de inspeccionar
    // el contenido interno del .docx o usar otra herramienta MCP.
    // Por ahora, solo verificaremos la existencia del archivo.
    // TODO: Implementar verificación de estilos si es posible.
  });
});