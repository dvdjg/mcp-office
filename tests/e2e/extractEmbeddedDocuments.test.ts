import * as fs from 'fs/promises';
import * as path from 'path';
import fetch from 'node-fetch';

const MCP_SERVER_URL = 'http://localhost:3000'; // Asumiendo que el servidor se ejecuta en localhost:3000

describe('Caso de Uso 6: Extraer Documentos Incrustados', () => {
  const inputFilePath = path.join(__dirname, '../fixtures/Composición.docx');
  const outputDirPath = path.join(__dirname, '../fixtures/extracted_objects');

  beforeAll(async () => {
    // Configuración: Asegurarse de que el archivo de entrada existe.
    // Este archivo debería ser parte de los fixtures del proyecto y contener objetos incrustados.
    // Podríamos añadir código aquí para crear un archivo .docx de prueba con objetos incrustados si fuera necesario.
    // Asegurarse de que el directorio de salida no existe antes de la prueba.
    try {
      await fs.rm(outputDirPath, { recursive: true });
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        console.error(`Error al limpiar el directorio de salida ${outputDirPath}:`, error);
      }
    }
  });

  afterAll(async () => {
    // Limpieza: Eliminar el directorio de salida y su contenido.
    try {
      await fs.rm(outputDirPath, { recursive: true });
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        console.error(`Error al eliminar el directorio de salida ${outputDirPath}:`, error);
      }
    }
  });

  test('Debería extraer archivos incrustados de Composición.docx y verificar que se guardan', async () => {
    // Ejecutar la herramienta MCP para extraer objetos incrustados
    const response = await fetch(`${MCP_SERVER_URL}/tool`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        tool_name: 'word/embedded-objects',
        arguments: {
          input_path: inputFilePath,
          output_dir: outputDirPath,
        },
      }),
    });

    // Verificar que la solicitud fue exitosa
    expect(response.ok).toBe(true);

    const result = await response.json();

    // Verificar que la herramienta se ejecutó correctamente
    expect(result).toHaveProperty('success', true); // Ajustar según la estructura de respuesta real
    // Asumiendo que la respuesta incluye información sobre los archivos extraídos, por ejemplo:
    // expect(result).toHaveProperty('extracted_files');
    // expect(result.extracted_files.length).toBeGreaterThan(0); // Verificar que se extrajo al menos un archivo

    // Verificar que el directorio de salida fue creado
    await expect(fs.access(outputDirPath)).resolves.toBeUndefined();

    // Verificar que se crearon archivos dentro del directorio de salida.
    // Esto requiere conocimiento de cuántos y qué tipo de archivos se esperan.
    // Por ahora, solo verificaremos que el directorio no está vacío.
    const extractedFiles = await fs.readdir(outputDirPath);
    expect(extractedFiles.length).toBeGreaterThan(0);

    // TODO: Implementar verificación más específica de los archivos extraídos si es posible.
  });
});