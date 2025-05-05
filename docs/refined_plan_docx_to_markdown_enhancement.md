# Plan Refinado: Mejora de la herramienta `word/markdown/export` (Enfoque Programático)

**Fecha:** 2025-05-05

**Autor:** Roo (AI Architect)

**Estado:** Aprobado

## 1. Objetivo

Transformar la herramienta existente `word/markdown/export` en un conversor completo que traduzca con precisión la estructura de documentos Word (incluyendo formato, tablas e imágenes) a Markdown, basándose en el enfoque programático y abordando los TODOs identificados.

## 2. Decisión de Enfoque

Se continuará con el enfoque de **transformación programática** utilizando la API COM de Word, tal como se describe en el plan original (`docs/plan_docx_to_markdown_enhancement.md`) y se refleja en el código actual (`src/tools/word/markdown.tool.ts`). Este enfoque ofrece un control más determinista sobre la conversión y se alinea con el trabajo ya iniciado. El enfoque de extracción a un formato normalizado para generación por IA se considera una posible mejora futura para casos complejos, pero no el método principal en este momento.

## 3. Tareas Prioritarias (TODOs a Abordar)

Los siguientes TODOs identificados en `src/tools/word/markdown.tool.ts` serán abordados, priorizando de la siguiente manera:

1.  **Formato Básico:** Implementar la detección y conversión de formato de texto (negrita, cursiva, tachado).
2.  **Encabezados:** Implementar la detección y conversión de encabezados (basado en estilos de Word como "Heading 1", "Heading 2", etc., o niveles de esquema) a la sintaxis Markdown (`#`, `##`, etc.).
3.  **Listas:** Implementar la detección y conversión de listas (viñetas y numeradas) a la sintaxis Markdown.
4.  **Imágenes:** Implementar la lógica de extracción de imágenes (formas inline y potencialmente otras), guardarlas en el directorio especificado (`imageDir`), generar nombres de archivo únicos y crear los enlaces Markdown relativos correctos. Esto requiere un manejo cuidadoso de la API COM para la extracción de datos de imagen.
5.  **Tablas:** Implementar la conversión de tablas. Para tablas simples, generar sintaxis Markdown estándar. Para tablas con celdas combinadas (`colspan`/`rowspan`) o formato complejo, generar sintaxis HTML `<table>`.
6.  **Comentarios:** Integrar el manejo de comentarios según la opción especificada (`ignore`, `append`, `inline`).
7.  **Otros Elementos:** Considerar el manejo de otros elementos potenciales del documento que puedan ser relevantes para la conversión a Markdown (por ejemplo, cuadros de texto, formas no inline que contengan texto o imágenes).

## 4. Pasos de Implementación Detallados

1.  **Revisión:** Re-leer `docs/plan_docx_to_markdown_enhancement.md` y los `TODO`s en `src/tools/word/markdown.tool.ts`.
2.  **Implementar Lógica de Recorrido:** Asegurar que la función `exportToMarkdown` itera de forma fiable a través de los elementos del documento (párrafos, tablas, formas inline) en el orden visual correcto, utilizando las propiedades `Start` y `End` de los objetos `Range` para controlar el avance.
3.  **Completar `handleParagraph`:**
    *   Añadir lógica para detectar `Paragraph.Range.Font.Bold`, `.Italic`, `.StrikeThrough`, etc., y aplicar la sintaxis Markdown correspondiente.
    *   Añadir lógica para detectar `Paragraph.Range.ListFormat` y generar la sintaxis de lista Markdown (`* `, `- `, `1. `, `2. `).
    *   Añadir lógica para detectar estilos de encabezado (`Paragraph.Style.NameLocal`) o niveles de esquema y generar la sintaxis de encabezado Markdown (`#`, `##`, etc.).
    *   Integrar la llamada a `handleImage` para cada `InlineShape` encontrado dentro del rango del párrafo, reemplazando el texto o marcador de posición con el enlace Markdown de la imagen.
4.  **Completar `handleTable`:**
    *   Implementar la lógica para iterar sobre `Table.Rows` y `Table.Columns`.
    *   Si `tableFormat` es 'markdown', generar la sintaxis de tabla Markdown (`| Header | Header |`, `|---|---|`, `| Cell | Cell |`). Manejar casos simples.
    *   Si `tableFormat` es 'html', generar la sintaxis HTML `<table>`, `<tr>`, `<td>`, utilizando `Cell.MergeInfo` o lógica similar para detectar `colspan` y `rowspan`.
5.  **Completar `handleImage`:**
    *   Implementar la extracción de datos de imagen utilizando métodos COM como `InlineShape.Select()` seguido de `wordApp.Selection.CopyAsPicture()` y pegado desde el portapapeles, o explorando `OLEFormat.Object.SaveAs()`.
    *   Guardar los datos de la imagen en un archivo dentro del `imageDir` especificado, utilizando `fs-extra`.
    *   Generar nombres de archivo únicos (por ejemplo, usando un contador o un hash).
    *   Extraer el texto alternativo (`AlternativeText`) de la forma de imagen si está disponible.
    *   Construir y devolver el enlace Markdown `![alt text](ruta/relativa/imagen.png)` utilizando `path.relative`.
6.  **Integrar Manejo de Comentarios:** Iterar a través de los comentarios del documento (`doc.Comments`) y, según el valor de `commentsOption`, ignorarlos, añadirlos al final del documento Markdown, o intentar insertarlos inline cerca de su ubicación original (esto último puede ser complejo).
7.  **Manejar Otros Elementos:** Añadir lógica para detectar y procesar otros tipos de formas o elementos del documento que puedan contener contenido relevante para la conversión.
8.  **Completar Lógica de Compresión:** Finalizar la implementación de la creación del archivo ZIP utilizando la librería `archiver`, incluyendo el archivo Markdown generado y el directorio de imágenes.
9.  **Pruebas:** Añadir pruebas unitarias exhaustivas para las funciones auxiliares (`handleParagraph`, `handleTable`, `handleImage`) y pruebas de integración para la función `exportToMarkdown` completa, utilizando documentos Word de prueba con diversas estructuras y formatos.
10. **Documentación:** Actualizar la descripción de la herramienta en `wordMarkdownExportTool` y añadir comentarios detallados en el código, especialmente en las secciones que interactúan con la API COM.
11. **Revisión y Refactorización:** Revisar el código implementado para asegurar la robustez, eficiencia y el correcto manejo y liberación de los objetos COM para evitar fugas de memoria.

## 5. Diagrama Mermaid (Flujo de Trabajo Mejorado)

```mermaid
graph TD
    subgraph Inicialización
        A[Inicio: Solicitud de Usuario (filePath, output, imageDir, tableFormat, zipOutput, ...)] --> B{Validar Parámetros};
        B --> C{Resolver Rutas (DOCX de entrada, MD de salida, Directorio de Imágenes)};
        C --> D{Inicializar Aplicación Word y Abrir DOCX vía COM};
        D --> E{Crear Directorio de Imágenes si no existe};
        E --> F[Inicializar Cadena de Salida Markdown y Contador de Imágenes];
    end

    subgraph Bucle de Procesamiento del Documento
        F --> G{Iterar Elementos del Documento (Párrafos, Tablas, Formas)};
        G -- Párrafo --> H[handleParagraph];
        H --> I{Añadir Markdown del Párrafo a la Salida};
        H -- Imagen Encontrada --> J[handleImage];
        J --> K{Añadir Enlace de Imagen a la Salida};

        G -- Tabla --> L[handleTable];
        L --> M{Añadir Markdown/HTML de la Tabla a la Salida};

        G -- Otro Elemento --> N{Manejar Otro Elemento (Ej: Formas)};
        N --> O{Añadir Contenido a la Salida (si aplica)};


        I --> P{Más Elementos?};
        M --> P;
        K --> P;
        O --> P;
        P -- Sí --> G;
    end

    subgraph Finalización
        P -- No --> Q[Manejar Comentarios (Ignorar/Añadir/Inline)];
        Q --> R[Escribir Cadena de Salida Markdown Final a Archivo MD];
        R --> S{zipOutput == true?};
        S -- Sí --> T{Crear Archivo ZIP};
        T --> U[Añadir Archivo MD al ZIP];
        U --> V[Añadir Directorio de Imágenes al ZIP];
        V --> W[Retornar Ruta del Archivo ZIP];
        S -- No --> X[Retornar Ruta del Archivo MD];
    end

    subgraph Limpieza
        W --> Y{Cerrar DOCX y Liberar Objetos COM};
        X --> Y;
        Y --> Z[Fin];
    end

    style J fill:#f9d,stroke:#333,stroke-width:2px
    style L fill:#f9d,stroke:#333,stroke-width:2px
    style N fill:#f9d,stroke:#333,stroke-width:2px
```

## 6. Revisión

Este plan refinado detalla los pasos necesarios para completar la mejora de la herramienta `word/markdown/export` siguiendo el enfoque programático. Aborda los TODOs identificados y proporciona una estructura clara para la implementación.