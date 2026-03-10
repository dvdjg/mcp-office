# Test Roadmap

## Objetivo
Recuperar cobertura funcional útil después de la migración a Jest + ESM, priorizando validaciones reales sobre artefactos Office y dejando explícito qué escenarios siguen pendientes por depender de COM, PowerPoint real o del servidor MCP publicado.

## Cobertura funcional ya activa

### Word -> Markdown y lectura de documentos
- `tests/integration/word/realDocumentConversions.integration.test.ts`
  - Verifica extracción de texto desde un `.docx` complejo generado al vuelo.
  - Verifica exportación a Markdown con encabezados, listas, formato inline y tablas complejas en `html_tables`.
  - Verifica empaquetado ZIP del resultado exportado.
- `tests/integration/word/text.integration.test.ts`
  - Verifica `word/text/get` en ruta de librería para documento completo.
  - Verifica contrato de error para rangos que requieren COM.
  - Verifica que `insert/modify/delete` siguen marcados explícitamente como pendientes en ruta de librería.

### Tablas Word en ruta de librería
- `tests/integration/word/tables.integration.test.ts`
  - Verifica creación de tablas nuevas con `insertTable`.
  - Verifica creación de tablas con contenido mediante `insertTableFromArray`.
  - Verifica lectura posterior del contenido generado.
  - Verifica contratos de error para escenarios no soportados todavía.
- `tests/integration/word/tableAndText.integration.test.ts`
  - Verifica pipeline `insertTableFromArray` -> `exportToMarkdown`.

## Escenarios que siguen pendientes

### COM / Microsoft Word real
- `word/text/insert`, `word/text/modify`, `word/text/delete`
  - TODO: validar edición real sobre documentos existentes y creación de documentos nuevos con Office instalado.
- `word/page`
  - TODO: validar márgenes, orientación, tamaños personalizados y conteo de páginas en Word real.
- `word/saveActiveWordAsMarkdown`
  - TODO: validar selección de documento activo único, múltiples documentos activos y escritura de salida real desde sesión Word abierta.
- `word/applyAutoTitles`
  - TODO: validar integración real con extracción de texto, sugerencias AI y aplicación de estilos sobre párrafos concretos.
- `word/concludeStoryInDocument`
  - TODO: validar resolución de documento activo, extracción del contexto correcto e inserción del cierre generado.

### PowerPoint / conversiones Word -> PPT
- `office/word-to-powerpoint`
  - TODO: validar segmentación por niveles de heading, número de diapositivas, títulos generados y reparto de párrafos entre slides.
- `office/adaptWordToPowerpoint`
  - TODO: validar selección automática del Word activo, derivación de ruta `.pptx`, y fidelidad de estructuras complejas en salida.
- `powerpoint/slides`
  - TODO: validar `add/delete/set/getText/getSlideCount` contra presentaciones reales.

### Extracción estructurada de tablas complejas
- `word/tables/extractData` en ruta de librería
  - TODO: implementar parser HTML fiable para tablas exportadas por Mammoth y cubrir merged cells.

### E2E contra servidor MCP real
- `tests/e2e/word/*.test.ts`
  - TODO: unificar harness para arrancar/parar servidor automáticamente y evitar dependencia manual de `http://localhost:3000`.
  - TODO: ejecutar la misma matriz en modo COM y modo librería cuando el entorno lo permita.

## Criterio de siguiente fase
1. Añadir harness de servidor para E2E.
2. Separar suites `library`, `com`, y `server-e2e`.
3. Añadir fixtures complejos específicos para:
   - listas multinivel
   - tablas con merged cells
   - encabezados/footers
   - imágenes embebidas
   - saltos de sección y orientación mixta
4. Añadir aserciones sobre fidelidad de conversión Word -> PowerPoint.
