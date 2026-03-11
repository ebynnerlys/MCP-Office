## Informe Técnico: Editor IA para Word, Excel y PowerPoint con MCP

### Resumen Ejecutivo

La opción recomendada para construir un editor asistido por IA para Word, Excel y PowerPoint en entorno Windows es:

- Python como lenguaje principal.
- Un servidor MCP local como capa de herramientas para la IA.
- Automatización de Office mediante COM usando `pywin32`.
- Un diseño basado en herramientas deterministas, no en edición libre por prompt sobre archivos binarios.

Esta combinación ofrece el mejor equilibrio entre:

- velocidad de desarrollo,
- capacidad real de edición sobre documentos existentes,
- integración con IA,
- posibilidad de evolucionar después a una arquitectura más robusta.

Node.js también es viable, especialmente para construir el servidor MCP y una capa API, pero para automatización profunda de Office en Windows Python sigue siendo más cómodo y más fiable en la práctica.

### Objetivo del Sistema

Construir un sistema capaz de:

- leer documentos de Word, libros de Excel y presentaciones de PowerPoint,
- entender su estructura,
- ejecutar cambios solicitados por lenguaje natural,
- exponer esas capacidades a una IA mediante MCP,
- mantener seguridad, trazabilidad y capacidad de recuperación ante errores.

### Recomendación Principal

#### Stack recomendado

- Lenguaje principal: Python 3.11+
- Automatización Office: `pywin32`
- Servidor MCP: Python
- Entorno: Windows 10/11 con Microsoft Office instalado
- Modelo IA: externo al sistema, consumiendo las tools MCP
- Librerías auxiliares recomendadas:
  - `pydantic`
  - `python-dotenv`
  - `tenacity`
  - `rich`
  - `loguru` o `structlog`

#### Motivo de esta elección

Python permite construir rápido, integrar IA fácilmente y trabajar bien con automatización local de Office. `pywin32` da acceso directo al modelo COM de Word, Excel y PowerPoint, que es la vía más potente si se necesita editar archivos existentes con buena fidelidad.

### Arquitectura Recomendada

#### Principio clave

La IA no debe editar archivos directamente.

La IA debe:

- inspeccionar estructura,
- decidir una acción,
- llamar una herramienta MCP,
- validar el resultado,
- guardar copia o exportar vista previa.

#### Capas del sistema

1. Capa IA

- recibe la instrucción del usuario,
- decide el plan,
- usa las tools del MCP.

2. Servidor MCP

- expone herramientas estables,
- valida parámetros,
- controla permisos,
- registra operaciones.

3. Capa de automatización Office

- Word por COM,
- Excel por COM,
- PowerPoint por COM.

4. Capa de seguridad y control

- copias de seguridad,
- guardado incremental,
- logs,
- validación de rutas,
- confirmación para operaciones destructivas.

### Por qué COM es la mejor base en Windows

El mayor problema de Word, Excel y PowerPoint no es abrir el archivo, sino editarlo con fidelidad cuando ya existe y tiene estructura compleja.

COM aporta:

- acceso al modelo real de Office,
- compatibilidad con documentos existentes,
- soporte para elementos complejos,
- exportación a PDF e imágenes,
- acceso a selección, rangos, shapes, gráficos, tablas y estilos.

Frente a librerías puras de archivo, COM suele ser mejor cuando el caso de uso es edición real, no solo generación desde cero.

### Herramientas MCP Recomendadas

#### Word

- `word_open_document(path)`
- `word_get_structure(path)`
- `word_get_selection(path)`
- `word_get_paragraphs(path)`
- `word_replace_text(path, find, replace)`
- `word_rewrite_paragraph(path, paragraph_id, instruction)`
- `word_apply_style(path, range_id, style_name)`
- `word_insert_comment(path, range_id, comment)`
- `word_track_changes(path, enabled)`
- `word_export_pdf(path, out_path)`
- `word_save_as(path, out_path)`

#### Excel

- `excel_open_workbook(path)`
- `excel_list_sheets(path)`
- `excel_get_used_range(path, sheet)`
- `excel_read_range(path, sheet, range)`
- `excel_write_range(path, sheet, range, values)`
- `excel_get_formulas(path, sheet, range)`
- `excel_add_formula(path, sheet, cell, formula)`
- `excel_create_chart(path, sheet, source_range, chart_type)`
- `excel_summarize_sheet(path, sheet)`
- `excel_export_pdf(path, out_path)`
- `excel_save_as(path, out_path)`

#### PowerPoint

- `ppt_open_presentation(path)`
- `ppt_list_slides(path)`
- `ppt_get_slide_text(path, slide_index)`
- `ppt_get_slide_shapes(path, slide_index)`
- `ppt_replace_text(path, slide_index, find, replace)`
- `ppt_add_slide(path, layout)`
- `ppt_insert_image(path, slide_index, image_path, x, y, w, h)`
- `ppt_apply_theme(path, theme_path)`
- `ppt_export_pdf(path, out_path)`
- `ppt_export_slide_images(path, out_dir)`
- `ppt_save_as(path, out_path)`

### Diseño Correcto de las Respuestas de las Tools

Cada tool no debe devolver solo "hecho". Debe devolver estructura útil para la IA.

Ejemplos:

- Word: títulos, secciones, tablas, comentarios, cambios detectados.
- Excel: hojas, rangos usados, fórmulas, errores, nombres de columnas.
- PowerPoint: número de slides, placeholders, shapes, notas, títulos.

Esto permite que la IA trabaje con contexto estructurado en lugar de actuar a ciegas.

### Flujo de Trabajo Recomendado

#### Flujo seguro por defecto

1. Abrir documento.
2. Inspeccionar estructura.
3. Crear copia de trabajo.
4. Ejecutar cambios sobre la copia.
5. Exportar vista previa si aplica.
6. Devolver resumen de cambios.
7. Guardar resultado final.

#### Ejemplo

Petición del usuario:

"Resume este Excel y crea una presentación con 5 diapositivas ejecutivas."

La IA debería:

1. leer hojas relevantes del Excel,
2. detectar KPIs,
3. generar una estructura narrativa,
4. abrir o crear PowerPoint,
5. crear las diapositivas,
6. insertar gráficos o tablas,
7. exportar a PDF o PNG para revisión.

### Seguridad y Fiabilidad

#### Imprescindible desde el inicio

- `save-as` antes de editar,
- carpeta temporal de trabajo,
- validación estricta de rutas,
- allowlist de directorios si el sistema crece,
- logs por operación,
- resumen de cambios realizados,
- manejo de excepciones COM,
- liberación correcta de procesos Office,
- posibilidad de modo simulación en operaciones grandes.

#### Riesgos a controlar

- procesos huérfanos de Office,
- archivos bloqueados por el usuario,
- documentos corruptos o protegidos,
- operaciones destructivas sin confirmación,
- respuestas IA ambiguas.

### Estructura de Proyecto Recomendada

```text
office-ai-mcp/
  src/
    server.py
    config.py
    tools/
      word_tools.py
      excel_tools.py
      powerpoint_tools.py
    services/
      word_service.py
      excel_service.py
      powerpoint_service.py
    models/
      requests.py
      responses.py
    utils/
      paths.py
      logging.py
      backups.py
      com_cleanup.py
  tests/
  docs/
  .env
  requirements.txt
  README.md
```

### Dependencias Recomendadas en Python

```txt
pywin32
pydantic
python-dotenv
tenacity
rich
loguru
```

Opcionales según evolución:

- `fastapi` si además quieres una API HTTP local
- `pytest` para pruebas
- `mypy` si quieres más rigor estático
- `ruff` para lint y formato

### Valoración de Python frente a Node.js

#### Opción 1: Python + pywin32 + MCP

##### Ventajas

- mejor ergonomía para COM en Windows,
- desarrollo rápido,
- ecosistema fuerte para IA,
- buena capacidad de scripting y automatización,
- muy adecuado para prototipo serio que luego pueda crecer.

##### Desventajas

- menos tipado fuerte que C#,
- despliegue algo menos estructurado si el proyecto crece mucho,
- COM sigue siendo COM: hay que tratar bien errores y limpieza.

##### Veredicto

Es la mejor opción para empezar si tu prioridad es construir algo útil pronto y con buena capacidad real sobre Office.

#### Opción 2: Node.js + MCP + COM

##### Variante A: Node.js + `winax`

###### Ventajas

- todo el servidor en JavaScript o TypeScript,
- muy natural para ecosistema MCP si ya trabajas con Node,
- buena experiencia en tooling web y servicios.

###### Desventajas

- la integración COM en Node suele ser menos cómoda,
- más fricción para depurar automatización Office compleja,
- menos maduro para este caso que Python o C#.

###### Veredicto

Viable, pero no es la opción que elegiría como base principal si el corazón del sistema es Office por COM.

##### Variante B: Node.js MCP + PowerShell o C# como capa COM

###### Ventajas

- Node sirve bien como orquestador y capa MCP,
- PowerShell o C# hacen la parte fuerte de Office,
- arquitectura limpia si quieres separar responsabilidades.

###### Desventajas

- introduces dos o tres tecnologías,
- más complejidad operativa,
- más puntos de fallo y más coste de mantenimiento.

###### Veredicto

Buena opción si ya tienes ecosistema Node.js y quieres mantener el servidor MCP en JavaScript, pero no es la ruta más simple para un primer sistema.

#### Opción 3: Node.js con Office Add-ins

##### Ventajas

- gran integración UI dentro de Word, Excel y PowerPoint,
- buena experiencia de usuario,
- muy adecuado para paneles laterales y comandos.

##### Desventajas

- `Office.js` no sustituye a COM para automatización profunda,
- peor para flujos entre aplicaciones,
- más orientado a experiencia in-app que a automatización general.

##### Veredicto

Muy bueno como interfaz de usuario, pero no lo elegiría como único motor si lo que quieres es un editor IA completo entre las tres apps.

### Comparativa resumida

| Opción                     | Velocidad de desarrollo | Edición profunda Office | Complejidad | Recomendación              |
| -------------------------- | ----------------------: | ----------------------: | ----------: | -------------------------- |
| Python + pywin32 + MCP     |                    Alta |                    Alta |       Media | Recomendada                |
| Node + winax               |                   Media |              Media/Alta |  Media/Alta | Válida, pero no preferente |
| Node + PowerShell/C# + MCP |                   Media |                    Alta |        Alta | Buena si ya vives en Node  |
| Office Add-ins con Node    |                   Media |                   Media |       Media | Mejor como capa UI         |

### Recomendación Estratégica

#### Recomendación inmediata

Construir el primer sistema con:

- Python,
- `pywin32`,
- MCP local,
- herramientas separadas por Word, Excel y PowerPoint,
- copias de seguridad automáticas,
- exportación de vista previa en PDF o imágenes.

#### Recomendación a medio plazo

Si el proyecto crece:

- mantener Python como motor de automatización,
- añadir una API local o backend central,
- crear después add-ins si necesitas experiencia dentro de Office,
- valorar C# para partes críticas si aparecen problemas de robustez.

### MVP Recomendado

#### Fase 1

- Word: leer estructura, reemplazar texto, aplicar estilos, exportar PDF.

#### Fase 2

- Excel: leer hojas, escribir rangos, fórmulas, gráficos simples.

#### Fase 3

- PowerPoint: leer slides, reemplazar texto, insertar imágenes, crear slides, exportar PDF.

#### Fase 4

- workflows entre aplicaciones.

Ejemplo:

- leer Excel,
- redactar resumen en Word,
- generar presentación en PowerPoint.

### Qué no recomendaría

- editar binarios Office directamente desde prompts,
- usar solo OCR o control de ratón o teclado,
- diseñar una tool única tipo `edit_office_document(prompt)`,
- empezar con tres frontends distintos sin backend común,
- prescindir de copias de seguridad.

### Conclusión Final

Si eliges Python, estás eligiendo la opción más equilibrada para construir un editor IA serio para Office en Windows.

La mejor decisión técnica inicial es:

- Python como base,
- MCP como interfaz para la IA,
- `pywin32` y COM como motor real de edición,
- herramientas pequeñas, explícitas y seguras.

Node.js sigue siendo una alternativa válida, especialmente si quieres un servidor MCP en JavaScript o una futura capa UI o web, pero para la edición profunda y fiable de Word, Excel y PowerPoint Python parte con ventaja clara.
