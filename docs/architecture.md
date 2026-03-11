# Arquitectura Inicial

El proyecto sigue la separacion propuesta en el informe:

- `server.py` crea el servidor MCP y registra las tools.
- `tools/` expone funciones MCP pequenas y estables.
- `services/` encapsula la automatizacion Office por COM.
- `models/` contiene modelos de entrada y salida.
- `utils/` concentra validacion de rutas, logging, backups y limpieza COM.

## Principios aplicados

- La IA no modifica binarios Office directamente.
- Toda ruta de entrada se valida antes de tocar disco.
- Las operaciones de escritura pueden crear backup previo.
- Las respuestas devuelven estructura util, no solo mensajes planos.
- Las dependencias COM se cargan de forma perezosa para no romper pruebas en entornos sin Office.

## Alcance de este MVP tecnico

- Word: inspeccion basica, reemplazo global y exportacion PDF.
- Excel: listado de hojas, lectura o escritura de rangos y exportacion PDF.
- PowerPoint: listado de slides, shapes y texto, presets visuales, lectura y escritura de transiciones, inspeccion y limpieza de animaciones, formato tipografico, colores de relleno y borde, fondo de diapositiva, soporte estructurado y creacion de tablas, charts y SmartArt, actualizacion de datos, leyenda y ejes en charts, estilo de series, creacion de AutoShapes y conectores libres o anclados a shapes, reemplazo en diapositivas, alta de slides, insercion de imagenes, aplicacion de temas, guardado y exportaciones.

## Layouts de slide soportados

- `title`
- `title_and_text`
- `two_column_text`
- `table`
- `chart`
- `title_only`
- `blank`
- `section_header`
- `two_content`
- `content_with_caption`
- `picture_with_caption`

Tambien se admite un entero como `layout` si necesitas pasar directamente un identificador COM valido.

## Riesgos que ya quedan encapsulados

- errores transitorios de COM con reintentos simples
- rutas fuera de directorios permitidos
- ausencia de carpetas de trabajo y backup
- procesos Office abiertos durante la operacion
