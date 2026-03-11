# TODO Word MCP

Cobertura funcional objetivo: inventario exhaustivo, por bloques, de las capacidades que un MCP de Word debería cubrir para aproximarse al 100% práctico de automatización vía Office COM.

## 1. Gestión del documento y archivos

- `word_open_document`
  - Abrir un documento con opciones de solo lectura o reparación.
- `word_close_document`
  - Cerrar un documento abierto.
- `word_save`
  - Guardar cambios.
- `word_save_as`
  - Guardar con otro nombre o formato.
- `word_save_copy`
  - Guardar una copia sin cambiar la ruta activa.
- `word_compare_documents`
  - Comparar dos documentos.
- `word_combine_documents`
  - Combinar revisiones de varios documentos.
- `word_split_document`
  - Dividir un documento por secciones, encabezados o páginas.
- `word_merge_documents`
  - Unir documentos en uno solo.
- `word_get_document_properties`
  - Leer metadatos.
- `word_set_document_properties`
  - Editar propiedades.
- `word_add_custom_property`
  - Propiedades personalizadas.
- `word_remove_custom_property`
  - Eliminar propiedades personalizadas.

## 2. Estructura y navegación del documento

- `word_list_sections`
  - Listar secciones del documento.
- `word_get_section_content`
  - Leer el contenido de una sección concreta.
- `word_insert_section_break`
  - Insertar un salto de sección.
- `word_remove_section_break`
  - Eliminar un salto de sección.
- `word_insert_page_break`
  - Insertar salto de página.
- `word_remove_page_break`
  - Eliminar salto de página.
- `word_get_document_outline`
  - Obtener el esquema jerárquico por títulos.
- `word_go_to_bookmark`
  - Resolver y navegar a marcadores.
- `word_list_bookmarks`
  - Enumerar marcadores del documento.
- `word_create_bookmark`
  - Crear un marcador.
- `word_delete_bookmark`
  - Eliminar un marcador.

## 3. Texto y reemplazo avanzado

- `word_find_text`
  - Buscar texto en todo el documento.
- `word_replace_text_advanced`
  - Reemplazo con opciones de mayúsculas, palabras completas o regex si aplica.
- `word_replace_in_range`
  - Reemplazar texto solo dentro de un rango.
- `word_insert_text_at_bookmark`
  - Insertar texto en un marcador.
- `word_append_paragraph`
  - Añadir un párrafo al final del documento o de una sección.
- `word_prepend_paragraph`
  - Insertar un párrafo al inicio.
- `word_insert_paragraph_after`
  - Insertar párrafo después de otro.
- `word_insert_paragraph_before`
  - Insertar párrafo antes de otro.
- `word_delete_paragraph`
  - Eliminar párrafos por índice.
- `word_merge_paragraphs`
  - Combinar párrafos consecutivos.
- `word_split_paragraph`
  - Dividir un párrafo en una posición dada.

## 4. Formato de texto y párrafo

- `word_set_text_style`
  - Aplicar fuente, tamaño, color y énfasis a un rango.
- `word_set_paragraph_style`
  - Aplicar estilos de párrafo predefinidos.
- `word_set_heading_level`
  - Convertir un párrafo en Título 1, 2, 3, etc.
- `word_set_alignment`
  - Alinear párrafos a izquierda, derecha, centro o justificado.
- `word_set_line_spacing`
  - Configurar interlineado.
- `word_set_paragraph_spacing`
  - Ajustar espacios antes y después.
- `word_set_indentation`
  - Cambiar sangrías izquierda, derecha y primera línea.
- `word_set_tabs`
  - Configurar tabulaciones.
- `word_set_shading`
  - Aplicar sombreado de párrafo o rango.
- `word_set_border`
  - Bordes de párrafo o rango.
- `word_clear_formatting`
  - Limpiar formato directo.
- `word_copy_format`
  - Copiar formato desde un rango a otro.

## 5. Listas y numeración

- `word_create_bulleted_list`
  - Crear listas con viñetas.
- `word_create_numbered_list`
  - Crear listas numeradas.
- `word_set_list_level`
  - Cambiar nivel de una lista.
- `word_restart_numbering`
  - Reiniciar numeración.
- `word_continue_numbering`
  - Continuar numeración previa.
- `word_convert_to_outline_list`
  - Crear listas multinivel.

## 6. Tablas

- `word_insert_table`
  - Crear una tabla nueva.
- `word_delete_table`
  - Eliminar una tabla.
- `word_get_table`
  - Inspeccionar una tabla concreta.
- `word_set_table_cell_text`
  - Editar texto de una celda.
- `word_set_table_cell_style`
  - Formato de celdas.
- `word_add_table_row`
  - Insertar fila.
- `word_add_table_column`
  - Insertar columna.
- `word_delete_table_row`
  - Eliminar fila.
- `word_delete_table_column`
  - Eliminar columna.
- `word_merge_table_cells`
  - Combinar celdas.
- `word_split_table_cells`
  - Dividir celdas.
- `word_sort_table`
  - Ordenar una tabla.
- `word_table_from_csv`
  - Crear tabla desde CSV.
- `word_convert_table_to_text`
  - Convertir tabla a texto.
- `word_repeat_table_header`
  - Repetir fila de encabezado en saltos de página.

## 7. Encabezados, pies, secciones y configuración de página

- `word_get_headers_footers`
  - Leer encabezados y pies.
- `word_set_header_text`
  - Modificar encabezado.
- `word_set_footer_text`
  - Modificar pie de página.
- `word_insert_page_number`
  - Insertar numeración de páginas.
- `word_set_page_number_format`
  - Configurar formato de numeración.
- `word_link_headers_to_previous`
  - Vincular encabezados/pies entre secciones.
- `word_set_page_orientation`
  - Cambiar entre vertical y horizontal.
- `word_set_page_margins`
  - Configurar márgenes.
- `word_set_page_size`
  - Ajustar tamaño de papel.
- `word_set_columns_layout`
  - Definir diseño de columnas.

## 8. Imágenes, formas y objetos

- `word_insert_image`
  - Insertar imagen.
- `word_replace_image`
  - Sustituir imagen existente.
- `word_resize_image`
  - Cambiar tamaño.
- `word_crop_image`
  - Recortar imagen.
- `word_set_image_wrap`
  - Ajustar modo de flujo del texto alrededor.
- `word_insert_shape`
  - Insertar una forma.
- `word_set_shape_text`
  - Editar texto de una forma.
- `word_delete_shape`
  - Eliminar formas.
- `word_align_shapes`
  - Alinear formas.
- `word_group_shapes`
  - Agrupar formas.
- `word_ungroup_shapes`
  - Desagrupar formas.
- `word_embed_object`
  - Incrustar un objeto OLE.

## 9. Comentarios, control de cambios y revisión

- `word_list_comments`
  - Listar comentarios.
- `word_add_comment`
  - Añadir comentario.
- `word_edit_comment`
  - Editar comentario.
- `word_delete_comment`
  - Eliminar comentario.
- `word_reply_comment`
  - Responder comentarios.
- `word_enable_track_changes`
  - Activar control de cambios.
- `word_disable_track_changes`
  - Desactivar control de cambios.
- `word_accept_all_changes`
  - Aceptar todos los cambios.
- `word_reject_all_changes`
  - Rechazar todos los cambios.
- `word_list_revisions`
  - Enumerar revisiones.
- `word_accept_revision`
  - Aceptar una revisión concreta.
- `word_reject_revision`
  - Rechazar una revisión concreta.

## 10. Referencias, TOC, notas y campos

- `word_insert_table_of_contents`
  - Insertar tabla de contenido.
- `word_update_table_of_contents`
  - Actualizar TOC.
- `word_insert_footnote`
  - Insertar nota al pie.
- `word_insert_endnote`
  - Insertar nota final.
- `word_insert_cross_reference`
  - Añadir referencia cruzada.
- `word_insert_caption`
  - Añadir leyendas a tablas o imágenes.
- `word_update_fields`
  - Actualizar todos los campos.
- `word_insert_field`
  - Insertar un campo específico.
- `word_lock_field`
  - Bloquear un campo.
- `word_unlock_field`
  - Desbloquear un campo.
- `word_insert_index`
  - Insertar índice analítico.
- `word_insert_bibliography`
  - Insertar bibliografía si está soportado.

## 11. Estilos, plantillas y temas

- `word_list_styles`
  - Listar estilos disponibles.
- `word_apply_style`
  - Aplicar un estilo existente.
- `word_create_style`
  - Crear estilos personalizados.
- `word_modify_style`
  - Editar un estilo.
- `word_delete_style`
  - Eliminar un estilo personalizado.
- `word_apply_template`
  - Aplicar una plantilla.
- `word_attach_template`
  - Asociar plantilla DOTX/DOTM.
- `word_copy_styles_from_template`
  - Importar estilos de otra plantilla.

## 12. Formularios, controles de contenido y automatización documental

- `word_list_content_controls`
  - Enumerar controles de contenido.
- `word_insert_content_control`
  - Crear un control de contenido.
- `word_set_content_control_value`
  - Rellenar un control.
- `word_remove_content_control`
  - Eliminar un control.
- `word_get_merge_fields`
  - Leer campos de combinación.
- `word_insert_merge_field`
  - Insertar campo.
- `word_fill_merge_fields`
  - Completar plantilla con datos.
- `word_mail_merge_to_documents`
  - Generar documentos desde datos.
- `word_mail_merge_to_pdf`
  - Exportar lotes a PDF.
- `word_run_macro`
  - Ejecutar una macro VBA si se permite.

## 13. Protección, firmas y permisos

- `word_protect_document`
  - Proteger contra edición.
- `word_unprotect_document`
  - Quitar protección.
- `word_restrict_editing`
  - Limitar tipos de edición.
- `word_allow_only_comments`
  - Restringir a comentarios.
- `word_add_digital_signature`
  - Aplicar firma si está soportado.
- `word_remove_digital_signature`
  - Quitar firma.

## 14. Exportación y salidas

- `word_export_range_pdf`
  - Exportar parte del documento a PDF.
- `word_export_to_html`
  - Exportar a HTML.
- `word_export_to_markdown`
  - Exportar a Markdown.
- `word_export_comments`
  - Volcar comentarios a JSON/Markdown.
- `word_export_outline`
  - Exportar estructura por títulos.
- `word_export_plain_text`
  - Exportar a TXT.
- `word_export_images`
  - Extraer imágenes del documento.

## 15. Calidad, accesibilidad y lenguaje

- `word_check_accessibility`
  - Revisar accesibilidad básica.
- `word_spellcheck`
  - Comprobación ortográfica.
- `word_grammar_check`
  - Revisión gramatical si está soportada.
- `word_set_proofing_language`
  - Configurar idioma de corrección.
- `word_find_empty_headings`
  - Detectar títulos vacíos.
- `word_find_broken_references`
  - Detectar referencias rotas.
- `word_add_alt_text`
  - Añadir texto alternativo a imágenes y formas.
- `word_get_readability_stats`
  - Obtener métricas de legibilidad.

## 16. Vistas, ventanas y experiencia de edición

- `word_set_view_mode`
  - Cambiar entre diseño de impresión, lectura, esquema o borrador.
- `word_get_view_mode`
  - Leer la vista activa.
- `word_zoom_view`
  - Ajustar zoom.
- `word_show_hide_markup`
  - Mostrar u ocultar marcas de revisión.
- `word_arrange_windows`
  - Gestionar ventanas abiertas.

## 17. Operaciones masivas y productividad

- `word_batch_update`
  - Varias operaciones en una sola llamada.
- `word_clone_format`
  - Copiar formato entre rangos.
- `word_cleanup_formatting`
  - Normalizar estilos y formato.
- `word_remove_empty_paragraphs`
  - Limpiar párrafos vacíos redundantes.
- `word_generate_change_report`
  - Resumen de cambios hechos por el MCP.
- `word_transaction`
  - Ejecutar varias operaciones con rollback si falla una.

## Priorización recomendada

1. `word_find_text`
2. `word_set_text_style`
3. `word_insert_table`
4. `word_list_comments`
5. `word_add_comment`
6. `word_enable_track_changes`
7. `word_accept_all_changes`
8. `word_insert_image`
9. `word_insert_table_of_contents`
10. `word_save_as`
