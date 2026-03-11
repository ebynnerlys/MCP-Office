# TODO Excel MCP

Cobertura funcional objetivo: inventario exhaustivo, por bloques, de las capacidades que un MCP de Excel debería cubrir para aproximarse al 100% práctico de automatización vía Office COM.

## 1. Gestión del libro y archivos

- `excel_open_workbook`
  - Abrir un libro con opciones de solo lectura o reparación.
- `excel_close_workbook`
  - Cerrar un libro.
- `excel_save`
  - Guardar cambios.
- `excel_save_as`
  - Guardar con otro nombre o formato.
- `excel_save_copy`
  - Guardar una copia sin cambiar la ruta activa.
- `excel_get_workbook_metadata`
  - Leer propiedades del libro, autor, hojas y ruta.
- `excel_set_workbook_properties`
  - Editar metadatos del libro.
- `excel_add_custom_property`
  - Añadir propiedad personalizada.
- `excel_remove_custom_property`
  - Eliminar propiedad personalizada.
- `excel_compare_workbooks`
  - Comparar dos libros si el flujo lo permite.

## 2. Hojas, estructura y organización

- `excel_add_sheet`
  - Crear una hoja nueva.
- `excel_delete_sheet`
  - Eliminar una hoja.
- `excel_rename_sheet`
  - Renombrar una hoja.
- `excel_move_sheet`
  - Cambiar el orden de una hoja.
- `excel_copy_sheet`
  - Duplicar una hoja dentro del libro o hacia otro libro.
- `excel_hide_sheet`
  - Ocultar una hoja.
- `excel_unhide_sheet`
  - Mostrar una hoja oculta.
- `excel_set_active_sheet`
  - Cambiar la hoja activa.
- `excel_set_sheet_tab_color`
  - Cambiar color de la pestaña.
- `excel_list_sheet_visibility`
  - Ver estados visible/oculta/muy oculta.

## 3. Rangos, celdas y manipulación base

- `excel_get_used_range`
  - Obtener el rango usado real de una hoja.
- `excel_clear_range`
  - Limpiar contenido o formato de un rango.
- `excel_clear_contents`
  - Limpiar solo valores y fórmulas.
- `excel_clear_comments`
  - Limpiar comentarios de un rango.
- `excel_copy_range`
  - Copiar un rango.
- `excel_paste_range`
  - Pegar un rango con opciones.
- `excel_insert_rows`
  - Insertar filas.
- `excel_insert_columns`
  - Insertar columnas.
- `excel_delete_rows`
  - Eliminar filas.
- `excel_delete_columns`
  - Eliminar columnas.
- `excel_merge_cells`
  - Combinar celdas.
- `excel_unmerge_cells`
  - Separar celdas combinadas.
- `excel_autofit_rows`
  - Ajustar alto automáticamente.
- `excel_autofit_columns`
  - Ajustar ancho automáticamente.
- `excel_fill_down`
  - Rellenar hacia abajo.
- `excel_fill_right`
  - Rellenar a la derecha.
- `excel_transpose_range`
  - Transponer datos.
- `excel_remove_duplicates`
  - Eliminar filas duplicadas.

## 4. Fórmulas, nombres y cálculo

- `excel_set_formula`
  - Escribir una fórmula en una o varias celdas.
- `excel_get_formula`
  - Leer fórmulas.
- `excel_set_array_formula`
  - Escribir fórmulas matriciales.
- `excel_convert_formulas_to_values`
  - Sustituir fórmulas por valores.
- `excel_recalculate_workbook`
  - Recalcular todo el libro.
- `excel_recalculate_sheet`
  - Recalcular una hoja.
- `excel_trace_precedents`
  - Inspeccionar dependencias previas.
- `excel_trace_dependents`
  - Inspeccionar dependientes.
- `excel_list_named_ranges`
  - Listar nombres definidos.
- `excel_create_named_range`
  - Crear un rango con nombre.
- `excel_delete_named_range`
  - Eliminar un nombre definido.
- `excel_update_named_range`
  - Cambiar la referencia de un nombre.

## 5. Formato, estilo y apariencia de hoja

- `excel_set_number_format`
  - Formato numérico.
- `excel_set_font_style`
  - Fuente, tamaño, color, negrita, etc.
- `excel_set_fill_color`
  - Color de relleno.
- `excel_set_border_style`
  - Bordes de celdas.
- `excel_set_alignment`
  - Alineación horizontal y vertical.
- `excel_set_wrap_text`
  - Ajuste de texto.
- `excel_set_text_rotation`
  - Rotación del texto.
- `excel_apply_style`
  - Aplicar estilo o formato predefinido.
- `excel_clear_formatting`
  - Quitar formatos.
- `excel_copy_format`
  - Copiar formato entre rangos.
- `excel_format_as_table`
  - Convertir un rango en tabla con estilo.
- `excel_set_row_height`
  - Ajustar altura fija.
- `excel_set_column_width`
  - Ajustar ancho fijo.

## 6. Tablas estructuradas de Excel

- `excel_create_table`
  - Crear una tabla estructurada.
- `excel_resize_table`
  - Cambiar tamaño de la tabla.
- `excel_rename_table`
  - Renombrar una tabla.
- `excel_list_tables`
  - Listar tablas del libro.
- `excel_get_table_data`
  - Leer datos de una tabla.
- `excel_append_table_rows`
  - Añadir filas a una tabla.
- `excel_delete_table_rows`
  - Eliminar filas concretas.
- `excel_set_table_style`
  - Cambiar estilo visual de la tabla.
- `excel_toggle_total_row`
  - Activar/desactivar fila de totales.
- `excel_convert_table_to_range`
  - Convertir tabla a rango normal.

## 7. Filtros, ordenación, agrupación y vistas

- `excel_apply_filter`
  - Filtrar un rango o tabla.
- `excel_clear_filter`
  - Limpiar filtros.
- `excel_sort_range`
  - Ordenar un rango.
- `excel_sort_table`
  - Ordenar una tabla.
- `excel_group_rows`
  - Agrupar filas.
- `excel_group_columns`
  - Agrupar columnas.
- `excel_ungroup_rows`
  - Desagrupar filas.
- `excel_ungroup_columns`
  - Desagrupar columnas.
- `excel_create_custom_view`
  - Crear vista personalizada si se soporta.
- `excel_freeze_panes`
  - Inmovilizar paneles.
- `excel_unfreeze_panes`
  - Liberar paneles.
- `excel_split_window`
  - Dividir ventana.

## 8. Charts, sparklines y visualización

- `excel_add_chart`
  - Crear gráficos.
- `excel_list_charts`
  - Listar gráficos existentes.
- `excel_set_chart_title`
  - Cambiar el título.
- `excel_set_chart_data`
  - Actualizar series y categorías.
- `excel_set_chart_style`
  - Aplicar estilos visuales.
- `excel_set_chart_axis`
  - Configurar ejes.
- `excel_delete_chart`
  - Eliminar gráfico.
- `excel_export_chart_image`
  - Exportar gráfico como imagen.
- `excel_move_chart`
  - Mover el gráfico a otra hoja o posición.
- `excel_add_sparkline`
  - Insertar minigráficos.
- `excel_clear_sparkline`
  - Eliminar minigráficos.

## 9. Pivot tables, análisis y escenarios

- `excel_create_pivot_table`
  - Crear una tabla dinámica.
- `excel_list_pivot_tables`
  - Listar tablas dinámicas.
- `excel_refresh_pivot_table`
  - Refrescar una tabla dinámica.
- `excel_refresh_all_pivots`
  - Refrescar todas.
- `excel_configure_pivot_fields`
  - Configurar filas, columnas, valores y filtros.
- `excel_set_pivot_layout`
  - Cambiar diseño de presentación.
- `excel_create_pivot_chart`
  - Crear gráfico dinámico.
- `excel_create_slicer`
  - Añadir segmentación de datos.
- `excel_remove_slicer`
  - Eliminar segmentación.
- `excel_run_goal_seek`
  - Ejecutar buscar objetivo.
- `excel_run_solver`
  - Ejecutar Solver si está disponible.
- `excel_manage_scenarios`
  - Crear o actualizar escenarios.

## 10. Validación, reglas y calidad de datos

- `excel_add_data_validation`
  - Añadir validación de datos.
- `excel_remove_data_validation`
  - Eliminar validación.
- `excel_add_conditional_format`
  - Añadir formato condicional.
- `excel_clear_conditional_format`
  - Limpiar reglas.
- `excel_list_conditional_formats`
  - Listar reglas activas.
- `excel_circle_invalid_data`
  - Marcar celdas inválidas.
- `excel_find_data_quality_issues`
  - Detectar nulos, fórmulas rotas o tipos inconsistentes.

## 11. Comentarios, notas y revisión

- `excel_list_comments`
  - Listar comentarios por celda.
- `excel_add_comment`
  - Añadir comentario.
- `excel_edit_comment`
  - Editar comentario.
- `excel_delete_comment`
  - Eliminar comentario.
- `excel_reply_comment`
  - Responder comentarios si aplica.
- `excel_list_notes`
  - Leer notas heredadas.
- `excel_set_note`
  - Crear o modificar nota.
- `excel_delete_note`
  - Eliminar nota.
- `excel_show_hide_comments`
  - Mostrar u ocultar comentarios.

## 12. Imágenes, formas, objetos y diseño visual

- `excel_insert_image`
  - Insertar imagen en una hoja.
- `excel_replace_image`
  - Sustituir imagen existente.
- `excel_resize_shape`
  - Redimensionar shapes.
- `excel_delete_shape`
  - Eliminar shapes.
- `excel_list_shapes`
  - Listar imágenes y formas.
- `excel_insert_textbox`
  - Añadir cuadro de texto.
- `excel_set_shape_text`
  - Cambiar texto de shapes.
- `excel_align_shapes`
  - Alinear shapes.
- `excel_group_shapes`
  - Agrupar shapes.
- `excel_ungroup_shapes`
  - Desagrupar shapes.
- `excel_embed_object`
  - Insertar objeto OLE.

## 13. Protección, permisos y compartición

- `excel_protect_sheet`
  - Proteger una hoja.
- `excel_unprotect_sheet`
  - Quitar protección.
- `excel_protect_workbook`
  - Proteger estructura del libro.
- `excel_unprotect_workbook`
  - Quitar protección del libro.
- `excel_lock_range`
  - Bloquear un rango.
- `excel_unlock_range`
  - Desbloquear un rango.
- `excel_allow_edit_range`
  - Definir rangos editables dentro de hoja protegida.
- `excel_share_workbook`
  - Configurar compartición si el modelo lo permite.

## 14. Impresión, diseño de página y salida

- `excel_set_print_area`
  - Definir área de impresión.
- `excel_clear_print_area`
  - Limpiar área de impresión.
- `excel_set_page_orientation`
  - Horizontal o vertical.
- `excel_set_page_margins`
  - Márgenes de página.
- `excel_set_repeat_rows_columns`
  - Repetir filas/columnas al imprimir.
- `excel_insert_page_break`
  - Insertar salto de página.
- `excel_remove_page_break`
  - Eliminar salto.
- `excel_set_header_footer`
  - Encabezados y pies de impresión.
- `excel_export_sheet_pdf`
  - Exportar una sola hoja a PDF.
- `excel_export_range_pdf`
  - Exportar un rango a PDF.

## 15. Importación, conexiones e integración externa

- `excel_import_csv`
  - Importar CSV a una hoja.
- `excel_export_csv`
  - Exportar una hoja o tabla a CSV.
- `excel_import_text_file`
  - Importar TXT delimitado.
- `excel_list_external_links`
  - Listar enlaces externos.
- `excel_refresh_external_links`
  - Actualizar vínculos externos.
- `excel_break_external_links`
  - Romper vínculos.
- `excel_list_queries`
  - Listar consultas y conexiones.
- `excel_refresh_query`
  - Refrescar consulta concreta.
- `excel_refresh_all_queries`
  - Refrescar todo.
- `excel_link_to_powerpoint_or_word`
  - Preparar salidas para integración cruzada.
- `excel_run_macro`
  - Ejecutar una macro VBA si se permite.

## 16. Calidad, perfilado y productividad

- `excel_find_empty_rows`
  - Detectar filas vacías.
- `excel_find_duplicates`
  - Detectar duplicados.
- `excel_profile_range`
  - Resumen de tipos, nulos, mínimos y máximos.
- `excel_detect_formula_errors`
  - Detectar `#VALUE!`, `#REF!`, etc.
- `excel_check_accessibility`
  - Revisar accesibilidad básica del libro.
- `excel_add_alt_text`
  - Añadir texto alternativo a gráficos y shapes.
- `excel_batch_update`
  - Varias operaciones en una sola llamada.
- `excel_cleanup_workbook`
  - Limpiar formatos y hojas innecesarias.
- `excel_generate_summary_sheet`
  - Crear una hoja resumen.
- `excel_transaction`
  - Ejecutar varias operaciones con rollback si falla una.

## Priorización recomendada

1. `excel_set_formula`
2. `excel_clear_range`
3. `excel_add_sheet`
4. `excel_set_number_format`
5. `excel_create_table`
6. `excel_apply_filter`
7. `excel_sort_range`
8. `excel_add_chart`
9. `excel_add_comment`
10. `excel_export_sheet_pdf`
