# TODO PowerPoint MCP

Cobertura funcional objetivo: inventario exhaustivo, por bloques, de las capacidades que un MCP de PowerPoint debería cubrir para aproximarse al 100% práctico de automatización vía Office COM.

## 1. Gestión de presentaciones y archivos

- `ppt_open_presentation`
  - Abrir una presentación explícitamente con opciones de solo lectura o reparación.
- `ppt_close_presentation`
  - Cerrar una presentación abierta.
- [x] `ppt_save`
  - Guardar cambios en el archivo actual.
- `ppt_save_as`
  - Guardar con otro nombre, ruta o formato.
- [x] `ppt_save_copy`
  - Guardar una copia sin cambiar la ruta activa.
- `ppt_merge_presentations`
  - Unir varias presentaciones en una sola.
- `ppt_split_presentation`
  - Exportar grupos de diapositivas a archivos separados.
- `ppt_compare_presentations`
  - Comparar dos `.pptx`.
- `ppt_repair_presentation`
  - Intentar abrir y reparar archivos dañados si COM lo soporta.
- [x] `ppt_get_document_properties`
  - Leer propiedades del archivo.
- [x] `ppt_set_document_properties`
  - Modificar autor, título, asunto, etiquetas.
- `ppt_add_custom_property`
  - Propiedades personalizadas.
- `ppt_remove_custom_property`
  - Eliminar propiedades personalizadas.
- [x] `ppt_get_file_links`
  - Inspeccionar vínculos externos presentes en el archivo.

## 2. Gestión de diapositivas

- [x] `ppt_duplicate_slide`
  - Duplicar una diapositiva existente.
- [x] `ppt_delete_slide`
  - Eliminar una o varias diapositivas.
- [x] `ppt_move_slide`
  - Cambiar el orden de una diapositiva.
- `ppt_copy_slide_to_presentation`
  - Copiar una diapositiva a otro archivo `.pptx`.
- [x] `ppt_hide_slide`
  - Ocultar una diapositiva en modo presentación.
- [x] `ppt_unhide_slide`
  - Volver a mostrar una diapositiva.
- [x] `ppt_set_slide_name`
  - Asignar o cambiar un nombre interno a la diapositiva.
- [x] `ppt_get_slide_metadata`
  - Leer información ampliada: id, nombre, layout, sección, oculta/no oculta.
- `ppt_add_section`
  - Crear secciones de diapositivas.
- `ppt_rename_section`
  - Renombrar secciones.
- `ppt_move_slide_to_section`
  - Reubicar una diapositiva en una sección.
- `ppt_delete_section`
  - Eliminar secciones.
- [x] `ppt_get_slide_summary_extended`
  - Resumen ampliado con texto, tablas, charts, notas y animaciones.

## 3. Layouts, masters, placeholders y temas

- [x] `ppt_list_layouts`
  - Listar layouts disponibles en la presentación.
- [x] `ppt_apply_layout`
  - Cambiar el layout de una diapositiva existente.
- [x] `ppt_get_slide_layout`
  - Consultar el layout actual.
- `ppt_list_masters`
  - Listar patrones de diapositivas.
- `ppt_get_master_details`
  - Leer colores, fuentes, placeholders y layout del patrón.
- `ppt_set_master_background`
  - Cambiar fondo del patrón.
- `ppt_set_master_fonts`
  - Configurar tipografías del tema.
- `ppt_set_master_colors`
  - Cambiar paleta del tema.
- `ppt_apply_theme_variant`
  - Aplicar variantes del tema activo.
- `ppt_extract_theme`
  - Exportar o resumir el tema activo.
- [x] `ppt_reset_slide_to_layout`
  - Restablecer una slide a su layout original.
- [x] `ppt_list_placeholders`
  - Enumerar placeholders de una slide o layout.
- `ppt_fill_placeholder`
  - Escribir contenido en un placeholder por tipo.
- `ppt_replace_placeholder_with_shape`
  - Sustituir placeholder por contenido final.
- `ppt_restore_placeholder`
  - Restaurar placeholders eliminados.

## 4. Texto, tipografía y contenido narrativo

- [x] `ppt_find_text`
  - Buscar texto en toda la presentación.
- [x] `ppt_replace_text_all`
  - Reemplazar texto en toda la presentación, no solo en una slide.
- [x] `ppt_set_slide_title`
  - Cambiar solo el título de una diapositiva.
- [x] `ppt_get_shape_text_runs`
  - Obtener fragmentos de texto con formato por rango.
- [x] `ppt_set_text_range_style`
  - Dar formato a una parte concreta del texto.
- [x] `ppt_insert_bullets`
  - Crear listas con viñetas en un shape.
- [x] `ppt_set_bullet_style`
  - Modificar sangría, nivel, símbolo y espaciado de viñetas.
- [x] `ppt_set_paragraph_spacing`
  - Ajustar espaciado antes/después e interlineado.
- [x] `ppt_set_textbox_margins`
  - Cambiar márgenes internos del cuadro de texto.
- [x] `ppt_set_text_direction`
  - Cambiar orientación del texto.
- [x] `ppt_set_autofit`
  - Ajustar texto automáticamente al shape.
- [x] `ppt_set_proofing_language`
  - Configurar idioma de corrección.
- [x] `ppt_spellcheck_slide`
  - Revisión ortográfica de una slide.
- [x] `ppt_spellcheck_presentation`
  - Revisión completa.
- [x] `ppt_translate_text`
  - Traducir texto si Office lo soporta.

## 5. Notas, comentarios y revisión

- [x] `ppt_get_presenter_notes_all`
  - Extraer todas las notas del presentador.
- [x] `ppt_find_in_notes`
  - Buscar texto dentro de las notas.
- `ppt_set_notes_style`
  - Dar formato al texto de notas.
- [x] `ppt_replace_notes_text`
  - Reemplazar texto dentro de notas.
- `ppt_list_comments`
  - Listar comentarios por diapositiva.
- `ppt_add_comment`
  - Añadir un comentario si la API COM lo permite.
- `ppt_edit_comment`
  - Editar un comentario existente.
- `ppt_delete_comment`
  - Eliminar comentarios.
- `ppt_resolve_comment`
  - Marcar comentarios como resueltos, si está soportado.
- `ppt_list_ink_annotations`
  - Inspeccionar anotaciones manuscritas o trazos.
- `ppt_delete_ink_annotations`
  - Limpiar anotaciones de tinta.
- `ppt_export_review_data`
  - Exportar comentarios y anotaciones a JSON/Markdown.

## 6. Shapes, geometría y composición visual

- [x] `ppt_duplicate_shape`
  - Duplicar un elemento individual.
- [x] `ppt_delete_shape`
  - Eliminar shapes por índice o nombre.
- [x] `ppt_rename_shape`
  - Renombrar un shape para flujos más robustos.
- [x] `ppt_find_shapes`
  - Buscar shapes por nombre, tipo o texto.
- [x] `ppt_group_shapes`
  - Agrupar varios elementos.
- [x] `ppt_ungroup_shapes`
  - Desagrupar un grupo.
- [x] `ppt_align_shapes`
  - Alinear a izquierda, derecha, centro, arriba, abajo, medio.
- [x] `ppt_distribute_shapes`
  - Distribuir horizontal o verticalmente.
- [x] `ppt_resize_shapes`
  - Igualar ancho, alto o ambos.
- [x] `ppt_rotate_shape`
  - Rotar un shape.
- [x] `ppt_flip_shape`
  - Voltear horizontal o verticalmente.
- [x] `ppt_set_shape_position`
  - Mover un shape a coordenadas exactas.
- [x] `ppt_set_shape_size`
  - Redimensionar un shape.
- [x] `ppt_lock_aspect_ratio`
  - Bloquear proporción de un elemento.
- [x] `ppt_set_shape_visibility`
  - Mostrar u ocultar un shape.
- [x] `ppt_bring_to_front`
  - Traer al frente.
- [x] `ppt_send_to_back`
  - Enviar al fondo.
- [x] `ppt_bring_forward`
  - Avanzar una capa.
- [x] `ppt_send_backward`
  - Retroceder una capa.
- [x] `ppt_set_shape_shadow`
  - Aplicar sombra.
- [x] `ppt_set_shape_glow`
  - Aplicar resplandor.
- [x] `ppt_set_shape_reflection`
  - Aplicar reflejo.
- [x] `ppt_set_shape_soft_edges`
  - Ajustar bordes suaves.
- [x] `ppt_set_shape_3d`
  - Aplicar efectos 3D.
- [x] `ppt_merge_shapes`
  - Unir, fragmentar, combinar o restar formas.
- [x] `ppt_crop_shape_to_content`
  - Recortar al contenido visible.

## 7. Imágenes, iconos y multimedia

- [x] `ppt_replace_image`
  - Sustituir una imagen conservando posición y tamaño.
- [x] `ppt_crop_image`
  - Recortar imagen.
- [x] `ppt_reset_image`
  - Restaurar recorte o ajustes.
- [x] `ppt_set_image_transparency`
  - Ajustar transparencia.
- [x] `ppt_set_image_brightness`
  - Brillo de imagen.
- [x] `ppt_set_image_contrast`
  - Contraste de imagen.
- `ppt_compress_images`
  - Reducir peso del archivo.
- `ppt_insert_icon`
  - Insertar iconos del catálogo Office.
- [x] `ppt_insert_svg`
  - Insertar SVG.
- [x] `ppt_add_video`
  - Insertar vídeo.
- [x] `ppt_set_video_playback`
  - Configurar reproducción automática, loop, volumen.
- [x] `ppt_trim_video`
  - Ajustar recorte temporal del vídeo.
- [x] `ppt_add_audio`
  - Insertar audio.
- [x] `ppt_set_audio_playback`
  - Configurar audio entre diapositivas, autoplay, loop.
- `ppt_add_screen_recording`
  - Insertar grabación de pantalla si la API lo soporta.
- [x] `ppt_extract_media_inventory`
  - Inventariar archivos multimedia usados.

## 8. Tablas

- [x] `ppt_add_row_to_table`
  - Insertar filas.
- [x] `ppt_add_column_to_table`
  - Insertar columnas.
- [x] `ppt_delete_row_from_table`
  - Eliminar filas.
- [x] `ppt_delete_column_from_table`
  - Eliminar columnas.
- [x] `ppt_merge_table_cells`
  - Combinar celdas.
- [x] `ppt_split_table_cells`
  - Dividir celdas.
- [x] `ppt_set_table_style`
  - Aplicar estilo de tabla.
- [x] `ppt_set_table_cell_style`
  - Formato de relleno, borde y texto por celda.
- [x] `ppt_set_table_row_style`
  - Formato por fila.
- [x] `ppt_set_table_column_style`
  - Formato por columna.
- [x] `ppt_autofit_table`
  - Ajustar tamaño al contenido.
- [x] `ppt_distribute_table_rows`
  - Distribuir alturas.
- [x] `ppt_distribute_table_columns`
  - Distribuir anchos.
- [x] `ppt_sort_table`
  - Ordenar filas por columna.
- [x] `ppt_table_from_csv`
  - Crear tabla desde CSV.
- [x] `ppt_table_from_excel_range`
  - Crear o refrescar una tabla desde Excel.

## 9. Charts

- [x] `ppt_refresh_chart`
  - Refrescar chart embebido.
- [x] `ppt_set_chart_axis_scale`
  - Definir mínimo, máximo y unidades.
- [x] `ppt_set_chart_series_order`
  - Reordenar series.
- [x] `ppt_add_chart_series`
  - Añadir una serie nueva.
- [x] `ppt_delete_chart_series`
  - Eliminar series.
- [x] `ppt_set_chart_data_labels`
  - Configurar etiquetas de datos.
- [x] `ppt_set_chart_gridlines`
  - Activar o desactivar rejillas.
- [x] `ppt_set_chart_colors`
  - Paleta por serie o chart completo.
- [x] `ppt_change_chart_type`
  - Convertir a otro tipo de chart.
- [x] `ppt_link_chart_to_excel`
  - Vincular chart a un archivo Excel.
- [x] `ppt_break_chart_link`
  - Romper vínculo externo.
- [x] `ppt_export_chart_data`
  - Extraer los datos del gráfico.

## 10. SmartArt, diagramas y conectores

- `ppt_add_smartart_node`
  - Añadir nodos.
- `ppt_delete_smartart_node`
  - Eliminar nodos.
- `ppt_promote_smartart_node`
  - Subir nivel jerárquico.
- `ppt_demote_smartart_node`
  - Bajar nivel jerárquico.
- `ppt_reorder_smartart_node`
  - Reordenar nodos.
- `ppt_set_smartart_style`
  - Cambiar estilo visual.
- `ppt_set_smartart_color_theme`
  - Cambiar combinación de colores.
- `ppt_convert_smartart_to_shapes`
  - Convertir a shapes editables.
- `ppt_reroute_connectors`
  - Recalcular conectores entre shapes.

## 11. Animaciones, transiciones y temporización

- `ppt_reorder_animation`
  - Cambiar el orden en la secuencia.
- `ppt_delete_animation`
  - Eliminar una animación específica.
- `ppt_update_animation`
  - Editar efecto, trigger o tiempos sin recrear.
- `ppt_add_motion_path`
  - Trayectorias de movimiento.
- `ppt_set_animation_repeat`
  - Repetición y rebobinado.
- `ppt_set_animation_sound`
  - Sonido asociado a animaciones.
- `ppt_set_animation_after_effect`
  - Efecto posterior a la animación.
- `ppt_set_transition_sound`
  - Sonido en transición.
- `ppt_set_transition_duration`
  - Duración precisa de transición.
- `ppt_apply_transition_to_all`
  - Replicar transición a toda la presentación.
- `ppt_copy_animations`
  - Copiar animaciones de un shape o slide a otro.
- `ppt_set_rehearsed_timings`
  - Ajustar tiempos ensayados por diapositiva.

## 12. Enlaces, acciones e interactividad

- `ppt_add_hyperlink`
  - Añadir hipervínculos a texto o shapes.
- `ppt_remove_hyperlink`
  - Eliminar hipervínculos.
- `ppt_set_action_on_click`
  - Definir navegación o acción al hacer clic.
- `ppt_set_action_on_hover`
  - Definir acción al pasar el cursor.
- `ppt_link_to_slide`
  - Crear navegación interna entre slides.
- `ppt_link_to_file`
  - Abrir archivo externo.
- `ppt_link_to_url`
  - Abrir una URL.
- `ppt_create_zoom_link`
  - Crear resumen, section zoom o slide zoom si COM lo soporta.

## 13. Slideshow, grabación y experiencia de presentación

- `ppt_start_slideshow`
  - Iniciar presentación.
- `ppt_stop_slideshow`
  - Detener presentación.
- `ppt_go_to_slide`
  - Navegar a una slide en modo presentación.
- `ppt_get_slideshow_state`
  - Estado actual de la presentación en ejecución.
- `ppt_set_kiosk_mode`
  - Configurar modo kiosco.
- `ppt_set_loop_until_esc`
  - Reproducción en bucle.
- `ppt_set_presenter_view`
  - Configurar vista del presentador.
- `ppt_record_narration`
  - Gestionar narración grabada si está soportado.
- `ppt_clear_recorded_timings`
  - Eliminar narraciones o tiempos grabados.
- `ppt_manage_custom_show`
  - Crear y editar presentaciones personalizadas.

## 14. Datos externos, OLE y automatización cruzada

- `ppt_embed_excel`
  - Insertar hoja/rango de Excel embebido.
- `ppt_link_excel_range`
  - Vincular rango externo.
- `ppt_refresh_links`
  - Actualizar todos los vínculos.
- `ppt_break_links`
  - Romper vínculos externos.
- `ppt_embed_object`
  - Insertar objetos OLE.
- `ppt_update_ole_object`
  - Actualizar un objeto incrustado.
- `ppt_embed_word_document`
  - Insertar contenido Word incrustado.
- `ppt_run_macro`
  - Ejecutar una macro VBA de PowerPoint si se permite.

## 15. Exportación y distribución

- `ppt_export_slide_as_pdf`
  - Exportar una sola diapositiva a PDF.
- `ppt_export_range_as_pdf`
  - Exportar un rango de slides.
- `ppt_export_shape_as_image`
  - Exportar un shape como imagen.
- `ppt_export_notes_to_markdown`
  - Volcar notas a Markdown.
- `ppt_export_outline`
  - Exportar esquema de títulos y bullets.
- `ppt_export_handout`
  - Exportar versión tipo folleto.
- `ppt_export_video`
  - Exportar presentación a vídeo.
- `ppt_package_for_cd`
  - Empaquetado portable, si COM lo permite.
- `ppt_publish_online_package`
  - Preparar paquete web o compartible si está disponible.

## 16. Calidad, accesibilidad y gobierno visual

- `ppt_check_accessibility`
  - Revisión de accesibilidad básica.
- `ppt_add_alt_text`
  - Añadir texto alternativo a imágenes y shapes.
- `ppt_get_alt_text`
  - Leer texto alternativo.
- `ppt_detect_overflow_text`
  - Detectar cuadros con texto recortado.
- `ppt_detect_low_contrast`
  - Detectar bajo contraste visual.
- `ppt_find_missing_titles`
  - Detectar slides sin título.
- `ppt_find_empty_tables`
  - Detectar tablas vacías.
- `ppt_find_unused_media`
  - Detectar multimedia no utilizada.
- `ppt_scan_branding_issues`
  - Detectar colores/fuentes fuera de la guía visual.
- `ppt_validate_presentation`
  - Detectar problemas comunes: slides vacías, textos desbordados, links rotos.
- `ppt_get_presentation_stats`
  - Número de slides, shapes, notas, imágenes, charts y multimedia.

## 17. Vistas, ventanas y experiencia de edición

- `ppt_set_view_mode`
  - Cambiar a vista normal, clasificador, patrón o notas.
- `ppt_get_view_mode`
  - Consultar la vista activa.
- `ppt_zoom_view`
  - Ajustar zoom de la ventana.
- `ppt_focus_slide`
  - Llevar una diapositiva al foco de edición.
- `ppt_arrange_windows`
  - Gestionar disposición de ventanas abiertas.

## 18. Operaciones masivas, transacciones y productividad

- `ppt_batch_update`
  - Ejecutar varias operaciones en una sola llamada.
- `ppt_clone_style_between_shapes`
  - Copiar formato entre shapes.
- `ppt_clone_slide_style`
  - Copiar estilo de una slide a otra.
- `ppt_apply_template_to_presentation`
  - Aplicar una presentación modelo.
- `ppt_generate_agenda_slide`
  - Crear slide agenda desde títulos existentes.
- `ppt_generate_section_dividers`
  - Crear diapositivas separadoras.
- `ppt_normalize_deck`
  - Normalizar fuentes, colores, posiciones y tamaños.
- `ppt_cleanup_unused_masters`
  - Limpiar patrones/layouts sin uso.
- `ppt_remove_empty_placeholders`
  - Limpiar placeholders vacíos.
- `ppt_auto_layout_shapes`
  - Reordenar elementos automáticamente.
- `ppt_generate_change_report`
  - Resumen de cambios realizados por el MCP.
- `ppt_rollback_last_operation`
  - Revertir última operación, apoyado en backups.
- `ppt_transaction`
  - Ejecutar varias operaciones con rollback si falla una.

## Priorización recomendada

1. `ppt_duplicate_slide`
2. `ppt_delete_slide`
3. `ppt_move_slide`
4. `ppt_align_shapes`
5. `ppt_group_shapes`
6. `ppt_find_text`
7. `ppt_replace_text_all`
8. `ppt_list_comments`
9. `ppt_add_comment`
10. `ppt_export_notes_to_markdown`
