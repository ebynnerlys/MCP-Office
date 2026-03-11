# PowerPoint MCP

Guía específica del servidor MCP de PowerPoint.

## Servidor

- Comando CLI: `office-ai-mcp-powerpoint`
- Módulo Python: `python -m office_ai_mcp.powerpoint_server`
- Servidor en VS Code: `officeAiPowerPoint`

La configuración de VS Code está en [../.vscode/mcp.json](../.vscode/mcp.json).

## Alcance

Este servidor expone únicamente herramientas de PowerPoint, además de las herramientas de sistema compartidas:

- `server_status`
- `create_working_backup`

## Herramientas disponibles

- `ppt_list_slides`
- `ppt_save`
- `ppt_save_copy`
- `ppt_get_document_properties`
- `ppt_set_document_properties`
- `ppt_get_file_links`
- `ppt_duplicate_slide`
- `ppt_delete_slide`
- `ppt_move_slide`
- `ppt_hide_slide`
- `ppt_unhide_slide`
- `ppt_set_slide_name`
- `ppt_get_slide_metadata`
- `ppt_get_slide_summary_extended`
- `ppt_list_layouts`
- `ppt_get_slide_layout`
- `ppt_apply_layout`
- `ppt_reset_slide_to_layout`
- `ppt_list_placeholders`
- `ppt_get_slide_shapes`
- `ppt_duplicate_shape`
- `ppt_delete_shape`
- `ppt_rename_shape`
- `ppt_find_shapes`
- `ppt_group_shapes`
- `ppt_ungroup_shapes`
- `ppt_align_shapes`
- `ppt_distribute_shapes`
- `ppt_resize_shapes`
- `ppt_rotate_shape`
- `ppt_flip_shape`
- `ppt_set_shape_position`
- `ppt_set_shape_size`
- `ppt_lock_aspect_ratio`
- `ppt_set_shape_visibility`
- `ppt_bring_to_front`
- `ppt_send_to_back`
- `ppt_bring_forward`
- `ppt_send_backward`
- `ppt_set_shape_shadow`
- `ppt_set_shape_glow`
- `ppt_set_shape_reflection`
- `ppt_set_shape_soft_edges`
- `ppt_set_shape_3d`
- `ppt_merge_shapes`
- `ppt_crop_shape_to_content`
- `ppt_get_slide_text`
- `ppt_find_text`
- `ppt_replace_text_all`
- `ppt_set_slide_title`
- `ppt_get_shape_text_runs`
- `ppt_set_text_range_style`
- `ppt_insert_bullets`
- `ppt_set_bullet_style`
- `ppt_set_paragraph_spacing`
- `ppt_set_textbox_margins`
- `ppt_set_text_direction`
- `ppt_set_autofit`
- `ppt_set_proofing_language`
- `ppt_spellcheck_slide`
- `ppt_spellcheck_presentation`
- `ppt_translate_text`
- `ppt_get_slide_notes`
- `ppt_get_presenter_notes_all`
- `ppt_find_in_notes`
- `ppt_replace_notes_text`
- `ppt_get_slide_transition`
- `ppt_get_slide_animations`
- `ppt_apply_style_preset`
- `ppt_set_slide_transition`
- `ppt_set_shape_text_style`
- `ppt_set_shape_fill`
- `ppt_set_shape_line`
- `ppt_set_slide_background`
- `ppt_add_shape`
- `ppt_add_connector`
- `ppt_connect_shapes`
- `ppt_add_table`
- `ppt_get_slide_tables`
- `ppt_set_table_cell_text`
- `ppt_add_row_to_table`
- `ppt_add_column_to_table`
- `ppt_delete_row_from_table`
- `ppt_delete_column_from_table`
- `ppt_merge_table_cells`
- `ppt_split_table_cells`
- `ppt_set_table_style`
- `ppt_set_table_cell_style`
- `ppt_set_table_row_style`
- `ppt_set_table_column_style`
- `ppt_autofit_table`
- `ppt_distribute_table_rows`
- `ppt_distribute_table_columns`
- `ppt_sort_table`
- `ppt_table_from_csv`
- `ppt_table_from_excel_range`
- `ppt_add_chart`
- `ppt_get_slide_charts`
- `ppt_set_chart_title`
- `ppt_set_chart_data`
- `ppt_set_chart_series_style`
- `ppt_set_chart_layout`
- `ppt_refresh_chart`
- `ppt_set_chart_axis_scale`
- `ppt_set_chart_series_order`
- `ppt_add_chart_series`
- `ppt_delete_chart_series`
- `ppt_set_chart_data_labels`
- `ppt_set_chart_gridlines`
- `ppt_set_chart_colors`
- `ppt_change_chart_type`
- `ppt_export_chart_data`
- `ppt_link_chart_to_excel`
- `ppt_break_chart_link`
- `ppt_add_smartart`
- `ppt_get_slide_smartart`
- `ppt_set_smartart_node_text`
- `ppt_add_shape_animation`
- `ppt_add_element_animation`
- `ppt_clear_slide_animations`
- `ppt_replace_text`
- `ppt_set_slide_notes`
- `ppt_add_slide`
- `ppt_insert_image`
- `ppt_insert_svg`
- `ppt_replace_image`
- `ppt_crop_image`
- `ppt_reset_image`
- `ppt_set_image_transparency`
- `ppt_set_image_brightness`
- `ppt_set_image_contrast`
- `ppt_add_video`
- `ppt_set_video_playback`
- `ppt_trim_video`
- `ppt_add_audio`
- `ppt_set_audio_playback`
- `ppt_extract_media_inventory`
- `ppt_apply_theme`
- `ppt_export_pdf`
- `ppt_export_slide_images`
- `ppt_save_as`

## Arranque rápido

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
pip install -e .
office-ai-mcp-powerpoint --transport stdio
```

Para HTTP:

```powershell
office-ai-mcp-powerpoint --transport streamable-http --host 127.0.0.1 --port 8000
```

## Prueba rápida

Script de inspección:

```powershell
.\.venv\Scripts\python.exe .\scripts\test_mcp_powerpoint.py ".\docs\Presentación - HobbyConnect 2.pptx" --slide-index 1
```

Script visual:

```powershell
.\.venv\Scripts\python.exe .\scripts\test_mcp_powerpoint_visual.py ".\tmp\hobbyconnect-working.pptx" --slide-index 67 --title-shape-index 1 --body-shape-index 2
```

## Casos de uso frecuentes

### Leer notas de una slide

```json
{
  "tool": "ppt_get_slide_notes",
  "arguments": {
    "path": "C:/presentaciones/demo.pptx",
    "slide_index": 3
  }
}
```

### Actualizar notas del presentador

```json
{
  "tool": "ppt_set_slide_notes",
  "arguments": {
    "path": "C:/presentaciones/demo.pptx",
    "slide_index": 3,
    "text": "Recordar explicar la diferencia entre margen bruto y margen neto.",
    "append": false,
    "create_backup": true
  }
}
```

### Animar el texto de un elemento

```json
{
  "tool": "ppt_add_element_animation",
  "arguments": {
    "path": "C:/presentaciones/demo.pptx",
    "slide_index": 1,
    "shape_index": 2,
    "target_kind": "text",
    "animation_level": "all_text_levels",
    "effect": "wipe",
    "trigger": "after_previous",
    "duration_seconds": 0.8,
    "delay_seconds": 0.2,
    "create_backup": true
  }
}
```

## Referencias internas

- Guía general: [../README.md](../README.md)
- Roadmap PowerPoint: [PowerPoint.TODO.md](PowerPoint.TODO.md)
