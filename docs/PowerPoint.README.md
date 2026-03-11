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
- `ppt_get_slide_shapes`
- `ppt_get_slide_text`
- `ppt_get_slide_notes`
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
- `ppt_add_chart`
- `ppt_get_slide_charts`
- `ppt_set_chart_title`
- `ppt_set_chart_data`
- `ppt_set_chart_series_style`
- `ppt_set_chart_layout`
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
