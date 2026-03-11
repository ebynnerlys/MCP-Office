# Excel MCP

Guía específica del servidor MCP de Excel.

## Servidor

- Comando CLI: `office-ai-mcp-excel`
- Módulo Python: `python -m office_ai_mcp.excel_server`
- Servidor en VS Code: `officeAiExcel`

La configuración de VS Code está en [../.vscode/mcp.json](../.vscode/mcp.json).

## Alcance

Este servidor expone únicamente herramientas de Excel, además de las herramientas de sistema compartidas:

- `server_status`
- `create_working_backup`

## Herramientas disponibles

- `excel_list_sheets`
- `excel_read_range`
- `excel_write_range`
- `excel_export_pdf`

## Arranque rápido

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
pip install -e .
office-ai-mcp-excel --transport stdio
```

Para HTTP:

```powershell
office-ai-mcp-excel --transport streamable-http --host 127.0.0.1 --port 8000
```

## Casos de uso frecuentes

### Listar hojas

```json
{
  "tool": "excel_list_sheets",
  "arguments": {
    "path": "C:/datos/ventas.xlsx"
  }
}
```

### Leer un rango

```json
{
  "tool": "excel_read_range",
  "arguments": {
    "path": "C:/datos/ventas.xlsx",
    "sheet": "Resumen",
    "cell_range": "A1:D10"
  }
}
```

### Escribir un rango

```json
{
  "tool": "excel_write_range",
  "arguments": {
    "path": "C:/datos/ventas.xlsx",
    "sheet": "Resumen",
    "cell_range": "A1:B2",
    "values": [
      ["Mes", "Total"],
      ["Enero", 1250]
    ],
    "create_backup": true
  }
}
```

### Exportar a PDF

```json
{
  "tool": "excel_export_pdf",
  "arguments": {
    "path": "C:/datos/ventas.xlsx",
    "out_path": "C:/datos/ventas.pdf"
  }
}
```

## Referencias internas

- Guía general: [../README.md](../README.md)
- Roadmap Excel: [Excel.TODO.md](Excel.TODO.md)
