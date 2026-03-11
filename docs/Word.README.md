# Word MCP

Guía específica del servidor MCP de Word.

## Servidor

- Comando CLI: `office-ai-mcp-word`
- Módulo Python: `python -m office_ai_mcp.word_server`
- Servidor en VS Code: `officeAiWord`

La configuración de VS Code está en [../.vscode/mcp.json](../.vscode/mcp.json).

## Alcance

Este servidor expone únicamente herramientas de Word, además de las herramientas de sistema compartidas:

- `server_status`
- `create_working_backup`

## Herramientas disponibles

- `word_get_structure`
- `word_replace_text`
- `word_export_pdf`

## Arranque rápido

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
pip install -e .
office-ai-mcp-word --transport stdio
```

Para HTTP:

```powershell
office-ai-mcp-word --transport streamable-http --host 127.0.0.1 --port 8000
```

## Casos de uso frecuentes

### Inspeccionar la estructura de un documento

```json
{
  "tool": "word_get_structure",
  "arguments": {
    "path": "C:/documentos/informe.docx"
  }
}
```

### Reemplazar texto

```json
{
  "tool": "word_replace_text",
  "arguments": {
    "path": "C:/documentos/informe.docx",
    "find": "Cliente X",
    "replace": "Cliente Y",
    "create_backup": true
  }
}
```

### Exportar a PDF

```json
{
  "tool": "word_export_pdf",
  "arguments": {
    "path": "C:/documentos/informe.docx",
    "out_path": "C:/documentos/informe.pdf"
  }
}
```

## Referencias internas

- Guía general: [../README.md](../README.md)
- Roadmap Word: [Word.TODO.md](Word.TODO.md)
