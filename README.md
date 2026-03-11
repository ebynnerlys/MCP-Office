# Office AI MCP

Proyecto MCP para automatización de Office en Windows mediante COM y `pywin32`, dividido en tres servidores independientes.

## Requisitos

- Windows 10 u 11
- Python 3.11 o superior
- Microsoft Office de escritorio instalado

## Servidores disponibles

- `office-ai-mcp-powerpoint`
- `office-ai-mcp-word`
- `office-ai-mcp-excel`

El servidor combinado `office-ai-mcp` ya no se usa.

## Arranque rápido

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
pip install -e .
office-ai-mcp-powerpoint --transport stdio
```

## Configuración en VS Code

Los tres servidores están definidos en [ .vscode/mcp.json ](.vscode/mcp.json).

## Documentación por servidor

### PowerPoint

- [docs/PowerPoint.README.md](docs/PowerPoint.README.md)
- [docs/PowerPoint.TODO.md](docs/PowerPoint.TODO.md)

### Word

- [docs/Word.README.md](docs/Word.README.md)
- [docs/Word.TODO.md](docs/Word.TODO.md)

### Excel

- [docs/Excel.README.md](docs/Excel.README.md)
- [docs/Excel.TODO.md](docs/Excel.TODO.md)

## Estructura

```text
.
├── docs/
├── src/office_ai_mcp/
│   ├── models/
│   ├── services/
│   ├── tools/
│   └── utils/
├── tests/
├── informe-editor-ia-office-python.md
├── pyproject.toml
└── requirements.txt
```
