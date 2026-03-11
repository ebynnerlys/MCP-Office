from __future__ import annotations

from pathlib import Path

from office_ai_mcp.config import Settings
from office_ai_mcp.utils.backups import create_backup
from office_ai_mcp.utils.paths import ensure_directory, validate_file_path


class OfficeService:
    allowed_suffixes: tuple[str, ...] = ()

    def __init__(self, settings: Settings) -> None:
        self.settings = settings

    def resolve_document_path(self, path: str | Path) -> Path:
        return validate_file_path(
            path,
            allowed_roots=self.settings.allowed_roots,
            allowed_suffixes=self.allowed_suffixes,
            must_exist=True,
        )

    def resolve_output_path(self, path: str | Path, *, allowed_suffixes: tuple[str, ...]) -> Path:
        target = validate_file_path(
            path,
            allowed_roots=self.settings.allowed_roots,
            allowed_suffixes=allowed_suffixes,
            must_exist=False,
        )
        target.parent.mkdir(parents=True, exist_ok=True)
        return target

    def maybe_create_backup(self, path: Path, create_copy: bool) -> str | None:
        if not create_copy:
            return None
        backup_dir = ensure_directory(self.settings.backup_dir)
        return str(create_backup(path, backup_dir))
