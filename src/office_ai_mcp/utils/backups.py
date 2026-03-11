from __future__ import annotations

import shutil
from datetime import datetime, timezone
from pathlib import Path

from office_ai_mcp.utils.paths import ensure_directory, validate_file_path


def create_backup(source_path: str | Path, backup_root: str | Path) -> Path:
    source = validate_file_path(source_path, must_exist=True)
    destination_root = ensure_directory(backup_root)
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    destination = destination_root / f"{source.stem}-{timestamp}{source.suffix}"
    shutil.copy2(source, destination)
    return destination
