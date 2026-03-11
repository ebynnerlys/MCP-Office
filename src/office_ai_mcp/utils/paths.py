from __future__ import annotations

from pathlib import Path
from typing import Iterable


def resolve_workspace_path(raw_path: str | Path) -> Path:
    path = Path(raw_path).expanduser()
    if path.is_absolute():
        return path.resolve()
    return (Path.cwd() / path).resolve()


def ensure_directory(path: str | Path) -> Path:
    directory = resolve_workspace_path(path)
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def ensure_parent_directory(path: str | Path) -> Path:
    target = resolve_workspace_path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    return target


def validate_file_path(
    raw_path: str | Path,
    *,
    allowed_roots: Iterable[str] | None = None,
    allowed_suffixes: Iterable[str] | None = None,
    must_exist: bool = True,
) -> Path:
    path = resolve_workspace_path(raw_path)

    if allowed_roots:
        resolved_roots = [resolve_workspace_path(item) for item in allowed_roots]
        if not any(path.is_relative_to(root) for root in resolved_roots):
            raise ValueError(f"Path is outside allowed roots: {path}")

    if allowed_suffixes:
        normalized_suffixes = {suffix.lower() for suffix in allowed_suffixes}
        if path.suffix.lower() not in normalized_suffixes:
            raise ValueError(f"Unsupported file type: {path.suffix}")

    if must_exist and not path.exists():
        raise FileNotFoundError(path)

    return path
