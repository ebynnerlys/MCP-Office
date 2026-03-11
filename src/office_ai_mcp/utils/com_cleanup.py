from __future__ import annotations

import platform
from contextlib import contextmanager, suppress
from typing import Iterator


def ensure_windows() -> None:
    if platform.system() != "Windows":
        raise RuntimeError("Office automation via COM requires Windows")


@contextmanager
def office_application(prog_id: str, *, visible: bool = False) -> Iterator[object]:
    ensure_windows()

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("pywin32 is required for Office automation") from exc

    pythoncom.CoInitialize()
    application = None
    try:
        application = win32com.client.DispatchEx(prog_id)
        with suppress(Exception):
            application.Visible = visible
        yield application
    finally:
        if application is not None:
            with suppress(Exception):
                application.DisplayAlerts = False
            with suppress(Exception):
                application.Quit()
        pythoncom.CoUninitialize()
