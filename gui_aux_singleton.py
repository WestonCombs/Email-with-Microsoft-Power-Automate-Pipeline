"""Single-instance guard + console detach for Excel-launched GUI helpers (Windows)."""

from __future__ import annotations

import atexit
import os
import subprocess
import sys
import tempfile
from pathlib import Path

_PID = Path(tempfile.gettempdir()) / "email_sorter_aux_gui.pid"


def detach_console_win32() -> None:
    """Detach from the console so ``python.exe`` does not leave a black window behind tk."""
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.kernel32.FreeConsole()
    except Exception:
        pass


def kill_previous_aux_gui() -> None:
    """Terminate the previous tracking viewer or gift-link helper, if still running."""
    if sys.platform != "win32" or not _PID.is_file():
        return
    try:
        old_pid = int(_PID.read_text(encoding="utf-8").strip())
    except (ValueError, OSError):
        _PID.unlink(missing_ok=True)
        return
    if old_pid <= 0 or old_pid == os.getpid():
        return
    kwargs: dict = {
        "capture_output": True,
        "stdin": subprocess.DEVNULL,
    }
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW  # type: ignore[assignment]
    subprocess.run(
        ["taskkill", "/PID", str(old_pid), "/F", "/T"],
        **kwargs,
    )
    _PID.unlink(missing_ok=True)


def register_current_aux_gui() -> None:
    """Kill any prior aux GUI, then record this process so the next launch replaces us."""
    kill_previous_aux_gui()
    try:
        _PID.write_text(str(os.getpid()), encoding="utf-8")
    except OSError:
        return

    def _cleanup() -> None:
        try:
            if _PID.is_file() and _PID.read_text(encoding="utf-8").strip() == str(os.getpid()):
                _PID.unlink(missing_ok=True)
        except OSError:
            pass

    atexit.register(_cleanup)
