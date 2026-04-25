"""Check whether proof-of-delivery capture (isolated Chrome + Ctrl+Shift+P) can start."""

from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

_INVALID_FS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

_PY_FILES = Path(__file__).resolve().parent.parent
if str(_PY_FILES) not in sys.path:
    sys.path.insert(0, str(_PY_FILES))

from pdfCaptureFromChrome.launch_mitm_chrome import find_chrome_executable  # noqa: E402


def mitmdump_on_path() -> bool:
    """True if ``mitmdump`` is on PATH (for manual / archival ``BASE_DIR/mitm_pdf_capture/run_pdf_capture.py``)."""
    return shutil.which("mitmdump") is not None


def _websocket_client_installed() -> bool:
    try:
        import websocket  # noqa: F401
    except ImportError:
        return False
    return True


def pdf_capture_environment_ready() -> tuple[bool, str | None]:
    """
    Returns (ok, user_message_if_not_ok).

    The viewer no longer uses mitmproxy. Requires Google Chrome, ``websocket-client``,
    and (for the global hotkey) Windows.
    """
    if sys.platform != "win32":
        return (
            False,
            "PDF capture (Ctrl+Shift+P) is only supported on Windows in this build.",
        )
    if find_chrome_executable() is None:
        return (
            False,
            "Google Chrome (chrome.exe) was not found. Install Google Chrome, "
            "or if it is a non-standard path, the html capture code must be told where it is.",
        )
    if not _websocket_client_installed():
        return (
            False,
            "The websocket-client package is required. Install with:\n\n"
            "  pip install websocket-client",
        )
    return True, None


def sanitize_filename_token(s: str) -> str:
    """Single segment safe for use inside a PDF basename."""
    t = " ".join((s or "").strip().split())
    t = _INVALID_FS.sub("_", t)
    return t or "unknown"
