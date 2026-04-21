"""Check whether PDF capture can start."""

from __future__ import annotations

import re
import shutil

_INVALID_FS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def mitmdump_on_path() -> bool:
    return shutil.which("mitmdump") is not None


def pdf_capture_environment_ready() -> tuple[bool, str | None]:
    """
    Returns (ok, user_message_if_not_ok).

    Only ``mitmdump`` is required up front. Certificate trust is handled during the
    actual browser session.
    """
    if not mitmdump_on_path():
        return (
            False,
            "mitmdump was not found on PATH. Install mitmproxy (pip install mitmproxy) "
            "or add its Scripts folder to PATH.\n\n"
            "Use Email Sorter -> Settings -> Begin -> MITM wizard for a guided setup.",
        )
    return True, None


def sanitize_filename_token(s: str) -> str:
    """Single segment safe for use inside a PDF basename."""
    t = " ".join((s or "").strip().split())
    t = _INVALID_FS.sub("_", t)
    return t or "unknown"
