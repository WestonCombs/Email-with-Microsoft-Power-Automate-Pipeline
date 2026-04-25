"""Console UTF-8 setup and safe printable strings (Windows cp1252 / odd Unicode from email)."""

from __future__ import annotations

import sys
from typing import TextIO


def _reconfigure_stream(stream: TextIO | None) -> None:
    if stream is None:
        return
    reconfigure = getattr(stream, "reconfigure", None)
    if not callable(reconfigure):
        return
    try:
        reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def configure_stdio_utf8() -> None:
    """Prefer UTF-8 for stdout/stderr; idempotent; never raises."""
    for name in ("stdout", "stderr"):
        _reconfigure_stream(getattr(sys, name, None))


def console_safe_text(value: object) -> str:
    """Text that can be passed to print() without UnicodeEncodeError for current stdout."""
    text = "" if value is None else str(value)
    encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
    try:
        text.encode(encoding)
        return text
    except UnicodeEncodeError:
        return text.encode(encoding, errors="replace").decode(encoding, errors="replace")
