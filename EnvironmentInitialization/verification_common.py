"""Shared path safety and logging for folder/file verification (EnvironmentInitialization)."""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path


# Same directory as FolderVerification.py / fileVerification.py
_CONSOLE_LOG = Path(__file__).resolve().parent / "console_log"


def verification_log(message: str) -> None:
    """Append one line to ``EnvironmentInitialization/console_log`` and print to stdout."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {message}\n"
    try:
        _CONSOLE_LOG.parent.mkdir(parents=True, exist_ok=True)
        with open(_CONSOLE_LOG, "a", encoding="utf-8") as f:
            f.write(line)
    except OSError:
        pass
    print(message, flush=True)


def resolved_root(project_root: Path | str) -> Path:
    return Path(project_root).expanduser().resolve()


def path_under_root(resolved_base: Path, relative_path: str) -> Path:
    """Join *relative_path* onto *resolved_base* without ``..`` escape.

    Does not require intermediate directories to exist (avoids flaky ``Path.resolve()``
    behavior on Windows when parents are missing).
    """
    rel = relative_path.strip()
    if not rel:
        raise ValueError("relative_path must be non-empty")
    tail = Path(rel)
    if tail.is_absolute():
        raise ValueError("relative_path must be relative to BASE_DIR (not an absolute path)")
    if ".." in tail.parts:
        raise ValueError("relative_path must stay under BASE_DIR (no '..')")

    base_s = os.path.normpath(str(resolved_base))
    combined = os.path.normpath(os.path.join(base_s, *tail.parts))
    base_n = os.path.normcase(base_s)
    combined_n = os.path.normcase(combined)
    if combined_n != base_n and not combined_n.startswith(base_n + os.sep):
        raise ValueError("relative_path must stay under BASE_DIR (no '..')")
    return Path(combined)


def ensure_parent_chain_exists(leaf: Path) -> None:
    """Create every missing parent directory up to *leaf*'s parent (``mkdir -p``)."""
    leaf.parent.mkdir(parents=True, exist_ok=True)
