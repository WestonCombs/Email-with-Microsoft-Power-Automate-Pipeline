"""Resolve the project root (historically ``BASE_DIR``) without storing it in ``.env``.

Layout: ``<project>/<python_files>/…`` — project root is the parent of the ``python_files``
directory containing this repo's scripts.
"""

from __future__ import annotations

import os
from pathlib import Path

_SHARED_DIR = Path(__file__).resolve().parent
_PYTHON_FILES_DIR = _SHARED_DIR.parent


def python_files_dir() -> Path:
    """Absolute path to the ``python_files`` directory."""
    return _PYTHON_FILES_DIR


def inferred_base_dir() -> Path:
    """Infer project root from repository layout."""
    pf = _PYTHON_FILES_DIR
    if pf.name.lower() == "python_files":
        return pf.parent.resolve()
    return pf.resolve()


def ensure_base_dir_in_environ() -> Path:
    """Set ``os.environ[\"BASE_DIR\"]`` to the inferred project root and return it."""
    base = inferred_base_dir()
    os.environ["BASE_DIR"] = str(base)
    return base
