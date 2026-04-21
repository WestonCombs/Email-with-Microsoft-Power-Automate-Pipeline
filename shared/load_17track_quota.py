"""Lazy import for ``17Track/quota_cache`` — folder name starts with a digit (not importable as a package)."""

from __future__ import annotations

import importlib.util
from pathlib import Path
from types import ModuleType

_cached: ModuleType | None = None


def get_17track_quota_module() -> ModuleType:
    """Load and cache ``python_files/17Track/quota_cache.py``."""
    global _cached
    if _cached is None:
        root = Path(__file__).resolve().parent.parent
        path = root / "17Track" / "quota_cache.py"
        spec = importlib.util.spec_from_file_location(
            "email_sorter_17track_quota_cache", path
        )
        if spec is None or spec.loader is None:
            raise ImportError(f"Cannot load 17TRACK quota module from {path}")
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        _cached = mod
    return _cached
