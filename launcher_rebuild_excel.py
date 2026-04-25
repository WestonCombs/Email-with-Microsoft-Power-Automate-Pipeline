"""Rebuild orders workbook in a subprocess (launcher reads stdout for progress lines)."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

_PYTHON_FILES_DIR = Path(__file__).resolve().parent

sys.path.insert(0, str(_PYTHON_FILES_DIR))
from shared.stdio_utf8 import configure_stdio_utf8, console_safe_text  # noqa: E402

configure_stdio_utf8()


def _load_create_excel_module():
    mod_path = _PYTHON_FILES_DIR / "createExcelDocument" / "createExcelDocument.py"
    spec = importlib.util.spec_from_file_location("_email_sorter_ced_launcher", mod_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load {mod_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main() -> None:
    if len(sys.argv) < 2:
        print("usage: launcher_rebuild_excel.py <excel_output_path>", file=sys.stderr)
        sys.exit(2)
    out = sys.argv[1]
    from shared.settings_store import apply_runtime_settings_from_json

    apply_runtime_settings_from_json()
    mod = _load_create_excel_module()
    mod.rebuild_orders_workbook(out)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nERROR: {console_safe_text(e)}", file=sys.stderr)
        sys.exit(1)
