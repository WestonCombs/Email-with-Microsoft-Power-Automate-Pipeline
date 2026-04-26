from __future__ import annotations

import argparse
import os
import sys
import traceback
from datetime import datetime, timezone
from pathlib import Path

_PYTHON_FILES_DIR = Path(__file__).resolve().parent.parent
if str(_PYTHON_FILES_DIR) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES_DIR))

from shared.stdio_utf8 import configure_stdio_utf8, console_safe_text

configure_stdio_utf8()

from shared.excel_user_edits import (  # noqa: E402
    ALLOWED_EXCEL_USER_EDIT_FIELDS,
    record_excel_user_edit,
)
from shared.project_paths import ensure_base_dir_in_environ  # noqa: E402
from shared.settings_store import apply_runtime_settings_from_json  # noqa: E402

_USER_EDIT_LOG_NAME = "email_sorter_user_edit.log"


def _user_edit_log_path() -> Path:
    base = os.environ.get("TEMP") or os.environ.get("TMP") or str(Path.cwd())
    return Path(base) / _USER_EDIT_LOG_NAME


def _append_user_edit_log(line: str) -> None:
    path = _user_edit_log_path()
    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    try:
        with path.open("a", encoding="utf-8", newline="\n") as fh:
            fh.write(f"{stamp} [python] {line}\n")
    except OSError:
        pass


def _user_edit_result_path(context_tsv: Path) -> Path:
    return context_tsv.with_suffix(context_tsv.suffix + ".out.tsv")


def _read_context_tsv(path: Path) -> dict[str, str]:
    out: dict[str, str] = {}
    text = path.read_text(encoding="utf-8-sig")
    for raw_line in text.splitlines():
        if not raw_line.strip() or "\t" not in raw_line:
            continue
        key, value = raw_line.split("\t", 1)
        out[key.strip()] = value.strip()
    return out


def _write_result_tsv(path: Path, summary: dict) -> None:
    def _clean_piece(value: object) -> str:
        text = "" if value is None else str(value)
        return text.replace("\r", " ").replace("\n", " ").replace("\t", " ")

    lines = [
        f"mode\t{_clean_piece(summary.get('mode'))}\n",
        f"display_value_kind\t{_clean_piece(summary.get('display_value_kind'))}\n",
        f"display_value\t{_clean_piece(summary.get('display_value'))}\n",
    ]
    path.write_text("".join(lines), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Apply a user edit from Orders.xlsm to Email Sorter JSON."
    )
    parser.add_argument("context_tsv", type=Path)
    args = parser.parse_args(argv)

    try:
        apply_runtime_settings_from_json()
        project_root = ensure_base_dir_in_environ()
        ctx = _read_context_tsv(args.context_tsv)
        field = ctx.get("field", "")
        order_no = ctx.get("order_number", "")
        src_uri = ctx.get("source_uri", "")
        _append_user_edit_log(
            f"context_tsv={args.context_tsv} project_root={project_root} "
            f"field={field!r} order_number_len={len(order_no)} source_uri_len={len(src_uri)}"
        )
        if field not in ALLOWED_EXCEL_USER_EDIT_FIELDS:
            raise ValueError(f"Unsupported field: {field}")
        summary = record_excel_user_edit(
            project_root,
            field=field,
            raw_value=ctx.get("value", ""),
            order_number=order_no,
            source_uri=src_uri,
        )
        _write_result_tsv(_user_edit_result_path(args.context_tsv), summary)
        _append_user_edit_log(
            "OK mode=%s matched_records=%s changed=%s"
            % (
                summary.get("mode"),
                summary.get("matched_records"),
                summary.get("changed_files"),
            )
        )
        return 0
    except Exception as exc:
        tb = traceback.format_exc()
        err_path = args.context_tsv.with_suffix(args.context_tsv.suffix + ".err.txt")
        try:
            err_path.write_text(
                console_safe_text(exc) + "\n\n" + console_safe_text(tb),
                encoding="utf-8",
            )
        except OSError:
            pass
        _append_user_edit_log(
            "FAIL %s: %s — full traceback: %s"
            % (type(exc).__name__, exc, err_path)
        )
        print(console_safe_text(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
