from __future__ import annotations

import argparse
import json
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
from shared.fix_flagged import sync_fix_flagged_requests  # noqa: E402
from shared.order_store import set_record_excel_flagged  # noqa: E402
from shared.project_paths import ensure_base_dir_in_environ  # noqa: E402
from shared.settings_store import apply_runtime_settings_from_json  # noqa: E402

_USER_EDIT_LOG_NAME = "email_sorter_user_edit.log"
FLAGGED_FIELD = "excel_flagged"
LEGACY_ACTIVE_FIELD = "excel_active"


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


def _sync_email_asset_names_after_edit(project_root: Path) -> None:
    results_path = project_root / "email_contents" / "json" / "results.json"
    if not results_path.is_file():
        return
    from grabbingImportantEmailContent.grabbingImportantEmailContent import (
        apply_order_company_consensus_and_sync,
    )

    records = json.loads(results_path.read_text(encoding="utf-8"))
    if not isinstance(records, list):
        return
    apply_order_company_consensus_and_sync(records, project_root)
    results_path.write_text(
        json.dumps(records, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def _parse_flagged_value(value: object) -> bool:
    text = str(value or "").strip().casefold()
    return text in {"1", "true", "yes", "y", "active", "flagged", "checked", "x"}


def _record_flagged_edit(project_root: Path, ctx: dict[str, str]) -> dict:
    record_id = ctx.get("record_id", "")
    if not record_id:
        raise ValueError("Flagged edits need a Record ID. Regenerate the workbook and try again.")
    flagged = _parse_flagged_value(ctx.get("value", ""))
    result = set_record_excel_flagged(
        project_root,
        record_id=record_id,
        flagged=flagged,
        source_uri=ctx.get("source_uri", ""),
        order_number=ctx.get("order_number", ""),
        sync_open_excel=False,
    )
    sync_fix_flagged_requests(project_root)
    return {
        "field": FLAGGED_FIELD,
        "value": flagged,
        "record_id": record_id,
        "matched_records": result.get("changed_records", 1),
        "changed_files": result.get("changed_files") or [
            str(project_root / "email_contents" / "json" / "results.json")
        ],
        "mode": "modified" if flagged else "cleared",
        "display_value": "True" if flagged else "",
        "display_value_kind": "text" if flagged else "blank",
    }


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
        if field == LEGACY_ACTIVE_FIELD:
            field = FLAGGED_FIELD
        if field != FLAGGED_FIELD and field not in ALLOWED_EXCEL_USER_EDIT_FIELDS:
            raise ValueError(f"Unsupported field: {field}")
        if field == FLAGGED_FIELD:
            summary = _record_flagged_edit(project_root, ctx)
        else:
            summary = record_excel_user_edit(
                project_root,
                field=field,
                raw_value=ctx.get("value", ""),
                order_number=order_no,
                source_uri=src_uri,
                record_id=ctx.get("record_id", ""),
            )
        try:
            _sync_email_asset_names_after_edit(project_root)
        except Exception as sync_exc:
            _append_user_edit_log(
                f"WARN filename sync skipped after edit: {type(sync_exc).__name__}: {sync_exc}"
            )
        try:
            sync_fix_flagged_requests(project_root)
        except Exception as sync_exc:
            _append_user_edit_log(
                f"WARN Fix Flagged sync skipped after edit: {type(sync_exc).__name__}: {sync_exc}"
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
