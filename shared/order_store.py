from __future__ import annotations

import hashlib
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname

from shared.excel_user_edits import (
    PURCHASE_DATETIME_USER_EDIT_CONFIDENCE,
    PURCHASE_DATETIME_USER_EDIT_SOURCE,
    coerce_user_edit_value,
    display_value_for_excel,
    modified_key,
    record_identity,
)

RECORD_ID_FIELD = "_record_id"
RECORD_ID_HEADER = "Record ID"

EXCEL_OWNED_FIELDS = (
    "company",
    "order_number",
    "purchase_datetime",
    "email_category",
    "subtotal_amount",
    "total_amount_paid",
    "tax_paid",
    "gift_card_amount",
    "accounting",
)

EXCEL_FIELD_HEADERS = {
    "excel_flagged": "Flagged",
    "excel_active": "Active",
    "company": "Company",
    "order_number": "Order Number",
    "purchase_datetime": "Purchase Date",
    "email_category": "Category",
    "subtotal_amount": "Subtotal",
    "total_amount_paid": "Total Paid",
    "tax_paid": "Tax Paid",
    "gift_card_amount": "GC Paid",
    "accounting": "Accounting",
}


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def results_json_path(project_root: Path) -> Path:
    return project_root / "email_contents" / "json" / "results.json"


def proof_of_delivery_json_path(project_root: Path) -> Path:
    return project_root / "email_contents" / "json" / "proof_of_delivery.json"


def load_json_records(path: Path) -> list[dict]:
    if not path.is_file():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    return [item for item in payload if isinstance(item, dict)] if isinstance(payload, list) else []


def save_json_records(path: Path, records: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8")


def stable_record_id(record: dict) -> str:
    existing = str(record.get(RECORD_ID_FIELD) or "").strip()
    if existing:
        return existing
    raw = record_identity(record)
    digest = hashlib.sha1(raw.encode("utf-8", errors="ignore")).hexdigest()[:20]
    return f"ord_{digest}"


def ensure_record_ids(records: list[dict]) -> bool:
    changed = False
    seen: set[str] = set()
    for idx, record in enumerate(records):
        if not isinstance(record, dict):
            continue
        record_id = stable_record_id(record)
        if record_id in seen:
            raw = f"{record_id}|{idx}|{record_identity(record)}"
            record_id = "ord_" + hashlib.sha1(raw.encode("utf-8", errors="ignore")).hexdigest()[:20]
        seen.add(record_id)
        if record.get(RECORD_ID_FIELD) != record_id:
            record[RECORD_ID_FIELD] = record_id
            changed = True
    return changed


def ensure_results_record_ids(project_root: Path) -> bool:
    path = results_json_path(project_root)
    records = load_json_records(path)
    if not records:
        return False
    changed = ensure_record_ids(records)
    if changed:
        save_json_records(path, records)
    return changed


def load_results_records(project_root: Path, *, persist_ids: bool = True) -> list[dict]:
    path = results_json_path(project_root)
    records = load_json_records(path)
    changed = ensure_record_ids(records)
    if changed and persist_ids:
        save_json_records(path, records)
    return records


def save_results_records(project_root: Path, records: list[dict]) -> None:
    ensure_record_ids(records)
    save_json_records(results_json_path(project_root), records)


def find_record_index_by_record_id(records: list[dict], record_id: str) -> int | None:
    wanted = str(record_id or "").strip()
    if not wanted:
        return None
    for idx, record in enumerate(records):
        if str(record.get(RECORD_ID_FIELD) or "").strip() == wanted:
            return idx
    return None


def find_record_index_by_identity(records: list[dict], wanted_identity: str) -> int | None:
    for idx, record in enumerate(records):
        if record_identity(record) == wanted_identity:
            return idx
    return None


def _source_identity(value: object) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    parsed = urlparse(raw)
    if parsed.scheme.casefold() == "file":
        path = unquote(url2pathname(parsed.path or ""))
        path = path.replace("\\", "/")
        if len(path) >= 3 and path[0] == "/" and path[2] == ":":
            path = path[1:]
        return "file:" + path.casefold()
    return unquote(raw).replace("\\", "/").casefold()


def record_matches_source_uri(record: dict, source_uri: str) -> bool:
    wanted = _source_identity(source_uri)
    if not wanted:
        return False
    for key in ("source_file_link", "source_file"):
        if _source_identity(record.get(key)) == wanted:
            return True
    return False


def find_record_index_for_excel_context(
    records: list[dict],
    *,
    record_id: str = "",
    source_uri: str = "",
    order_number: str = "",
    allow_order_fallback: bool = True,
) -> int | None:
    idx = find_record_index_by_record_id(records, record_id)
    if idx is not None:
        return idx

    wanted_id = str(record_id or "").strip()
    if wanted_id:
        for idx, record in enumerate(records):
            if stable_record_id(record) == wanted_id:
                return idx

    source_matches = [
        idx
        for idx, record in enumerate(records)
        if record_matches_source_uri(record, source_uri)
    ]
    if len(source_matches) == 1:
        return source_matches[0]

    wanted_order = str(order_number or "").strip()
    if allow_order_fallback and wanted_order:
        order_matches = [
            idx
            for idx, record in enumerate(records)
            if str(record.get("order_number") or "").strip() == wanted_order
        ]
        if len(order_matches) == 1:
            return order_matches[0]
    return None


def coerce_excel_owned_value(field: str, value: object) -> object:
    if field not in EXCEL_OWNED_FIELDS:
        raise ValueError(f"Unsupported Excel-owned field: {field}")
    return coerce_user_edit_value(field, value)


def update_excel_owned_fields(
    project_root: Path,
    *,
    record_id: str,
    updates: dict[str, object],
    fanout_order_fields: set[str] | None = None,
    sync_open_excel: bool = True,
) -> dict:
    records = load_results_records(project_root)
    idx = find_record_index_by_record_id(records, record_id)
    if idx is None:
        raise ValueError("Could not match this item to a row in results.json.")

    cleaned_updates = {
        field: coerce_excel_owned_value(field, value)
        for field, value in updates.items()
        if field in EXCEL_OWNED_FIELDS
    }
    if not cleaned_updates:
        return {"changed_records": 0, "record": records[idx], "records": records}

    fanout_order_fields = fanout_order_fields or set()
    now = _utc_now()
    target = records[idx]
    original_order = str(target.get("order_number") or "").strip()
    changed_record_ids: set[str] = set()

    for record in records:
        same_target = str(record.get(RECORD_ID_FIELD) or "") == record_id
        same_order = bool(original_order) and str(record.get("order_number") or "").strip() == original_order
        fields_to_apply: dict[str, object] = {}
        for field, value in cleaned_updates.items():
            if same_target or (field in fanout_order_fields and same_order):
                fields_to_apply[field] = value
        if not fields_to_apply:
            continue

        for field, value in fields_to_apply.items():
            record[field] = value
            record[modified_key(field)] = True
            if field == "purchase_datetime":
                record["purchase_datetime_source"] = PURCHASE_DATETIME_USER_EDIT_SOURCE
                record["purchase_datetime_confidence"] = PURCHASE_DATETIME_USER_EDIT_CONFIDENCE
        record["user_modified"] = True
        record["user_modified_at"] = now
        changed_record_ids.add(str(record.get(RECORD_ID_FIELD) or ""))

    save_results_records(project_root, records)

    if sync_open_excel:
        try_sync_open_excel_records(
            project_root,
            records=records,
            record_ids=changed_record_ids,
            fields=set(cleaned_updates),
            show_success=True,
        )

    refreshed_idx = find_record_index_by_record_id(records, record_id)
    return {
        "changed_records": len(changed_record_ids),
        "record": records[refreshed_idx if refreshed_idx is not None else idx],
        "records": records,
    }


def resolve_orders_workbook_candidates(project_root: Path) -> list[Path]:
    raw = os.getenv("EXCEL_OUTPUT_PATH")
    candidates: list[Path] = []
    if raw:
        configured = Path(raw).expanduser().resolve()
        candidates.append(configured)
        if configured.suffix.lower() == ".xlsm":
            candidates.append(configured.with_suffix(".xlsx"))
        elif configured.suffix.lower() == ".xlsx":
            candidates.append(configured.with_suffix(".xlsm"))
    candidates.extend(
        [
            project_root / "email_contents" / "orders.xlsm",
            project_root / "email_contents" / "orders.xlsx",
        ]
    )
    deduped: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate.expanduser().resolve()).casefold()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(candidate.expanduser().resolve())
    return deduped


def _header_map(ws) -> dict[str, int]:
    header_row = 2
    try:
        row_two_markers = {
            str(ws.Cells(header_row, col).Value or "").strip().casefold()
            for col in range(1, 5)
        }
        if not row_two_markers.intersection({"flagged", "active", "category", "order number"}):
            header_row = 1
        last_col = int(ws.Cells(header_row, ws.Columns.Count).End(-4159).Column)  # xlToLeft
    except Exception:
        return {}
    headers: dict[str, int] = {}
    for col in range(1, last_col + 1):
        value = str(ws.Cells(header_row, col).Value or "").strip()
        if value:
            headers[value.casefold()] = col
    return headers


def _find_record_row(ws, headers: dict[str, int], record: dict) -> int | None:
    record_col = headers.get(RECORD_ID_HEADER.casefold())
    source_col = 29
    order_col = headers.get(EXCEL_FIELD_HEADERS["order_number"].casefold())
    wanted_id = str(record.get(RECORD_ID_FIELD) or "").strip()
    wanted_source = str(record.get("source_file_link") or record.get("source_file") or "").strip()
    wanted_order = str(record.get("order_number") or "").strip()
    try:
        scan_col = record_col or order_col or source_col or 1
        last_row = int(ws.Cells(ws.Rows.Count, scan_col).End(-4162).Row)  # xlUp
    except Exception:
        return None
    for row in range(3, last_row + 1):
        if record_col and wanted_id:
            got = str(ws.Cells(row, record_col).Value or "").strip()
            if got == wanted_id:
                return row
        if wanted_source:
            got_source = str(ws.Cells(row, source_col).Value or "").strip()
            if got_source == wanted_source:
                return row
        if order_col and wanted_order and not wanted_source:
            got_order = str(ws.Cells(row, order_col).Value or "").strip()
            if got_order == wanted_order:
                return row
    return None


def _run_success_rainbow(excel, workbook_name: str) -> None:
    safe_name = workbook_name.replace("'", "''")
    try:
        excel.Run(f"'{safe_name}'!EmailSorter_ShowSuccessRainbowForActiveSheet")
    except Exception:
        pass


def is_record_excel_flagged(record: dict) -> bool:
    return bool(record.get("excel_flagged") or record.get("excel_active"))


def excel_display_value(record: dict, field: str) -> object:
    if field in {"excel_flagged", "excel_active"}:
        return "True" if is_record_excel_flagged(record) else None
    return display_value_for_excel(record, field, record.get(field))


def try_sync_open_excel_records(
    project_root: Path,
    *,
    records: list[dict],
    record_ids: set[str],
    fields: set[str],
    show_success: bool = False,
) -> bool:
    if sys.platform != "win32" or not record_ids or not fields:
        return False
    try:
        from giftcardInvoiceLink.excel_link_sync import find_workbook_and_application
    except Exception:
        return False

    targets = {str(record_id) for record_id in record_ids if record_id}
    records_by_id = {
        str(record.get(RECORD_ID_FIELD) or ""): record
        for record in records
        if str(record.get(RECORD_ID_FIELD) or "") in targets
    }
    if not records_by_id:
        return False

    for path in resolve_orders_workbook_candidates(project_root):
        wb = None
        excel = None
        try:
            wb, excel = find_workbook_and_application(str(path))
        except Exception:
            wb = None
        if wb is None:
            continue
        try:
            ws = wb.Worksheets("Orders")
            headers = _header_map(ws)
            updated = False
            old_events = None
            if excel is not None:
                try:
                    old_events = excel.EnableEvents
                    excel.EnableEvents = False
                except Exception:
                    old_events = None
            for record in records_by_id.values():
                row = _find_record_row(ws, headers, record)
                if row is None:
                    continue
                for field in fields:
                    header = EXCEL_FIELD_HEADERS.get(field)
                    if not header:
                        continue
                    col = headers.get(header.casefold())
                    if not col:
                        continue
                    value = excel_display_value(record, field)
                    ws.Cells(row, col).Value = value
                    updated = True
            if excel is not None and old_events is not None:
                try:
                    excel.EnableEvents = old_events
                except Exception:
                    pass
            if updated and excel is not None:
                if show_success:
                    _run_success_rainbow(excel, str(wb.Name))
                return True
        except Exception:
            if excel is not None:
                try:
                    excel.EnableEvents = True
                except Exception:
                    pass
            continue
    return False


def try_remove_open_excel_record(
    project_root: Path,
    *,
    record: dict,
    show_success: bool = False,
) -> bool:
    if sys.platform != "win32":
        return False
    try:
        from giftcardInvoiceLink.excel_link_sync import find_workbook_and_application
    except Exception:
        return False

    for path in resolve_orders_workbook_candidates(project_root):
        wb = None
        excel = None
        try:
            wb, excel = find_workbook_and_application(str(path))
        except Exception:
            wb = None
        if wb is None:
            continue
        try:
            ws = wb.Worksheets("Orders")
            headers = _header_map(ws)
            row = _find_record_row(ws, headers, record)
            if row is None:
                continue
            old_events = None
            old_alerts = None
            if excel is not None:
                try:
                    old_events = excel.EnableEvents
                    excel.EnableEvents = False
                except Exception:
                    old_events = None
                try:
                    old_alerts = excel.DisplayAlerts
                    excel.DisplayAlerts = False
                except Exception:
                    old_alerts = None
            ws.Rows(row).Delete()
            if excel is not None:
                if old_events is not None:
                    try:
                        excel.EnableEvents = old_events
                    except Exception:
                        pass
                if old_alerts is not None:
                    try:
                        excel.DisplayAlerts = old_alerts
                    except Exception:
                        pass
                if show_success:
                    _run_success_rainbow(excel, str(wb.Name))
            return True
        except Exception:
            if excel is not None:
                try:
                    excel.EnableEvents = True
                except Exception:
                    pass
                try:
                    excel.DisplayAlerts = True
                except Exception:
                    pass
            continue
    return False


def set_record_excel_flagged(
    project_root: Path,
    *,
    record_id: str,
    flagged: bool = True,
    source_uri: str = "",
    order_number: str = "",
    sync_open_excel: bool = True,
) -> dict:
    result_path = results_json_path(project_root)
    pod_path = proof_of_delivery_json_path(project_root)
    file_records: list[tuple[Path, list[dict]]] = [
        (result_path, load_json_records(result_path)),
        (pod_path, load_json_records(pod_path)),
    ]
    matched_path: Path | None = None
    matched_records: list[dict] | None = None
    idx: int | None = None
    for path, records in file_records:
        if not records:
            continue
        changed_ids = ensure_record_ids(records)
        idx = find_record_index_for_excel_context(
            records,
            record_id=record_id,
            source_uri=source_uri,
            order_number=order_number,
            allow_order_fallback=False,
        )
        if idx is not None:
            matched_path = path
            matched_records = records
            if changed_ids:
                save_json_records(path, records)
            break
        if changed_ids:
            save_json_records(path, records)
    if matched_path is None:
        wanted_order = str(order_number or "").strip()
        order_matches: list[tuple[Path, list[dict], int]] = []
        if wanted_order:
            for path, records in file_records:
                for record_idx, record in enumerate(records):
                    if str(record.get("order_number") or "").strip() == wanted_order:
                        order_matches.append((path, records, record_idx))
        if len(order_matches) == 1:
            matched_path, matched_records, idx = order_matches[0]
    if matched_path is None or matched_records is None or idx is None:
        raise ValueError("Could not match this item to a row in results.json or proof_of_delivery.json.")

    record = matched_records[idx]
    if flagged:
        record["excel_flagged"] = True
        record.pop("excel_active", None)
    else:
        record.pop("excel_flagged", None)
        record.pop("excel_active", None)
    ensure_record_ids(matched_records)
    save_json_records(matched_path, matched_records)

    changed_ids = {str(record.get(RECORD_ID_FIELD) or "")}
    if sync_open_excel:
        try_sync_open_excel_records(
            project_root,
            records=matched_records,
            record_ids=changed_ids,
            fields={"excel_flagged", "excel_active"},
            show_success=True,
        )

    return {
        "changed_records": 1,
        "record": record,
        "records": matched_records,
        "changed_files": [str(matched_path)],
    }
