from __future__ import annotations

import hashlib
import json
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname

from shared.excel_user_edits import (
    coerce_user_edit_value,
    modified_key,
    original_value_for_user_clear,
    record_excel_user_edit,
    record_identity,
)
from shared.order_store import (
    RECORD_ID_FIELD,
    find_record_index_by_record_id,
    is_record_excel_flagged,
    load_results_records,
    proof_of_delivery_json_path,
    save_results_records,
    set_record_excel_flagged,
    try_remove_open_excel_record,
    try_sync_open_excel_records,
)

FIX_FLAGGED_REQUESTS_JSON_NAME = "fix_flagged_requests.json"
LEGACY_REQUESTS_JSON_NAME = "help_ai_requests.json"
REQUEST_TYPE_INVOICE_TOTALS = "invoice_totals_ambiguous"
REQUEST_TYPE_EXCEL_FLAGGED = "excel_flagged"
REQUEST_TYPE_POD_REVIEW = "pod_capture_needs_review"
ACTIVE_STATUSES = {"active", ""}
MANUAL_EDIT_FIELDS = (
    "company",
    "order_number",
    "purchase_datetime",
    "subtotal_amount",
    "total_amount_paid",
    "tax_paid",
    "gift_card_amount",
    "email_category",
)


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def results_json_path(project_root: Path) -> Path:
    return project_root / "email_contents" / "json" / "results.json"


def fix_flagged_requests_path(project_root: Path) -> Path:
    return project_root / "email_contents" / "json" / FIX_FLAGGED_REQUESTS_JSON_NAME


def legacy_requests_path(project_root: Path) -> Path:
    return project_root / "email_contents" / "json" / LEGACY_REQUESTS_JSON_NAME


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


def load_results(project_root: Path) -> list[dict]:
    return load_results_records(project_root)


def load_review_records(project_root: Path) -> list[dict]:
    return load_results_records(project_root) + load_json_records(proof_of_delivery_json_path(project_root))


def save_results(project_root: Path, records: list[dict]) -> None:
    save_results_records(project_root, records)


def load_requests(project_root: Path) -> list[dict]:
    path = fix_flagged_requests_path(project_root)
    if path.is_file():
        return load_json_records(path)
    return load_json_records(legacy_requests_path(project_root))


def save_requests(project_root: Path, requests: list[dict]) -> None:
    save_json_records(fix_flagged_requests_path(project_root), requests)


def active_requests(requests: list[dict]) -> list[dict]:
    return [
        req
        for req in requests
        if str(req.get("status") or "active").strip().lower() in ACTIVE_STATUSES
    ]


def stable_request_id(record: dict, request_type: str) -> str:
    raw = f"{request_type}|{record_identity(record)}"
    return hashlib.sha1(raw.encode("utf-8", errors="ignore")).hexdigest()[:16]


def _numeric_amount(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    text = str(value).strip()
    if not text:
        return None
    normalized = (
        text.replace("$", "")
        .replace(",", "")
        .replace(" ", "")
        .replace("\u00a0", "")
    )
    if normalized.startswith("(") and normalized.endswith(")"):
        normalized = "-" + normalized[1:-1]
    try:
        return round(float(normalized), 2)
    except ValueError:
        return None


def _boolish(value: object) -> bool | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "y", "paid"}:
        return True
    if text in {"0", "false", "no", "n", "not paid", "none"}:
        return False
    return None


def _money_text(value: object) -> str:
    amount = _numeric_amount(value)
    return "" if amount is None else f"{amount:.2f}"


def _field_snapshot(record: dict) -> dict[str, object]:
    return {field: record.get(field) for field in MANUAL_EDIT_FIELDS}


def audit_invoice_totals(record: dict) -> dict | None:
    category = str(record.get("email_category") or "").strip()
    if category not in {"Invoice", "Gift Card"}:
        return None

    subtotal = _numeric_amount(record.get("subtotal_amount"))
    total = _numeric_amount(record.get("total_amount_paid"))
    tax = _numeric_amount(record.get("tax_paid"))
    gift = _numeric_amount(record.get("gift_card_amount"))

    reasons: list[str] = []
    model_reason = str(record.get("invoice_total_review_reason") or "").strip()
    if bool(record.get("invoice_total_needs_review")):
        reasons.append(model_reason or "AI marked the invoice totals as ambiguous.")

    if subtotal is not None and total is not None:
        expected = subtotal + (tax or 0.0) - (gift or 0.0)
        if abs(expected - total) > 0.03:
            reasons.append(
                "Subtotal, tax, gift card, and total do not reconcile "
                f"({subtotal:.2f} + {(tax or 0.0):.2f} - {(gift or 0.0):.2f} = {expected:.2f}, "
                f"but total is {total:.2f})."
            )

    if not reasons:
        return None

    return {
        "type": REQUEST_TYPE_INVOICE_TOTALS,
        "title": "Invoice totals need review",
        "reason": " ".join(reasons),
        "fields": [
            "subtotal_amount",
            "tax_paid",
            "gift_card_amount",
            "total_amount_paid",
        ],
    }


def flagged_review_request(record: dict) -> dict:
    return {
        "type": REQUEST_TYPE_EXCEL_FLAGGED,
        "title": "Excel row flagged for review",
        "reason": "This row was flagged in Excel for Fix Flagged review.",
        "fields": list(MANUAL_EDIT_FIELDS),
    }


def pod_review_request(record: dict) -> dict | None:
    if not bool(record.get("pod_review_required")):
        return None
    if str(record.get("pod_review_status") or "active").strip().lower() not in ACTIVE_STATUSES:
        return None
    reason = str(record.get("pod_review_reason") or "").strip()
    tracking_number = str(record.get("pod_tracking_number") or "").strip()
    if not tracking_number:
        raw_numbers = record.get("tracking_numbers")
        if isinstance(raw_numbers, list) and raw_numbers:
            tracking_number = str(raw_numbers[0] or "").strip()
    if not reason:
        reason = (
            "The POD page was captured, but the tracking details were not visible enough "
            "for automatic approval."
        )
    if tracking_number:
        reason = f"Tracking {tracking_number}: {reason}"
    return {
        "type": REQUEST_TYPE_POD_REVIEW,
        "title": "POD capture needs review",
        "reason": reason,
        "fields": list(MANUAL_EDIT_FIELDS),
    }


def sync_fix_flagged_requests_for_records(project_root: Path, records: list[dict]) -> list[dict]:
    requests = load_requests(project_root)
    by_id = {str(req.get("id") or ""): req for req in requests if req.get("id")}
    current_invoice_request_ids: set[str] = set()
    active_invoice_issue_ids: set[str] = set()
    current_flagged_request_ids: set[str] = set()
    active_flagged_issue_ids: set[str] = set()
    current_pod_request_ids: set[str] = set()
    active_pod_issue_ids: set[str] = set()
    changed = False
    now = _utc_now()

    for record in records:
        if not isinstance(record, dict):
            continue
        invoice_request_id = stable_request_id(record, REQUEST_TYPE_INVOICE_TOTALS)
        current_invoice_request_ids.add(invoice_request_id)
        flagged_request_id = stable_request_id(record, REQUEST_TYPE_EXCEL_FLAGGED)
        current_flagged_request_ids.add(flagged_request_id)
        pod_request_id = stable_request_id(record, REQUEST_TYPE_POD_REVIEW)
        current_pod_request_ids.add(pod_request_id)
        issues: list[dict] = []
        pod_issue = pod_review_request(record)
        if pod_issue:
            issues.append(pod_issue)
        elif is_record_excel_flagged(record):
            issues.append(flagged_review_request(record))
        issue = audit_invoice_totals(record)
        if issue:
            issues.append(issue)
        for issue in issues:
            request_id = stable_request_id(record, str(issue["type"]))
            if issue["type"] == REQUEST_TYPE_INVOICE_TOTALS:
                active_invoice_issue_ids.add(request_id)
            if issue["type"] == REQUEST_TYPE_EXCEL_FLAGGED:
                active_flagged_issue_ids.add(request_id)
            if issue["type"] == REQUEST_TYPE_POD_REVIEW:
                active_pod_issue_ids.add(request_id)
            existing = by_id.get(request_id)
            if (
                issue["type"] != REQUEST_TYPE_EXCEL_FLAGGED
                and existing is not None
                and str(existing.get("status") or "active").lower() not in ACTIVE_STATUSES
            ):
                continue
            payload = {
                "id": request_id,
                "status": "active",
                "type": issue["type"],
                "title": issue["title"],
                "reason": issue["reason"],
                "fields": issue["fields"],
                "record_identity": record_identity(record),
                "order_number": str(record.get("order_number") or "").strip(),
                "company": str(record.get("company") or "").strip(),
                "source_file_link": str(record.get("source_file_link") or "").strip(),
                "source_file": str(record.get("source_file") or "").strip(),
                "subject": str(record.get("subject") or "").strip(),
                "field_values": _field_snapshot(record),
                "updated_at": now,
            }
            if existing is None:
                payload["created_at"] = now
                requests.append(payload)
                by_id[request_id] = payload
                changed = True
            else:
                created = existing.get("created_at") or now
                if any(existing.get(k) != v for k, v in payload.items()):
                    existing.update(payload)
                    existing["created_at"] = created
                    changed = True

    resolved_invoice_ids = current_invoice_request_ids - active_invoice_issue_ids
    resolved_flagged_ids = current_flagged_request_ids - active_flagged_issue_ids
    resolved_pod_ids = current_pod_request_ids - active_pod_issue_ids
    for req in requests:
        request_id = str(req.get("id") or "")
        request_type = str(req.get("type") or "")
        if request_type == REQUEST_TYPE_INVOICE_TOTALS:
            should_resolve = request_id in resolved_invoice_ids
        elif request_type == REQUEST_TYPE_EXCEL_FLAGGED:
            should_resolve = request_id in resolved_flagged_ids
        elif request_type == REQUEST_TYPE_POD_REVIEW:
            should_resolve = request_id in resolved_pod_ids
        else:
            should_resolve = False
        if not should_resolve:
            continue
        if str(req.get("status") or "active").strip().lower() not in ACTIVE_STATUSES:
            continue
        req["status"] = "resolved"
        req["resolved_at"] = now
        changed = True

    if changed:
        save_requests(project_root, requests)
    return requests


def sync_fix_flagged_requests(project_root: Path) -> list[dict]:
    return sync_fix_flagged_requests_for_records(project_root, load_review_records(project_root))


def find_record_index(records: list[dict], wanted_identity: str) -> int | None:
    for idx, record in enumerate(records):
        if record_identity(record) == wanted_identity:
            return idx
    return None


def _coerce_manual_value(field: str, value: object) -> object:
    text = "" if value is None else str(value).strip()
    if not text:
        return None
    if field == "email_category":
        return coerce_user_edit_value(field, text)
    if field in {"company", "order_number"}:
        return text
    if field == "purchase_datetime":
        return coerce_user_edit_value(field, text)
    if field in {"subtotal_amount", "total_amount_paid", "tax_paid", "gift_card_amount"}:
        amount = _numeric_amount(text)
        if amount is None:
            raise ValueError(f"{field} must be a number-like value")
        return amount
    raise ValueError(f"Unsupported Fix Flagged field: {field}")


def _blank_manual_value(value: object) -> bool:
    return str(value or "").strip() == ""


def _record_ids_for_reset_sync(
    records: list[dict],
    record: dict,
    fields: set[str],
) -> set[str]:
    if fields.intersection({"company", "purchase_datetime"}):
        order_number = str(record.get("order_number") or "").strip()
        if order_number:
            return {
                str(item.get(RECORD_ID_FIELD) or "")
                for item in records
                if str(item.get("order_number") or "").strip() == order_number
            }
    return {str(record.get(RECORD_ID_FIELD) or "")}


def save_record_updates(
    project_root: Path,
    *,
    record_id: str,
    updates: dict[str, object],
    request_id: str | None = None,
    status: str | None = None,
) -> dict:
    records = load_results(project_root)
    records_path = results_json_path(project_root)
    is_pod_record = False
    idx = find_record_index(records, record_id)
    if idx is None:
        records_path = proof_of_delivery_json_path(project_root)
        records = load_json_records(records_path)
        idx = find_record_index(records, record_id)
        is_pod_record = idx is not None
    if idx is None:
        raise ValueError("Could not match this item to a row in results.json or proof_of_delivery.json.")
    should_clear_flag = False
    should_clear_pod_review = False
    if request_id and status == "resolved":
        for req in load_requests(project_root):
            if str(req.get("id") or "") == request_id:
                request_type = str(req.get("type") or "")
                should_clear_flag = request_type in {REQUEST_TYPE_EXCEL_FLAGGED, REQUEST_TYPE_POD_REVIEW}
                should_clear_pod_review = request_type == REQUEST_TYPE_POD_REVIEW
                break

    submitted_updates = {
        field: value
        for field, value in updates.items()
        if field in MANUAL_EDIT_FIELDS
    }
    if not submitted_updates:
        if should_clear_flag:
            if should_clear_pod_review:
                records[idx]["pod_review_required"] = False
                records[idx]["pod_review_status"] = "resolved"
                records[idx]["pod_review_resolved_at"] = _utc_now()
                save_json_records(proof_of_delivery_json_path(project_root), records)
            result = set_record_excel_flagged(
                project_root,
                record_id=str(records[idx].get(RECORD_ID_FIELD) or ""),
                flagged=False,
                sync_open_excel=True,
            )
            update_request_status(project_root, request_id=request_id, status=status)
            sync_fix_flagged_requests_for_records(project_root, result["records"])
            return {"changed_records": result.get("changed_records", 1), "record": result["record"]}
        if request_id and status:
            update_request_status(project_root, request_id=request_id, status=status)
        sync_fix_flagged_requests_for_records(project_root, records)
        return {"changed_records": 0, "record": records[idx]}

    store_record_id = str(records[idx].get(RECORD_ID_FIELD) or "").strip()
    if not store_record_id:
        save_results(project_root, records)
        records = load_results(project_root)
        idx = find_record_index(records, record_id)
        if idx is None:
            raise ValueError("Could not match this item to a row in results.json.")
        store_record_id = str(records[idx].get(RECORD_ID_FIELD) or "").strip()
    original_record = records[idx]
    order_number = str(original_record.get("order_number") or "").strip()
    source_uri = str(original_record.get("source_file_link") or original_record.get("source_file") or "").strip()
    changed_fields: set[str] = set()
    if is_pod_record:
        now = _utc_now()
        for field, value in submitted_updates.items():
            if _blank_manual_value(value):
                raw_value = original_value_for_user_clear(records[idx], field)
                records[idx].pop(modified_key(field), None)
            else:
                raw_value = _coerce_manual_value(field, value)
                records[idx][modified_key(field)] = True
            records[idx][field] = raw_value
            changed_fields.add(field)
        records[idx]["user_modified"] = True
        records[idx]["user_modified_at"] = now
        if request_id and status:
            if should_clear_pod_review:
                records[idx]["pod_review_required"] = False
                records[idx]["pod_review_status"] = "resolved"
                records[idx]["pod_review_resolved_at"] = now
            if should_clear_flag:
                records[idx].pop("excel_flagged", None)
                records[idx].pop("excel_active", None)
                changed_fields.update({"excel_flagged", "excel_active"})
            update_request_status(project_root, request_id=request_id, status=status)
        save_json_records(records_path, records)
        try_sync_open_excel_records(
            project_root,
            records=records,
            record_ids={str(records[idx].get(RECORD_ID_FIELD) or "")},
            fields=changed_fields,
            show_success=True,
        )
        sync_fix_flagged_requests_for_records(project_root, records)
        return {"changed_records": 1, "record": records[idx], "records": records}

    for field, value in submitted_updates.items():
        raw_value = "" if _blank_manual_value(value) else _coerce_manual_value(field, value)
        record_excel_user_edit(
            project_root,
            field=field,
            raw_value=raw_value,
            order_number=order_number,
            source_uri=source_uri,
            record_id=store_record_id,
        )
        changed_fields.add(field)

    records = load_results(project_root)
    idx = find_record_index_by_record_id(records, store_record_id)
    if idx is None:
        idx = find_record_index(records, record_id)
    if idx is None:
        raise ValueError("Could not match this item to a row in results.json.")
    changed_ids = _record_ids_for_reset_sync(records, records[idx], changed_fields)
    try_sync_open_excel_records(
        project_root,
        records=records,
        record_ids=changed_ids,
        fields=changed_fields,
        show_success=True,
    )
    changed_records = len([record_id for record_id in changed_ids if record_id])
    result = {"changed_records": changed_records, "record": records[idx], "records": records}

    if request_id and status:
        if should_clear_flag:
            if should_clear_pod_review:
                records[idx]["pod_review_required"] = False
                records[idx]["pod_review_status"] = "resolved"
                records[idx]["pod_review_resolved_at"] = _utc_now()
                save_json_records(proof_of_delivery_json_path(project_root), records)
            result = set_record_excel_flagged(
                project_root,
                record_id=str(records[idx].get(RECORD_ID_FIELD) or ""),
                flagged=False,
                sync_open_excel=True,
            )
            records = result["records"]
        update_request_status(project_root, request_id=request_id, status=status)
    sync_fix_flagged_requests_for_records(project_root, result["records"])
    return {"changed_records": changed_records, "record": result["record"]}


def mark_record_excel_flagged(project_root: Path, *, record_id: str, flagged: bool = True) -> dict:
    records = load_results(project_root)
    idx = find_record_index(records, record_id)
    if idx is None:
        raise ValueError("Could not match this item to a row in results.json.")
    store_record_id = str(records[idx].get(RECORD_ID_FIELD) or "").strip()
    if not store_record_id:
        save_results(project_root, records)
        records = load_results(project_root)
        idx = find_record_index(records, record_id)
        if idx is None:
            raise ValueError("Could not match this item to a row in results.json.")
        store_record_id = str(records[idx].get(RECORD_ID_FIELD) or "").strip()
    result = set_record_excel_flagged(
        project_root,
        record_id=store_record_id,
        flagged=flagged,
        sync_open_excel=True,
    )
    sync_fix_flagged_requests_for_records(project_root, result["records"])
    return result


def update_request_status(project_root: Path, *, request_id: str, status: str) -> None:
    requests = load_requests(project_root)
    now = _utc_now()
    changed = False
    for req in requests:
        if str(req.get("id") or "") == request_id:
            req["status"] = status
            req["resolved_at" if status == "resolved" else "updated_at"] = now
            changed = True
            break
    if changed:
        save_requests(project_root, requests)


def skip_record_eternally(
    project_root: Path,
    *,
    record_id: str,
    request_id: str | None = None,
) -> dict[str, object]:
    file_records: list[tuple[Path, list[dict]]] = [
        (results_json_path(project_root), load_results(project_root)),
        (proof_of_delivery_json_path(project_root), load_json_records(proof_of_delivery_json_path(project_root))),
    ]
    matched_path: Path | None = None
    matched_records: list[dict] | None = None
    matched_record: dict | None = None
    matched_idx: int | None = None

    for path, records in file_records:
        idx = find_record_index(records, record_id)
        if idx is None:
            continue
        matched_path = path
        matched_records = records
        matched_record = dict(records[idx])
        matched_idx = idx
        break

    if matched_path is None or matched_records is None or matched_record is None or matched_idx is None:
        raise ValueError("Could not match this item to a row in results.json or proof_of_delivery.json.")

    removed_excel_row = try_remove_open_excel_record(
        project_root,
        record=matched_record,
        show_success=True,
    )
    is_pod_review_record = (
        matched_path == proof_of_delivery_json_path(project_root)
        and bool(matched_record.get("pod_review_required"))
    )
    if is_pod_review_record:
        now = _utc_now()
        matched_records[matched_idx]["pod_review_required"] = False
        matched_records[matched_idx]["pod_review_status"] = "resolved"
        matched_records[matched_idx]["pod_review_resolved_at"] = now
        matched_records[matched_idx]["pod_review_skipped_eternally"] = True
        matched_records[matched_idx].pop("excel_flagged", None)
        matched_records[matched_idx].pop("excel_active", None)
        save_json_records(matched_path, matched_records)
    else:
        del matched_records[matched_idx]
        if matched_path == results_json_path(project_root):
            save_results(project_root, matched_records)
        else:
            save_json_records(matched_path, matched_records)

    requests = load_requests(project_root)
    removed_requests = 0
    if request_id:
        before = len(requests)
        requests = [req for req in requests if str(req.get("id") or "") != request_id]
        removed_requests += before - len(requests)
    record_ident = record_identity(matched_record)
    before = len(requests)
    requests = [
        req
        for req in requests
        if not (
            str(req.get("type") or "") == REQUEST_TYPE_POD_REVIEW
            and str(req.get("record_identity") or "") == record_ident
        )
    ]
    removed_requests += before - len(requests)
    save_requests(project_root, requests)

    sync_fix_flagged_requests_for_records(project_root, load_review_records(project_root))
    return {
        "removed_records": 0 if is_pod_review_record else 1,
        "resolved_records": 1 if is_pod_review_record else 0,
        "removed_requests": removed_requests,
        "removed_excel_row": removed_excel_row,
        "changed_files": [str(matched_path)],
    }


def _unique_existing_backup_path(path: Path, marker: str) -> Path:
    candidate = path.with_name(f"{path.stem}{marker}{path.suffix}")
    if not candidate.exists():
        return candidate
    idx = 2
    while True:
        candidate = path.with_name(f"{path.stem}{marker} ({idx}){path.suffix}")
        if not candidate.exists():
            return candidate
        idx += 1


def resolve_pod_review_with_manual_scan(
    project_root: Path,
    *,
    record_id: str,
    scanned_pdf_path: str | Path,
    request_id: str | None = None,
) -> dict[str, object]:
    """Promote one manual rescan PDF into the matched POD review row."""
    records_path = proof_of_delivery_json_path(project_root)
    records = load_json_records(records_path)
    idx = find_record_index(records, record_id)
    if idx is None:
        raise ValueError("Could not match this POD review item to proof_of_delivery.json.")

    record = records[idx]
    scan_path = Path(scanned_pdf_path).expanduser().resolve()
    if not scan_path.is_file():
        raise FileNotFoundError(f"Manual rescan PDF was not found: {scan_path}")

    tracking_number = str(record.get("pod_tracking_number") or "").strip()
    if not tracking_number:
        raw_numbers = record.get("tracking_numbers")
        if isinstance(raw_numbers, list) and raw_numbers:
            tracking_number = str(raw_numbers[0] or "").strip()
    carrier = str(record.get("pod_carrier") or record.get("carrier") or "").strip()
    expected_path = None
    if tracking_number:
        try:
            from proofOfDelivery.pod_data import expected_pod_pdf_path

            expected_path = expected_pod_pdf_path(
                project_root,
                record.get("company"),
                record.get("pod_source_purchase_datetime") or record.get("purchase_datetime"),
                tracking_number,
                carrier,
            ).resolve()
        except Exception:
            expected_path = None
    final_path = expected_path or scan_path
    final_path.parent.mkdir(parents=True, exist_ok=True)

    if scan_path != final_path:
        if final_path.exists():
            backup_path = _unique_existing_backup_path(final_path, "_manual_rescan_replaced")
            shutil.move(str(final_path), str(backup_path))
        shutil.move(str(scan_path), str(final_path))

    now = _utc_now()
    record["source_file"] = str(final_path)
    try:
        record["source_file_link"] = final_path.as_uri()
    except ValueError:
        record["source_file_link"] = str(final_path)
    record["pod_generated_file_name"] = final_path.name
    record["pod_expected_file_name"] = final_path.name if expected_path is not None else record.get("pod_expected_file_name")
    record["pod_capture_review_pdf"] = str(final_path)
    record["latest_tracking_info_visible"] = True
    record["pod_review_required"] = False
    record["pod_review_status"] = "resolved"
    record["pod_review_resolved_at"] = now
    record["pod_review_manual_rescan_accepted_at"] = now
    record.pop("excel_flagged", None)
    record.pop("excel_active", None)
    save_json_records(records_path, records)

    if request_id:
        update_request_status(project_root, request_id=request_id, status="resolved")

    record_key = str(record.get(RECORD_ID_FIELD) or "")
    if record_key:
        try_sync_open_excel_records(
            project_root,
            records=records,
            record_ids={record_key},
            fields={
                "source_file",
                "source_file_link",
                "pod_generated_file_name",
                "pod_expected_file_name",
                "pod_capture_review_pdf",
                "excel_flagged",
                "excel_active",
            },
            show_success=True,
        )
    sync_fix_flagged_requests_for_records(project_root, load_review_records(project_root))
    return {
        "record": record,
        "final_pdf": str(final_path),
        "changed_files": [str(records_path), str(final_path)],
    }


def _path_from_file_uri(value: object) -> Path | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    try:
        parsed = urlparse(raw)
    except Exception:
        return None
    if parsed.scheme != "file":
        return None
    try:
        local_path = url2pathname(unquote(parsed.path))
    except Exception:
        return None
    try:
        return Path(local_path).expanduser().resolve()
    except OSError:
        return None


def html_path_for_record(project_root: Path, record: dict) -> Path | None:
    source_path = _path_from_file_uri(record.get("source_file_link"))
    if source_path is None and record.get("source_file"):
        try:
            source_path = Path(str(record.get("source_file"))).expanduser().resolve()
        except OSError:
            source_path = None
    if source_path is not None and source_path.suffix.lower() in {".html", ".htm"} and source_path.is_file():
        return source_path
    if source_path is not None:
        candidate = project_root / "email_contents" / "html" / f"{source_path.stem}.html"
        if candidate.is_file():
            return candidate.resolve()
    return None


def reprocess_record_with_ai(
    project_root: Path,
    *,
    record_id: str,
    keep_fields: set[str] | None = None,
) -> dict:
    records = load_results(project_root)
    idx = find_record_index(records, record_id)
    if idx is None:
        raise ValueError("Could not match this item to a row in results.json.")
    old = records[idx]
    html_path = html_path_for_record(project_root, old)
    if html_path is None:
        raise FileNotFoundError("Archived HTML for this order was not found.")

    from grabbingImportantEmailContent.grabbingImportantEmailContent import (
        apply_order_company_consensus_and_sync,
        process_file,
    )

    new_record = process_file(
        html_path,
        str(old.get("subject") or "") or None,
        str(old.get("sender_name") or "") or None,
        str(old.get("email") or "") or None,
        str(old.get("email_datetime") or "") or None,
        str(old.get("email_datetime_source") or "") or None,
    )
    new_record.pop("_timings", None)
    for key in (
        RECORD_ID_FIELD,
        "source_file",
        "source_file_link",
        "content_hash",
        "duplicate_on_last_run",
        "accounting",
        "excel_flagged",
        "excel_active",
    ):
        if key in old:
            new_record[key] = old[key]
    for field in keep_fields or set():
        if field in old:
            new_record[field] = old[field]
        marker = modified_key(field)
        if marker in old:
            new_record[marker] = old[marker]
        if field == "purchase_datetime":
            for meta_field in ("purchase_datetime_source", "purchase_datetime_confidence"):
                if meta_field in old:
                    new_record[meta_field] = old[meta_field]
    if any(str(key).startswith("modified_") and bool(value) for key, value in new_record.items()):
        if "user_modified" in old:
            new_record["user_modified"] = old["user_modified"]
        if "user_modified_at" in old:
            new_record["user_modified_at"] = old["user_modified_at"]
    new_record["reprocessed_using_ai"] = True
    new_record["reprocessed_at"] = _utc_now()
    records[idx] = new_record
    apply_order_company_consensus_and_sync(records, project_root)
    save_results(project_root, records)
    sync_fix_flagged_requests_for_records(project_root, records)
    return records[idx]


def preview_reprocess_record_with_ai(
    project_root: Path,
    *,
    record_id: str,
) -> dict:
    records = load_results(project_root)
    idx = find_record_index(records, record_id)
    if idx is None:
        raise ValueError("Could not match this item to a row in results.json.")
    old = records[idx]
    html_path = html_path_for_record(project_root, old)
    if html_path is None:
        raise FileNotFoundError("Archived HTML for this order was not found.")

    from grabbingImportantEmailContent.grabbingImportantEmailContent import process_file

    new_record = process_file(
        html_path,
        str(old.get("subject") or "") or None,
        str(old.get("sender_name") or "") or None,
        str(old.get("email") or "") or None,
        str(old.get("email_datetime") or "") or None,
        str(old.get("email_datetime_source") or "") or None,
    )
    new_record.pop("_timings", None)
    return new_record
