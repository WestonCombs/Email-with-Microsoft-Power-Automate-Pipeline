from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname

from shared.excel_user_edits import (
    PURCHASE_DATETIME_USER_EDIT_CONFIDENCE,
    PURCHASE_DATETIME_USER_EDIT_SOURCE,
    coerce_user_edit_value,
    modified_key,
    record_identity,
)

HELP_AI_REQUESTS_JSON_NAME = "help_ai_requests.json"
REQUEST_TYPE_INVOICE_TOTALS = "invoice_totals_ambiguous"
ACTIVE_STATUSES = {"active", ""}
MANUAL_EDIT_FIELDS = (
    "company",
    "order_number",
    "purchase_datetime",
    "subtotal_amount",
    "total_amount_paid",
    "tax_paid",
    "gift_card_amount",
    "tax_was_paid",
    "email_category",
)


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def results_json_path(project_root: Path) -> Path:
    return project_root / "email_contents" / "json" / "results.json"


def help_ai_requests_path(project_root: Path) -> Path:
    return project_root / "email_contents" / "json" / HELP_AI_REQUESTS_JSON_NAME


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
    return load_json_records(results_json_path(project_root))


def save_results(project_root: Path, records: list[dict]) -> None:
    save_json_records(results_json_path(project_root), records)


def load_requests(project_root: Path) -> list[dict]:
    return load_json_records(help_ai_requests_path(project_root))


def save_requests(project_root: Path, requests: list[dict]) -> None:
    save_json_records(help_ai_requests_path(project_root), requests)


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
    tax_was_paid = _boolish(record.get("tax_was_paid"))

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

    if tax_was_paid is True and tax is None:
        reasons.append("AI thinks tax was paid but did not extract a tax amount.")
    if tax_was_paid is False and tax is not None and tax > 0:
        reasons.append("AI marked tax as not paid, but a positive tax amount is present.")

    if not reasons:
        return None

    return {
        "type": REQUEST_TYPE_INVOICE_TOTALS,
        "title": "Invoice totals need review",
        "reason": " ".join(reasons),
        "fields": [
            "subtotal_amount",
            "tax_paid",
            "tax_was_paid",
            "gift_card_amount",
            "total_amount_paid",
        ],
    }


def sync_help_ai_requests_for_records(project_root: Path, records: list[dict]) -> list[dict]:
    requests = load_requests(project_root)
    by_id = {str(req.get("id") or ""): req for req in requests if req.get("id")}
    changed = False
    now = _utc_now()

    for record in records:
        if not isinstance(record, dict):
            continue
        issue = audit_invoice_totals(record)
        if not issue:
            continue
        request_id = stable_request_id(record, str(issue["type"]))
        existing = by_id.get(request_id)
        if existing is not None and str(existing.get("status") or "active").lower() not in ACTIVE_STATUSES:
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

    if changed:
        save_requests(project_root, requests)
    return requests


def sync_help_ai_requests(project_root: Path) -> list[dict]:
    return sync_help_ai_requests_for_records(project_root, load_results(project_root))


def find_record_index(records: list[dict], wanted_identity: str) -> int | None:
    for idx, record in enumerate(records):
        if record_identity(record) == wanted_identity:
            return idx
    return None


def _coerce_manual_value(field: str, value: object) -> object:
    text = "" if value is None else str(value).strip()
    if not text:
        return None
    if field in {"company", "order_number", "email_category"}:
        return text
    if field == "purchase_datetime":
        return coerce_user_edit_value(field, text)
    if field == "tax_was_paid":
        parsed = _boolish(text)
        if parsed is None:
            raise ValueError("tax_was_paid must be yes/no, true/false, or blank")
        return parsed
    if field in {"subtotal_amount", "total_amount_paid", "tax_paid", "gift_card_amount"}:
        amount = _numeric_amount(text)
        if amount is None:
            raise ValueError(f"{field} must be a number-like value")
        return amount
    raise ValueError(f"Unsupported Help AI field: {field}")


def save_record_updates(
    project_root: Path,
    *,
    record_id: str,
    updates: dict[str, object],
    request_id: str | None = None,
    status: str | None = None,
) -> dict:
    records = load_results(project_root)
    idx = find_record_index(records, record_id)
    if idx is None:
        raise ValueError("Could not match this item to a row in results.json.")

    cleaned_updates = {
        field: _coerce_manual_value(field, value)
        for field, value in updates.items()
        if field in MANUAL_EDIT_FIELDS
    }
    if not cleaned_updates:
        if request_id and status:
            update_request_status(project_root, request_id=request_id, status=status)
        sync_help_ai_requests_for_records(project_root, records)
        return {"changed_records": 0, "record": records[idx]}

    now = _utc_now()
    target = records[idx]
    order_for_date = str(target.get("order_number") or "").strip()
    date_fanout = "purchase_datetime" in cleaned_updates and order_for_date

    changed = 0
    for record in records:
        same_target = record_identity(record) == record_id
        same_order_for_date = (
            date_fanout
            and str(record.get("order_number") or "").strip() == order_for_date
        )
        if not same_target and not same_order_for_date:
            continue
        fields_to_apply = cleaned_updates
        if same_order_for_date and not same_target:
            fields_to_apply = {"purchase_datetime": cleaned_updates["purchase_datetime"]}
        for field, value in fields_to_apply.items():
            record[field] = value
            record[modified_key(field)] = True
            if field == "purchase_datetime":
                record["purchase_datetime_source"] = PURCHASE_DATETIME_USER_EDIT_SOURCE
                record["purchase_datetime_confidence"] = PURCHASE_DATETIME_USER_EDIT_CONFIDENCE
        record["user_modified"] = True
        record["user_modified_at"] = now
        changed += 1

    save_results(project_root, records)

    if request_id and status:
        update_request_status(project_root, request_id=request_id, status=status)
    sync_help_ai_requests_for_records(project_root, records)
    return {"changed_records": changed, "record": records[idx]}


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


def reprocess_record_with_ai(project_root: Path, *, record_id: str) -> dict:
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
        "source_file",
        "source_file_link",
        "content_hash",
        "duplicate_on_last_run",
        "accounting",
    ):
        if key in old:
            new_record[key] = old[key]
    new_record["reprocessed_using_ai"] = True
    new_record["reprocessed_at"] = _utc_now()
    records[idx] = new_record
    apply_order_company_consensus_and_sync(records, project_root)
    save_results(project_root, records)
    sync_help_ai_requests_for_records(project_root, records)
    return records[idx]
