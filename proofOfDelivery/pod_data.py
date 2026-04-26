from __future__ import annotations

import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname

from trackingNumbersViewer.mitm_readiness import sanitize_filename_token

POD_CATEGORY = "POD"
AUTOMATION_HUB_CATEGORY = "Automation Hub"
AUTOMATION_HUB_ORDER_LABEL = "POD Automation"
AUTOMATION_HUB_COMPANY_LABEL = "Proof of Delivery"
AUTOMATION_HUB_STATUS_LABEL = "Process Remaining PODs"
POD_HUB_MODE = "remaining_pod_hub"
PROOF_OF_DELIVERY_JSON_NAME = "proof_of_delivery.json"
_LEGACY_CATEGORY_SUFFIX_MAP = {
    "Invoice": "INVOICE",
    "Shipped": "SHIPPED",
    "Delivered": "DELIVERED",
}


def project_root_from_env() -> Path:
    from shared.project_paths import ensure_base_dir_in_environ

    return ensure_base_dir_in_environ()


def results_json_path(project_root: Path) -> Path:
    return project_root / "email_contents" / "json" / "results.json"


def proof_of_delivery_json_path(project_root: Path) -> Path:
    return project_root / "email_contents" / "json" / PROOF_OF_DELIVERY_JSON_NAME


def pdf_output_dir(project_root: Path) -> Path:
    return project_root / "email_contents" / "pdf"


def clean_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str) and value.strip().lower() == "null":
        return None
    return value


def _normalized_text(value: object) -> str:
    return " ".join(str(clean_value(value) or "").strip().split()).casefold()


def _normalize_tracking_number(value: object) -> str:
    text = str(clean_value(value) or "").strip()
    return "".join(ch for ch in text if not ch.isspace())


def _record_tracking_number(record: dict) -> str:
    pod_num = _normalize_tracking_number(record.get("pod_tracking_number"))
    if pod_num:
        return pod_num
    nums = tracking_numbers_for_record(record)
    if nums:
        return _normalize_tracking_number(nums[0])
    return ""


def _path_key(path: Path) -> str:
    try:
        resolved = path.expanduser().resolve()
    except OSError:
        resolved = path.expanduser()
    return str(resolved).replace("/", "\\").casefold()


def _safe_path(value: object) -> Path | None:
    raw = str(clean_value(value) or "").strip()
    if not raw:
        return None
    try:
        return Path(raw).expanduser().resolve()
    except OSError:
        return None


def _path_from_file_uri(value: object) -> Path | None:
    raw = str(clean_value(value) or "").strip()
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


def is_pod_record(record: object) -> bool:
    return isinstance(record, dict) and str(record.get("email_category") or "").strip() == POD_CATEGORY


def is_automation_hub_record(record: object) -> bool:
    return (
        isinstance(record, dict)
        and str(record.get("email_category") or "").strip() == AUTOMATION_HUB_CATEGORY
    )


def load_json_records(path: Path) -> list[dict]:
    if not path.is_file():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    return payload if isinstance(payload, list) else []


def save_json_records(path: Path, records: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8")


def load_results_records(project_root: Path) -> list[dict]:
    return [r for r in load_json_records(results_json_path(project_root)) if isinstance(r, dict)]


def load_proof_of_delivery_records(project_root: Path) -> list[dict]:
    return [r for r in load_json_records(proof_of_delivery_json_path(project_root)) if isinstance(r, dict)]


def _carrier_display_for_number(tracking_number: str) -> str:
    try:
        from trackingNumbersViewer.seventeen_track_smart import carrier_display_for_number

        return str(carrier_display_for_number(tracking_number) or "").strip() or "Unknown"
    except Exception:
        return "Unknown"


def _purchase_date_token(value: object) -> str:
    raw = str(clean_value(value) or "").strip()
    if not raw:
        return "nodate"
    if len(raw) >= 10 and raw[4:5] == "-" and raw[7:8] == "-":
        return sanitize_filename_token(raw[:10])
    return sanitize_filename_token(raw.split()[0])


def _tracking_last4_token(value: object) -> str:
    raw = str(clean_value(value) or "").strip()
    compact = "".join(ch for ch in raw if not ch.isspace())
    if len(compact) >= 4:
        return sanitize_filename_token(compact[-4:])
    if compact:
        return sanitize_filename_token(compact.zfill(4))
    return "0000"


def _order_last4_token(value: object) -> str:
    raw = str(clean_value(value) or "").strip()
    digits = re.sub(r"\D", "", raw)
    if len(digits) >= 4:
        return digits[-4:]
    if digits:
        return digits.zfill(4)
    return "0000"


def pod_pdf_basename(
    company: object,
    purchase_datetime: object,
    tracking_number: object,
    carrier_display: object,
) -> str:
    company_tok = sanitize_filename_token(clean_value(company) or "Unknown")
    date_tok = _purchase_date_token(purchase_datetime)
    tracking_last4 = _tracking_last4_token(tracking_number)
    return f"DOC {company_tok} {date_tok} TRACKING_INV_{tracking_last4}"


def legacy_pod_pdf_basename(
    company: object,
    purchase_datetime: object,
    tracking_number: object,
    carrier_display: object,
) -> str:
    company_tok = sanitize_filename_token(clean_value(company) or "Unknown")
    date_tok = _purchase_date_token(purchase_datetime)
    track_tok = sanitize_filename_token(clean_value(tracking_number) or "")
    carrier_tok = sanitize_filename_token(clean_value(carrier_display) or "Unknown")
    return f"DOC {company_tok} {date_tok} {track_tok}_FROM_{carrier_tok}"


def expected_pod_pdf_path(
    project_root: Path,
    company: object,
    purchase_datetime: object,
    tracking_number: object,
    carrier_display: object,
) -> Path:
    return pdf_output_dir(project_root) / f"{pod_pdf_basename(company, purchase_datetime, tracking_number, carrier_display)}.pdf"


def legacy_expected_pod_pdf_path(
    project_root: Path,
    company: object,
    purchase_datetime: object,
    tracking_number: object,
    carrier_display: object,
) -> Path:
    return pdf_output_dir(project_root) / f"{legacy_pod_pdf_basename(company, purchase_datetime, tracking_number, carrier_display)}.pdf"


def legacy_email_capture_pdf_basename(
    company: object,
    purchase_datetime: object,
    order_number: object,
    category: object,
) -> str:
    company_tok = sanitize_filename_token(clean_value(company) or "Unknown")
    date_tok = _purchase_date_token(purchase_datetime)
    order_last4 = _order_last4_token(order_number)
    suffix = _LEGACY_CATEGORY_SUFFIX_MAP.get(str(clean_value(category) or "").strip())
    if suffix:
        return f"DOC {company_tok} {date_tok} {suffix}_{order_last4}"
    return f"DOC {company_tok} {date_tok}_{order_last4}"


def first_existing_pdf_named(project_root: Path, basename: str) -> Path | None:
    out_dir = pdf_output_dir(project_root)
    direct = out_dir / f"{basename}.pdf"
    if direct.is_file():
        return direct
    collision_re = re.compile(rf"^{re.escape(basename)} \(\d+\)\.pdf$", re.IGNORECASE)
    try:
        for path in out_dir.glob("*.pdf"):
            if path.is_file() and collision_re.match(path.name):
                return path
    except OSError:
        return None
    return None


def _all_existing_pdf_named(project_root: Path, basename: str) -> list[Path]:
    out_dir = pdf_output_dir(project_root)
    out: list[Path] = []
    exact = out_dir / f"{basename}.pdf"
    if exact.is_file():
        out.append(exact)
    collision_re = re.compile(rf"^{re.escape(basename)} \(\d+\)\.pdf$", re.IGNORECASE)
    try:
        for path in out_dir.glob("*.pdf"):
            if path.is_file() and collision_re.match(path.name):
                out.append(path)
    except OSError:
        pass
    return out


def first_existing_legacy_email_capture_pdf_path(
    project_root: Path,
    company: object,
    purchase_datetime: object,
    order_number: object,
    category: object,
) -> Path | None:
    return first_existing_pdf_named(
        project_root,
        legacy_email_capture_pdf_basename(company, purchase_datetime, order_number, category),
    )


def first_existing_pod_pdf_path(
    project_root: Path,
    company: object,
    purchase_datetime: object,
    tracking_number: object,
    carrier_display: object,
) -> Path | None:
    for basename in (
        pod_pdf_basename(company, purchase_datetime, tracking_number, carrier_display),
        legacy_pod_pdf_basename(company, purchase_datetime, tracking_number, carrier_display),
    ):
        path = first_existing_pdf_named(project_root, basename)
        if path is not None:
            return path
    return None


def first_existing_capture_pdf_path(
    project_root: Path,
    company: object,
    purchase_datetime: object,
    tracking_number: object,
    carrier_display: object,
    order_number: object,
    category: object,
) -> Path | None:
    pod_path = first_existing_pod_pdf_path(
        project_root,
        company,
        purchase_datetime,
        tracking_number,
        carrier_display,
    )
    if pod_path is not None:
        return pod_path
    return first_existing_legacy_email_capture_pdf_path(
        project_root,
        company,
        purchase_datetime,
        order_number,
        category,
    )


def tracking_numbers_for_record(record: object) -> list[str]:
    if not isinstance(record, dict):
        return []
    raw = record.get("tracking_numbers")
    if not isinstance(raw, list):
        return []
    out: list[str] = []
    seen: set[str] = set()
    for item in raw:
        value = str(item or "").strip()
        if not value or value in seen:
            continue
        seen.add(value)
        out.append(value)
    return out


def automation_hub_record() -> dict:
    return {
        "email_category": AUTOMATION_HUB_CATEGORY,
        "order_number": AUTOMATION_HUB_ORDER_LABEL,
        "company": AUTOMATION_HUB_COMPANY_LABEL,
        "email": "Workbook-wide POD actions",
        "purchase_datetime": None,
        "source_file": None,
        "source_file_link": None,
        "total_amount_paid": None,
        "tax_paid": None,
        "tracking_numbers": [],
        "tracking_links": [],
        "tracking_numbers_link_confirmed": [],
        "pod_hub_mode": POD_HUB_MODE,
    }


def _base_record_company(record: dict) -> str:
    return (
        str(
            clean_value(record.get("company"))
            or clean_value(record.get("llm_obtained_company"))
            or ""
        ).strip()
        or "Unknown"
    )


def _pod_record_identity(record: dict) -> str:
    return str(clean_value(record.get("source_file_link")) or "").strip()


def discover_proof_of_delivery_records(project_root: Path, base_records: list[dict]) -> list[dict]:
    discovered: list[dict] = []
    seen_links: set[str] = set()
    for source_index, record in enumerate(base_records):
        if is_pod_record(record) or is_automation_hub_record(record):
            continue
        company = _base_record_company(record)
        purchase_datetime = clean_value(record.get("purchase_datetime"))
        order_number = clean_value(record.get("order_number"))
        source_category = clean_value(record.get("email_category"))
        source_email = clean_value(record.get("email"))
        for tracking_number in tracking_numbers_for_record(record):
            carrier_display = _carrier_display_for_number(tracking_number)
            pdf_path = expected_pod_pdf_path(
                project_root,
                company,
                purchase_datetime,
                tracking_number,
                carrier_display,
            )
            existing_pdf_path = first_existing_capture_pdf_path(
                project_root,
                company,
                purchase_datetime,
                tracking_number,
                carrier_display,
                order_number,
                source_category,
            )
            if existing_pdf_path is None:
                continue
            pdf_path = existing_pdf_path
            source_file_link = pdf_path.resolve().as_uri()
            if source_file_link in seen_links:
                continue
            seen_links.add(source_file_link)
            discovered.append(
                {
                    "email_category": POD_CATEGORY,
                    "order_number": order_number,
                    "purchase_datetime": None,
                    "company": company,
                    "email": None,
                    "source_file": str(pdf_path.resolve()),
                    "source_file_link": source_file_link,
                    "subject": f"Proof of delivery for {company} tracking {tracking_number}",
                    "total_amount_paid": None,
                    "tax_paid": None,
                    "tracking_numbers": [tracking_number],
                    "tracking_links": [],
                    "tracking_numbers_link_confirmed": [True],
                    "pod_tracking_number": tracking_number,
                    "pod_carrier": carrier_display,
                    "pod_expected_file_name": pdf_path.name,
                    "pod_generated_file_name": pdf_path.name,
                    "pod_source_category": source_category,
                    "pod_source_email": source_email,
                    "pod_source_purchase_datetime": purchase_datetime,
                    "pod_source_index": source_index,
                }
            )
    return discovered


def missing_proof_of_delivery_records(project_root: Path) -> list[dict]:
    base_records = load_results_records(project_root)
    desired = discover_proof_of_delivery_records(project_root, base_records)
    existing = {
        _pod_record_identity(r)
        for r in load_proof_of_delivery_records(project_root)
        if _pod_record_identity(r)
    }
    return [record for record in desired if _pod_record_identity(record) not in existing]


def sync_proof_of_delivery_records(project_root: Path) -> tuple[list[dict], bool]:
    base_records = load_results_records(project_root)
    desired = discover_proof_of_delivery_records(project_root, base_records)
    current = load_proof_of_delivery_records(project_root)
    desired_ids = {_pod_record_identity(r) for r in desired if _pod_record_identity(r)}
    for record in current:
        if not isinstance(record, dict):
            continue
        if not record.get("user_modified") and not any(
            str(k).startswith("modified_") and bool(v) for k, v in record.items()
        ):
            continue
        ident = _pod_record_identity(record)
        if ident and ident not in desired_ids:
            desired.append(record)
            desired_ids.add(ident)
    changed = json.dumps(current, sort_keys=True, ensure_ascii=False) != json.dumps(
        desired,
        sort_keys=True,
        ensure_ascii=False,
    )
    if changed:
        save_json_records(proof_of_delivery_json_path(project_root), desired)
    return desired, changed


def merge_excel_records(
    base_records: list[dict],
    pod_records: list[dict],
    *,
    include_automation_hub: bool = True,
) -> list[dict]:
    merged: list[dict] = []
    pod_by_order: dict[str, list[dict]] = {}
    pod_no_order: list[dict] = []

    for record in pod_records:
        order_number = str(clean_value(record.get("order_number")) or "").strip()
        if order_number:
            pod_by_order.setdefault(order_number, []).append(record)
        else:
            pod_no_order.append(record)

    if include_automation_hub:
        merged.append(automation_hub_record())

    consumed_orders: set[str] = set()
    for index, record in enumerate(base_records):
        merged.append(record)
        order_number = str(clean_value(record.get("order_number")) or "").strip()
        next_order = ""
        if index + 1 < len(base_records):
            next_order = str(clean_value(base_records[index + 1].get("order_number")) or "").strip()
        if order_number and order_number != next_order and order_number not in consumed_orders:
            merged.extend(pod_by_order.get(order_number, []))
            consumed_orders.add(order_number)

    for order_number, records in pod_by_order.items():
        if order_number not in consumed_orders:
            merged.extend(records)
    merged.extend(pod_no_order)
    return merged


def load_excel_records(
    project_root: Path,
    *,
    include_automation_hub: bool = True,
    sync_pod_json: bool = True,
) -> list[dict]:
    base_records = load_results_records(project_root)
    if sync_pod_json:
        pod_records, _changed = sync_proof_of_delivery_records(project_root)
    else:
        pod_records = load_proof_of_delivery_records(project_root)
    return merge_excel_records(
        base_records,
        pod_records,
        include_automation_hub=include_automation_hub,
    )


def remaining_pod_candidates(project_root: Path) -> list[dict]:
    try:
        from trackingNumbersViewer.seventeen_track_smart import tracking_is_greyed_out
    except Exception:
        def tracking_is_greyed_out(_tracking_number: str) -> bool:
            return False

    base_records = load_results_records(project_root)
    seen_numbers: set[str] = set()
    out: list[dict] = []
    for record in base_records:
        if is_pod_record(record) or is_automation_hub_record(record):
            continue
        company = _base_record_company(record)
        purchase_datetime = clean_value(record.get("purchase_datetime"))
        order_number = clean_value(record.get("order_number"))
        source_category = clean_value(record.get("email_category"))
        for tracking_number in tracking_numbers_for_record(record):
            if tracking_number in seen_numbers:
                continue
            seen_numbers.add(tracking_number)
            if tracking_is_greyed_out(tracking_number):
                continue
            carrier_display = _carrier_display_for_number(tracking_number)
            pdf_path = expected_pod_pdf_path(
                project_root,
                company,
                purchase_datetime,
                tracking_number,
                carrier_display,
            )
            if first_existing_capture_pdf_path(
                project_root,
                company,
                purchase_datetime,
                tracking_number,
                carrier_display,
                order_number,
                source_category,
            ) is not None:
                continue
            out.append(
                {
                    "tracking_number": tracking_number,
                    "carrier": carrier_display,
                    "company": company,
                    "purchase_datetime": purchase_datetime,
                    "order_number": order_number,
                    "category": source_category,
                    "expected_pdf_path": str(pdf_path.resolve()),
                }
            )
    return out


def pod_status_viewer_rows(project_root: Path) -> list[dict]:
    """All tracking rows for the POD/status viewer, including processed and grey rows."""
    try:
        from trackingNumbersViewer.seventeen_track_smart import tracking_is_greyed_out
    except Exception:
        def tracking_is_greyed_out(_tracking_number: str) -> bool:
            return False

    base_records = load_results_records(project_root)
    seen_numbers: set[str] = set()
    out: list[dict] = []
    for record in base_records:
        if is_pod_record(record) or is_automation_hub_record(record):
            continue
        company = _base_record_company(record)
        purchase_datetime = clean_value(record.get("purchase_datetime"))
        order_number = clean_value(record.get("order_number"))
        source_category = clean_value(record.get("email_category"))
        for tracking_number in tracking_numbers_for_record(record):
            if tracking_number in seen_numbers:
                continue
            seen_numbers.add(tracking_number)
            carrier_display = _carrier_display_for_number(tracking_number)
            pdf_path = expected_pod_pdf_path(
                project_root,
                company,
                purchase_datetime,
                tracking_number,
                carrier_display,
            )
            out.append(
                {
                    "tracking_number": tracking_number,
                    "carrier": carrier_display,
                    "company": company,
                    "purchase_datetime": purchase_datetime,
                    "order_number": order_number,
                    "category": source_category,
                    "expected_pdf_path": str(pdf_path.resolve()),
                    "greyed_out": bool(tracking_is_greyed_out(tracking_number)),
                }
            )
    return out


def delete_processed_tracking_artifacts(
    project_root: Path,
    *,
    tracking_number: object,
    company: object = "",
    purchase_datetime: object = "",
    carrier_display: object = "",
    order_number: object = "",
    category: object = "",
    processed_pdf_path: str | Path | None = None,
) -> dict[str, int]:
    """Delete capture artifacts for one tracking row and forget related JSON references."""
    target_tracking = _normalize_tracking_number(tracking_number)
    target_order = _normalized_text(order_number)
    target_category = _normalized_text(category)
    candidate_paths: dict[str, Path] = {}

    def add_candidate(path: Path | None) -> None:
        if path is None:
            return
        key = _path_key(path)
        candidate_paths[key] = path

    if isinstance(processed_pdf_path, Path):
        add_candidate(processed_pdf_path)
    elif processed_pdf_path:
        add_candidate(_safe_path(processed_pdf_path))

    add_candidate(
        first_existing_capture_pdf_path(
            project_root,
            company,
            purchase_datetime,
            tracking_number,
            carrier_display,
            order_number,
            category,
        )
    )
    for basename in (
        pod_pdf_basename(company, purchase_datetime, tracking_number, carrier_display),
        legacy_pod_pdf_basename(company, purchase_datetime, tracking_number, carrier_display),
        legacy_email_capture_pdf_basename(company, purchase_datetime, order_number, category),
    ):
        for path in _all_existing_pdf_named(project_root, basename):
            add_candidate(path)

    pod_records = load_proof_of_delivery_records(project_root)
    for record in pod_records:
        record_tracking = _record_tracking_number(record)
        if target_tracking and record_tracking != target_tracking:
            continue
        if target_order and _normalized_text(record.get("order_number")) not in ("", target_order):
            continue
        if target_category and _normalized_text(record.get("pod_source_category") or record.get("email_category")) not in ("", target_category):
            continue
        add_candidate(_safe_path(record.get("source_file")))
        add_candidate(_path_from_file_uri(record.get("source_file_link")))

    audit_file = project_root / "email_contents" / "json" / "tracking_pdf_audit.json"
    if audit_file.is_file():
        try:
            payload = json.loads(audit_file.read_text(encoding="utf-8"))
            audit_entries = [entry for entry in payload if isinstance(entry, dict)] if isinstance(payload, list) else []
        except (OSError, json.JSONDecodeError):
            audit_entries = []
    else:
        audit_entries = []
    for entry in audit_entries:
        entry_tracking = _normalize_tracking_number(entry.get("tracking_number"))
        if target_tracking and entry_tracking != target_tracking:
            continue
        if target_order and _normalized_text(entry.get("order_number")) not in ("", target_order):
            continue
        if target_category and _normalized_text(entry.get("category")) not in ("", target_category):
            continue
        add_candidate(_safe_path(entry.get("path")))

    deleted_paths: set[str] = set()
    deleted_pdf_count = 0
    for key, path in list(candidate_paths.items()):
        try:
            if path.is_file():
                path.unlink()
                deleted_pdf_count += 1
                deleted_paths.add(key)
        except OSError:
            continue

    if not deleted_paths:
        for key in candidate_paths:
            deleted_paths.add(key)

    kept_pod_records: list[dict] = []
    removed_pod_records = 0
    for record in pod_records:
        record_path = _safe_path(record.get("source_file"))
        record_uri_path = _path_from_file_uri(record.get("source_file_link"))
        record_tracking = _record_tracking_number(record)
        path_match = (
            (record_path is not None and _path_key(record_path) in deleted_paths)
            or (record_uri_path is not None and _path_key(record_uri_path) in deleted_paths)
        )
        tracking_match = bool(target_tracking) and record_tracking == target_tracking
        order_match = (not target_order) or (_normalized_text(record.get("order_number")) == target_order)
        if (path_match or tracking_match) and order_match:
            removed_pod_records += 1
            continue
        kept_pod_records.append(record)

    if removed_pod_records:
        save_json_records(proof_of_delivery_json_path(project_root), kept_pod_records)

    kept_audit_entries: list[dict[str, Any]] = []
    removed_audit_entries = 0
    for entry in audit_entries:
        entry_path = _safe_path(entry.get("path"))
        entry_tracking = _normalize_tracking_number(entry.get("tracking_number"))
        path_match = entry_path is not None and _path_key(entry_path) in deleted_paths
        tracking_match = bool(target_tracking) and entry_tracking == target_tracking
        order_match = (not target_order) or (
            _normalized_text(entry.get("order_number")) in ("", target_order)
        )
        if (path_match or tracking_match) and order_match:
            removed_audit_entries += 1
            continue
        kept_audit_entries.append(entry)

    if removed_audit_entries:
        audit_file.parent.mkdir(parents=True, exist_ok=True)
        audit_file.write_text(
            json.dumps(kept_audit_entries, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    return {
        "deleted_pdfs": deleted_pdf_count,
        "removed_pod_records": removed_pod_records,
        "removed_audit_entries": removed_audit_entries,
    }


def parse_sortable_datetime(value: object) -> datetime | None:
    raw = str(clean_value(value) or "").strip()
    if not raw:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None
