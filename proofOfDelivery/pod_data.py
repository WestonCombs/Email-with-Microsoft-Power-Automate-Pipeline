from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from trackingNumbersViewer.mitm_readiness import sanitize_filename_token

POD_CATEGORY = "POD"
AUTOMATION_HUB_CATEGORY = "Automation Hub"
AUTOMATION_HUB_ORDER_LABEL = "POD Automation"
AUTOMATION_HUB_COMPANY_LABEL = "Proof of Delivery"
AUTOMATION_HUB_STATUS_LABEL = "Process Remaining PODs"
POD_HUB_MODE = "remaining_pod_hub"
PROOF_OF_DELIVERY_JSON_NAME = "proof_of_delivery.json"


def project_root_from_env() -> Path:
    base_raw = (os.getenv("BASE_DIR") or "").strip()
    if not base_raw:
        raise ValueError(
            'BASE_DIR is not set. Set it in Email Sorter → Settings ("Project folder on disk") and Save.'
        )
    return Path(base_raw).expanduser().resolve()


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


def pod_pdf_basename(
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
    return str(clean_value(record.get("company")) or "").strip() or "Unknown"


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
            if not pdf_path.is_file():
                continue
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
            if pdf_path.is_file():
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
