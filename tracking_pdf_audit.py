"""Audit logging for assisted tracking PDF captures."""

from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


AUDIT_FILE_NAME = "tracking_pdf_audit.json"


def _project_root() -> Path:
    """Return the project root used by the rest of the email sorter."""
    from shared.project_paths import ensure_base_dir_in_environ

    return ensure_base_dir_in_environ()


def audit_path() -> Path:
    """Return the JSON audit file path, creating the parent directory if needed."""
    path = _project_root() / "email_contents" / "json" / AUDIT_FILE_NAME
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _normalize_tracking_number(value: object) -> str:
    text = str(value or "").strip()
    return "".join(ch for ch in text if not ch.isspace())


def _record_text(record: dict, *keys: str) -> str:
    for key in keys:
        value = record.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return ""


def _order_last4(record: dict) -> str:
    raw = _record_text(record, "order_number")
    digits = re.sub(r"\D", "", raw)
    if len(digits) >= 4:
        return digits[-4:]
    if digits:
        return digits.zfill(4)
    return "0000"


def _load_audit_entries(path: Path) -> list[dict[str, Any]]:
    if not path.is_file():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    return []


def load_tracking_pdf_audit_entries() -> list[dict[str, Any]]:
    """Return audit entries newest-first."""
    entries = _load_audit_entries(audit_path())

    def sort_key(entry: dict[str, Any]) -> str:
        return str(entry.get("timestamp_captured") or "")

    entries.sort(key=sort_key, reverse=True)
    return entries


def _confidence_value(validation: dict) -> int | float:
    raw = validation.get("confidence", 0)
    try:
        value = float(raw)
    except (TypeError, ValueError):
        return 0
    if value.is_integer():
        return int(value)
    return value


def log_tracking_pdf(pdf_path: str, record: dict, validation: dict) -> None:
    """Append one tracking PDF capture audit entry to ``tracking_pdf_audit.json``."""
    path = Path(pdf_path).expanduser().resolve()
    validation = validation if isinstance(validation, dict) else {}
    entry = {
        "filename": path.name,
        "company": _record_text(record, "company") or "Unknown",
        "order_number": _record_text(record, "order_number"),
        "order_last4": _order_last4(record),
        "category": _record_text(record, "email_category", "category") or "Unknown",
        "purchase_datetime": _record_text(record, "purchase_datetime"),
        "tracking_number": _normalize_tracking_number(record.get("tracking_number")),
        "timestamp_captured": datetime.now(timezone.utc).isoformat(),
        "path": str(path),
        "latest_tracking_info_visible": bool(validation.get("latest_tracking_info_visible")),
        "confidence": _confidence_value(validation),
        "status_found": validation.get("status_found", "Unknown") or "Unknown",
        "latest_update_found": validation.get("latest_update_found"),
        "reason": validation.get("reason", "") or "",
    }

    target = audit_path()
    entries = _load_audit_entries(target)
    entries.append(entry)
    tmp = target.with_suffix(".tmp")
    tmp.write_text(json.dumps(entries, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(target)
