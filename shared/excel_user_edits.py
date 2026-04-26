from __future__ import annotations

from collections import Counter, defaultdict
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ALLOWED_EXCEL_USER_EDIT_FIELDS = ("company", "total_amount_paid", "tax_paid")
ALLOWED_EXCEL_USER_EDIT_LABELS = ("Company", "Total Paid", "Tax Paid")
EXCEL_USER_EDITS_JSON_NAME = "excel_user_edits.json"
LLM_OBTAINED_COMPANY_FIELD = "llm_obtained_company"
ORIGINAL_LLM_OBTAINED_COMPANY_FIELD = "original_llm_obtained_company"


def excel_user_edits_path(project_root: Path) -> Path:
    return project_root / "email_contents" / "json" / EXCEL_USER_EDITS_JSON_NAME


def modified_key(field: str) -> str:
    return f"modified_{field}"


def is_modified(record: dict, field: str) -> bool:
    return bool(record.get(modified_key(field)))


def strip_excel_modified_marker(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    text = value.strip()
    while text.endswith("*"):
        text = text[:-1].rstrip()
    return text


def _clean_record_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        text = value.strip()
        if not text or text.lower() == "null":
            return None
    return value


def _infer_company_from_subject(subject: Any) -> str | None:
    subject = _clean_record_value(subject)
    if not isinstance(subject, str):
        return None

    normalized = subject
    while True:
        updated = re.sub(
            r"^\s*(?:fw|fwd|re)\s*:\s*", "", normalized, flags=re.IGNORECASE
        )
        if updated == normalized:
            break
        normalized = updated

    patterns = [
        r"your\s+(.+?)\s+order(?:\b|:)",
        r"order\s+from\s+(.+?)(?:\b|:)",
        r"thanks\s+.+?\s+for\s+your\s+purchase\s+with\s+(.+?)(?:\b|!|\.|:)",
    ]
    for pattern in patterns:
        match = re.search(pattern, normalized, flags=re.IGNORECASE)
        if match:
            company = _clean_record_value(match.group(1))
            if isinstance(company, str) and company:
                return company.strip(" -,:;.!?")
    return None


def _infer_company_from_source_file(source_file: Any) -> str | None:
    if source_file is None:
        return None
    raw = str(source_file).strip()
    if not raw:
        return None
    try:
        stem = Path(raw).stem
    except OSError:
        return None
    stem = re.sub(r" \(\d+\)$", "", stem).strip()
    patterns = (
        r"^DOC (.+?) \d{4}-\d{2}-\d{2} (?:INVOICE|SHIPPED|DELIVERED)_\d{4}$",
        r"^DOC (.+?) \d{4}-\d{2}-\d{2} TRACKING_INV_\d{4}$",
        r"^DOC (.+?) \d{4}-\d{2}-\d{2} .+_FROM_.+$",
        r"^DOC (.+?) \d{4}-\d{2}-\d{2}_\d{4}$",
        r"^(.+?) \d{4}-\d{2}-\d{2}_\d{4}$",
    )
    for pattern in patterns:
        match = re.match(pattern, stem, flags=re.IGNORECASE)
        if not match:
            continue
        company = _clean_record_value(match.group(1))
        if isinstance(company, str) and company:
            return company
    return None


def _normalized_company_vote_key(company: Any) -> str:
    text = _clean_record_value(company)
    if not isinstance(text, str) or not text:
        return ""
    normalized = text.casefold()
    normalized = re.sub(r"\s+", " ", normalized)
    normalized = normalized.replace("&", " and ")
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


def _company_consensus_value(values: list[str]) -> str | None:
    key_votes: Counter[str] = Counter()
    originals_by_vote_key: dict[str, list[str]] = defaultdict(list)
    for value in values:
        cleaned = _clean_record_value(value)
        if not isinstance(cleaned, str) or not cleaned:
            continue
        vote_key = _normalized_company_vote_key(cleaned)
        if not vote_key:
            continue
        key_votes[vote_key] += 1
        originals_by_vote_key[vote_key].append(cleaned)
    if not key_votes:
        return None
    winning_vote_key = sorted(
        key_votes.items(),
        key=lambda item: (
            -item[1],
            -max((len(x) for x in originals_by_vote_key[item[0]]), default=0),
            item[0],
        ),
    )[0][0]
    originals = Counter(originals_by_vote_key[winning_vote_key])
    return sorted(
        originals.items(),
        key=_company_display_sort_key,
    )[0][0]


def _company_display_sort_key(item: tuple[str, int]) -> tuple[int, int, int, str, str]:
    value, count = item
    has_alpha = any(ch.isalpha() for ch in value)
    all_caps_penalty = 1 if has_alpha and value.upper() == value else 0
    return (-count, all_caps_penalty, -len(value), value.casefold(), value)


def _company_baseline_candidates(
    record: dict, overlay: dict | None = None
) -> list[str]:
    candidates: list[str] = []
    seen_vote_keys: set[str] = set()

    values: list[Any] = [
        record.get(ORIGINAL_LLM_OBTAINED_COMPANY_FIELD),
        record.get(LLM_OBTAINED_COMPANY_FIELD),
    ]
    if overlay is not None:
        values.append(_original_value_for_clear(overlay, record, "company"))
    if not is_modified(record, "company"):
        values.append(record.get("company"))
    values.append(_infer_company_from_subject(record.get("subject")))

    if not any(_normalized_company_vote_key(v) for v in values):
        values.append(_infer_company_from_source_file(record.get("source_file")))

    for value in values:
        cleaned = _clean_record_value(value)
        if not isinstance(cleaned, str) or not cleaned:
            continue
        vote_key = _normalized_company_vote_key(cleaned)
        if not vote_key or vote_key in seen_vote_keys:
            continue
        seen_vote_keys.add(vote_key)
        candidates.append(cleaned)

    return candidates


def ensure_llm_obtained_company_fields(
    records: list[dict], overlay: dict | None = None
) -> bool:
    changed = False
    order_candidates: dict[str, list[str]] = defaultdict(list)
    record_candidates: dict[int, str | None] = {}

    for record in records:
        if not isinstance(record, dict):
            continue
        if (
            ORIGINAL_LLM_OBTAINED_COMPANY_FIELD not in record
            and _clean_record_value(record.get(LLM_OBTAINED_COMPANY_FIELD))
        ):
            record[ORIGINAL_LLM_OBTAINED_COMPANY_FIELD] = _clean_record_value(
                record.get(LLM_OBTAINED_COMPANY_FIELD)
            )
            changed = True
        candidates = _company_baseline_candidates(record, overlay)
        record_candidates[id(record)] = candidates[0] if candidates else None
        order_number = str(record.get("order_number") or "").strip()
        if order_number:
            order_candidates[order_number].extend(candidates)

    order_winners = {
        order_number: _company_consensus_value(values)
        for order_number, values in order_candidates.items()
    }

    for record in records:
        if not isinstance(record, dict):
            continue
        order_number = str(record.get("order_number") or "").strip()
        desired_llm_company = order_winners.get(order_number) or record_candidates.get(id(record))
        if (
            LLM_OBTAINED_COMPANY_FIELD not in record
            or _clean_record_value(record.get(LLM_OBTAINED_COMPANY_FIELD)) != desired_llm_company
        ):
            record[LLM_OBTAINED_COMPANY_FIELD] = desired_llm_company
            changed = True
        if not is_modified(record, "company"):
            if _clean_record_value(record.get("company")) != desired_llm_company:
                record["company"] = desired_llm_company
                changed = True

    return changed


def company_display_value(record: dict) -> Any:
    explicit_company = _clean_record_value(record.get("company"))
    if explicit_company:
        return explicit_company
    return _clean_record_value(record.get(LLM_OBTAINED_COMPANY_FIELD))


def display_value_for_field(record: dict, field: str) -> Any:
    if field == "company":
        return company_display_value(record)
    return _clean_record_value(record.get(field))


def display_value_kind(value: Any) -> str:
    if value is None:
        return "blank"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return "number"
    return "text"


def display_value_for_excel(record: dict, field: str, value: Any) -> Any:
    if not is_modified(record, field):
        return value
    if value is None:
        return "*"
    text = str(value)
    if text.rstrip().endswith("*"):
        return text
    return f"{text}*"


def coerce_user_edit_value(field: str, raw_value: Any) -> Any:
    if field not in ALLOWED_EXCEL_USER_EDIT_FIELDS:
        raise ValueError(f"Unsupported Excel user-edit field: {field}")

    value = strip_excel_modified_marker(raw_value)
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None

    if field == "company":
        return text

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
    except ValueError as exc:
        raise ValueError(f"{field} must be a number-like value, got {text!r}") from exc


def record_identity(record: dict) -> str:
    for key in ("source_file_link", "source_file", "content_hash"):
        value = record.get(key)
        if value is not None and str(value).strip():
            return f"{key}:{str(value).strip()}"
    parts = [
        str(record.get("order_number") or "").strip(),
        str(record.get("email_category") or "").strip(),
        str(record.get("purchase_datetime") or "").strip(),
        str(record.get("subject") or "").strip(),
        str(record.get("email") or "").strip(),
    ]
    return "fallback:" + "|".join(parts)


def record_matches_source_uri(record: dict, source_uri: str) -> bool:
    want = str(source_uri or "").strip()
    if not want:
        return False
    for key in ("source_file_link", "source_file"):
        value = record.get(key)
        if value is not None and str(value).strip() == want:
            return True
    return False


def load_json_records(path: Path) -> list[dict]:
    if not path.is_file():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    return [item for item in data if isinstance(item, dict)] if isinstance(data, list) else []


def save_json_records(path: Path, records: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(records, indent=2, ensure_ascii=False), encoding="utf-8")


def _empty_overlay() -> dict:
    return {"version": 1, "records": {}, "order_company": {}}


def load_user_edit_overlay(project_root: Path) -> dict:
    path = excel_user_edits_path(project_root)
    if not path.is_file():
        return _empty_overlay()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return _empty_overlay()
    if not isinstance(payload, dict):
        return _empty_overlay()
    payload.setdefault("version", 1)
    if not isinstance(payload.get("records"), dict):
        payload["records"] = {}
    if not isinstance(payload.get("order_company"), dict):
        payload["order_company"] = {}
    return payload


def save_user_edit_overlay(project_root: Path, overlay: dict) -> None:
    path = excel_user_edits_path(project_root)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(overlay, indent=2, ensure_ascii=False), encoding="utf-8")


def apply_user_edits_to_records(project_root: Path, records: list[dict]) -> list[dict]:
    overlay = load_user_edit_overlay(project_root)
    records_overlay = overlay.get("records") if isinstance(overlay, dict) else {}
    order_company = overlay.get("order_company") if isinstance(overlay, dict) else {}
    if not isinstance(records_overlay, dict):
        records_overlay = {}
    if not isinstance(order_company, dict):
        order_company = {}
    ensure_llm_obtained_company_fields(records, overlay)
    overlay_changed = False

    records_by_order: dict[str, list[dict]] = defaultdict(list)
    for record in records:
        if not isinstance(record, dict):
            continue
        order_number = str(record.get("order_number") or "").strip()
        if order_number:
            records_by_order[order_number].append(record)

    for order_number, company_edit in list(order_company.items()):
        if not isinstance(company_edit, dict) or company_edit.get("value") is None:
            continue
        order_records = records_by_order.get(str(order_number))
        if order_records and all(
            _company_value_matches_stored_unmodified(
                record, overlay, company_edit.get("value")
            )
            for record in order_records
        ):
            order_company.pop(order_number, None)
            overlay_changed = True

    for record in records:
        if not isinstance(record, dict):
            continue
        if is_modified(record, "company") and _company_value_matches_stored_unmodified(
            record, overlay, record.get("company")
        ):
            restored_company = _clean_record_value(record.get(LLM_OBTAINED_COMPANY_FIELD))
            if restored_company is None:
                restored_company = _original_value_for_clear(overlay, record, "company")
            record["company"] = restored_company
            _clear_modified_field_state(record, "company")

        order_number = str(record.get("order_number") or "").strip()
        company_edit = order_company.get(order_number)
        if (
            isinstance(company_edit, dict)
            and "value" in company_edit
            and company_edit.get("value") is not None
        ):
            if not _company_value_matches_stored_unmodified(
                record, overlay, company_edit.get("value")
            ):
                record["company"] = company_edit.get("value")
                record[modified_key("company")] = True
                record["user_modified"] = True

        item = records_overlay.get(record_identity(record))
        if not isinstance(item, dict):
            continue
        values = item.get("values")
        if not isinstance(values, dict):
            continue
        for field in ALLOWED_EXCEL_USER_EDIT_FIELDS:
            if field in values and values[field] is not None:
                if field == "company" and _company_value_matches_stored_unmodified(
                    record, overlay, values[field]
                ):
                    _clear_record_overlay(overlay, record, field)
                    overlay_changed = True
                    continue
                record[field] = values[field]
                record[modified_key(field)] = True
                record["user_modified"] = True

    if overlay_changed:
        save_user_edit_overlay(project_root, overlay)

    return records


def apply_user_edits_to_json_files(project_root: Path) -> None:
    results_path, pod_path = _json_paths(project_root)
    for path in (results_path, pod_path):
        records = load_json_records(path)
        if not records:
            continue
        before = json.dumps(records, sort_keys=True, ensure_ascii=False)
        apply_user_edits_to_records(project_root, records)
        after = json.dumps(records, sort_keys=True, ensure_ascii=False)
        if before != after:
            save_json_records(path, records)


def _json_paths(project_root: Path) -> tuple[Path, Path]:
    json_dir = project_root / "email_contents" / "json"
    return json_dir / "results.json", json_dir / "proof_of_delivery.json"


def _set_modified_value(record: dict, field: str, value: Any, timestamp: str) -> None:
    record[field] = value
    record[modified_key(field)] = True
    record["user_modified"] = True
    record["user_modified_at"] = timestamp


def _refresh_user_modified_state(record: dict, timestamp: str) -> None:
    any_modified = any(
        str(key).startswith("modified_") and bool(value)
        for key, value in record.items()
    )
    if any_modified:
        record["user_modified"] = True
        record["user_modified_at"] = timestamp
        return
    record.pop("user_modified", None)
    record.pop("user_modified_at", None)


def _clear_modified_field_state(record: dict, field: str) -> None:
    record.pop(modified_key(field), None)
    any_modified = any(
        str(key).startswith("modified_") and bool(value)
        for key, value in record.items()
    )
    if not any_modified:
        record.pop("user_modified", None)
        record.pop("user_modified_at", None)


def _restore_unmodified_value(
    record: dict, field: str, value: Any, timestamp: str
) -> None:
    record[field] = value
    record.pop(modified_key(field), None)
    _refresh_user_modified_state(record, timestamp)


def _ensure_overlay_item(overlay: dict, record: dict) -> dict:
    records_overlay = overlay.setdefault("records", {})
    key = record_identity(record)
    item = records_overlay.setdefault(key, {})
    item["order_number"] = str(record.get("order_number") or "").strip()
    if record.get("source_file_link"):
        item["source_file_link"] = record.get("source_file_link")
    return item


def _remember_original_value(overlay: dict, record: dict, field: str) -> None:
    item = _ensure_overlay_item(overlay, record)
    original_values = item.setdefault("original_values", {})
    if field not in original_values:
        original_values[field] = record.get(field)


def _original_value_for_clear(overlay: dict, record: dict, field: str) -> Any:
    records_overlay = overlay.get("records")
    if not isinstance(records_overlay, dict):
        return record.get(field)
    item = records_overlay.get(record_identity(record))
    if not isinstance(item, dict):
        return record.get(field)
    original_values = item.get("original_values")
    if not isinstance(original_values, dict):
        return record.get(field)
    if field not in original_values:
        return record.get(field)
    return original_values.get(field)


def _company_stored_unmodified_values(record: dict, overlay: dict) -> list[Any]:
    values: list[Any] = [
        record.get(ORIGINAL_LLM_OBTAINED_COMPANY_FIELD),
        record.get(LLM_OBTAINED_COMPANY_FIELD),
        _original_value_for_clear(overlay, record, "company"),
    ]
    if not is_modified(record, "company"):
        values.append(record.get("company"))
    return values


def _company_value_matches_stored_unmodified(
    record: dict, overlay: dict, value: Any
) -> bool:
    wanted = _normalized_company_vote_key(value)
    if not wanted:
        return False
    return any(
        _normalized_company_vote_key(candidate) == wanted
        for candidate in _company_stored_unmodified_values(record, overlay)
    )


def _update_record_overlay(overlay: dict, record: dict, field: str, value: Any, timestamp: str) -> None:
    item = _ensure_overlay_item(overlay, record)
    item["updated_at"] = timestamp
    values = item.setdefault("values", {})
    values[field] = value


def _clear_record_overlay(overlay: dict, record: dict, field: str) -> None:
    records_overlay = overlay.get("records")
    if not isinstance(records_overlay, dict):
        return
    key = record_identity(record)
    item = records_overlay.get(key)
    if not isinstance(item, dict):
        return
    values = item.get("values")
    if isinstance(values, dict):
        values.pop(field, None)
        if not values:
            item.pop("values", None)
    original_values = item.get("original_values")
    if isinstance(original_values, dict):
        original_values.pop(field, None)
        if not original_values:
            item.pop("original_values", None)
    if "values" not in item:
        records_overlay.pop(key, None)


def record_excel_user_edit(
    project_root: Path,
    *,
    field: str,
    raw_value: Any,
    order_number: str = "",
    source_uri: str = "",
) -> dict:
    clear_requested = str(strip_excel_modified_marker(raw_value) or "").strip() == ""
    value = None if clear_requested else coerce_user_edit_value(field, raw_value)
    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    results_path, pod_path = _json_paths(project_root)
    file_records = [
        (results_path, load_json_records(results_path)),
        (pod_path, load_json_records(pod_path)),
    ]
    overlay = load_user_edit_overlay(project_root)
    matched = 0
    changed_files: list[str] = []
    result_record: dict | None = None
    result_record_source_match = False
    changed_by_path: dict[Path, bool] = {}
    matched_entries: list[tuple[Path, dict]] = []

    clean_order = str(order_number or "").strip()
    for path, records in file_records:
        changed_by_path[path] = ensure_llm_obtained_company_fields(records, overlay)
        for record in records:
            if field == "company" and clean_order:
                match = str(record.get("order_number") or "").strip() == clean_order
            else:
                match = record_matches_source_uri(record, source_uri)
            if not match:
                continue
            matched_entries.append((path, record))

    if not matched_entries:
        raise ValueError("Could not match the edited Excel row to a JSON record.")

    reset_requested = clear_requested or (
        field == "company"
        and any(
            _company_value_matches_stored_unmodified(record, overlay, value)
            for _path, record in matched_entries
        )
    )

    matched_by_path: dict[Path, list[dict]] = defaultdict(list)
    for path, record in matched_entries:
        matched_by_path[path].append(record)

    for path, records in file_records:
        changed = changed_by_path.get(path, False)
        for record in matched_by_path.get(path, []):
            if reset_requested:
                if field == "company":
                    restored_value = _clean_record_value(record.get(LLM_OBTAINED_COMPANY_FIELD))
                    if restored_value is None:
                        restored_value = _original_value_for_clear(overlay, record, field)
                else:
                    restored_value = _original_value_for_clear(overlay, record, field)
                _restore_unmodified_value(record, field, restored_value, timestamp)
                _clear_record_overlay(overlay, record, field)
            else:
                _remember_original_value(overlay, record, field)
                _set_modified_value(record, field, value, timestamp)
                _update_record_overlay(overlay, record, field, value, timestamp)
            matched += 1
            changed = True
            is_source_match = record_matches_source_uri(record, source_uri)
            if result_record is None or (is_source_match and not result_record_source_match):
                result_record = record
                result_record_source_match = is_source_match
        if changed:
            save_json_records(path, records)
            changed_files.append(str(path))

    if reset_requested:
        order_company = overlay.get("order_company")
        if field == "company" and clean_order and isinstance(order_company, dict):
            order_company.pop(clean_order, None)
    elif field == "company" and clean_order:
        overlay.setdefault("order_company", {})[clean_order] = {
            "value": value,
            "updated_at": timestamp,
        }
    save_user_edit_overlay(project_root, overlay)
    display_value = display_value_for_field(result_record or {}, field)
    return {
        "field": field,
        "value": None if reset_requested else value,
        "order_number": clean_order,
        "source_uri": source_uri,
        "matched_records": matched,
        "changed_files": changed_files,
        "mode": "cleared" if reset_requested else "modified",
        "display_value": display_value,
        "display_value_kind": display_value_kind(display_value),
    }
