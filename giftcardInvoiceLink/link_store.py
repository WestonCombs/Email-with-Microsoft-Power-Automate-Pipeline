"""Persist gift-card ↔ order-number links beside ``results.json``."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class GiftOrderEdge:
    gift_key: str
    order_number: str


def clean_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str) and value.strip().lower() == "null":
        return None
    return value


def stable_record_key(record: dict, index: int) -> str:
    """Stable id for a JSON row (order_number + source path fallback)."""
    on = clean_value(record.get("order_number"))
    on = str(on).strip() if on is not None else ""
    sf = clean_value(record.get("source_file_link"))
    sf = str(sf).strip() if sf is not None else ""
    if not on and not sf:
        return f"__idx_{index}"
    return f"{on}\x1f{sf}"


def normalized_order_number(record: dict) -> str:
    v = clean_value(record.get("order_number"))
    if v is None:
        return ""
    return str(v).strip()


def index_for_key(records: list[dict], want_key: str) -> int | None:
    for i, r in enumerate(records):
        if stable_record_key(r, i) == want_key:
            return i
    return None


def links_path_for_project_root(project_root: Path) -> Path:
    return project_root / "email_contents" / "json" / "gift_invoice_links.json"


def load_edges(path: Path, records: list[dict] | None = None) -> list[GiftOrderEdge]:
    """Load edges. Migrates legacy ``invoice_key`` rows using *records* when provided."""
    if not path.is_file():
        return []
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    edges: list[GiftOrderEdge] = []
    for item in raw.get("edges", []):
        if not isinstance(item, dict):
            continue
        g = item.get("gift_key")
        if not isinstance(g, str) or not g:
            continue
        on = item.get("order_number")
        if isinstance(on, str) and on.strip():
            edges.append(GiftOrderEdge(gift_key=g, order_number=on.strip()))
            continue
        inv = item.get("invoice_key")
        if isinstance(inv, str) and inv and records:
            oi = index_for_key(records, inv)
            if oi is not None:
                ordn = normalized_order_number(records[oi])
                if ordn:
                    edges.append(GiftOrderEdge(gift_key=g, order_number=ordn))
    return edges


def save_edges(path: Path, edges: list[GiftOrderEdge]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "edges": [
            {"gift_key": e.gift_key, "order_number": e.order_number} for e in edges
        ],
    }
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def add_edge(
    edges: list[GiftOrderEdge], gift_key: str, order_number: str
) -> list[GiftOrderEdge]:
    on = order_number.strip()
    if not on:
        return edges
    for e in edges:
        if e.gift_key == gift_key and e.order_number == on:
            return edges
    return [*edges, GiftOrderEdge(gift_key=gift_key, order_number=on)]


def remove_edge(
    edges: list[GiftOrderEdge], gift_key: str, order_number: str
) -> list[GiftOrderEdge]:
    on = order_number.strip()
    return [e for e in edges if not (e.gift_key == gift_key and e.order_number == on)]


def remove_all_edges_for_gift(edges: list[GiftOrderEdge], gift_key: str) -> list[GiftOrderEdge]:
    return [e for e in edges if e.gift_key != gift_key]


def remove_all_edges_for_order_number(
    edges: list[GiftOrderEdge], order_number: str
) -> list[GiftOrderEdge]:
    on = order_number.strip()
    return [e for e in edges if e.order_number != on]


def gift_order_link_label(
    category: str | None,
    gift_key: str,
    order_num: str,
    edges: list[GiftOrderEdge],
) -> str | None:
    """Cell text for Invoice link column. *order_num* = normalized order for this row."""
    cat = (category or "").strip() if isinstance(category, str) else ""
    if cat == "Automation Hub":
        return None
    if cat == "Gift Card":
        return "Linked" if any(e.gift_key == gift_key for e in edges) else "Link to order"
    if not cat:
        return None
    if not order_num:
        return None
    return (
        "Linked"
        if any(e.order_number == order_num for e in edges)
        else "Link to Gift Card"
    )
