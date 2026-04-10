"""Update ``Invoice link`` column on an open Excel workbook via COM."""

from __future__ import annotations

import json
import sys
from pathlib import Path

if sys.platform != "win32":
    raise RuntimeError("excel_link_sync requires Windows")

from giftcardInvoiceLink.link_store import (
    gift_order_link_label,
    load_edges,
    links_path_for_project_root,
    normalized_order_number,
    stable_record_key,
)


def _com_header_col(ws, want: str) -> int:
    """1-based column index or 0."""
    last = 1
    for c in range(1, 40):
        v = ws.Cells(1, c).Value
        if v is None or (isinstance(v, str) and not v.strip()):
            continue
        last = c
    want_l = want.strip().lower()
    for c in range(1, last + 1):
        h = ws.Cells(1, c).Value
        if h is None:
            continue
        if str(h).strip().lower() == want_l:
            return c
    return 0


def _sheet_name_for_ref(name: str) -> str:
    if "'" in name:
        return "'" + name.replace("'", "''") + "'"
    if " " in name or not name.replace("_", "").isalnum():
        return "'" + name + "'"
    return name


def _col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def sync_workbook_invoice_links(
    workbook_com,
    *,
    project_root: Path,
    sheet_name: str = "Orders",
    data_start_row: int = 2,
) -> bool:
    """Refresh Invoice link column from JSON + link file. Returns True if updated."""
    json_path = project_root / "email_contents" / "json" / "results.json"
    if not json_path.is_file():
        return False
    records: list[dict] = json.loads(json_path.read_text(encoding="utf-8"))
    if not isinstance(records, list):
        return False
    edges = load_edges(links_path_for_project_root(project_root), records)

    try:
        ws = workbook_com.Worksheets(sheet_name)
    except Exception:
        return False

    col_invoice = _com_header_col(ws, "Invoice link")
    col_cat = _com_header_col(ws, "Category")
    if col_invoice == 0 or col_cat == 0:
        return False

    sn = _sheet_name_for_ref(str(ws.Name))
    n = len(records)
    for i in range(n):
        row = data_start_row + i
        rec = records[i]
        key = stable_record_key(rec, i)
        cat = rec.get("email_category")
        ordn = normalized_order_number(rec)

        act = gift_order_link_label(
            cat if isinstance(cat, str) else None,
            key,
            ordn,
            edges,
        )
        rng_act = ws.Cells(row, col_invoice)
        while rng_act.Hyperlinks.Count > 0:
            rng_act.Hyperlinks(1).Delete()

        if act:
            rng_act.Value = act
            sub = f"#{sn}!{_col_letter(col_cat)}{row}"
            rng_act.Hyperlinks.Add(Anchor=rng_act, Address="", SubAddress=sub, TextToDisplay=act)
            try:
                cat_cell = ws.Cells(row, col_cat)
                rng_act.Interior.Color = cat_cell.Interior.Color
            except Exception:
                pass
        else:
            rng_act.Value = None

    return True


def find_workbook_by_path(excel_app, path: str):
    """Return Workbook COM object whose FullName matches *path*."""
    from pathlib import Path as P

    want = str(P(path).resolve())
    try:
        n = int(excel_app.Workbooks.Count)
    except Exception:
        return None
    for i in range(1, n + 1):
        try:
            wb = excel_app.Workbooks(i)
            cur = str(P(str(wb.FullName)).resolve())
        except Exception:
            continue
        if cur.lower() == want.lower():
            return wb
    return None
