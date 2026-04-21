"""Update ``Invoice link`` column on an open Excel workbook via COM."""

from __future__ import annotations

import contextlib
import os
import re
import sys
import time
from pathlib import Path

if sys.platform != "win32":
    raise RuntimeError("excel_link_sync requires Windows")

# Excel Constants — Borders (avoid win32com.client.constants dependency)
_XL_EDGE_LEFT = 7
_XL_EDGE_TOP = 8
_XL_EDGE_BOTTOM = 9
_XL_EDGE_RIGHT = 10
_XL_CONTINUOUS = 1  # xlContinuous
_XL_HAIRLINE = 1  # xlHairline border weight (matches openpyxl hair side)

from giftcardInvoiceLink.link_store import (
    gift_order_link_label,
    load_edges,
    links_path_for_project_root,
    normalized_order_number,
    stable_record_key,
)
from proofOfDelivery.pod_data import load_excel_records


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


def _norm_subaddress(value: object) -> str:
    """Normalize hyperlink SubAddress for comparison (Excel formats vary)."""
    s = str(value or "").strip()
    if s.startswith("#"):
        s = s[1:]
    s = s.replace("$", "")
    s = re.sub(r"\s+", "", s)
    return s.casefold()


def _invoice_cell_matches(
    rng_act,
    want_text: str | None,
    want_sub_norm: str | None,
) -> bool:
    """True if the cell already shows *want_text* and the same in-sheet jump target."""
    cur = rng_act.Value
    cur_text = str(cur).strip() if cur is not None else ""
    want_t = (want_text or "").strip()

    if want_t:
        if cur_text != want_t:
            return False
        try:
            n = int(rng_act.Hyperlinks.Count)
        except Exception:
            return False
        if n < 1:
            return False
        try:
            sub = rng_act.Hyperlinks(1).SubAddress
        except Exception:
            return False
        if want_sub_norm is None:
            return False
        return _norm_subaddress(sub) == want_sub_norm
    try:
        hc = int(rng_act.Hyperlinks.Count)
    except Exception:
        return False
    return cur_text == "" and hc == 0


def _hairline_all_edges(rng_act) -> None:
    """Restore hairline borders on a single cell (COM clears them when rewriting values)."""
    for edge in (_XL_EDGE_LEFT, _XL_EDGE_TOP, _XL_EDGE_BOTTOM, _XL_EDGE_RIGHT):
        try:
            b = rng_act.Borders(edge)
            b.LineStyle = _XL_CONTINUOUS
            b.Weight = _XL_HAIRLINE
            b.ColorIndex = 1
        except Exception:
            pass


@contextlib.contextmanager
def _excel_quiet_automation(app):
    """Freeze user interaction during COM updates; always restore prior Application state."""
    try:
        old_scr = bool(app.ScreenUpdating)
        old_evt = bool(app.EnableEvents)
        old_int = bool(app.Interactive)
    except Exception:
        old_scr, old_evt, old_int = True, True, True
    try:
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.Interactive = False
        yield
    finally:
        try:
            app.Interactive = old_int
            app.EnableEvents = old_evt
            app.ScreenUpdating = old_scr
        except Exception:
            pass


def sync_workbook_invoice_links(
    workbook_com,
    *,
    project_root: Path,
    sheet_name: str = "Orders",
    data_start_row: int = 2,
) -> bool:
    """Refresh Invoice link column from JSON + link file. Returns True if updated."""
    records = load_excel_records(project_root, include_automation_hub=True, sync_pod_json=False)
    if not isinstance(records, list) or not records:
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

    app = workbook_com.Application
    with _excel_quiet_automation(app):
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

            want_sub_norm: str | None = None
            if act:
                sub_raw = f"#{sn}!{_col_letter(col_cat)}{row}"
                want_sub_norm = _norm_subaddress(sub_raw)

            if act and _invoice_cell_matches(rng_act, act, want_sub_norm):
                _hairline_all_edges(rng_act)
                continue
            if not act and _invoice_cell_matches(rng_act, None, None):
                _hairline_all_edges(rng_act)
                continue

            while rng_act.Hyperlinks.Count > 0:
                rng_act.Hyperlinks(1).Delete()

            if act:
                rng_act.Value = act
                sub = f"#{sn}!{_col_letter(col_cat)}{row}"
                rng_act.Hyperlinks.Add(
                    Anchor=rng_act, Address="", SubAddress=sub, TextToDisplay=act
                )
                try:
                    cat_cell = ws.Cells(row, col_cat)
                    rng_act.Interior.Color = cat_cell.Interior.Color
                except Exception:
                    pass
            else:
                rng_act.Value = None

            _hairline_all_edges(rng_act)

    return True


def _safe_resolve(path_str: str) -> str:
    """Best-effort absolute path; avoids failing when a segment is temporarily unavailable."""
    p = Path(path_str)
    try:
        return str(p.resolve())
    except (OSError, RuntimeError, ValueError):
        return str(p)


def _workbook_paths_match(want: str, current: str) -> bool:
    """True if *want* and *current* refer to the same workbook file (Windows)."""
    wa = os.path.normcase(os.path.normpath(_safe_resolve(want)))
    wb = os.path.normcase(os.path.normpath(_safe_resolve(current)))
    if wa == wb:
        return True
    try:
        if os.path.samefile(want, current):
            return True
    except OSError:
        pass
    try:
        if os.path.samefile(wa, wb):
            return True
    except OSError:
        pass
    return False


def find_workbook_by_path(excel_app, path: str):
    """Return Workbook COM object whose FullName matches *path*."""
    try:
        n = int(excel_app.Workbooks.Count)
    except Exception:
        return None
    for i in range(1, n + 1):
        try:
            wb = excel_app.Workbooks(i)
            cur = str(wb.FullName)
        except Exception:
            continue
        if _workbook_paths_match(path, cur):
            return wb
    return None


def _find_workbook_via_rot(want_path: str):
    """
    Find an open Workbook COM object by path without relying on GetActiveObject's Excel instance.

    When multiple ``Excel.exe`` processes are running, ``GetActiveObject`` may attach to an
    instance that does not own the workbook the user clicked from; ROT still exposes the file.
    """
    import pythoncom
    import win32com.client

    try:
        ctx = pythoncom.CreateBindCtx(0)
        rot = pythoncom.GetRunningObjectTable()
    except Exception:
        return None
    try:
        enum = rot.EnumRunning()
    except Exception:
        return None
    if enum is None:
        return None
    while True:
        try:
            mons = enum.Next(1)
        except pythoncom.com_error:
            break
        if not mons:
            break
        mon = mons[0]
        try:
            dname = mon.GetDisplayName(ctx, None)
        except TypeError:
            try:
                dname = mon.GetDisplayName(ctx, mon)
            except Exception:
                continue
        except Exception:
            continue
        low = dname.lower()
        if ".xls" not in low:
            continue
        try:
            obj = rot.GetObject(mon)
            wb = win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
        except Exception:
            continue
        try:
            fn = str(wb.FullName)
        except Exception:
            continue
        if _workbook_paths_match(want_path, fn):
            return wb
    return None


def find_workbook_and_application(
    want_path: str,
    *,
    poll_interval_sec: float = 0.1,
    max_wait_sec: float = 1.5,
) -> tuple[object | None, object | None]:
    """
    Return ``(workbook, excel_application)`` for *want_path*, polling briefly.

    Tries the active Excel instance first, then the Running Object Table (other instances).
    """
    import win32com.client

    deadline = time.monotonic() + max_wait_sec
    last_excel = None
    while time.monotonic() < deadline:
        try:
            last_excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            last_excel = None

        if last_excel is not None:
            wb = find_workbook_by_path(last_excel, want_path)
            if wb is not None:
                return wb, last_excel

        wb_rot = _find_workbook_via_rot(want_path)
        if wb_rot is not None:
            try:
                return wb_rot, wb_rot.Application
            except Exception:
                return wb_rot, last_excel

        time.sleep(poll_interval_sec)

    return None, last_excel
