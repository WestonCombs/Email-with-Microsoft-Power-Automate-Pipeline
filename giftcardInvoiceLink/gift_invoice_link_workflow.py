"""
Launched from Excel VBA when the user follows the ``Invoice link`` hyperlink.

Args: ``<workbook_full_path> <excel_row_1based>``

Links a **Gift Card** row to an **order number** (any non–gift-card row with that order shares
the same linked state). Legacy ``invoice_key`` edges are migrated when ``results.json`` is read.
"""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path

if sys.platform != "win32":
    print("This helper requires Windows + Excel.", file=sys.stderr)
    sys.exit(1)

from dotenv import load_dotenv

_PYTHON_FILES = Path(__file__).resolve().parent.parent
if str(_PYTHON_FILES) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES))

load_dotenv(_PYTHON_FILES / ".env")

import os

from gui_aux_singleton import detach_console_win32, register_current_aux_gui

_base = os.getenv("BASE_DIR")
if not _base:
    print("BASE_DIR missing from .env", file=sys.stderr)
    sys.exit(1)

PROJECT_ROOT = Path(_base).expanduser().resolve()
JSON_PATH = PROJECT_ROOT / "email_contents" / "json" / "results.json"

from tkinter import Tk, messagebox

from giftcardInvoiceLink.excel_link_sync import find_workbook_by_path, sync_workbook_invoice_links
from giftcardInvoiceLink.link_store import (
    add_edge,
    clean_value,
    index_for_key,
    load_edges,
    links_path_for_project_root,
    normalized_order_number,
    remove_all_edges_for_gift,
    remove_all_edges_for_order_number,
    save_edges,
    stable_record_key,
)


def _load_records() -> list[dict]:
    if not JSON_PATH.is_file():
        return []
    data = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    return data if isinstance(data, list) else []


def _category(ws, row: int, col_cat: int) -> str:
    v = ws.Cells(row, col_cat).Value
    if v is None:
        return ""
    return str(v).strip()


def _com_header_col(ws, want: str) -> int:
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


def _order_summary(records: list[dict], idx: int) -> str:
    if idx < 0 or idx >= len(records):
        return "?"
    r = records[idx]
    on = clean_value(r.get("order_number"))
    co = clean_value(r.get("company"))
    parts = []
    if on:
        parts.append(f"order {on}")
    if co:
        parts.append(str(co))
    return ", ".join(parts) if parts else f"row {idx + 2}"


def _wait_different_row(excel, initial_row: int) -> int | None:
    deadline = time.time() + 600.0
    while time.time() < deadline:
        try:
            sel = excel.Selection
            if hasattr(sel, "Row"):
                r = int(sel.Row)
                if r != initial_row:
                    return r
        except Exception:
            pass
        time.sleep(0.25)
    return None


def _remove_flow(
    records: list[dict],
    origin_row: int,
    ws,
    col_cat: int,
) -> None:
    if origin_row < 2 or origin_row > len(records) + 1:
        messagebox.showerror("Invoice link", "That row is outside the data range.")
        return
    idx = origin_row - 2
    key = stable_record_key(records[idx], idx)
    cat = _category(ws, origin_row, col_cat)
    link_path = links_path_for_project_root(PROJECT_ROOT)
    edges = load_edges(link_path, records)

    if cat == "Gift Card":
        nums = sorted({e.order_number for e in edges if e.gift_key == key})
        if not nums:
            messagebox.showinfo("Invoice link", "No links are stored for this gift card.")
            return
        msg = (
            "Remove all links from this gift card?\n\n"
            "Linked order numbers:\n• " + "\n• ".join(nums)
        )
        r = messagebox.askyesnocancel("Remove gift / order links", msg)
        if r is not True:
            return
        new_edges = remove_all_edges_for_gift(edges, key)
        save_edges(link_path, new_edges)
        return

    on = normalized_order_number(records[idx])
    if not on:
        messagebox.showinfo("Invoice link", "This row has no order number to unlink.")
        return

    rel = [e for e in edges if e.order_number == on]
    if not rel:
        messagebox.showinfo("Invoice link", "No links are stored for this order number.")
        return

    gift_lines = []
    for e in rel:
        gi = index_for_key(records, e.gift_key)
        if gi is not None:
            gift_lines.append(_order_summary(records, gi))
        else:
            gift_lines.append("(gift card row)")

    msg = (
        f"Remove all gift-card links for order number {on}?\n\n"
        "Linked gift card row(s):\n• " + "\n• ".join(gift_lines)
    )
    r = messagebox.askyesnocancel("Remove gift / order links", msg)
    if r is not True:
        return
    new_edges = remove_all_edges_for_order_number(edges, on)
    save_edges(link_path, new_edges)


def _add_flow(
    records: list[dict],
    origin_row: int,
    excel,
    ws,
    col_cat: int,
) -> None:
    if origin_row < 2 or origin_row > len(records) + 1:
        messagebox.showerror("Invoice link", "That row is outside the data range.")
        return
    oidx = origin_row - 2
    ocat = _category(ws, origin_row, col_cat)

    target_row = _wait_different_row(excel, origin_row)
    if target_row is None:
        messagebox.showwarning("Invoice link", "Timed out waiting for a new row selection.")
        return
    if target_row < 2 or target_row > len(records) + 1:
        messagebox.showerror("Invoice link", "Selected row is outside the data range.")
        return

    tidx = target_row - 2
    tcat = _category(ws, target_row, col_cat)

    if ocat == tcat:
        messagebox.showerror("Invoice link", "Select a different kind of row (gift vs order line).")
        return

    if ocat != "Gift Card" and tcat != "Gift Card":
        messagebox.showerror(
            "Invoice link",
            "One row must be Category “Gift Card” and the other must not be.",
        )
        return

    if ocat == "Gift Card":
        gidx, xidx = oidx, tidx
    else:
        gidx, xidx = tidx, oidx

    order_on = normalized_order_number(records[xidx])
    if not order_on:
        messagebox.showerror(
            "Invoice link",
            "The selected row has no order number. Pick a row with an order number.",
        )
        return

    g_key = stable_record_key(records[gidx], gidx)

    msg = (
        "Create this link?\n\n"
        f"Gift card: {_order_summary(records, gidx)}\n"
        f"Order number (all rows with this number will show as linked): {order_on}"
    )
    r = messagebox.askyesnocancel("Confirm gift / order link", msg)
    if r is not True:
        return

    link_path = links_path_for_project_root(PROJECT_ROOT)
    edges = load_edges(link_path, records)
    edges = add_edge(edges, g_key, order_on)
    save_edges(link_path, edges)


def main() -> None:
    detach_console_win32()
    register_current_aux_gui()

    if len(sys.argv) < 3:
        print(
            "Usage: python gift_invoice_link_workflow.py <workbook.xlsx|xlsm> <row>",
            file=sys.stderr,
        )
        sys.exit(1)

    wb_path = str(Path(sys.argv[1]).resolve())
    try:
        origin_row = int(sys.argv[2])
    except ValueError:
        sys.exit(1)

    records = _load_records()
    if not records:
        root = Tk()
        root.withdraw()
        messagebox.showerror("Invoice link", f"No data or missing JSON:\n{JSON_PATH}")
        root.destroy()
        return

    try:
        import win32com.client

        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        root = Tk()
        root.withdraw()
        messagebox.showerror("Invoice link", "Excel is not running.")
        root.destroy()
        return

    wb = find_workbook_by_path(excel, wb_path)
    if wb is None:
        root = Tk()
        root.withdraw()
        messagebox.showerror(
            "Invoice link",
            "Open this workbook in Excel and keep it as the file you clicked from.",
        )
        root.destroy()
        return

    try:
        ws = wb.Worksheets("Orders")
    except Exception:
        root = Tk()
        root.withdraw()
        messagebox.showerror("Invoice link", "No sheet named Orders.")
        root.destroy()
        return

    col_cat = _com_header_col(ws, "Category")
    col_inv = _com_header_col(ws, "Invoice link")
    if col_cat == 0 or col_inv == 0:
        root = Tk()
        root.withdraw()
        messagebox.showerror("Invoice link", "Missing Category or Invoice link column headers.")
        root.destroy()
        return

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    try:
        cell_txt = ws.Cells(origin_row, col_inv).Value
        txt = str(cell_txt).strip() if cell_txt is not None else ""

        if txt == "Linked":
            _remove_flow(records, origin_row, ws, col_cat)
        elif txt in ("Link to order", "Link to Gift Card"):
            _add_flow(records, origin_row, excel, ws, col_cat)
        else:
            messagebox.showinfo(
                "Invoice link",
                "This row has no link action (need Gift Card or a row with an order number).",
            )
    finally:
        root.destroy()

    try:
        wb2 = find_workbook_by_path(excel, wb_path)
        if wb2 is not None:
            sync_workbook_invoice_links(wb2, project_root=PROJECT_ROOT)
            wb2.Save()
    except Exception as ex:
        root = Tk()
        root.withdraw()
        messagebox.showwarning(
            "Invoice link",
            f"Links were saved, but Excel could not refresh automatically:\n{ex}",
        )
        root.destroy()


if __name__ == "__main__":
    main()
