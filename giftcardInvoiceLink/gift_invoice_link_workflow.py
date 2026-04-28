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

_PYTHON_FILES = Path(__file__).resolve().parent.parent
if str(_PYTHON_FILES) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES))

from shared.settings_store import apply_runtime_settings_from_json

apply_runtime_settings_from_json()

import os

from shared.gui_aux_singleton import detach_console_win32, register_current_aux_gui
from shared.project_paths import ensure_base_dir_in_environ

PROJECT_ROOT = ensure_base_dir_in_environ()
JSON_PATH = PROJECT_ROOT / "email_contents" / "json" / "results.json"

_REFRESH_ATTEMPTS = 3
_REFRESH_BACKOFF_SEC = 0.35

_MSG_REFRESH_FAILED = (
    "Your link changes were saved on disk, but Excel could not finish refreshing "
    "every Invoice link cell.\n\n"
    "Wait until processing finishes - do not click other rows or Invoice links while "
    "Excel updates. Then try again, or save and reopen the workbook.\n\n"
    "Technical detail:\n"
)

from tkinter import Tk, messagebox

from giftcardInvoiceLink.excel_link_sync import (
    find_header_columns,
    find_workbook_and_application,
    find_workbook_by_path,
    sync_workbook_invoice_links,
)
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
from proofOfDelivery.pod_data import load_excel_records


def _refresh_invoice_links_with_retries(excel, wb_path: str) -> None:
    """Run COM refresh + Save with backoff (recovers from races with UI interaction)."""
    last: BaseException | None = None
    for attempt in range(_REFRESH_ATTEMPTS):
        try:
            wb2 = find_workbook_by_path(excel, wb_path)
            if wb2 is None:
                raise RuntimeError(
                    "The workbook was closed or is no longer the same file you clicked from."
                )
            sync_workbook_invoice_links(wb2, project_root=PROJECT_ROOT)
            wb2.Save()
            return
        except Exception as e:
            last = e
            time.sleep(_REFRESH_BACKOFF_SEC * (attempt + 1))
    assert last is not None
    raise last


def _load_records() -> list[dict]:
    return load_excel_records(PROJECT_ROOT, include_automation_hub=False, sync_pod_json=False)


def _category(ws, row: int, col_cat: int) -> str:
    v = ws.Cells(row, col_cat).Value
    if v is None:
        return ""
    return str(v).strip()


def _record_index_for_row(records: list[dict], row: int, data_start_row: int) -> int | None:
    idx = row - data_start_row
    if idx < 0 or idx >= len(records):
        return None
    return idx


def _order_summary(records: list[dict], idx: int, data_start_row: int | None = None) -> str:
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
    display_row = idx + data_start_row if data_start_row is not None else idx + 2
    return ", ".join(parts) if parts else f"row {display_row}"


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
    data_start_row: int,
    ws,
    col_cat: int,
) -> None:
    idx = _record_index_for_row(records, origin_row, data_start_row)
    if idx is None:
        messagebox.showerror("Invoice link", "That row is outside the data range.")
        return
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
            gift_lines.append(_order_summary(records, gi, data_start_row))
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
    data_start_row: int,
    excel,
    ws,
    col_cat: int,
) -> None:
    oidx = _record_index_for_row(records, origin_row, data_start_row)
    if oidx is None:
        messagebox.showerror("Invoice link", "That row is outside the data range.")
        return
    ocat = _category(ws, origin_row, col_cat)

    target_row = _wait_different_row(excel, origin_row)
    if target_row is None:
        messagebox.showwarning("Invoice link", "Timed out waiting for a new row selection.")
        return
    tidx = _record_index_for_row(records, target_row, data_start_row)
    if tidx is None:
        messagebox.showerror("Invoice link", "Selected row is outside the data range.")
        return

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
        f"Gift card: {_order_summary(records, gidx, data_start_row)}\n"
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

    import win32com.client

    wb, excel = find_workbook_and_application(wb_path)
    if wb is None:
        root = Tk()
        root.withdraw()
        try:
            win32com.client.GetActiveObject("Excel.Application")
            messagebox.showerror(
                "Invoice link",
                "Excel is open, but this script could not match your orders file to any "
                "open workbook.\n\n"
                "Close all Excel windows, open only your orders workbook, then try the "
                "Invoice link again. Using several Excel windows at once can cause this.",
            )
        except Exception:
            messagebox.showerror("Invoice link", "Excel is not running.")
        root.destroy()
        return

    if excel is None:
        try:
            excel = wb.Application
        except Exception:
            root = Tk()
            root.withdraw()
            messagebox.showerror(
                "Invoice link",
                "Found the workbook but could not attach to Excel. Close other Office apps "
                "and try again.",
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

    header_row, header_cols = find_header_columns(ws, "Category", "Invoice link")
    col_cat = header_cols.get("Category", 0)
    col_inv = header_cols.get("Invoice link", 0)
    if col_cat == 0 or col_inv == 0:
        root = Tk()
        root.withdraw()
        messagebox.showerror("Invoice link", "Missing Category or Invoice link column headers.")
        root.destroy()
        return
    data_start_row = header_row + 1

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    try:
        cell_txt = ws.Cells(origin_row, col_inv).Value
        txt = str(cell_txt).strip() if cell_txt is not None else ""

        if txt == "Linked":
            _remove_flow(records, origin_row, data_start_row, ws, col_cat)
        elif txt in ("Link to order", "Link to Gift Card"):
            _add_flow(records, origin_row, data_start_row, excel, ws, col_cat)
        else:
            messagebox.showinfo(
                "Invoice link",
                "This row has no link action (need Gift Card or a row with an order number).",
            )
    finally:
        root.destroy()

    try:
        _refresh_invoice_links_with_retries(excel, wb_path)
        root = Tk()
        root.withdraw()
        messagebox.showinfo(
            "Invoice link",
            "Invoice Link column is now up to date, whether you made a change or not.",
        )
        root.destroy()
    except Exception as ex:
        root = Tk()
        root.withdraw()
        messagebox.showwarning("Invoice link", _MSG_REFRESH_FAILED + str(ex))
        root.destroy()


if __name__ == "__main__":
    main()
