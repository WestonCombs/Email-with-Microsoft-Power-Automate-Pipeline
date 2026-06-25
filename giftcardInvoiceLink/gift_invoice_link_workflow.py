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
_WAITING_BLUE = (0, 122, 255)
_CONFIRMED_GREEN = (52, 199, 89)
_DEFAULT_TOP_ORANGE = (244, 177, 131)
_ROW_RAINBOW_SECONDS = 0.45
_ROW_RAINBOW_FRAME_SECONDS = 0.075
_RAINBOW_PALETTE = (
    (255, 59, 48),
    (255, 149, 0),
    (255, 214, 10),
    (52, 199, 89),
    (0, 122, 255),
    (88, 86, 214),
    (191, 90, 242),
)

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


def _excel_rgb(red: int, green: int, blue: int) -> int:
    """Return the COLORREF integer Excel expects for Interior.Color."""
    return int(red) + (int(green) * 256) + (int(blue) * 65536)


def _last_visible_action_col(ws, header_row: int) -> int:
    try:
        used = ws.UsedRange
        last_col = int(used.Column) + int(used.Columns.Count) - 1
    except Exception:
        last_col = 1
    if last_col < 1:
        last_col = 1

    last_visible = 1
    for col in range(1, last_col + 1):
        try:
            if bool(ws.Columns(col).Hidden):
                continue
            row1 = ws.Cells(1, col).Value
            head = ws.Cells(header_row, col).Value
            if str(row1 or "").strip() or str(head or "").strip():
                last_visible = col
        except Exception:
            continue
    return max(last_visible, 1)


def _set_top_row_color(ws, header_row: int, color_rgb: tuple[int, int, int]) -> None:
    color = _excel_rgb(*color_rgb)
    last_col = _last_visible_action_col(ws, header_row)
    for col in range(1, last_col + 1):
        try:
            if bool(ws.Columns(col).Hidden):
                continue
            cell = ws.Cells(1, col)
            cell.Interior.Pattern = 1
            cell.Interior.Color = color
            cell.Interior.TintAndShade = 0
            cell.Interior.PatternTintAndShade = 0
        except Exception:
            continue


def _set_status(excel, text: str | None) -> None:
    try:
        excel.StatusBar = False if text is None else text
    except Exception:
        pass


def _row_display_color(ws, row: int, fallback_col: int) -> int:
    try:
        return int(ws.Cells(row, fallback_col).Interior.Color)
    except Exception:
        return _excel_rgb(255, 255, 255)


def _excel_col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _contiguous_spans(values: list[int]) -> list[tuple[int, int]]:
    if not values:
        return []
    spans: list[tuple[int, int]] = []
    start = prev = values[0]
    for value in values[1:]:
        if value == prev + 1:
            prev = value
            continue
        spans.append((start, prev))
        start = prev = value
    spans.append((start, prev))
    return spans


def _visible_col_spans(ws, last_col: int) -> list[tuple[int, int]]:
    visible: list[int] = []
    for col in range(1, last_col + 1):
        try:
            if not bool(ws.Columns(col).Hidden):
                visible.append(col)
        except Exception:
            continue
    return _contiguous_spans(visible)


def _cell_address(row: int, col: int) -> str:
    return f"${_excel_col_letter(col)}${row}"


def _range_address(row_start: int, row_end: int, col_start: int, col_end: int) -> str:
    first = _cell_address(row_start, col_start)
    last = _cell_address(row_end, col_end)
    return first if first == last else f"{first}:{last}"


def _range_addresses(
    row_spans: list[tuple[int, int]],
    col_spans: list[tuple[int, int]],
) -> list[str]:
    return [
        _range_address(row_start, row_end, col_start, col_end)
        for row_start, row_end in row_spans
        for col_start, col_end in col_spans
    ]


def _range_chunks(addresses: list[str], limit: int = 220) -> list[str]:
    chunks: list[str] = []
    current: list[str] = []
    current_len = 0
    for address in addresses:
        added_len = len(address) + (1 if current else 0)
        if current and current_len + added_len > limit:
            chunks.append(",".join(current))
            current = [address]
            current_len = len(address)
        else:
            current.append(address)
            current_len += added_len
    if current:
        chunks.append(",".join(current))
    return chunks


def _apply_interior_to_addresses(
    ws,
    addresses: list[str],
    *,
    pattern: int,
    color: int,
    tint: float = 0.0,
    pattern_tint: float = 0.0,
) -> None:
    for chunk in _range_chunks(addresses):
        try:
            interior = ws.Range(chunk).Interior
            interior.Pattern = pattern
            interior.Color = color
            interior.TintAndShade = tint
            interior.PatternTintAndShade = pattern_tint
        except Exception:
            continue


def _snapshot_rows(
    ws,
    rows: list[int],
    visible_col_spans: list[tuple[int, int]],
) -> dict[tuple[int, int, float, float], list[str]]:
    saved: dict[tuple[int, int, float, float], list[str]] = {}
    for row in rows:
        for col_start, col_end in visible_col_spans:
            for col in range(col_start, col_end + 1):
                try:
                    interior = ws.Cells(row, col).Interior
                    style = (
                        int(interior.Pattern),
                        int(interior.Color),
                        float(interior.TintAndShade),
                        float(interior.PatternTintAndShade),
                    )
                    saved.setdefault(style, []).append(_cell_address(row, col))
                except Exception:
                    continue
    return saved


def _restore_row_snapshot(
    ws,
    saved: dict[tuple[int, int, float, float], list[str]],
) -> None:
    for (pattern, color, tint, pattern_tint), addresses in saved.items():
        _apply_interior_to_addresses(
            ws,
            addresses,
            pattern=pattern,
            color=color,
            tint=tint,
            pattern_tint=pattern_tint,
        )


def _rows_for_link_flash(
    records: list[dict],
    gidx: int,
    xidx: int,
    data_start_row: int,
    linked_order_number: str,
) -> list[int]:
    rows = {data_start_row + gidx, data_start_row + xidx}
    order_nums = {linked_order_number}
    gift_order = normalized_order_number(records[gidx])
    if gift_order:
        order_nums.add(gift_order)
    for i, record in enumerate(records):
        if normalized_order_number(record) in order_nums:
            rows.add(data_start_row + i)
    return sorted(rows)


def _rows_for_unlink_flash(
    records: list[dict],
    gift_indices: list[int],
    order_numbers: list[str],
    data_start_row: int,
) -> list[int]:
    rows = {data_start_row + idx for idx in gift_indices if 0 <= idx < len(records)}
    order_set = {
        str(order_number).strip()
        for order_number in order_numbers
        if str(order_number).strip()
    }
    for i, record in enumerate(records):
        if normalized_order_number(record) in order_set:
            rows.add(data_start_row + i)
    return sorted(rows)


def _flash_rows_rainbow(
    excel,
    ws,
    rows: list[int],
    header_row: int,
    fallback_col: int,
) -> None:
    if not rows:
        return
    last_col = _last_visible_action_col(ws, header_row)
    visible_col_spans = _visible_col_spans(ws, last_col)
    flash_addresses = _range_addresses(_contiguous_spans(rows), visible_col_spans)
    saved = _snapshot_rows(ws, rows, visible_col_spans)
    started = time.time()
    frame = 0
    try:
        while time.time() - started < _ROW_RAINBOW_SECONDS:
            color = _excel_rgb(*_RAINBOW_PALETTE[frame % len(_RAINBOW_PALETTE)])
            _apply_interior_to_addresses(
                ws,
                flash_addresses,
                pattern=1,
                color=color,
            )
            try:
                excel.ScreenUpdating = True
            except Exception:
                pass
            time.sleep(_ROW_RAINBOW_FRAME_SECONDS)
            frame += 1
    finally:
        if saved:
            _restore_row_snapshot(ws, saved)
        else:
            _apply_interior_to_addresses(
                ws,
                flash_addresses,
                pattern=1,
                color=_row_display_color(ws, rows[0], fallback_col),
            )


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
    header_row: int,
    excel,
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
        affected_rows = _rows_for_unlink_flash(records, [idx], nums, data_start_row)
        new_edges = remove_all_edges_for_gift(edges, key)
        save_edges(link_path, new_edges)
        _set_top_row_color(ws, header_row, _CONFIRMED_GREEN)
        _set_status(excel, "Invoice Link removed.")
        _flash_rows_rainbow(excel, ws, affected_rows, header_row, col_cat)
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
    gift_indices = [
        gi
        for e in rel
        if (gi := index_for_key(records, e.gift_key)) is not None
    ]
    affected_rows = _rows_for_unlink_flash(records, gift_indices, [on], data_start_row)
    new_edges = remove_all_edges_for_order_number(edges, on)
    save_edges(link_path, new_edges)
    _set_top_row_color(ws, header_row, _CONFIRMED_GREEN)
    _set_status(excel, "Invoice Link removed.")
    _flash_rows_rainbow(excel, ws, affected_rows, header_row, col_cat)


def _add_flow(
    records: list[dict],
    origin_row: int,
    data_start_row: int,
    header_row: int,
    excel,
    ws,
    col_cat: int,
) -> None:
    oidx = _record_index_for_row(records, origin_row, data_start_row)
    if oidx is None:
        messagebox.showerror("Invoice link", "That row is outside the data range.")
        return
    ocat = _category(ws, origin_row, col_cat)

    _set_top_row_color(ws, header_row, _WAITING_BLUE)
    _set_status(
        excel,
        "Invoice Link: waiting for the matching gift card/order row. Click the row to link.",
    )
    target_row = _wait_different_row(excel, origin_row)
    if target_row is None:
        _set_top_row_color(ws, header_row, _DEFAULT_TOP_ORANGE)
        _set_status(excel, None)
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
        _set_top_row_color(ws, header_row, _DEFAULT_TOP_ORANGE)
        _set_status(excel, None)
        return

    link_path = links_path_for_project_root(PROJECT_ROOT)
    edges = load_edges(link_path, records)
    edges = add_edge(edges, g_key, order_on)
    save_edges(link_path, edges)
    _set_top_row_color(ws, header_row, _CONFIRMED_GREEN)
    _set_status(excel, "Invoice Link confirmed.")
    _flash_rows_rainbow(
        excel,
        ws,
        _rows_for_link_flash(records, gidx, xidx, data_start_row, order_on),
        header_row,
        col_cat,
    )


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
            _remove_flow(records, origin_row, data_start_row, header_row, excel, ws, col_cat)
        elif txt in ("Link to order", "Link to Gift Card"):
            _add_flow(records, origin_row, data_start_row, header_row, excel, ws, col_cat)
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
    finally:
        try:
            _set_top_row_color(ws, header_row, _DEFAULT_TOP_ORANGE)
            _set_status(excel, None)
        except Exception:
            pass


if __name__ == "__main__":
    main()
