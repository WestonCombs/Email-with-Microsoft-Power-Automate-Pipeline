"""Right-click context menu on ``ttk.Treeview`` cells: copy the clicked cell to the clipboard."""

from __future__ import annotations

import sys
import tkinter as tk
from tkinter import ttk


def treeview_cell_text(tree: ttk.Treeview, item: str, col_id: str) -> str:
    """Return display text for *item* at Treeview column *col_id* (``#0``, ``#1``, …)."""
    if not item or not col_id:
        return ""
    try:
        n = int(col_id[1:], 10)
    except (ValueError, IndexError):
        return ""
    if n == 0:
        return str(tree.item(item, "text") or "")
    vals = tree.item(item, "values")
    i = n - 1
    if 0 <= i < len(vals):
        return str(vals[i])
    return ""


def treeview_row_text_tsv(tree: ttk.Treeview, item: str) -> str:
    """Tree column text + all value columns, tab-separated (for optional “copy row”)."""
    if not item:
        return ""
    parts: list[str] = []
    t0 = str(tree.item(item, "text") or "").strip()
    if t0:
        parts.append(t0)
    vals = tree.item(item, "values")
    parts.extend(str(v) for v in vals)
    return "\t".join(parts)


def bind_treeview_copy_menu(
    tree: ttk.Treeview,
    toplevel: tk.Misc,
    extra_commands: list[tuple[str, object]] | None = None,
) -> None:
    """Bind right-click on a data cell or tree column: menu entries **Copy cell** and **Copy row**."""
    menu = tk.Menu(toplevel, tearoff=0)
    state: dict[str, str | None] = {"item": None, "col": None}

    def _clipboard_set(text: str) -> None:
        try:
            toplevel.clipboard_clear()
            toplevel.clipboard_append(text)
            toplevel.update_idletasks()
        except tk.TclError:
            pass

    def copy_cell() -> None:
        item = state["item"]
        col = state["col"]
        if not item or not col:
            return
        _clipboard_set(treeview_cell_text(tree, item, col))

    def copy_row() -> None:
        item = state["item"]
        if not item:
            return
        _clipboard_set(treeview_row_text_tsv(tree, item))

    menu.add_command(label="Copy cell", command=copy_cell)
    menu.add_command(label="Copy row", command=copy_row)
    if extra_commands:
        menu.add_separator()
        for label, callback in extra_commands:
            menu.add_command(label=label, command=callback)

    def on_button(event: tk.Event) -> None:
        region = tree.identify_region(event.x, event.y)
        if region not in ("cell", "tree"):
            return
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or not col:
            return
        try:
            tree.selection_set(item)
            tree.focus(item)
        except tk.TclError:
            pass
        state["item"] = item
        state["col"] = col
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    tree.bind("<Button-3>", on_button)
    if sys.platform == "darwin":
        tree.bind("<Button-2>", on_button)
        tree.bind("<Control-Button-1>", on_button)
