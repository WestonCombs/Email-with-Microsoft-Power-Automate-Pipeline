"""
Grid of tracking numbers from Excel (JSON ``tracking_numbers``).

**web**: double-click a row to open that carrier's public tracking page in the browser.
**order**: one row per distinct tracking ID with **Frequency** = occurrences across all
lines in the order block (concatenated from Excel).

Rows are **color-coded** when Excel writes ``number<TAB>0|1`` per line:
**Green** — ID also appears on a URL classified as a shipment-tracking link (same heuristics as
“View Tracking Links”). **Yellow** — found only from body/regex/LLM (or legacy files without flags).

Usage:
  python tracking_numbers_viewer.py <path_to_numbers.txt> web
  python tracking_numbers_viewer.py <path_to_numbers.txt> order
"""
from __future__ import annotations

import argparse
import sys
import webbrowser
from collections import Counter
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, ttk

_PYTHON_FILES_DIR = Path(__file__).resolve().parent.parent
if str(_PYTHON_FILES_DIR) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES_DIR))

from shared.gui_aux_singleton import detach_console_win32, register_current_aux_gui
from shared.gui_treeview_copy import bind_treeview_copy_menu
from htmlHandler.carrier_urls import infer_carrier, public_tracking_url
from shared.ui_dark_theme import UI_BG, UI_FG, UI_FG_DIM, dark_tk_scrollbar, setup_dark_theme

# Green = also on classified tracking URL; yellow = text/LLM only or unknown
_ROW_LINK_OK = "link_ok"
_ROW_TEXT_ONLY = "text_only"


def _parse_line(line: str) -> tuple[str, bool | None]:
    line = line.strip()
    if not line:
        return "", None
    conf: bool | None = None
    if "\t" in line:
        num, _, rest = line.partition("\t")
        num = num.strip()
        f = rest.strip()
        if f == "1":
            conf = True
        elif f == "0":
            conf = False
    else:
        num = line
    return num, conf


def _load_tracking_file(path: Path) -> tuple[list[str], list[bool | None]]:
    """Return tracking numbers and optional link-cross-check flags (True/False/None per row). Deduped."""
    raw = path.read_text(encoding="utf-8-sig")
    numbers: list[str] = []
    confirmed: list[bool | None] = []
    seen: set[str] = set()
    for line in raw.splitlines():
        num, conf = _parse_line(line)
        if not num or num in seen:
            continue
        seen.add(num)
        numbers.append(num)
        confirmed.append(conf)
    return numbers, confirmed


def _load_tracking_lines_allow_dupes(path: Path) -> list[tuple[str, bool | None]]:
    """Each non-empty line becomes one entry; duplicates allowed (for order aggregate)."""
    raw = path.read_text(encoding="utf-8-sig")
    out: list[tuple[str, bool | None]] = []
    for line in raw.splitlines():
        num, conf = _parse_line(line)
        if num:
            out.append((num, conf))
    return out


def _aggregate_counts_and_link(
    lines: list[tuple[str, bool | None]],
) -> tuple[dict[str, int], dict[str, bool]]:
    counts: Counter[str] = Counter()
    link_any: dict[str, bool] = {}
    for num, conf in lines:
        counts[num] += 1
        ok = conf is True
        link_any[num] = link_any.get(num, False) or ok
    return dict(counts), link_any


def _load_context_tsv(path: Path) -> dict[str, str]:
    if not path.is_file():
        return {}
    out: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or "\t" not in line:
            continue
        key, _, val = line.partition("\t")
        k = key.strip()
        if k:
            out[k] = val.strip()
    return out


class TrackingNumbersViewerApp:
    def __init__(
        self,
        numbers: list[str],
        link_confirmed: list[bool | None],
        context: dict[str, str],
    ) -> None:
        self._numbers = list(numbers)
        self._link_confirmed = list(link_confirmed)
        while len(self._link_confirmed) < len(self._numbers):
            self._link_confirmed.append(None)
        self._context = dict(context)

        self._root = tk.Tk()
        self._root.title("Tracking Numbers")
        self._root.minsize(720, 380)
        self._root.geometry("920x500")
        setup_dark_theme(self._root)

        self._frm = ttk.Frame(self._root, padding=8)
        self._frm.pack(fill=tk.BOTH, expand=True)

        bits = []
        for k in ("company", "order_number", "category", "purchase_datetime"):
            v = self._context.get(k)
            if v:
                bits.append(f"{k.replace('_', ' ').title()}: {v}")
        tn = self._context.get("tracking_numbers")
        if tn:
            bits.append(f"Tracking Numbers: {tn}")
        ctx_line = "  |  ".join(bits) if bits else ""

        if ctx_line:
            ttk.Label(self._frm, text=ctx_line, wraplength=880, foreground=UI_FG_DIM).pack(
                fill=tk.X, anchor=tk.W, pady=(0, 6)
            )

        legend = (
            "Row colors:  "
            "Green = also found on a classified tracking URL (same pipeline as “View Tracking Links”).  "
            "Yellow = from body/regex/LLM only, or not matched to a tracking link — treat as less certain."
        )
        ttk.Label(self._frm, text=legend, wraplength=880, foreground=UI_FG_DIM).pack(
            fill=tk.X, anchor=tk.W, pady=(0, 6)
        )

        hint = "Double-click a row (or press Enter): open the carrier tracking page in your browser."
        ttk.Label(self._frm, text=hint, wraplength=880, foreground=UI_FG).pack(
            fill=tk.X, anchor=tk.W, pady=(0, 4)
        )

        tree_outer = tk.Frame(self._frm, bg=UI_BG)
        tree_outer.pack(fill=tk.BOTH, expand=True)

        cols = ("idx", "tracking_number", "carrier", "source", "note")
        self._tree = ttk.Treeview(
            tree_outer,
            columns=cols,
            show="headings",
            selectmode="browse",
            height=min(18, max(6, len(self._numbers))),
        )
        self._tree.heading("idx", text="#")
        self._tree.heading("tracking_number", text="Tracking Number")
        self._tree.heading("carrier", text="Carrier (Guess)")
        self._tree.heading("source", text="Link Match")
        self._tree.heading("note", text="Opens")

        self._tree.column("idx", width=40, anchor=tk.CENTER)
        self._tree.column("tracking_number", width=220, anchor=tk.W)
        self._tree.column("carrier", width=88, anchor=tk.CENTER)
        self._tree.column("source", width=120, anchor=tk.CENTER)
        self._tree.column("note", width=360, anchor=tk.W)

        self._configure_tree_style()
        self._tree.tag_configure(_ROW_LINK_OK, background="#166534", foreground="#ecfdf5")
        self._tree.tag_configure(_ROW_TEXT_ONLY, background="#a16207", foreground="#fefce8")

        scroll = dark_tk_scrollbar(tree_outer, tk.VERTICAL, self._tree.yview)
        self._tree.configure(yscrollcommand=scroll.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        for i, num in enumerate(self._numbers, start=1):
            car = infer_carrier(num)
            lc = self._link_confirmed[i - 1] if i - 1 < len(self._link_confirmed) else None
            if lc is True:
                src = "Link + Text"
                tag = _ROW_LINK_OK
            else:
                src = "Text / LLM Only" if lc is False else "Unknown (Old File)"
                tag = _ROW_TEXT_ONLY
            url = public_tracking_url(num)
            note = (url[:72] + "…") if len(url) > 72 else url
            self._tree.insert(
                "",
                tk.END,
                iid=str(i - 1),
                values=(i, num, car, src, note),
                tags=(tag,),
            )

        self._tree.bind("<Double-1>", self._on_double)
        self._tree.bind("<Return>", self._on_return)
        bind_treeview_copy_menu(self._tree, self._root)

        btn_row = ttk.Frame(self._frm)
        btn_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_row, text="Close", command=self._root.destroy).pack(side=tk.RIGHT)

        if not self._numbers:
            ttk.Label(self._frm, text="No tracking numbers in this file.", foreground=UI_FG_DIM).pack()

    def _configure_tree_style(self) -> None:
        setup_dark_theme(self._root)

    def _row_index(self, iid: str) -> int | None:
        try:
            return int(iid)
        except ValueError:
            return None

    def _on_double(self, _event: tk.Event) -> None:
        self._activate()

    def _on_return(self, _event: tk.Event) -> None:
        self._activate()

    def _activate(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        idx = self._row_index(sel[0])
        if idx is None or idx < 0 or idx >= len(self._numbers):
            return
        num = self._numbers[idx]
        url = public_tracking_url(num)
        if not url:
            messagebox.showerror("Tracking", "Could not build a URL for this number.")
            return
        webbrowser.open(url)

    def run(self) -> None:
        self._root.mainloop()


class TrackingNumbersOrderViewerApp:
    """One row per distinct tracking ID; Frequency = occurrences across the order block file."""

    def __init__(
        self,
        counts: dict[str, int],
        link_any: dict[str, bool],
        context: dict[str, str],
    ) -> None:
        self._counts = counts
        self._link_any = link_any
        self._context = dict(context)
        nums_sorted = sorted(counts.keys(), key=lambda n: (-counts[n], n))
        self._numbers = nums_sorted

        self._root = tk.Tk()
        self._root.title("Tracking Numbers — All For Order")
        self._root.minsize(760, 400)
        self._root.geometry("960x520")
        setup_dark_theme(self._root)

        self._frm = ttk.Frame(self._root, padding=8)
        self._frm.pack(fill=tk.BOTH, expand=True)

        bits = []
        for k in ("company", "order_number", "category", "purchase_datetime"):
            v = self._context.get(k)
            if v:
                bits.append(f"{k.replace('_', ' ').title()}: {v}")
        tn = self._context.get("tracking_numbers")
        if tn:
            bits.append(f"Tracking Numbers (This Row): {tn}")
        ctx_line = "  |  ".join(bits) if bits else ""

        if ctx_line:
            ttk.Label(self._frm, text=ctx_line, wraplength=900, foreground=UI_FG_DIM).pack(
                fill=tk.X, anchor=tk.W, pady=(0, 6)
            )

        ttk.Label(
            self._frm,
            text=(
                "Frequency counts how many row-level tracking slots listed each ID across "
                "all lines merged for this order. One row per distinct ID."
            ),
            wraplength=900,
            foreground=UI_FG_DIM,
        ).pack(fill=tk.X, anchor=tk.W, pady=(0, 6))

        legend = (
            "Row colors:  "
            "Green = at least one line had link cross-check.  "
            "Yellow = text/LLM only on every line for that ID."
        )
        ttk.Label(self._frm, text=legend, wraplength=900, foreground=UI_FG_DIM).pack(
            fill=tk.X, anchor=tk.W, pady=(0, 6)
        )

        hint = "Double-click a row (or press Enter): open the carrier tracking page in your browser."
        ttk.Label(self._frm, text=hint, wraplength=900, foreground=UI_FG).pack(
            fill=tk.X, anchor=tk.W, pady=(0, 4)
        )

        tree_outer = tk.Frame(self._frm, bg=UI_BG)
        tree_outer.pack(fill=tk.BOTH, expand=True)

        cols = ("idx", "tracking_number", "carrier", "frequency", "source", "note")
        self._tree = ttk.Treeview(
            tree_outer,
            columns=cols,
            show="headings",
            selectmode="browse",
            height=min(18, max(6, len(self._numbers))),
        )
        self._tree.heading("idx", text="#")
        self._tree.heading("tracking_number", text="Tracking Number")
        self._tree.heading("carrier", text="Carrier (Guess)")
        self._tree.heading("frequency", text="Frequency")
        self._tree.heading("source", text="Link Match")
        self._tree.heading("note", text="Opens")

        self._tree.column("idx", width=40, anchor=tk.CENTER)
        self._tree.column("tracking_number", width=200, anchor=tk.W)
        self._tree.column("carrier", width=88, anchor=tk.CENTER)
        self._tree.column("frequency", width=72, anchor=tk.CENTER)
        self._tree.column("source", width=120, anchor=tk.CENTER)
        self._tree.column("note", width=320, anchor=tk.W)

        self._tree.tag_configure(_ROW_LINK_OK, background="#166534", foreground="#ecfdf5")
        self._tree.tag_configure(_ROW_TEXT_ONLY, background="#a16207", foreground="#fefce8")

        scroll = dark_tk_scrollbar(tree_outer, tk.VERTICAL, self._tree.yview)
        self._tree.configure(yscrollcommand=scroll.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        for i, num in enumerate(self._numbers, start=1):
            car = infer_carrier(num)
            freq = counts[num]
            lc = link_any.get(num, False)
            if lc:
                src = "Link + Text"
                tag = _ROW_LINK_OK
            else:
                src = "Text / LLM Only"
                tag = _ROW_TEXT_ONLY
            url = public_tracking_url(num)
            note = (url[:64] + "…") if len(url) > 64 else url
            self._tree.insert(
                "",
                tk.END,
                iid=str(i - 1),
                values=(i, num, car, freq, src, note),
                tags=(tag,),
            )

        self._tree.bind("<Double-1>", self._on_double)
        self._tree.bind("<Return>", self._on_return)
        bind_treeview_copy_menu(self._tree, self._root)

        btn_row = ttk.Frame(self._frm)
        btn_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_row, text="Close", command=self._root.destroy).pack(side=tk.RIGHT)

        if not self._numbers:
            ttk.Label(self._frm, text="No tracking numbers in this file.", foreground=UI_FG_DIM).pack()

    def _row_index(self, iid: str) -> int | None:
        try:
            return int(iid)
        except ValueError:
            return None

    def _on_double(self, _event: tk.Event) -> None:
        self._activate()

    def _on_return(self, _event: tk.Event) -> None:
        self._activate()

    def _activate(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        idx = self._row_index(sel[0])
        if idx is None or idx < 0 or idx >= len(self._numbers):
            return
        num = self._numbers[idx]
        url = public_tracking_url(num)
        if not url:
            messagebox.showerror("Tracking", "Could not build a URL for this number.")
            return
        webbrowser.open(url)

    def run(self) -> None:
        self._root.mainloop()


def main() -> int:
    detach_console_win32()
    register_current_aux_gui()

    try:
        from shared.settings_store import apply_runtime_settings_from_json

        apply_runtime_settings_from_json()
    except Exception:
        pass

    parser = argparse.ArgumentParser(description="View tracking numbers (carrier web URLs).")
    parser.add_argument(
        "number_file",
        nargs="?",
        type=Path,
        help="UTF-8 file: one tracking number per line, or number<TAB>0|1 for link cross-check",
    )
    parser.add_argument(
        "mode",
        nargs="?",
        default="web",
        choices=("web", "order"),
        help="web = per-row deduped list; order = aggregate counts (duplicate lines allowed)",
    )
    args = parser.parse_args()

    if args.number_file is None:
        app = TrackingNumbersViewerApp([], [], {})
        app.run()
        return 0

    p = args.number_file.expanduser().resolve()
    if not p.is_file():
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Tracking Numbers", f"File not found:\n{p}")
        root.destroy()
        return 1

    try:
        context = _load_context_tsv(p.with_suffix(".ctx.tsv"))
        if args.mode == "order":
            lines = _load_tracking_lines_allow_dupes(p)
            counts, link_any = _aggregate_counts_and_link(lines)
            app = TrackingNumbersOrderViewerApp(counts, link_any, context)
        else:
            numbers, link_confirmed = _load_tracking_file(p)
            app = TrackingNumbersViewerApp(numbers, link_confirmed, context)
    except OSError as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Tracking Numbers", f"Could not read file:\n{e}")
        root.destroy()
        return 1

    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
