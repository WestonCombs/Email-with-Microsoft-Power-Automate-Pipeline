"""
Grid view of tracking URLs (#, Item, Description, Delivered/Cancelled, full link).

- **Double-click anywhere on a row** (any column, including #) to open that row's tracking URL
  in your default browser. Press Enter with a row selected for the same action.
- **Opened** shows ✓ after a link has been opened successfully (not toggled by clicking the cell).
  **Delivered** / **Cancelled** are per-row, mutually exclusive; state is stored under
  ``email_contents/tracking_link_viewer_state/``.
- Optional row context: Excel writes ``<same_base>.ctx.tsv`` next to the URL list; shown in the window.

Links in the workbook come from JSON ``tracking_links``, filled by
``grabbingImportantEmailContent`` using ``htmlHandler.tracking_hrefs`` (``href_final_pairs`` →
``list_tracking_links_from_pairs``) after anchor extraction and redirect resolution.

Launched from Excel via **View tracking links**: VBA writes a temp UTF-8 ``.txt`` (one URL per line)
and an optional ``.ctx.tsv``. State is keyed by URL list + order context (not the temp path), so it
survives each launch.

Usage:
  python tracking_link_viewer.py <path_to_urls.txt>

Standalone (empty grid):
  python tracking_link_viewer.py
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import webbrowser
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, ttk

_PYTHON_FILES_DIR = Path(__file__).resolve().parent.parent

if str(_PYTHON_FILES_DIR) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES_DIR))

from shared.gui_aux_singleton import detach_console_win32, register_current_aux_gui
from shared.gui_treeview_copy import bind_treeview_copy_menu

# Per-viewer JSON files: ``BASE_DIR/email_contents/tracking_link_viewer_state/<sha256>.json``
_TRACKING_STATE_SUBDIR = "tracking_link_viewer_state"
_STATE_VERSION = 1

# Dark blue / grey theme (backgrounds + light text); blues slightly light for readability
_UI_BG = "#1e293b"
_UI_BG_PANEL = "#0f172a"
_UI_TREE_BG = "#334155"
_UI_TREE_HEAD = "#60a5fa"
# Hover/pressed: clam’s default can use a near-white active bg while fg stays light (unreadable).
_UI_TREE_HEAD_ACTIVE = "#2563eb"
_UI_FG = "#f8fafc"
_UI_FG_DIM = "#94a3b8"
_UI_SEL = "#3b82f6"
_UI_BTN = "#3b82f6"
_UI_BTN_ACTIVE = "#2563eb"
# Classic tk.Scrollbar (ttk scrollbars stay Windows-themed on many setups)
_UI_SCROLL_TROUGH = "#1e293b"
_UI_SCROLL_THUMB = "#64748b"
_UI_SCROLL_THUMB_ACTIVE = "#94a3b8"

# Leading gunk on the first line of temp files (VBA/Excel UTF-8) breaks urlparse if not stripped.
_URL_LEADING_INVISIBLE = "\ufeff\u200b\u200c\u200d\u2060"


def _strip_leading_url_gunk(s: str) -> str:
    s = (s or "").strip().lstrip(_URL_LEADING_INVISIBLE).strip()
    return s.replace("\n", " ").replace("\r", "")


def _normalize_url_for_browser(raw: str) -> str | None:
    """Return an absolute http(s) URL or None if *raw* is not openable (e.g. a bare token)."""
    u = _strip_leading_url_gunk(raw)
    if not u:
        return None
    low = u.lower()
    if low.startswith("http://") or low.startswith("https://"):
        candidate = u
    elif u.startswith("//"):
        candidate = "https:" + u
    elif "://" not in u:
        candidate = "https://" + u.lstrip("/")
    else:
        candidate = u
    try:
        p = urlparse(candidate)
    except Exception:
        return None
    if p.scheme not in ("http", "https"):
        return None
    host = (p.netloc or "").split("@")[-1].split(":")[0].strip().strip(".").lower()
    if not host:
        return None
    if host == "localhost" or host.startswith("[") or "." in host:
        return candidate
    octets = host.split(".")
    if len(octets) == 4 and all(o.isdigit() and 0 <= int(o) <= 255 for o in octets):
        return candidate
    return None


def _open_url(url: str) -> None:
    """Open *url* in the user's default browser.

    On Windows, ``webbrowser`` sometimes picks Edge for the first open and Chrome after;
    ``os.startfile`` uses the same HTTP handler as Explorer every time.
    """
    u = _normalize_url_for_browser(url)
    if not u:
        messagebox.showwarning(
            "Cannot open link",
            "This row is not a full web address. It should start with http:// or https:// "
            "(or look like www.example.com/…).\n\n"
            "If you only see random characters, the value may be a truncated fragment—not a "
            "complete URL. Check the original email or re-run the extraction pipeline.",
        )
        return
    if sys.platform == "win32":
        try:
            os.startfile(u)
        except OSError:
            webbrowser.open(u)
    else:
        webbrowser.open(u)


def _load_urls(path: Path) -> list[str]:
    # utf-8-sig drops a BOM on the first line so "https://..." parses (otherwise only row 1 breaks).
    raw = path.read_text(encoding="utf-8-sig")
    return [_strip_leading_url_gunk(line) for line in raw.splitlines() if line.strip()]


def _context_path_for_urls_file(urls_file: Path) -> Path:
    return urls_file.with_suffix(".ctx.tsv")


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


def _email_contents_project_root() -> Path | None:
    """Project root containing ``email_contents`` (``BASE_DIR``, set in Email Sorter → Settings)."""
    try:
        from dotenv import load_dotenv

        load_dotenv(_PYTHON_FILES_DIR / ".env", override=False)
    except ImportError:
        pass
    base = (os.getenv("BASE_DIR") or "").strip()
    if not base:
        return None
    p = Path(base).expanduser().resolve()
    return p if p.is_dir() else None


def _tracking_state_dir() -> Path | None:
    root = _email_contents_project_root()
    if root is None:
        return None
    d = root / "email_contents" / _TRACKING_STATE_SUBDIR
    try:
        d.mkdir(parents=True, exist_ok=True)
    except OSError:
        return None
    return d


def _context_tracking_numbers_value(context: dict[str, str]) -> str:
    raw = context.get("tracking_numbers")
    if raw is None:
        raw = context.get("tracking_number")
    if isinstance(raw, list):
        return ", ".join(str(x).strip() for x in raw if str(x).strip())
    return (raw or "").strip() if isinstance(raw, str) else ""


def _state_fingerprint(urls: list[str], context: dict[str, str]) -> str:
    normalized = [_strip_leading_url_gunk(u) for u in urls]
    payload = {
        "company": (context.get("company") or "").strip(),
        "order_number": (context.get("order_number") or "").strip(),
        "tracking_numbers": _context_tracking_numbers_value(context),
        "urls": normalized,
    }
    raw = json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def _state_path_for_links(urls: list[str], context: dict[str, str]) -> Path | None:
    if not urls:
        return None
    sd = _tracking_state_dir()
    if sd is None:
        return None
    return sd / f"{_state_fingerprint(urls, context)}.json"


def _load_row_state(
    path: Path | None,
    n: int,
) -> tuple[list[bool], list[bool], list[bool]]:
    opened = [False] * n
    delivered = [False] * n
    cancelled = [False] * n
    if path is None or not path.is_file():
        return opened, delivered, cancelled
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return opened, delivered, cancelled

    def _bool_list(key: str) -> list[bool]:
        arr = data.get(key)
        if not isinstance(arr, list):
            return [False] * n
        return [bool(arr[i]) if i < len(arr) else False for i in range(n)]

    opened = _bool_list("opened")
    delivered = _bool_list("delivered")
    cancelled = _bool_list("cancelled")
    for i in range(n):
        if delivered[i] and cancelled[i]:
            cancelled[i] = False
    return opened, delivered, cancelled


def _save_row_state(
    path: Path | None,
    opened: list[bool],
    delivered: list[bool],
    cancelled: list[bool],
) -> None:
    if path is None:
        return
    for i in range(len(delivered)):
        if delivered[i] and cancelled[i]:
            cancelled[i] = False
    payload = {
        "version": _STATE_VERSION,
        "opened": opened,
        "delivered": delivered,
        "cancelled": cancelled,
    }
    try:
        path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def _link_column_width_px(master: tk.Misc, urls: list[str]) -> int:
    """Pixel width for the link column from the longest URL so horizontal scroll can show it all."""
    if not urls:
        return 380
    f = tkfont.nametofont("TkDefaultFont")
    try:
        longest = max(urls, key=len)
        m = f.measure(longest.replace("\n", " ").replace("\r", ""))
    except (ValueError, tk.TclError):
        m = 380
    # Cell chrome / padding; cap avoids absurd values from pathological strings.
    return min(max(m + 56, 280), 8000)


def _heuristic_item_description(url: str) -> tuple[str, str]:
    """Best-effort labels without network or API calls."""
    try:
        p = urlparse(url)
        host = (p.netloc or "").lower()
        path_l = (p.path or "").lower()
        if not host:
            return "Tracking link", ""

        brand_guess = host
        if host.endswith(".narvar.com") or "narvar" in host:
            qs = parse_qs(p.query)
            retailer = (qs.get("retailer") or qs.get("r") or [None])[0]
            if retailer:
                try:
                    brand_guess = unquote(str(retailer)).replace("+", " ")
                except Exception:
                    brand_guess = str(retailer)
            item = "Shipment tracking"
            desc = f"Narvar — {brand_guess}" if brand_guess else "Narvar tracking page"
            return item, desc

        if "track" in path_l or "tracking" in path_l or "ship" in path_l:
            return "Shipment tracking", host

        if "ups.com" in host:
            return "UPS tracking", host
        if "fedex" in host:
            return "FedEx tracking", host
        if "usps" in host:
            return "USPS tracking", host
        if "dhl" in host:
            return "DHL tracking", host

        first = host.split(".")[0] if host else "link"
        return "Tracking link", first.replace("-", " ").title()
    except Exception:
        return "Tracking link", ""


def _setup_dark_theme(root: tk.Tk) -> ttk.Style:
    root.configure(bg=_UI_BG)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure("TFrame", background=_UI_BG)
    style.configure("TLabel", background=_UI_BG, foreground=_UI_FG)
    style.configure(
        "TButton",
        background=_UI_BTN,
        foreground=_UI_FG,
        borderwidth=0,
        focuscolor="none",
        padding=(14, 6),
    )
    style.map(
        "TButton",
        background=[("active", _UI_BTN_ACTIVE), ("disabled", "#475569")],
        foreground=[("disabled", _UI_FG_DIM)],
    )
    style.configure(
        "Treeview",
        background=_UI_TREE_BG,
        fieldbackground=_UI_TREE_BG,
        foreground=_UI_FG,
        borderwidth=0,
        rowheight=26,
    )
    style.configure(
        "Treeview.Heading",
        background=_UI_TREE_HEAD,
        foreground=_UI_FG,
        borderwidth=0,
        relief="flat",
        anchor="w",
    )
    style.map(
        "Treeview.Heading",
        background=[
            ("active", _UI_TREE_HEAD_ACTIVE),
            ("pressed", _UI_TREE_HEAD_ACTIVE),
        ],
        foreground=[
            ("active", _UI_FG),
            ("pressed", _UI_FG),
        ],
    )
    style.map(
        "Treeview",
        background=[("selected", _UI_SEL)],
        foreground=[("selected", _UI_FG)],
    )
    return style


def _dark_tk_scrollbar(master: tk.Misc, orient: str, command) -> tk.Scrollbar:
    return tk.Scrollbar(
        master,
        orient=orient,
        command=command,
        troughcolor=_UI_SCROLL_TROUGH,
        bg=_UI_SCROLL_THUMB,
        activebackground=_UI_SCROLL_THUMB_ACTIVE,
        highlightthickness=0,
        bd=0,
        width=14,
        elementborderwidth=1,
        jump=0,
    )


class TrackingLinkViewerApp:
    _CHECK_MARK = "\u2713"  # ✓

    def __init__(self, urls: list[str], context: dict[str, str]) -> None:
        self._urls = list(urls)
        self._context = dict(context)
        self._items = [_heuristic_item_description(u)[0] for u in self._urls]
        self._descs = [_heuristic_item_description(u)[1] for u in self._urls]
        self._state_path = _state_path_for_links(self._urls, self._context)
        o, d, c = _load_row_state(self._state_path, len(self._urls))
        self._opened = o
        self._delivered = d
        self._cancelled = c

        self._root = tk.Tk()
        self._root.title("Tracking links")
        self._root.minsize(640, 320)
        self._root.geometry("900x520")
        _setup_dark_theme(self._root)

        self._frm = ttk.Frame(self._root, padding=8)
        self._frm.pack(fill=tk.BOTH, expand=True)

        ctx_bits = []
        for k in ("company", "order_number", "category", "purchase_datetime"):
            v = self._context.get(k)
            if v:
                ctx_bits.append(f"{k.replace('_', ' ')}: {v}")
        tn_disp = _context_tracking_numbers_value(self._context)
        if tn_disp:
            ctx_bits.append(f"tracking numbers: {tn_disp}")
        ctx_line = "  |  ".join(ctx_bits) if ctx_bits else "No order context file (optional .ctx.tsv from Excel)."

        hint = ttk.Label(
            self._frm,
            text=(
                "Double-click anywhere on a row to open that row's tracking link in your browser "
                "(or select the row and press Enter).\n\n"
                "Click Delivered or Cancelled once per row (they clear each other).\n\n"
                "URLs are in the Tracking link column."
            ),
            wraplength=860,
            justify=tk.LEFT,
        )
        hint.pack(fill=tk.X, anchor=tk.W, pady=(0, 4))

        ctx_lbl = ttk.Label(self._frm, text=ctx_line, wraplength=860, foreground=_UI_FG_DIM)
        ctx_lbl.pack(fill=tk.X, pady=(0, 6))

        tree_frame = tk.Frame(self._frm, bg=_UI_BG)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        yscroll = _dark_tk_scrollbar(tree_frame, tk.VERTICAL, lambda *a: None)
        xscroll = _dark_tk_scrollbar(tree_frame, tk.HORIZONTAL, lambda *a: None)
        self._tree_scroll_y = yscroll
        self._tree_scroll_x = xscroll

        self._tree = ttk.Treeview(
            tree_frame,
            columns=("opened", "item", "description", "delivered", "cancelled", "link"),
            show="tree headings",
            yscrollcommand=yscroll.set,
            xscrollcommand=xscroll.set,
            selectmode="browse",
        )
        self._tree.grid(row=0, column=0, sticky="nsew")
        yscroll.config(command=self._tree.yview)
        xscroll.config(command=self._tree.xview)
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        self._tree.heading("#0", text="#", anchor=tk.W)
        self._tree.column("#0", width=44, stretch=False, anchor=tk.W)
        self._tree.heading("opened", text="Opened", anchor=tk.CENTER)
        self._tree.column("opened", width=56, stretch=False, minwidth=48, anchor=tk.CENTER)
        self._tree.heading("item", text="Item", anchor=tk.W)
        self._tree.column("item", width=180, stretch=False, minwidth=72, anchor=tk.W)
        self._tree.heading("description", text="Description", anchor=tk.W)
        self._tree.column("description", width=300, stretch=False, minwidth=96, anchor=tk.W)
        self._tree.heading("delivered", text="Delivered", anchor=tk.CENTER)
        self._tree.column("delivered", width=76, stretch=False, minwidth=64, anchor=tk.CENTER)
        self._tree.heading("cancelled", text="Cancelled", anchor=tk.CENTER)
        self._tree.column("cancelled", width=76, stretch=False, minwidth=64, anchor=tk.CENTER)
        self._tree.heading("link", text="Tracking link", anchor=tk.W)
        _link_w = _link_column_width_px(self._root, self._urls)
        self._tree.column("link", width=_link_w, stretch=False, minwidth=120, anchor=tk.W)

        for i, u in enumerate(self._urls):
            iid = str(i)
            self._tree.insert(
                "",
                tk.END,
                iid=iid,
                text=str(i + 1),
                values=self._row_values(i),
            )

        self._tree.bind("<Enter>", lambda _e: self._tree.focus_set())
        self._tree.bind("<Button-1>", self._on_tree_button1)
        self._tree.bind("<Double-1>", self._on_tree_double)
        self._tree.bind("<Return>", self._on_return_key)

        self._tree.bind("<MouseWheel>", self._on_mousewheel)
        self._tree.bind("<Button-4>", self._on_mousewheel_linux)
        self._tree.bind("<Button-5>", self._on_mousewheel_linux)
        bind_treeview_copy_menu(self._tree, self._root)

        btn_row = ttk.Frame(self._frm)
        btn_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_row, text="Close", command=self._root.destroy).pack(side=tk.RIGHT)

        if not urls:
            ttk.Label(
                self._frm,
                text="No tracking URLs in this file.",
                foreground=_UI_FG_DIM,
            ).pack(pady=4)

    def _row_index(self, iid: str) -> int | None:
        try:
            i = int(iid)
        except (TypeError, ValueError):
            return None
        if 0 <= i < len(self._urls):
            return i
        return None

    def _row_values(self, i: int) -> tuple[str, str, str, str, str, str]:
        m = self._CHECK_MARK
        return (
            m if self._opened[i] else "",
            self._items[i],
            self._descs[i],
            m if self._delivered[i] else "",
            m if self._cancelled[i] else "",
            self._urls[i],
        )

    def _tree_column_name_at(self, x: int) -> str | None:
        cid = self._tree.identify_column(x)
        if cid == "#0":
            return None
        try:
            idx = int(cid[1:]) - 1
        except ValueError:
            return None
        cols: tuple[str, ...] = self._tree["columns"]
        if not (0 <= idx < len(cols)):
            return None
        return cols[idx]

    def _sync_row_tree(self, idx: int) -> None:
        self._tree.item(str(idx), values=self._row_values(idx))

    def _persist_checkbox_state(self) -> None:
        _save_row_state(self._state_path, self._opened, self._delivered, self._cancelled)

    def _on_tree_button1(self, event: tk.Event) -> None:
        region = self._tree.identify_region(event.x, event.y)
        if region not in ("cell", "tree"):
            return
        col_name = self._tree_column_name_at(event.x)
        if col_name not in ("delivered", "cancelled"):
            return
        iid = self._tree.identify_row(event.y)
        if not iid:
            return
        idx = self._row_index(iid)
        if idx is None:
            return
        if col_name == "delivered":
            if self._delivered[idx]:
                self._delivered[idx] = False
            else:
                self._delivered[idx] = True
                self._cancelled[idx] = False
        else:
            if self._cancelled[idx]:
                self._cancelled[idx] = False
            else:
                self._cancelled[idx] = True
                self._delivered[idx] = False
        self._sync_row_tree(idx)
        self._persist_checkbox_state()

    def _open_url_tracked(self, idx: int) -> None:
        u = self._urls[idx]
        if _normalize_url_for_browser(u) is None:
            _open_url(u)
            return
        _open_url(u)
        if not self._opened[idx]:
            self._opened[idx] = True
            self._sync_row_tree(idx)
            self._persist_checkbox_state()

    def _on_tree_double(self, event: tk.Event) -> None:
        region = self._tree.identify_region(event.x, event.y)
        if region in ("nothing", "heading", "separator"):
            return
        iid = self._tree.identify_row(event.y)
        if not iid:
            return
        idx = self._row_index(iid)
        if idx is not None:
            self._open_url_tracked(idx)

    def _on_return_key(self, _event: tk.Event) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        idx = self._row_index(sel[0])
        if idx is not None:
            self._open_url_tracked(idx)

    def _on_mousewheel(self, event: tk.Event) -> None:
        if event.delta:
            self._tree.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event: tk.Event) -> None:
        if event.num == 4:
            self._tree.yview_scroll(-3, "units")
        elif event.num == 5:
            self._tree.yview_scroll(3, "units")

    def run(self) -> None:
        self._root.mainloop()


def main() -> int:
    detach_console_win32()
    register_current_aux_gui()

    # Match project .env location when launched from Excel (cwd may be System32).
    try:
        from dotenv import load_dotenv

        load_dotenv(_PYTHON_FILES_DIR / ".env", override=False)
    except ImportError:
        pass

    parser = argparse.ArgumentParser(description="View tracking URLs in a small grid.")
    parser.add_argument(
        "url_file",
        nargs="?",
        type=Path,
        help="UTF-8 file with one URL per line",
    )
    args = parser.parse_args()

    urls: list[str] = []
    context: dict[str, str] = {}
    if args.url_file is not None:
        p = args.url_file.expanduser().resolve()
        if not p.is_file():
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Tracking links", f"File not found:\n{p}")
            root.destroy()
            return 1
        try:
            urls = _load_urls(p)
            context = _load_context_tsv(_context_path_for_urls_file(p))
        except OSError as e:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Tracking links", f"Could not read file:\n{e}")
            root.destroy()
            return 1

    app = TrackingLinkViewerApp(urls, context)
    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
