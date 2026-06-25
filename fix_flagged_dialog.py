from __future__ import annotations

import queue
import subprocess
import sys
import tempfile
import threading
import webbrowser
import os
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname

from launcher_progress_ui import THEME
from shared.fix_flagged import (
    MANUAL_EDIT_FIELDS,
    REQUEST_TYPE_POD_REVIEW,
    active_requests,
    load_review_records,
    load_requests,
    mark_record_excel_flagged,
    preview_reprocess_record_with_ai,
    record_identity,
    save_record_updates,
    skip_record_eternally,
    sync_fix_flagged_requests,
)
from shared.project_paths import ensure_base_dir_in_environ
from shared.tk_launcher_theme import (
    SettingsStyleSwitch,
    make_flat_button,
    settings_entry_opts,
    settings_label_opts,
    theme_font,
)


FIELD_LABELS = {
    "company": "Company",
    "order_number": "Order Number",
    "purchase_datetime": "Purchase Date",
    "subtotal_amount": "Subtotal",
    "total_amount_paid": "Total Paid",
    "tax_paid": "Tax Paid",
    "gift_card_amount": "GC Paid",
    "email_category": "Category",
}

CATEGORY_VALUES = (
    "Delivered",
    "Invoice",
    "Shipped",
    "Gift Card",
    "Unknown",
    "POD",
)

FIELD_PLACEHOLDERS = {
    "company": "Jersey Mikes",
    "order_number": "AB1234",
    "purchase_datetime": "2026-02-20",
    "subtotal_amount": "00.00",
    "total_amount_paid": "00.00",
    "tax_paid": "00.00",
    "gift_card_amount": "00.00",
}


class FixFlaggedDialog:
    def __init__(self, parent: tk.Misc) -> None:
        self._parent = parent
        self._project_root = ensure_base_dir_in_environ()
        sync_fix_flagged_requests(self._project_root)
        self._requests = load_requests(self._project_root)
        self._records = load_review_records(self._project_root)
        self._items: list[dict] = []
        self._selected_item: dict | None = None
        self._selected_index: int | None = None
        self._suppress_select = False
        self._suppress_show_all_change = False
        self._data_signature: tuple[tuple[str, int, int], ...] | None = None
        self._work_q: queue.Queue[tuple[str, object]] = queue.Queue()
        self._placeholder_active: set[str] = set()
        self._entry_normal_fg = str(settings_entry_opts().get("fg") or THEME["fg"])
        self._entry_placeholder_fg = "#46505a"
        self._pdf_preview_photo: tk.PhotoImage | None = None
        self._current_pdf_path: Path | None = None
        self._pending_ai_preview: dict | None = None

        self._win = tk.Toplevel(parent)
        self._win.title("Fix Flagged")
        self._win.configure(bg=THEME["bg"])
        self._win.geometry("1650x900")
        self._win.minsize(1280, 760)
        self._configure_combobox_style()

        outer = tk.Frame(self._win, bg=THEME["bg"], padx=14, pady=14)
        outer.pack(fill=tk.BOTH, expand=True)

        top = tk.Frame(outer, bg=THEME["bg"])
        top.pack(fill=tk.X, pady=(0, 10))
        self._type_var = tk.StringVar(value="Fix Flagged")
        tk.Label(
            top,
            textvariable=self._type_var,
            fg=THEME["fg"],
            bg=THEME["bg"],
            font=theme_font("title"),
            anchor=tk.W,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)

        self._show_all_var = tk.IntVar(value=0)
        SettingsStyleSwitch(top, self._show_all_var).pack(side=tk.LEFT)
        tk.Label(top, text="Audit all orders", **settings_label_opts()).pack(side=tk.LEFT, padx=(8, 0))
        self._show_all_var.trace_add("write", self._on_show_all_changed)

        body = tk.Frame(outer, bg=THEME["bg"])
        body.pack(fill=tk.BOTH, expand=True)

        left = tk.Frame(body, bg=THEME["bg"])
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 12))
        self._count_var = tk.StringVar(value="")
        tk.Label(left, textvariable=self._count_var, **settings_label_opts()).pack(fill=tk.X, anchor=tk.W)
        self._list = tk.Listbox(
            left,
            width=38,
            height=22,
            bg=THEME["surface"],
            fg=THEME["fg"],
            selectbackground=THEME["run_accent"],
            selectforeground="#ffffff",
            highlightthickness=1,
            highlightbackground=THEME["border"],
            borderwidth=0,
            font=theme_font("body"),
        )
        self._list.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self._list.bind("<<ListboxSelect>>", self._on_select)
        self._list.bind("<Button-3>", self._show_list_context_menu)
        self._list.bind("<Button-2>", self._show_list_context_menu)
        self._list_menu = tk.Menu(self._win, tearoff=0)
        self._list_menu.add_command(label="Mark Flagged in Excel", command=lambda: self._mark_selected_flagged(True))
        self._list_menu.add_command(label="Clear Flagged in Excel", command=lambda: self._mark_selected_flagged(False))

        right = tk.Frame(body, bg=THEME["bg"])
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 12))
        preview_column = tk.Frame(body, bg=THEME["bg"], width=645)
        preview_column.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False)
        preview_column.pack_propagate(False)
        self._summary_var = tk.StringVar(value="Select a request or order.")
        tk.Label(
            right,
            textvariable=self._summary_var,
            wraplength=600,
            justify=tk.LEFT,
            anchor=tk.W,
            fg=THEME["muted"],
            bg=THEME["bg"],
            font=theme_font("body"),
        ).pack(fill=tk.X, anchor=tk.W, pady=(0, 10))

        field_grid = tk.Frame(right, bg=THEME["bg"])
        field_grid.pack(fill=tk.X)
        self._entries: dict[str, tk.Widget] = {}
        entry_opts = settings_entry_opts()
        label_opts = settings_label_opts()
        for idx, field in enumerate(MANUAL_EDIT_FIELDS):
            row = idx // 2
            col = (idx % 2) * 2
            tk.Label(field_grid, text=FIELD_LABELS.get(field, field), anchor=tk.W, **label_opts).grid(
                row=row * 2, column=col, sticky=tk.W, padx=(0 if col == 0 else 14, 6), pady=(0, 3)
            )
            if field == "email_category":
                ent = ttk.Combobox(
                    field_grid,
                    width=26,
                    values=CATEGORY_VALUES,
                    state="readonly",
                    font=theme_font("body"),
                    style="FixFlagged.TCombobox",
                )
            else:
                ent = tk.Entry(field_grid, width=28, **entry_opts)
                self._bind_placeholder(field, ent)
            ent.grid(row=row * 2 + 1, column=col, columnspan=2, sticky=tk.EW, padx=(0 if col == 0 else 14, 0), pady=(0, 8))
            self._entries[field] = ent
        field_grid.grid_columnconfigure(1, weight=1)
        field_grid.grid_columnconfigure(3, weight=1)

        self._status_var = tk.StringVar(value="")
        tk.Label(
            right,
            textvariable=self._status_var,
            anchor=tk.W,
            fg=THEME["muted"],
            bg=THEME["bg"],
            font=theme_font("body"),
        ).pack(fill=tk.X, pady=(8, 0))

        btns = tk.Frame(right, bg=THEME["bg"])
        btns.pack(fill=tk.X, pady=(12, 0))
        self._save_btn = make_flat_button(
            btns,
            text="Save Changes",
            command=self._save_only,
            bg=THEME["excel_accent"],
            active_bg=THEME["excel_accent_dim"],
        )
        self._save_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._reprocess_btn = make_flat_button(
            btns,
            text="Reprocess using AI",
            command=self._reprocess_selected,
            bg="#f59e0b",
            active_bg="#d97706",
            fg="#111827",
            active_fg="#111827",
        )
        self._reprocess_btn.pack(side=tk.LEFT)
        self._skip_eternal_btn = make_flat_button(
            btns,
            text="Skip Eternally",
            command=self._skip_eternally,
            bg=THEME["stop_fg"],
            active_bg="#da3633",
        )
        self._skip_eternal_pack_opts = {"side": tk.LEFT, "padx": (8, 0)}
        make_flat_button(
            btns,
            text="Close",
            command=self._close,
            bg=THEME["surface"],
            active_bg=THEME["track"],
            fg=THEME["fg"],
            active_fg=THEME["fg"],
        ).pack(side=tk.RIGHT)

        preview_panel = tk.Frame(
            preview_column,
            bg=THEME["surface"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            padx=12,
            pady=12,
        )
        preview_panel.pack(fill=tk.BOTH, expand=True)
        tk.Label(
            preview_panel,
            text="POD Preview",
            fg=THEME["fg"],
            bg=THEME["surface"],
            font=theme_font("button"),
            anchor=tk.W,
        ).pack(fill=tk.X, anchor=tk.W)
        preview_box = tk.Frame(
            preview_panel,
            bg=THEME["bg"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
        )
        preview_box.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self._pdf_preview_label = tk.Label(
            preview_box,
            text="No PDF\npreview",
            justify=tk.CENTER,
            fg=THEME["muted"],
            bg=THEME["bg"],
            font=theme_font("body"),
            cursor="hand2",
        )
        self._pdf_preview_label.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self._pdf_preview_label.bind("<Button-1>", lambda _event: self._open_current_pdf())

        review_panel = tk.Frame(
            right,
            bg=THEME["surface"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
            padx=14,
            pady=12,
        )
        review_panel.pack(fill=tk.BOTH, expand=True, pady=(14, 0))
        self._review_help_title_var = tk.StringVar(value="Review details")
        tk.Label(
            review_panel,
            textvariable=self._review_help_title_var,
            fg=THEME["fg"],
            bg=THEME["surface"],
            font=theme_font("button"),
            anchor=tk.W,
        ).pack(fill=tk.X, anchor=tk.W)

        review_body = tk.Frame(review_panel, bg=THEME["surface"])
        review_body.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        review_text = tk.Frame(review_body, bg=THEME["surface"])
        review_text.pack(fill=tk.BOTH, expand=True)
        self._review_help_var = tk.StringVar(value="")
        tk.Label(
            review_text,
            textvariable=self._review_help_var,
            wraplength=455,
            justify=tk.LEFT,
            anchor=tk.NW,
            fg=THEME["muted"],
            bg=THEME["surface"],
            font=theme_font("body"),
        ).pack(fill=tk.X, anchor=tk.NW)
        self._diff_var = tk.StringVar(value="")
        tk.Label(
            review_text,
            textvariable=self._diff_var,
            wraplength=455,
            justify=tk.LEFT,
            anchor=tk.NW,
            fg=THEME["fg"],
            bg=THEME["bg"],
            font=("Consolas", 9),
            padx=8,
            pady=6,
        ).pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self._diff_actions = tk.Frame(review_text, bg=THEME["surface"])
        self._accept_ai_btn = make_flat_button(
            self._diff_actions,
            text="✓ Accept AI Changes",
            command=self._accept_ai_preview,
            bg=THEME["excel_accent"],
            active_bg=THEME["excel_accent_dim"],
        )
        self._accept_ai_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._reject_ai_btn = make_flat_button(
            self._diff_actions,
            text="X Reject",
            command=self._reject_ai_preview,
            bg=THEME["surface"],
            active_bg=THEME["track"],
            fg=THEME["fg"],
            active_fg=THEME["fg"],
        )
        self._reject_ai_btn.pack(side=tk.LEFT)

        self._win.protocol("WM_DELETE_WINDOW", self._close)
        self._reload_items()
        self._pump_work_queue()
        self._poll_external_data_changes()

    def _configure_combobox_style(self) -> None:
        try:
            style = ttk.Style(self._win)
            try:
                style.theme_use("clam")
            except tk.TclError:
                pass
            style.configure(
                "FixFlagged.TCombobox",
                fieldbackground=THEME["surface"],
                background=THEME["surface"],
                foreground=THEME["fg"],
                arrowcolor=THEME["fg"],
                bordercolor=THEME["border"],
                lightcolor=THEME["border"],
                darkcolor=THEME["border"],
                selectbackground=THEME["surface"],
                selectforeground=THEME["fg"],
                insertcolor=THEME["fg"],
                relief="flat",
                padding=1,
            )
            style.map(
                "FixFlagged.TCombobox",
                fieldbackground=[("readonly", THEME["surface"]), ("focus", THEME["surface"])],
                background=[("active", THEME["track"]), ("readonly", THEME["surface"])],
                foreground=[("disabled", THEME["muted"]), ("readonly", THEME["fg"])],
                selectbackground=[("readonly", THEME["surface"])],
                selectforeground=[("readonly", THEME["fg"])],
            )
            self._win.option_add("*TCombobox*Listbox.background", THEME["surface"])
            self._win.option_add("*TCombobox*Listbox.foreground", THEME["fg"])
            self._win.option_add("*TCombobox*Listbox.selectBackground", THEME["run_accent_dim"])
            self._win.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
        except tk.TclError:
            pass

    def _reload_data(self) -> None:
        sync_fix_flagged_requests(self._project_root)
        self._requests = load_requests(self._project_root)
        self._records = load_review_records(self._project_root)
        self._data_signature = self._current_data_signature()

    def _current_data_signature(self) -> tuple[tuple[str, int, int], ...]:
        json_dir = self._project_root / "email_contents" / "json"
        paths = [
            json_dir / "results.json",
            json_dir / "fix_flagged_requests.json",
            json_dir / "help_ai_requests.json",
        ]
        signature: list[tuple[str, int, int]] = []
        for path in paths:
            try:
                stat = path.stat()
                signature.append((str(path), stat.st_mtime_ns, stat.st_size))
            except OSError:
                signature.append((str(path), 0, 0))
        return tuple(signature)

    def _poll_external_data_changes(self) -> None:
        try:
            signature = self._current_data_signature()
            if self._data_signature is None:
                self._data_signature = signature
            elif signature != self._data_signature:
                if self._has_unsaved_changes():
                    self._status_var.set(
                        "Saved data changed elsewhere. Save or discard this edit before refreshing."
                    )
                else:
                    self._reload_items(self._selected_index)
                    self._status_var.set("Updated from latest saved data.")
        except tk.TclError:
            return
        except Exception:
            pass
        try:
            self._win.after(1500, self._poll_external_data_changes)
        except tk.TclError:
            pass

    def _request_label(self, req: dict) -> str:
        company = str(req.get("company") or "Unknown").strip()
        order = str(req.get("order_number") or "no order").strip()
        prefix = "[POD]" if str(req.get("type") or "") == REQUEST_TYPE_POD_REVIEW else "[Fix]"
        return f"{prefix} {company} | {order}"

    def _record_label(self, record: dict) -> str:
        company = str(record.get("company") or "Unknown").strip()
        order = str(record.get("order_number") or "no order").strip()
        category = str(record.get("email_category") or "Unknown").strip()
        return f"{company} | {order} | {category}"

    def _is_pod_request(self, req: dict | None) -> bool:
        return bool(req) and str(req.get("type") or "") == REQUEST_TYPE_POD_REVIEW

    def _is_pod_record(self, record: dict | None) -> bool:
        if not isinstance(record, dict):
            return False
        category = str(record.get("email_category") or "").strip()
        return category == "POD" or bool(record.get("pod_review_required") or record.get("pod_tracking_number"))

    def _is_pod_item(self, record: dict | None, req: dict | None = None) -> bool:
        return self._is_pod_request(req) or self._is_pod_record(record)

    def _pod_request_for_record(self, record: dict) -> dict | None:
        wanted = record_identity(record)
        for req in active_requests(self._requests):
            if (
                str(req.get("type") or "") == REQUEST_TYPE_POD_REVIEW
                and str(req.get("record_identity") or "") == wanted
            ):
                return req
        return None

    def _set_action_labels(self, record: dict | None = None, req: dict | None = None) -> None:
        if self._is_pod_item(record, req):
            self._reprocess_btn.config(text="Rescan Manually")
            if not self._skip_eternal_btn.winfo_ismapped():
                self._skip_eternal_btn.pack(**self._skip_eternal_pack_opts)
            self._skip_eternal_btn.config(state=tk.NORMAL)
            return
        self._reprocess_btn.config(text="Reprocess using AI")
        self._skip_eternal_btn.pack_forget()

    def _set_review_help(self, *, title: str, body: str) -> None:
        self._review_help_title_var.set(title)
        self._review_help_var.set(body)
        if self._pending_ai_preview is None:
            self._diff_var.set("")
            self._diff_actions.pack_forget()

    def _path_from_file_uri(self, value: object) -> Path | None:
        raw = str(value or "").strip()
        if not raw:
            return None
        try:
            parsed = urlparse(raw)
        except Exception:
            return None
        if parsed.scheme != "file":
            return None
        try:
            local_path = url2pathname(unquote(parsed.path))
            return Path(local_path).expanduser().resolve()
        except Exception:
            return None

    def _candidate_pdf_paths(self, record: dict) -> list[Path]:
        out: list[Path] = []

        def add(path: Path | None) -> None:
            if path is None:
                return
            try:
                resolved = path.expanduser().resolve()
            except OSError:
                return
            if resolved.suffix.lower() == ".pdf" and resolved.is_file() and resolved not in out:
                out.append(resolved)

        for key in ("pod_capture_review_pdf", "source_file"):
            raw = str(record.get(key) or "").strip()
            if raw:
                add(Path(raw))
        add(self._path_from_file_uri(record.get("source_file_link")))
        for key in ("pod_generated_file_name", "pod_expected_file_name"):
            filename = str(record.get(key) or "").strip()
            if filename:
                add(self._project_root / "email_contents" / "pdf" / filename)
        source_raw = str(record.get("source_file") or "").strip()
        if source_raw:
            try:
                source_path = Path(source_raw).expanduser()
                if source_path.suffix.lower() in {".html", ".htm"}:
                    add(self._project_root / "email_contents" / "pdf" / f"{source_path.stem}.pdf")
            except OSError:
                pass
        return out

    def _pdf_path_for_record(self, record: dict) -> Path | None:
        paths = self._candidate_pdf_paths(record)
        return paths[0] if paths else None

    def _open_current_pdf(self) -> None:
        path = self._current_pdf_path
        if path is None or not path.is_file():
            return
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception:
            webbrowser.open(path.as_uri())

    def _render_pdf_preview_image(self, path: Path, *, max_width: int = 588, max_height: int = 900):
        try:
            from PIL import Image, ImageTk
        except Exception:
            return None

        image = None
        try:
            import fitz  # type: ignore

            doc = fitz.open(str(path))
            try:
                page = doc.load_page(0)
                rect = page.rect
                scale = min(max_width / max(float(rect.width), 1.0), max_height / max(float(rect.height), 1.0)) * 2.0
                pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            finally:
                doc.close()
        except Exception:
            try:
                import pypdfium2 as pdfium  # type: ignore

                pdf = pdfium.PdfDocument(str(path))
                try:
                    page = pdf[0]
                    bitmap = page.render(scale=1.8)
                    image = bitmap.to_pil()
                finally:
                    pdf.close()
            except Exception:
                image = None
        if image is None:
            return None
        image.thumbnail((max_width, max_height))
        return ImageTk.PhotoImage(image, master=self._win)

    def _update_pdf_preview(self, record: dict) -> None:
        path = self._pdf_path_for_record(record)
        self._current_pdf_path = path
        self._pdf_preview_photo = None
        if path is None:
            self._pdf_preview_label.config(
                image="",
                text="No PDF\npreview",
                fg=THEME["muted"],
                bg=THEME["bg"],
            )
            return
        photo = self._render_pdf_preview_image(path)
        if photo is None:
            self._pdf_preview_label.config(
                image="",
                text=f"PDF preview\nunavailable\n\n{path.name}",
                fg=THEME["muted"],
                bg=THEME["bg"],
            )
            return
        self._pdf_preview_photo = photo
        self._pdf_preview_label.config(image=photo, text="", bg="#ffffff")

    def _bind_placeholder(self, field: str, ent: tk.Entry) -> None:
        if field not in FIELD_PLACEHOLDERS:
            return

        def focus_in(_event: tk.Event) -> None:
            if field in self._placeholder_active:
                self._placeholder_active.discard(field)
                ent.delete(0, tk.END)
                ent.config(fg=self._entry_normal_fg)

        def focus_out(_event: tk.Event) -> None:
            if not ent.get().strip():
                self._show_placeholder(field)

        ent.bind("<FocusIn>", focus_in)
        ent.bind("<FocusOut>", focus_out)

    def _show_placeholder(self, field: str) -> None:
        ent = self._entries.get(field)
        if not isinstance(ent, tk.Entry):
            return
        placeholder = FIELD_PLACEHOLDERS.get(field)
        if not placeholder:
            return
        ent.delete(0, tk.END)
        ent.insert(0, placeholder)
        ent.config(fg=self._entry_placeholder_fg)
        self._placeholder_active.add(field)

    def _clear_placeholder(self, field: str) -> None:
        if field not in self._placeholder_active:
            return
        ent = self._entries.get(field)
        if isinstance(ent, tk.Entry):
            ent.delete(0, tk.END)
            ent.config(fg=self._entry_normal_fg)
        self._placeholder_active.discard(field)

    def _entry_text(self, field: str, ent: tk.Widget) -> str:
        if field in self._placeholder_active:
            return ""
        return ent.get().strip()

    def _reload_items(self, selected_index: int | None = None) -> None:
        self._reload_data()
        if selected_index is None:
            selected_index = self._selected_index
        self._list.delete(0, tk.END)
        active = active_requests(self._requests)
        self._items = [{"kind": "request", "request": req} for req in active]
        if self._show_all_var.get():
            self._items.extend({"kind": "record", "record": rec} for rec in self._records)
        for item in self._items:
            if item["kind"] == "request":
                self._list.insert(tk.END, self._request_label(item["request"]))
            else:
                self._list.insert(tk.END, self._record_label(item["record"]))
        self._count_var.set(
            f"{len(active)} active request(s)"
            + (f" | {len(self._records)} order item(s)" if self._show_all_var.get() else "")
        )
        if self._items:
            idx = selected_index if selected_index is not None else 0
            idx = max(0, min(idx, len(self._items) - 1))
            self._set_list_selection(idx)
            self._show_item(self._items[idx], idx)
        else:
            self._selected_item = None
            self._selected_index = None
            self._type_var.set("Fix Flagged")
            self._summary_var.set("No active Fix Flagged requests.")
            self._set_action_labels(None)
            self._set_review_help(
                title="All clear",
                body="No active Fix Flagged requests. New POD review captures will appear here automatically.",
            )
            self._clear_entries()
            self._update_pdf_preview({})

    def _on_select(self, _event: tk.Event) -> None:
        if self._suppress_select:
            return
        sel = self._list.curselection()
        if not sel:
            return
        idx = int(sel[0])
        if idx == self._selected_index:
            return
        if 0 <= idx < len(self._items):
            if not self._confirm_save_or_discard_changes():
                if self._selected_index is not None:
                    self._set_list_selection(self._selected_index)
                return
            self._show_item(self._items[idx], idx)

    def _on_show_all_changed(self, *_args: object) -> None:
        if self._suppress_show_all_change:
            return
        if self._confirm_save_or_discard_changes():
            self._reload_items(0)
            return
        self._suppress_show_all_change = True
        try:
            self._show_all_var.set(0 if self._show_all_var.get() else 1)
        finally:
            self._suppress_show_all_change = False

    def _set_list_selection(self, idx: int) -> None:
        self._suppress_select = True
        try:
            self._list.selection_clear(0, tk.END)
            self._list.selection_set(idx)
            self._list.activate(idx)
            self._list.see(idx)
        finally:
            self._suppress_select = False

    def _clear_entries(self) -> None:
        self._placeholder_active.clear()
        for field, ent in self._entries.items():
            if hasattr(ent, "set"):
                ent.set("")  # type: ignore[attr-defined]
            else:
                ent.delete(0, tk.END)
                ent.config(fg=self._entry_normal_fg)
                self._show_placeholder(field)

    def _set_entry_value(self, field: str, value: object) -> None:
        ent = self._entries[field]
        text = "" if value is None else str(value)
        self._clear_placeholder(field)
        if hasattr(ent, "set"):
            ent.set(text)  # type: ignore[attr-defined]
            return
        ent.delete(0, tk.END)
        ent.config(fg=self._entry_normal_fg)
        if text:
            ent.insert(0, text)
        else:
            self._show_placeholder(field)

    def _record_for_request(self, req: dict) -> dict | None:
        wanted = str(req.get("record_identity") or "")
        for record in self._records:
            if record_identity(record) == wanted:
                return record
        return None

    def _show_item(self, item: dict, idx: int | None = None) -> None:
        self._pending_ai_preview = None
        self._diff_var.set("")
        self._diff_actions.pack_forget()
        self._selected_item = item
        self._selected_index = idx
        self._clear_entries()
        if item["kind"] == "request":
            req = item["request"]
            record = self._record_for_request(req) or {}
            self._set_action_labels(record, req)
            self._type_var.set(f"Type: {req.get('type') or 'request'}")
            self._summary_var.set(
                f"{req.get('title') or 'Fix Flagged request'}\n"
                f"{req.get('reason') or ''}\n"
                f"Subject: {record.get('subject') or req.get('subject') or ''}"
            )
            if self._is_pod_request(req):
                self._set_review_help(
                    title="POD saved for review",
                    body=(
                        "The page PDF was captured and the row was flagged, but the audit could not "
                        "confirm usable tracking details. This often happens because email scanning "
                        "mistakes an order or reference number for a carrier tracking number. It can "
                        "also happen when old carrier tracking details are no longer available. Use "
                        "Rescan Manually to inspect the saved PDF/tracking page, or Skip Eternally to "
                        "clear this review item."
                    ),
                )
            else:
                self._set_review_help(
                    title="Review details",
                    body="Check the highlighted fields, save any corrections, then skip or resolve the request.",
                )
        else:
            record = item["record"]
            pod_req = self._pod_request_for_record(record)
            self._set_action_labels(record, pod_req)
            if self._is_pod_item(record, pod_req):
                self._type_var.set("Type: POD review")
                self._summary_var.set(
                    f"Existing POD item. Subject: {record.get('subject') or ''}\n"
                    "Use Rescan Manually to open the tracking page and decide whether to replace this POD."
                )
                self._set_review_help(
                    title="POD saved for review",
                    body=(
                        "This row is a proof-of-delivery item. Rescan Manually opens a visible Chrome "
                        "capture for just this tracking number, then waits for your replace or skip decision."
                    ),
                )
            else:
                self._type_var.set("Type: Manual order audit")
                self._summary_var.set(
                    f"Existing order item. Subject: {record.get('subject') or ''}\n"
                    "You can adjust fields manually or force a fresh AI extraction."
                )
                self._set_review_help(
                    title="Manual audit",
                    body=(
                        "AI can be useful for a fresh extraction, but it can also over-read numbers in "
                        "emails. Keep any fields you trust before reprocessing."
                    ),
                )
        for field, ent in self._entries.items():
            value = record.get(field)
            self._set_entry_value(field, value)
        self._update_pdf_preview(record)

    def _selected_record_and_request(self) -> tuple[dict, dict | None]:
        item = self._selected_item
        if not item:
            raise ValueError("Select an item first.")
        if item["kind"] == "request":
            req = item["request"]
            record = self._record_for_request(req)
            if record is None:
                raise ValueError("The request no longer matches an order in results.json.")
            return record, req
        return item["record"], None

    def _entry_updates(self) -> dict[str, object]:
        record, _req = self._selected_record_and_request()
        updates: dict[str, object] = {}
        for field, ent in self._entries.items():
            new_value = self._entry_text(field, ent)
            old = record.get(field)
            old_value = "" if old is None else str(old)
            if new_value != old_value:
                updates[field] = new_value
        return updates

    def _show_list_context_menu(self, event: tk.Event) -> None:
        idx = self._list.nearest(event.y)
        if idx < 0 or idx >= len(self._items):
            return
        if idx != self._selected_index:
            if not self._confirm_save_or_discard_changes():
                return
            self._set_list_selection(idx)
            self._show_item(self._items[idx], idx)
        try:
            record, _req = self._selected_record_and_request()
            is_flagged = bool(record.get("excel_flagged") or record.get("excel_active"))
            self._list_menu.entryconfig(0, state=tk.DISABLED if is_flagged else tk.NORMAL)
            self._list_menu.entryconfig(1, state=tk.NORMAL if is_flagged else tk.DISABLED)
            self._list_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self._list_menu.grab_release()

    def _mark_selected_flagged(self, flagged: bool) -> None:
        try:
            record, _req = self._selected_record_and_request()
            mark_record_excel_flagged(
                self._project_root,
                record_id=record_identity(record),
                flagged=flagged,
            )
            self._status_var.set("Marked Flagged in Excel." if flagged else "Cleared Flagged in Excel.")
            self._reload_items(self._selected_index)
        except Exception as exc:
            messagebox.showerror("Fix Flagged", str(exc), parent=self._win)

    def _has_unsaved_changes(self) -> bool:
        try:
            return bool(self._entry_updates())
        except Exception:
            return False

    def _save(self, *, status: str | None = None, reload_items: bool = True) -> None:
        record, req = self._selected_record_and_request()
        save_record_updates(
            self._project_root,
            record_id=record_identity(record),
            updates=self._entry_updates(),
            request_id=str(req.get("id")) if req and status else None,
            status=status,
        )
        self._status_var.set("Saved.")
        if reload_items:
            self._reload_items(self._selected_index)
        else:
            self._reload_data()

    def _save_only(self) -> None:
        try:
            self._save()
        except Exception as exc:
            messagebox.showerror("Fix Flagged", str(exc), parent=self._win)

    def _confirm_save_or_discard_changes(self) -> bool:
        if not self._has_unsaved_changes():
            return True
        choice = messagebox.askyesnocancel(
            "Unsaved Fix Flagged changes",
            "Save changes before leaving this item?\n\n"
            "Yes saves your changes. No discards them. Cancel returns to editing.",
            parent=self._win,
        )
        if choice is None:
            return False
        if choice:
            try:
                self._save(reload_items=False)
            except Exception as exc:
                messagebox.showerror("Fix Flagged", str(exc), parent=self._win)
                return False
        return True

    def _close(self) -> None:
        if self._confirm_save_or_discard_changes():
            self._win.destroy()

    def _set_busy(self, busy: bool, message: str = "") -> None:
        state = tk.DISABLED if busy else tk.NORMAL
        for btn in (self._save_btn, self._reprocess_btn, self._skip_eternal_btn):
            btn.config(state=state)
        if not busy:
            item = self._selected_item
            req = item.get("request") if isinstance(item, dict) and item.get("kind") == "request" else None
            record = None
            if isinstance(item, dict):
                if item.get("kind") == "request" and isinstance(req, dict):
                    record = self._record_for_request(req)
                elif item.get("kind") == "record":
                    record = item.get("record")
            self._set_action_labels(record, req)
        self._status_var.set(message)

    def _reprocess_selected(self) -> None:
        if not self._confirm_save_or_discard_changes():
            return
        try:
            record, req = self._selected_record_and_request()
        except Exception as exc:
            messagebox.showerror("Fix Flagged", str(exc), parent=self._win)
            return
        pod_req = req or self._pod_request_for_record(record)
        if self._is_pod_item(record, pod_req):
            self._rescan_pod_manually(record, pod_req)
            return
        record_id = record_identity(record)
        self._set_busy(True, "Reprocessing with AI for preview...")

        def worker() -> None:
            try:
                result = preview_reprocess_record_with_ai(
                    self._project_root,
                    record_id=record_id,
                )
                self._work_q.put(("ai_preview", {"record_id": record_id, "old": record, "new": result}))
            except Exception as exc:
                self._work_q.put(("err", exc))

        threading.Thread(target=worker, daemon=True).start()

    def _rescan_pod_manually(self, record: dict, req: dict | None = None) -> None:
        tracking_number = str(record.get("pod_tracking_number") or "").strip()
        if not tracking_number:
            raw_numbers = record.get("tracking_numbers")
            if isinstance(raw_numbers, list) and raw_numbers:
                tracking_number = str(raw_numbers[0] or "").strip()
        if tracking_number:
            tmp_dir = Path(tempfile.mkdtemp(prefix="email_sorter_pod_review_"))
            txt_path = tmp_dir / "pod_review_rescan.txt"
            ctx_path = txt_path.with_suffix(".ctx.tsv")
            txt_path.write_text(f"{tracking_number}\n", encoding="utf-8")
            context_values = {
                "company": record.get("company"),
                "order_number": record.get("order_number"),
                "purchase_datetime": record.get("pod_source_purchase_datetime") or record.get("purchase_datetime"),
                "category": record.get("pod_source_category") or record.get("email_category") or "POD",
                "email_category": record.get("pod_source_category") or record.get("email_category") or "POD",
                "email": record.get("pod_source_email") or record.get("email"),
                "tracking_numbers": tracking_number,
                "auto_start_pod_capture": "1",
                "pod_review_rescan": "1",
                "fix_flagged_record_identity": record_identity(record),
                "fix_flagged_request_id": req.get("id") if req else "",
                "source_file": record.get("source_file"),
                "pod_capture_review_pdf": record.get("pod_capture_review_pdf"),
            }
            lines = []
            for key, value in context_values.items():
                text = str(value or "").strip()
                if text:
                    lines.append(f"{key}\t{text}\n")
            ctx_path.write_text("".join(lines), encoding="utf-8")
            viewer = Path(__file__).resolve().parent / "trackingNumbersViewer" / "tracking_status_viewer.py"
            subprocess.Popen([sys.executable, str(viewer), str(txt_path)], cwd=str(Path(__file__).resolve().parent))
            self._status_var.set("Opened the POD capture workflow for this review item.")
            return
        self._open_current_pdf()
        self._status_var.set("Opened the saved review PDF. No tracking number was available for POD rescan.")

    def _skip_eternally(self) -> None:
        try:
            record, req = self._selected_record_and_request()
            pod_req = req or self._pod_request_for_record(record)
            if not self._is_pod_item(record, pod_req):
                return
            result = skip_record_eternally(
                self._project_root,
                record_id=record_identity(record),
                request_id=str(pod_req.get("id") or "") if pod_req else None,
            )
            removed_excel = bool(result.get("removed_excel_row"))
            self._status_var.set(
                "Skipped eternally and removed from JSON and Excel."
                if removed_excel
                else "Skipped eternally and removed from JSON. Rebuild/sync Excel to remove it there if it was not open."
            )
            self._reload_items(self._selected_index)
        except Exception as exc:
            messagebox.showerror("Fix Flagged", str(exc), parent=self._win)

    def _manual_field_values_from_record(self, record: dict) -> dict[str, object]:
        return {field: record.get(field) for field in MANUAL_EDIT_FIELDS}

    def _display_value(self, value: object) -> str:
        return "" if value is None else str(value)

    def _apply_field_values(self, values: dict[str, object]) -> None:
        for field in MANUAL_EDIT_FIELDS:
            self._set_entry_value(field, values.get(field))

    def _show_ai_preview(self, payload: dict) -> None:
        old_record = payload.get("old") if isinstance(payload.get("old"), dict) else {}
        new_record = payload.get("new") if isinstance(payload.get("new"), dict) else {}
        old_values = self._manual_field_values_from_record(old_record)
        new_values = self._manual_field_values_from_record(new_record)
        self._pending_ai_preview = {
            "record_id": str(payload.get("record_id") or record_identity(old_record)),
            "old_values": old_values,
            "new_values": new_values,
        }
        self._apply_field_values(new_values)
        lines = ["AI proposed field changes:"]
        changed = False
        for field in MANUAL_EDIT_FIELDS:
            old_text = self._display_value(old_values.get(field))
            new_text = self._display_value(new_values.get(field))
            if old_text == new_text:
                continue
            changed = True
            label = FIELD_LABELS.get(field, field)
            lines.append(f"- {label}: {old_text or '(blank)'}")
            lines.append(f"+ {label}: {new_text or '(blank)'}")
        if not changed:
            lines.append("No visible field changes returned.")
        self._diff_var.set("\n".join(lines))
        self._diff_actions.pack(fill=tk.X, pady=(8, 0))
        self._review_help_title_var.set("AI change preview")
        self._review_help_var.set(
            "Review the proposed values now shown in the fields. Use the checkmark to save them, or X to restore the previous values."
        )
        self._status_var.set("AI preview ready. Confirm or reject the proposed changes.")

    def _accept_ai_preview(self) -> None:
        if not self._pending_ai_preview:
            return
        try:
            record, _req = self._selected_record_and_request()
            save_record_updates(
                self._project_root,
                record_id=record_identity(record),
                updates=self._entry_updates(),
            )
            self._pending_ai_preview = None
            self._diff_var.set("")
            self._diff_actions.pack_forget()
            self._status_var.set("AI changes saved.")
            self._reload_items(self._selected_index)
        except Exception as exc:
            messagebox.showerror("Fix Flagged", str(exc), parent=self._win)

    def _reject_ai_preview(self) -> None:
        if not self._pending_ai_preview:
            return
        old_values = self._pending_ai_preview.get("old_values")
        self._pending_ai_preview = None
        if isinstance(old_values, dict):
            self._apply_field_values(old_values)
        self._diff_var.set("")
        self._diff_actions.pack_forget()
        self._status_var.set("AI changes rejected.")

    def _pump_work_queue(self) -> None:
        try:
            while True:
                kind, payload = self._work_q.get_nowait()
                self._set_busy(False)
                if kind == "ai_preview" and isinstance(payload, dict):
                    self._show_ai_preview(payload)
                elif kind == "ok":
                    self._status_var.set("Done.")
                    self._reload_items()
                else:
                    messagebox.showerror("Fix Flagged", str(payload), parent=self._win)
        except queue.Empty:
            pass
        try:
            self._win.after(100, self._pump_work_queue)
        except tk.TclError:
            pass
