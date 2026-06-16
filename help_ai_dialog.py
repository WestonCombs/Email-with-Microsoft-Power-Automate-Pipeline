from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from launcher_progress_ui import THEME
from shared.help_ai import (
    MANUAL_EDIT_FIELDS,
    active_requests,
    load_results,
    load_requests,
    record_identity,
    reprocess_record_with_ai,
    save_record_updates,
    sync_help_ai_requests,
    update_request_status,
)
from shared.project_paths import ensure_base_dir_in_environ
from shared.tk_launcher_theme import (
    SettingsStyleSwitch,
    add_button_hover,
    danger_colors,
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
    "gift_card_amount": "Gift Card",
    "tax_was_paid": "Tax Was Paid",
    "email_category": "Category",
}


class HelpAIDialog:
    def __init__(self, parent: tk.Misc) -> None:
        self._parent = parent
        self._project_root = ensure_base_dir_in_environ()
        sync_help_ai_requests(self._project_root)
        self._requests = load_requests(self._project_root)
        self._records = load_results(self._project_root)
        self._items: list[dict] = []
        self._selected_item: dict | None = None
        self._work_q: queue.Queue[tuple[str, object]] = queue.Queue()

        self._win = tk.Toplevel(parent)
        self._win.title("Help AI")
        self._win.configure(bg=THEME["bg"])
        self._win.geometry("980x640")
        self._win.minsize(850, 560)

        outer = tk.Frame(self._win, bg=THEME["bg"], padx=14, pady=14)
        outer.pack(fill=tk.BOTH, expand=True)

        top = tk.Frame(outer, bg=THEME["bg"])
        top.pack(fill=tk.X, pady=(0, 10))
        self._type_var = tk.StringVar(value="Help AI")
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
        self._show_all_var.trace_add("write", lambda *_: self._reload_items())

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

        right = tk.Frame(body, bg=THEME["bg"])
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
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
        self._entries: dict[str, tk.Entry] = {}
        entry_opts = settings_entry_opts()
        label_opts = settings_label_opts()
        for idx, field in enumerate(MANUAL_EDIT_FIELDS):
            row = idx // 2
            col = (idx % 2) * 2
            tk.Label(field_grid, text=FIELD_LABELS.get(field, field), anchor=tk.W, **label_opts).grid(
                row=row * 2, column=col, sticky=tk.W, padx=(0 if col == 0 else 14, 6), pady=(0, 3)
            )
            ent = tk.Entry(field_grid, width=28, **entry_opts)
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
        d_bg, d_active = danger_colors()
        self._save_btn = make_flat_button(
            btns,
            text="Save Changes",
            command=self._save_only,
            bg=THEME["excel_accent"],
            active_bg=THEME["excel_accent_dim"],
        )
        self._save_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._resolve_btn = make_flat_button(
            btns,
            text="Mark Resolved",
            command=self._resolve_selected,
            bg="#16a34a",
            active_bg="#15803d",
        )
        self._resolve_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._skip_btn = make_flat_button(
            btns,
            text="Skip",
            command=self._skip_selected,
            bg=d_bg,
            active_bg=d_active,
        )
        self._skip_btn.pack(side=tk.LEFT, padx=(0, 8))
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
        make_flat_button(
            btns,
            text="Close",
            command=self._win.destroy,
            bg=THEME["surface"],
            active_bg=THEME["track"],
            fg=THEME["fg"],
            active_fg=THEME["fg"],
        ).pack(side=tk.RIGHT)

        self._reload_items()
        self._pump_work_queue()

    def _reload_data(self) -> None:
        sync_help_ai_requests(self._project_root)
        self._requests = load_requests(self._project_root)
        self._records = load_results(self._project_root)

    def _request_label(self, req: dict) -> str:
        company = str(req.get("company") or "Unknown").strip()
        order = str(req.get("order_number") or "no order").strip()
        return f"[Fix] {company} | {order}"

    def _record_label(self, record: dict) -> str:
        company = str(record.get("company") or "Unknown").strip()
        order = str(record.get("order_number") or "no order").strip()
        category = str(record.get("email_category") or "Unknown").strip()
        return f"[Order] {company} | {order} | {category}"

    def _reload_items(self) -> None:
        self._reload_data()
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
            self._list.selection_set(0)
            self._show_item(self._items[0])
        else:
            self._selected_item = None
            self._type_var.set("Help AI")
            self._summary_var.set("No active Help AI requests.")
            self._clear_entries()

    def _on_select(self, _event: tk.Event) -> None:
        sel = self._list.curselection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self._items):
            self._show_item(self._items[idx])

    def _clear_entries(self) -> None:
        for ent in self._entries.values():
            ent.delete(0, tk.END)

    def _record_for_request(self, req: dict) -> dict | None:
        wanted = str(req.get("record_identity") or "")
        for record in self._records:
            if record_identity(record) == wanted:
                return record
        return None

    def _show_item(self, item: dict) -> None:
        self._selected_item = item
        self._clear_entries()
        if item["kind"] == "request":
            req = item["request"]
            record = self._record_for_request(req) or {}
            self._type_var.set(f"Type: {req.get('type') or 'request'}")
            self._summary_var.set(
                f"{req.get('title') or 'Help request'}\n"
                f"{req.get('reason') or ''}\n"
                f"Subject: {record.get('subject') or req.get('subject') or ''}"
            )
        else:
            record = item["record"]
            self._type_var.set("Type: Manual order audit")
            self._summary_var.set(
                f"Existing order item. Subject: {record.get('subject') or ''}\n"
                "You can adjust fields manually or force a fresh AI extraction."
            )
        for field, ent in self._entries.items():
            value = record.get(field)
            ent.insert(0, "" if value is None else str(value))

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
            new_value = ent.get().strip()
            old = record.get(field)
            old_value = "" if old is None else str(old)
            if new_value != old_value:
                updates[field] = new_value
        return updates

    def _save(self, *, status: str | None = None) -> None:
        record, req = self._selected_record_and_request()
        save_record_updates(
            self._project_root,
            record_id=record_identity(record),
            updates=self._entry_updates(),
            request_id=str(req.get("id")) if req and status else None,
            status=status,
        )
        self._status_var.set("Saved.")
        self._reload_items()

    def _save_only(self) -> None:
        try:
            self._save()
        except Exception as exc:
            messagebox.showerror("Help AI", str(exc), parent=self._win)

    def _resolve_selected(self) -> None:
        try:
            self._save(status="resolved")
        except Exception as exc:
            messagebox.showerror("Help AI", str(exc), parent=self._win)

    def _skip_selected(self) -> None:
        try:
            _record, req = self._selected_record_and_request()
            if not req:
                self._status_var.set("Nothing to skip for a normal order item.")
                return
            update_request_status(self._project_root, request_id=str(req.get("id")), status="skipped")
            self._status_var.set("Skipped.")
            self._reload_items()
        except Exception as exc:
            messagebox.showerror("Help AI", str(exc), parent=self._win)

    def _set_busy(self, busy: bool, message: str = "") -> None:
        state = tk.DISABLED if busy else tk.NORMAL
        for btn in (self._save_btn, self._resolve_btn, self._skip_btn, self._reprocess_btn):
            btn.config(state=state)
        self._status_var.set(message)

    def _reprocess_selected(self) -> None:
        try:
            record, _req = self._selected_record_and_request()
        except Exception as exc:
            messagebox.showerror("Help AI", str(exc), parent=self._win)
            return
        if not messagebox.askyesno(
            "Reprocess using AI",
            "Run this order through AI extraction again and overwrite its current stored fields?",
            parent=self._win,
        ):
            return
        record_id = record_identity(record)
        self._set_busy(True, "Reprocessing with AI...")

        def worker() -> None:
            try:
                result = reprocess_record_with_ai(self._project_root, record_id=record_id)
                self._work_q.put(("ok", result))
            except Exception as exc:
                self._work_q.put(("err", exc))

        threading.Thread(target=worker, daemon=True).start()

    def _pump_work_queue(self) -> None:
        try:
            while True:
                kind, payload = self._work_q.get_nowait()
                self._set_busy(False)
                if kind == "ok":
                    self._status_var.set("Reprocessed with AI.")
                    self._reload_items()
                else:
                    messagebox.showerror("Help AI", str(payload), parent=self._win)
        except queue.Empty:
            pass
        try:
            self._win.after(100, self._pump_work_queue)
        except tk.TclError:
            pass
