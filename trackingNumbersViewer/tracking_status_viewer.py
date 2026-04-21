"""
17TRACK shipping status viewer (smart cache + milestones).

Reads the same temp file format as ``tracking_numbers_viewer`` (one number per line;
optional ``number<TAB>0|1``). Optional ``.ctx.tsv`` beside the file supplies row context.

Usage:
  python tracking_status_viewer.py <path_to_numbers.txt>
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import webbrowser
from datetime import datetime
from functools import lru_cache
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

_PYTHON_FILES_DIR = Path(__file__).resolve().parent.parent
if str(_PYTHON_FILES_DIR) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES_DIR))

from shared.gui_aux_singleton import detach_console_win32, register_current_aux_gui
from shared.gui_treeview_copy import bind_treeview_copy_menu
from htmlHandler.carrier_urls import normalize_carrier_for_public_url, public_tracking_url
from pdfCaptureFromChrome.paths import PDF_CAPTURE_SESSION_LOG
from proofOfDelivery.pod_data import (
    POD_HUB_MODE,
    expected_pod_pdf_path,
    project_root_from_env,
    remaining_pod_candidates,
)
from trackingNumbersViewer.mitm_readiness import pdf_capture_environment_ready, sanitize_filename_token
from trackingNumbersViewer.seventeen_track_api import api_key_from_env
from trackingNumbersViewer.seventeen_track_smart import (
    carrier_display_for_number,
    extract_track_info,
    fetch_tracking_smart,
    load_cache,
    quick_status_from_cache,
    tracking_is_greyed_out,
)
from shared.ui_dark_theme import (
    UI_BG,
    UI_BG_PANEL,
    UI_FG,
    UI_FG_DIM,
    UI_TREE_BG,
    dark_tk_scrollbar,
    setup_dark_theme,
    style_text_widget,
)

_PDF_TOGGLE_ON = "#22c55e"
_PDF_TOGGLE_OFF = "#ef4444"
_PDF_TOGGLE_FG = "#ffffff"

_ROW_GREYED = "greyed_out"
_ROW_POD_COMPLETE = "pod_pdf_saved"

_NOTFOUND_STATUS_RE = re.compile(r"\bnot[\s_-]?found\b", re.IGNORECASE)


def _quick_status_indicates_notfound(quick_status: str) -> bool:
    return bool(_NOTFOUND_STATUS_RE.search(quick_status or ""))


def _load_tracking_file(path: Path) -> tuple[list[str], list[bool | None]]:
    raw = path.read_text(encoding="utf-8-sig")
    numbers: list[str] = []
    confirmed: list[bool | None] = []
    seen: set[str] = set()
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        if "\t" in line:
            num, _, _ = line.partition("\t")
            num = num.strip()
        else:
            num = line
        if not num or num in seen:
            continue
        seen.add(num)
        numbers.append(num)
        confirmed.append(None)
    return numbers, confirmed


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


def _normalize_ctx_text(value: object) -> str:
    return " ".join(str(value or "").strip().split()).casefold()


def _tracking_numbers_for_record(record: object) -> set[str]:
    if not isinstance(record, dict):
        return set()
    raw = record.get("tracking_numbers")
    if not isinstance(raw, list):
        return set()
    out: set[str] = set()
    for item in raw:
        s = str(item or "").strip()
        if s:
            out.add(_normalize_ctx_text(s))
    return out


@lru_cache(maxsize=1)
def _load_results_json_records() -> tuple[dict, ...]:
    base_raw = (os.getenv("BASE_DIR") or "").strip()
    if not base_raw:
        return ()
    path = Path(base_raw).expanduser().resolve() / "email_contents" / "json" / "results.json"
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return ()
    if not isinstance(payload, list):
        return ()
    return tuple(item for item in payload if isinstance(item, dict))


def _derive_company_from_project_data(context: dict[str, str], numbers: list[str]) -> str | None:
    records = _load_results_json_records()
    if not records:
        return None

    want_order = _normalize_ctx_text(context.get("order_number"))
    want_date = _normalize_ctx_text(context.get("purchase_datetime"))
    want_category = _normalize_ctx_text(context.get("category"))
    want_email = _normalize_ctx_text(context.get("email"))
    want_numbers = {_normalize_ctx_text(n) for n in numbers if str(n).strip()}

    best_score = 0
    best_company: str | None = None

    for record in records:
        company = " ".join(str(record.get("company") or "").strip().split())
        if not company:
            continue

        score = 0
        if want_order and _normalize_ctx_text(record.get("order_number")) == want_order:
            score += 10
        if want_date and _normalize_ctx_text(record.get("purchase_datetime")) == want_date:
            score += 3
        if want_category and _normalize_ctx_text(record.get("email_category")) == want_category:
            score += 2
        if want_email and _normalize_ctx_text(record.get("email")) == want_email:
            score += 1

        overlap = len(want_numbers & _tracking_numbers_for_record(record))
        if overlap:
            score += 6 + overlap

        if score > best_score:
            best_score = score
            best_company = company

    return best_company


def _event_time_sort_key(e: dict) -> float:
    raw = str(e.get("time_iso") or e.get("time_raw") or "").strip()
    if not raw:
        return 0.0
    try:
        s = raw.replace("Z", "+00:00")
        return datetime.fromisoformat(s).timestamp()
    except Exception:
        return 0.0


def _events_lines_from_track_info(track_info: dict | None) -> list[str]:
    if not isinstance(track_info, dict):
        return []
    rows: list[tuple[float, str]] = []
    prov = track_info.get("tracking", {}).get("providers", [])
    if not isinstance(prov, list):
        return []
    for p in prov:
        if not isinstance(p, dict):
            continue
        events = p.get("events", [])
        if not isinstance(events, list):
            continue
        for e in events:
            if not isinstance(e, dict):
                continue
            t = str(e.get("time_iso") or e.get("time_raw") or "").strip()
            desc = str(e.get("description") or e.get("stage") or "").strip()
            loc = str(e.get("location") or "").strip()
            chunk = " | ".join(x for x in (t, desc, loc) if x)
            if chunk:
                rows.append((_event_time_sort_key(e), chunk))
    rows.sort(key=lambda x: x[0])
    return [r[1] for r in rows]


class TrackingStatusViewerApp:
    def __init__(self, numbers: list[str], context: dict[str, str]) -> None:
        self._context = dict(context)
        self._hub_remaining_mode = str(self._context.get("pod_mode") or "").strip() == POD_HUB_MODE
        try:
            self._project_root = project_root_from_env()
        except Exception:
            self._project_root = None

        if not str(self._context.get("company") or "").strip():
            derived_company = _derive_company_from_project_data(self._context, numbers)
            if derived_company:
                self._context["company"] = derived_company

        self._row_infos = self._build_row_infos(numbers)
        self._numbers = [str(info["tracking_number"]) for info in self._row_infos]
        self._pdf_capture = self._hub_remaining_mode

        try:
            from dotenv import load_dotenv

            load_dotenv(_PYTHON_FILES_DIR / ".env", override=False)
        except ImportError:
            pass

        self._root = tk.Tk()
        self._root.title("Shipping status (17TRACK)")
        self._root.minsize(760, 400)
        self._root.geometry("960x520")
        setup_dark_theme(self._root)

        self._frm = ttk.Frame(self._root, padding=8)
        self._frm.pack(fill=tk.BOTH, expand=True)

        bits = []
        for k in ("company", "order_number", "category", "purchase_datetime"):
            v = self._context.get(k)
            if v:
                bits.append(f"{k.replace('_', ' ')}: {v}")
        tn = self._context.get("tracking_numbers")
        if tn:
            bits.append(f"tracking numbers: {tn}")
        ctx_line = "  |  ".join(bits) if bits else ""
        if ctx_line:
            ttk.Label(self._frm, text=ctx_line, wraplength=920, foreground=UI_FG_DIM).pack(
                fill=tk.X, anchor=tk.W, pady=(0, 6)
            )

        hint_frame = ttk.Frame(self._frm)
        hint_frame.pack(fill=tk.X, anchor=tk.W, pady=(0, 4))

        if self._hub_remaining_mode:
            ttk.Label(
                hint_frame,
                text=(
                    "This view lists every tracking number that still needs a proof-of-delivery PDF. "
                    "(Weston needs more time to perfect or adapt this feature.)"
                ),
                wraplength=920,
                foreground=UI_FG,
            ).pack(fill=tk.X, anchor=tk.W)
        else:
            ttk.Label(
                hint_frame,
                text="Double-click a row to open the tracking page,",
                wraplength=920,
                foreground=UI_FG,
            ).pack(fill=tk.X, anchor=tk.W)
            ttk.Label(
                hint_frame,
                text=(
                    "Turn on Toggle PDF Capture to run that process. "
                    "(Weston needs more time to perfect or adapt this feature.)"
                ),
                wraplength=920,
                foreground=UI_FG,
            ).pack(fill=tk.X, anchor=tk.W)

        ttk.Label(
            hint_frame,
            text=(
                #"Tracking numbers that were already processed appear with a green row highlight. "
                "NotFound tracking numbers appear with a grey row highlight for two weeks and expire "
                "after a final retry occurs two weeks after their first check."
            ),
            wraplength=920,
            foreground=UI_FG,
        ).pack(fill=tk.X, anchor=tk.W, pady=(4, 0))

        demo = ttk.Frame(hint_frame)
        demo.pack(fill=tk.X, anchor=tk.W, pady=(6, 0))
        ttk.Label(demo, text="Sample row colors: ").pack(side=tk.LEFT)
        tk.Label(
            demo,
            text=" GREEN ",
            bg="#166534",
            fg=UI_FG,
            padx=6,
            pady=2,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(demo, text="processed (beta feature)").pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(
            demo,
            text=" GREY ",
            bg="#475569",
            fg="#e2e8f0",
            padx=6,
            pady=2,
        ).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(demo, text="NotFound / temporary").pack(side=tk.LEFT)

        if not api_key_from_env():
            ttk.Label(
                self._frm,
                text=(
                    "No SEVENTEEN_TRACK_API_KEY in .env — only cached tracking data will appear. "
                    "Get a free key at 17track.net."
                ),
                wraplength=920,
                foreground="#fbbf24",
            ).pack(fill=tk.X, anchor=tk.W, pady=(0, 6))

        btn_row = ttk.Frame(self._frm)
        btn_row.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))
        ttk.Button(btn_row, text="View more", command=self._view_more).pack(side=tk.LEFT, padx=(0, 8))
        self._pdf_toggle_btn = tk.Button(
            btn_row,
            text="PDF Capture Locked On" if self._hub_remaining_mode else "Toggle PDF Capture",
            command=self._toggle_pdf_capture,
            bg=_PDF_TOGGLE_ON if self._pdf_capture else _PDF_TOGGLE_OFF,
            fg=_PDF_TOGGLE_FG,
            activebackground="#16a34a" if self._pdf_capture else "#dc2626",
            activeforeground=_PDF_TOGGLE_FG,
            relief=tk.RAISED,
            bd=2,
            cursor="hand2",
            state=tk.DISABLED if self._hub_remaining_mode else tk.NORMAL,
        )
        self._pdf_toggle_btn.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(btn_row, text="Close", command=self._root.destroy).pack(side=tk.RIGHT)

        tree_outer = tk.Frame(self._frm, bg=UI_BG)
        tree_outer.pack(fill=tk.BOTH, expand=True)

        cols = (
            ("idx", "tracking_number", "carrier", "quick_status")
            if self._hub_remaining_mode
            else ("idx", "tracking_number", "carrier", "quick_status", "already_processed")
        )
        self._tree = ttk.Treeview(
            tree_outer,
            columns=cols,
            show="headings",
            selectmode="browse",
            height=min(18, max(6, len(self._row_infos) or 6)),
        )
        self._tree.heading("idx", text="#")
        self._tree.heading("tracking_number", text="Tracking number")
        self._tree.heading("carrier", text="Carrier")
        self._tree.heading("quick_status", text="Shipping Status")
        if not self._hub_remaining_mode:
            self._tree.heading("already_processed", text="Already Processed")

        self._tree.column("idx", width=40, anchor=tk.CENTER)
        self._tree.column("tracking_number", width=200, anchor=tk.W)
        self._tree.column("carrier", width=100, anchor=tk.CENTER)
        self._tree.column("quick_status", width=550 if self._hub_remaining_mode else 410, anchor=tk.W)
        if not self._hub_remaining_mode:
            self._tree.column("already_processed", width=140, anchor=tk.CENTER)
        self._tree.tag_configure(
            _ROW_GREYED,
            background="#475569",
            foreground="#e2e8f0",
        )
        self._tree.tag_configure(
            _ROW_POD_COMPLETE,
            background="#166534",
            foreground=UI_FG,
        )

        scroll = dark_tk_scrollbar(tree_outer, tk.VERTICAL, self._tree.yview)
        self._tree.configure(yscrollcommand=scroll.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._pod_poll_after_id = None
        self._apply_pod_completion_layout(initial=True)

        key = api_key_from_env()
        if key and self._row_infos:
            for idx, info in enumerate(self._row_infos):
                num = str(info["tracking_number"])
                if load_cache(num) is not None:
                    continue
                try:
                    result = fetch_tracking_smart(
                        key,
                        num,
                        force_refresh=False,
                        purchase_datetime=info.get("purchase_datetime"),
                    )
                    info["quick_status"] = str(result.get("quick_status_label") or "—")
                    info["carrier"] = str(
                        result.get("carrier_display")
                        or info.get("carrier")
                        or carrier_display_for_number(num)
                    )
                    info["greyed_out"] = bool(result.get("greyed_out"))
                    self._render_row(idx, info)
                except Exception:
                    pass
                self._root.update_idletasks()

        self._apply_pod_completion_layout()

        if self._project_root is not None:
            self._schedule_pod_poll()

        self._tree.bind("<Double-1>", self._on_double)
        self._tree.bind("<Return>", lambda _e: self._activate_row_open_or_capture())
        bind_treeview_copy_menu(
            self._tree,
            self._root,
            extra_commands=[("Force Check Tracking Number", self._force_check_selected_tracking_number)],
        )

        if not self._row_infos:
            ttk.Label(self._frm, text="No tracking numbers in this file.", foreground=UI_FG_DIM).pack()

    def _build_row_infos(self, numbers: list[str]) -> list[dict[str, object]]:
        if self._hub_remaining_mode and self._project_root is not None:
            rows = remaining_pod_candidates(self._project_root)
            out: list[dict[str, object]] = []
            for row in rows:
                num = str(row.get("tracking_number") or "").strip()
                if not num:
                    continue
                if tracking_is_greyed_out(num):
                    continue
                out.append(
                    {
                        "tracking_number": num,
                        "carrier": str(row.get("carrier") or "").strip() or carrier_display_for_number(num),
                        "company": str(row.get("company") or "").strip(),
                        "purchase_datetime": str(row.get("purchase_datetime") or "").strip(),
                        "order_number": str(row.get("order_number") or "").strip(),
                        "quick_status": quick_status_from_cache(num) or "—",
                        "greyed_out": tracking_is_greyed_out(num),
                    }
                )
            return out

        out: list[dict[str, object]] = []
        company = str(self._context.get("company") or "").strip()
        purchase_datetime = str(self._context.get("purchase_datetime") or "").strip()
        order_number = str(self._context.get("order_number") or "").strip()
        for num in numbers:
            s_num = str(num or "").strip()
            if not s_num:
                continue
            if tracking_is_greyed_out(s_num):
                continue
            out.append(
                {
                    "tracking_number": s_num,
                    "carrier": carrier_display_for_number(s_num),
                    "company": company,
                    "purchase_datetime": purchase_datetime,
                    "order_number": order_number,
                    "quick_status": quick_status_from_cache(s_num) or "—",
                    "greyed_out": tracking_is_greyed_out(s_num),
                }
            )
        return out

    def _selected_index(self) -> int | None:
        sel = self._tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except ValueError:
            return None

    def _debug_positional_flag(self) -> str:
        v = (os.getenv("DEBUG_MODE") or "0").strip().lower()
        return "1" if v in ("1", "true", "yes") else "0"

    def _info_for_index(self, idx: int) -> dict[str, object]:
        return self._row_infos[idx]

    def _carrier_url_for_index(self, idx: int) -> str | None:
        info = self._info_for_index(idx)
        num = str(info.get("tracking_number") or "")
        carrier = str(info.get("carrier") or "") or carrier_display_for_number(num)
        return public_tracking_url(num, normalize_carrier_for_public_url(carrier, num))

    def _expected_pdf_path_for_info(self, info: dict[str, object]) -> Path | None:
        if self._project_root is None:
            return None
        num = str(info.get("tracking_number") or "").strip()
        if not num:
            return None
        return expected_pod_pdf_path(
            self._project_root,
            str(info.get("company") or "").strip() or "Unknown",
            str(info.get("purchase_datetime") or "").strip(),
            num,
            str(info.get("carrier") or "").strip() or carrier_display_for_number(num),
        )

    def _is_processed(self, info: dict[str, object]) -> bool:
        pdf_path = self._expected_pdf_path_for_info(info)
        return bool(pdf_path and pdf_path.is_file())

    def _pod_layout_enabled(self) -> bool:
        return self._project_root is not None and (self._hub_remaining_mode or self._pdf_capture)

    def _refresh_all_tree_rows(self) -> None:
        for item in self._tree.get_children():
            self._tree.delete(item)
        for idx, info in enumerate(self._row_infos):
            self._render_row(idx, info)

    def _apply_pod_completion_layout(self, *, initial: bool = False) -> None:
        if self._project_root is None:
            if initial:
                for idx, info in enumerate(self._row_infos):
                    self._render_row(idx, info)
            return

        sel_track: str | None = None
        sel_idx = self._selected_index()
        if sel_idx is not None and 0 <= sel_idx < len(self._row_infos):
            t = str(self._row_infos[sel_idx].get("tracking_number") or "").strip()
            sel_track = t or None

        old_order = tuple(str(x.get("tracking_number") or "") for x in self._row_infos)
        enumerated = list(enumerate(self._row_infos))
        enumerated.sort(key=lambda pair: (1 if self._is_processed(pair[1]) else 0, pair[0]))
        new_infos = [pair[1] for pair in enumerated]
        new_order = tuple(str(x.get("tracking_number") or "") for x in new_infos)
        order_changed = old_order != new_order

        self._row_infos = new_infos
        self._numbers = [str(info["tracking_number"]) for info in self._row_infos]

        for row in self._row_infos:
            row["pod_completed_row"] = bool(self._is_processed(row))

        if order_changed or initial:
            self._refresh_all_tree_rows()
        else:
            for idx, info in enumerate(self._row_infos):
                self._render_row(idx, info)

        if sel_track:
            for i, inf in enumerate(self._row_infos):
                if str(inf.get("tracking_number") or "").strip() == sel_track:
                    self._tree.selection_set(str(i))
                    self._tree.see(str(i))
                    break

    def _cancel_pod_poll(self) -> None:
        if self._pod_poll_after_id is not None:
            try:
                self._root.after_cancel(self._pod_poll_after_id)
            except tk.TclError:
                pass
            self._pod_poll_after_id = None

    def _schedule_pod_poll(self) -> None:
        self._cancel_pod_poll()
        if self._project_root is None:
            return
        delay_ms = 2000 if self._pod_layout_enabled() else 4000

        def tick() -> None:
            if self._project_root is None:
                self._pod_poll_after_id = None
                return
            self._apply_pod_completion_layout()
            self._pod_poll_after_id = self._root.after(delay_ms, tick)

        self._apply_pod_completion_layout()
        self._pod_poll_after_id = self._root.after(delay_ms, tick)

    def _pdf_basename_for_row(self, idx: int) -> str:
        info = self._info_for_index(idx)
        num = str(info.get("tracking_number") or "")
        carrier_ui = str(info.get("carrier") or "") or carrier_display_for_number(num)
        company = sanitize_filename_token(info.get("company") or "Unknown")
        raw_date = str(info.get("purchase_datetime") or "").strip()
        if not raw_date:
            date_tok = "nodate"
        elif len(raw_date) >= 10 and raw_date[4:5] == "-" and raw_date[7:8] == "-":
            date_tok = sanitize_filename_token(raw_date[:10])
        else:
            date_tok = sanitize_filename_token(raw_date.split()[0])
        track_tok = sanitize_filename_token(num)
        car_tok = sanitize_filename_token(carrier_ui)
        return f"DOC {company} {date_tok} {track_tok}_FROM_{car_tok}"

    def _render_row(self, idx: int, info: dict[str, object]) -> None:
        num = str(info.get("tracking_number") or "").strip()
        carrier = str(info.get("carrier") or "").strip() or carrier_display_for_number(num)
        quick_status = str(info.get("quick_status") or quick_status_from_cache(num) or "—")
        values: tuple[object, ...] = (
            idx + 1,
            num,
            carrier,
            quick_status,
        )
        if not self._hub_remaining_mode:
            values = values + ("Yes" if self._is_processed(info) else "No",)
        if self._is_processed(info):
            tags = (_ROW_POD_COMPLETE,)
        elif _quick_status_indicates_notfound(quick_status):
            tags = (_ROW_GREYED,)
        else:
            tags = ()
        if self._tree.exists(str(idx)):
            self._tree.item(str(idx), values=values, tags=tags)
        else:
            self._tree.insert("", tk.END, iid=str(idx), values=values, tags=tags)

    def _toggle_pdf_capture(self) -> None:
        if self._hub_remaining_mode:
            return
        if self._pdf_capture:
            self._pdf_capture = False
            self._pdf_toggle_btn.config(bg=_PDF_TOGGLE_OFF, activebackground="#dc2626")
            self._cancel_pod_poll()
            if self._project_root is not None:
                self._schedule_pod_poll()
            else:
                self._apply_pod_completion_layout()
            return
        ok, err = pdf_capture_environment_ready()
        if not ok:
            messagebox.showerror("PDF capture", err or "Setup incomplete.")
            return
        self._pdf_capture = True
        self._pdf_toggle_btn.config(bg=_PDF_TOGGLE_ON, activebackground="#16a34a")
        self._apply_pod_completion_layout()
        if self._project_root is not None:
            self._schedule_pod_poll()

    def _on_double(self, _event: tk.Event) -> None:
        self._activate_row_open_or_capture()

    def _open_carrier_in_browser_only(self) -> None:
        idx = self._selected_index()
        if idx is None or idx < 0 or idx >= len(self._row_infos):
            return
        url = self._carrier_url_for_index(idx)
        if not url:
            messagebox.showerror("Tracking", "Could not build a carrier URL for this number.")
            return
        webbrowser.open(url)

    def _activate_row_open_or_capture(self) -> None:
        idx = self._selected_index()
        if idx is None or idx < 0 or idx >= len(self._row_infos):
            return
        info = self._info_for_index(idx)
        if self._pdf_capture:
            if bool(info.get("greyed_out")):
                messagebox.showinfo(
                    "PDF capture",
                    "This tracking number is greyed out after a final automatic NotFound retry.\n\n"
                    "Right-click it and choose Force Check Tracking Number to retry 17TRACK.",
                )
                return
            ok, err = pdf_capture_environment_ready()
            if not ok:
                messagebox.showerror("PDF capture", err or "Setup incomplete.")
                return
            expected_pdf = self._expected_pdf_path_for_info(info)
            if expected_pdf is None:
                messagebox.showerror(
                    "PDF capture",
                    "Could not build the expected proof-of-delivery PDF path from BASE_DIR.",
                )
                return
            if expected_pdf.is_file():
                if bool(info.get("pod_completed_row")):
                    messagebox.showinfo(
                        "PDF capture",
                        "That proof-of-delivery file was already generated for this tracking number.",
                    )
                    return
                self._apply_pod_completion_layout()
                return
            url = self._carrier_url_for_index(idx)
            if not url:
                messagebox.showerror("Tracking", "Could not build a carrier URL for this number.")
                return
            base_raw = (os.getenv("BASE_DIR") or "").strip()
            if not base_raw:
                messagebox.showerror(
                    "PDF capture",
                    'BASE_DIR is not set.\nSet "Project folder on disk" in Email Sorter → Settings and Save.',
                )
                return
            pdf_dir = expected_pdf.parent
            try:
                expected_pdf.parent.mkdir(parents=True, exist_ok=True)
            except OSError as e:
                messagebox.showerror("PDF capture", f"Could not create output folder:\n{e}")
                return
            stem = expected_pdf.name
            dbg = self._debug_positional_flag()
            script = _PYTHON_FILES_DIR / "pdfCaptureFromChrome" / "run_pdf_capture.py"
            capture_cwd = _PYTHON_FILES_DIR / "pdfCaptureFromChrome"
            log_hint = PDF_CAPTURE_SESSION_LOG
            try:
                proc = subprocess.Popen(
                    [sys.executable, str(script), url, str(expected_pdf.parent), stem, dbg],
                    cwd=str(capture_cwd),
                )
                try:
                    with open(log_hint, "a", encoding="utf-8", newline="\n") as lf:
                        lf.write(
                            f"{datetime.now().isoformat(timespec='seconds')} "
                            f"[viewer] spawned run_pdf_capture pid={proc.pid} dbg_flag={dbg}\n"
                        )
                except OSError:
                    pass
            except OSError as e:
                messagebox.showerror("PDF capture", f"Could not start capture:\n{e}")
            return

        self._open_carrier_in_browser_only()

    def _view_more(self) -> None:
        idx = self._selected_index()
        if idx is None or idx < 0 or idx >= len(self._row_infos):
            messagebox.showinfo("Details", "Select a row first.")
            return
        num = str(self._info_for_index(idx).get("tracking_number") or "")
        c = load_cache(num)
        track_info = None
        raw_get: dict = {}
        if isinstance(c, dict):
            raw_get = c.get("last_get_response") if isinstance(c.get("last_get_response"), dict) else {}
            track_info = extract_track_info(raw_get, num)
        lines = _events_lines_from_track_info(track_info)
        body = (
            "\n".join(lines)
            if lines
            else "(No milestone events in cached data — open the workbook via the launcher Excel button to refresh caches.)"
        )

        dlg = tk.Toplevel(self._root)
        dlg.title(f"Details — {num}")
        setup_dark_theme(dlg)
        dlg.configure(bg=UI_BG_PANEL)
        dlg.geometry("720x480")
        ttk.Label(dlg, text="Milestones / events", foreground=UI_FG).pack(anchor=tk.W, padx=8, pady=(8, 4))
        t1 = scrolledtext.ScrolledText(dlg, height=12, width=86)
        style_text_widget(t1)
        t1.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        t1.insert(tk.END, body)
        t1.configure(state=tk.DISABLED)

        ttk.Label(dlg, text="Cached gettrackinfo (JSON)", foreground=UI_FG_DIM).pack(anchor=tk.W, padx=8)
        t2 = scrolledtext.ScrolledText(dlg, height=10, width=86)
        style_text_widget(t2)
        t2.pack(fill=tk.BOTH, expand=True, padx=8, pady=(4, 8))
        t2.insert(tk.END, json.dumps(raw_get, indent=2, ensure_ascii=False) if raw_get else "{}")
        t2.configure(state=tk.DISABLED)
        ttk.Button(dlg, text="Close", command=dlg.destroy).pack(anchor=tk.E, padx=8, pady=(0, 8))

    def _force_check_selected_tracking_number(self) -> None:
        idx = self._selected_index()
        if idx is None or idx < 0 or idx >= len(self._row_infos):
            return
        key = api_key_from_env()
        if not key:
            messagebox.showerror(
                "Force Check Tracking Number",
                "No SEVENTEEN_TRACK_API_KEY is configured in .env.",
            )
            return
        info = self._info_for_index(idx)
        num = str(info.get("tracking_number") or "").strip()
        if not num:
            return
        try:
            result = fetch_tracking_smart(
                key,
                num,
                force_refresh=True,
                purchase_datetime=info.get("purchase_datetime"),
            )
        except Exception as exc:
            messagebox.showerror("Force Check Tracking Number", f"17TRACK force check failed:\n{exc}")
            return

        info["quick_status"] = str(result.get("quick_status_label") or "—")
        info["carrier"] = str(result.get("carrier_display") or info.get("carrier") or carrier_display_for_number(num))
        info["greyed_out"] = bool(result.get("greyed_out"))
        if bool(result.get("greyed_out")) and tracking_is_greyed_out(num):
            self._row_infos.pop(idx)
            self._numbers = [str(x["tracking_number"]) for x in self._row_infos]
            self._refresh_all_tree_rows()
        elif self._project_root is not None:
            self._apply_pod_completion_layout()
        else:
            self._render_row(idx, info)

        if bool(result.get("greyed_out")):
            messagebox.showinfo(
                "Force Check Tracking Number",
                "17TRACK still returned NotFound for this tracking number.",
            )
        else:
            messagebox.showinfo(
                "Force Check Tracking Number",
                "17TRACK returned active tracking data and the row was restored.",
            )

    def run(self) -> None:
        self._root.mainloop()


def main() -> int:
    detach_console_win32()
    register_current_aux_gui()

    try:
        from dotenv import load_dotenv

        load_dotenv(_PYTHON_FILES_DIR / ".env", override=False)
    except ImportError:
        pass

    parser = argparse.ArgumentParser(description="View shipping status (17TRACK smart cache).")
    parser.add_argument(
        "number_file",
        nargs="?",
        type=Path,
        help="UTF-8 file: one tracking number per line",
    )
    args = parser.parse_args()

    numbers: list[str] = []
    context: dict[str, str] = {}
    if args.number_file is not None:
        p = args.number_file.expanduser().resolve()
        if not p.is_file():
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Shipping status", f"File not found:\n{p}")
            root.destroy()
            return 1
        try:
            numbers, _ = _load_tracking_file(p)
            context = _load_context_tsv(p.with_suffix(".ctx.tsv"))
        except OSError as e:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Shipping status", f"Could not read file:\n{e}")
            root.destroy()
            return 1

    app = TrackingStatusViewerApp(numbers, context)
    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
