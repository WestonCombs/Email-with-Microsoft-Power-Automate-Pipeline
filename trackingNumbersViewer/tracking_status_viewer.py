"""
Tracking/POD viewer with assisted proof-of-delivery capture.

Reads the same temp file format as ``tracking_numbers_viewer`` (one number per line;
optional ``number<TAB>0|1``). Optional ``.ctx.tsv`` beside the file supplies row context.

Usage:
  python tracking_status_viewer.py <path_to_numbers.txt>
"""
from __future__ import annotations

import argparse
import json
import os
import queue
import re
import sys
import threading
import time
import webbrowser
from collections.abc import Callable
from functools import lru_cache
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, ttk

_PYTHON_FILES_DIR = Path(__file__).resolve().parent.parent
if str(_PYTHON_FILES_DIR) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES_DIR))

from shared.gui_aux_singleton import detach_console_win32, register_current_aux_gui
from shared.gui_treeview_copy import treeview_cell_text, treeview_row_text_tsv
from htmlHandler.carrier_urls import (
    infer_carrier,
    normalize_carrier_for_public_url,
    public_tracking_url,
)
from proofOfDelivery.pod_data import (
    POD_HUB_MODE,
    delete_processed_tracking_artifacts,
    expected_pod_pdf_path,
    first_existing_pod_pdf_path,
    pod_status_viewer_rows,
    project_root_from_env,
    upsert_pod_review_record,
)
from pdfCaptureFromChrome.html_capture import HtmlCaptureController
from pdfCaptureFromChrome.html_capture.hotkey_win32 import CAPTURE_HOTKEY_LABEL
from pdfCaptureFromChrome.html_capture.pod_readiness import pod_readiness_debug_enabled
from pdfCaptureFromChrome.launch_mitm_chrome import terminate_isolated_capture_chrome
from tracking_pdf_audit import audit_path, load_tracking_pdf_audit_entries
from tracking_pdf_capture import (
    read_hands_free_capture_enabled,
    write_hands_free_capture_enabled,
)
from launcher_progress_ui import THEME
from shared.tk_launcher_theme import (
    SettingsStyleSwitch,
    apply_launcher_theme_root,
    configure_launcher_ttk_styles,
    danger_colors,
    launcher_scrollbar,
    make_flat_button,
    settings_label_opts,
    theme_font,
)

_ROW_GREYED = "greyed_out"
_ROW_POD_COMPLETE = "pod_pdf_saved"
_ROW_POD_RESOLVED = "pod_review_resolved"

_NOTFOUND_STATUS_RE = re.compile(r"\bnot[\s_-]?found\b", re.IGNORECASE)


def carrier_display_for_number(tracking_number: object) -> str:
    return infer_carrier(str(tracking_number or "")) or "Unknown"


def quick_status_from_cache(_tracking_number: object) -> str:
    return ""


def tracking_is_greyed_out(_tracking_number: object) -> bool:
    return False


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


def _normalize_tracking_number_text(value: object) -> str:
    return "".join(ch for ch in str(value or "").strip() if not ch.isspace())


def _order_last4_text(value: object) -> str:
    digits = re.sub(r"\D", "", str(value or "").strip())
    if len(digits) >= 4:
        return digits[-4:]
    if digits:
        return digits.zfill(4)
    return "0000"


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


def _unique_sibling_path(path: Path, marker: str) -> Path:
    suffix = path.suffix or ".pdf"
    candidate = path.with_name(f"{path.stem}{marker}{suffix}")
    if not candidate.exists():
        return candidate
    idx = 2
    while True:
        candidate = path.with_name(f"{path.stem}{marker} ({idx}){suffix}")
        if not candidate.exists():
            return candidate
        idx += 1


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


class TrackingStatusViewerApp:
    def __init__(self, numbers: list[str], context: dict[str, str]) -> None:
        self._context = dict(context)
        self._hub_remaining_mode = str(self._context.get("pod_mode") or "").strip() == POD_HUB_MODE
        self._review_rescan_mode = str(self._context.get("pod_review_rescan") or "").strip().lower() in {
            "1",
            "true",
            "yes",
        }
        self._auto_start_pod_capture = str(self._context.get("auto_start_pod_capture") or "").strip().lower() in {
            "1",
            "true",
            "yes",
        }
        try:
            self._project_root = project_root_from_env()
        except Exception:
            self._project_root = None

        if not str(self._context.get("company") or "").strip():
            derived_company = _derive_company_from_project_data(self._context, numbers)
            if derived_company:
                self._context["company"] = derived_company

        self._hands_free_sync_after_id: str | None = None
        self._hands_free_state_syncing = False
        self._hands_free_capture_running = False
        self._pod_excel_sync_running = False
        self._capture_controller: HtmlCaptureController | None = None
        self._batch_capture_active = False
        self._batch_capture_after_id: str | None = None
        self._audit_entries_cache: list[dict[str, object]] = []
        self._audit_entries_mtime_ns: int | None = None
        self._capture_instruction_window: tk.Toplevel | None = None
        self._capture_progress_window: tk.Toplevel | None = None
        self._capture_progress_status_var: tk.StringVar | None = None
        self._capture_progress_bar: ttk.Progressbar | None = None
        self._selector_window: tk.Toplevel | None = None
        self._selector_status_var: tk.StringVar | None = None
        self._selector_time_var: tk.StringVar | None = None
        self._selector_detail_var: tk.StringVar | None = None
        self._selector_target_id: str | None = None
        self._selector_pdf_saved = False
        self._selector_capture_btn: tk.Button | None = None
        self._selector_next_pod_btn: tk.Button | None = None
        self._review_rescan_decision_window: tk.Toplevel | None = None
        self._active_capture_info: dict[str, object] | None = None
        self._active_capture_expected_pdf: Path | None = None
        self._active_capture_output_pdf: Path | None = None

        try:
            from shared.settings_store import apply_runtime_settings_from_json

            apply_runtime_settings_from_json()
        except Exception:
            pass
        try:
            write_hands_free_capture_enabled(False)
        except OSError:
            pass
        try:
            terminate_isolated_capture_chrome()
        except Exception:
            pass

        self._root = tk.Tk()
        self._root.title(
            "Remaining PODs — proof of delivery"
            if self._hub_remaining_mode
            else "Tracking numbers"
        )
        if self._review_rescan_mode:
            self._root.title("POD review rescan")
        self._root.minsize(760, 400)
        self._root.geometry("960x520")
        apply_launcher_theme_root(self._root)
        configure_launcher_ttk_styles(self._root)

        loading_frame = ttk.Frame(self._root, style="Launcher.TFrame", padding=18)
        loading_frame.pack(fill=tk.BOTH, expand=True)
        tk.Label(
            loading_frame,
            text="Loading POD tracking rows...",
            fg=THEME["fg"],
            bg=THEME["bg"],
            font=theme_font("title"),
            anchor=tk.W,
        ).pack(fill=tk.X, anchor=tk.W)
        tk.Label(
            loading_frame,
            text="Checking saved POD PDFs and preparing the capture list.",
            fg=THEME["muted"],
            bg=THEME["bg"],
            font=theme_font("body"),
            anchor=tk.W,
        ).pack(fill=tk.X, anchor=tk.W, pady=(6, 12))
        self._load_status_var = tk.StringVar(value="Starting...")
        tk.Label(
            loading_frame,
            textvariable=self._load_status_var,
            fg=THEME["muted"],
            bg=THEME["bg"],
            font=theme_font("body"),
            anchor=tk.W,
        ).pack(fill=tk.X, anchor=tk.W, pady=(0, 6))
        loading_var = tk.DoubleVar(value=0.0)
        loading_bar = ttk.Progressbar(
            loading_frame,
            mode="determinate",
            variable=loading_var,
            maximum=100,
            length=420,
        )
        loading_bar.pack(fill=tk.X)
        self._root.update_idletasks()

        self._row_infos = self._load_row_infos_with_progress(numbers, loading_var)
        self._numbers = [str(info["tracking_number"]) for info in self._row_infos]

        loading_frame.destroy()

        self._frm = ttk.Frame(self._root, style="Launcher.TFrame", padding=8)
        self._frm.pack(fill=tk.BOTH, expand=True)

        bits = []
        for k in ("company", "category", "purchase_datetime"):
            v = self._context.get(k)
            if v:
                bits.append(f"{k.replace('_', ' ')}: {v}")
        tn = self._context.get("tracking_numbers")
        if tn:
            bits.append(f"tracking numbers: {tn}")
        ctx_line = "  |  ".join(bits) if bits else ""
        if ctx_line:
            tk.Label(
                self._frm,
                text=ctx_line,
                wraplength=920,
                anchor=tk.W,
                fg=THEME["muted"],
                bg=THEME["bg"],
                font=theme_font("body"),
            ).pack(fill=tk.X, anchor=tk.W, pady=(0, 6))

        hint_frame = tk.Frame(self._frm, bg=THEME["bg"])
        hint_frame.pack(fill=tk.X, anchor=tk.W, pady=(0, 4))

        tk.Label(
            hint_frame,
            text=(
                "Rows stay available for POD capture whenever the email parser found a tracking number. "
                "Saved POD PDFs appear in green."
            ),
            wraplength=920,
            anchor=tk.W,
            justify=tk.LEFT,
            fg=THEME["muted"],
            bg=THEME["bg"],
            font=theme_font("body"),
        ).pack(fill=tk.X, anchor=tk.W, pady=(4, 0))

        demo = tk.Frame(hint_frame, bg=THEME["bg"])
        demo.pack(fill=tk.X, anchor=tk.W, pady=(6, 0))
        tk.Label(demo, text="Sample row colors: ", **settings_label_opts()).pack(side=tk.LEFT)
        tk.Label(
            demo,
            text=" GREEN ",
            bg="#166534",
            fg=THEME["fg"],
            padx=6,
            pady=2,
        ).pack(side=tk.LEFT, padx=(0, 4))
        tk.Label(demo, text="processed", **settings_label_opts()).pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(
            demo,
            text=" DARK ",
            bg="#475569",
            fg="#e2e8f0",
            padx=6,
            pady=2,
        ).pack(side=tk.LEFT, padx=(0, 4))
        tk.Label(demo, text="needs POD", **settings_label_opts()).pack(side=tk.LEFT)

        if False:
            tk.Label(
                self._frm,
                text=(
                    "Live tracking validation has been removed. "
                    "Parser-found tracking numbers remain available for POD capture."
                ),
                wraplength=920,
                anchor=tk.W,
                fg="#fbbf24",
                bg=THEME["bg"],
                font=theme_font("body"),
            ).pack(fill=tk.X, anchor=tk.W, pady=(0, 6))

        btn_row = tk.Frame(self._frm, bg=THEME["bg"])
        btn_row.pack(side=tk.BOTTOM, fill=tk.X, pady=(8, 0))
        self._hands_free_pdf_var = tk.IntVar(value=0)
        self._hands_free_pdf_var.trace_add("write", self._on_hands_free_pdf_toggle)
        hf_row = tk.Frame(btn_row, bg=THEME["bg"])
        hf_row.pack(side=tk.LEFT, padx=(0, 8))
        SettingsStyleSwitch(hf_row, self._hands_free_pdf_var).pack(side=tk.LEFT)
        tk.Label(
            hf_row,
            text="Assisted PDF Capture",
            **settings_label_opts(),
        ).pack(side=tk.LEFT, padx=(10, 0))
        self._capture_status_var = tk.StringVar(value="")
        tk.Label(
            btn_row,
            textvariable=self._capture_status_var,
            anchor=tk.W,
            fg=THEME["muted"],
            bg=THEME["bg"],
            font=theme_font("body"),
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        self._update_capture_controls()
        d_bg, d_active = danger_colors()
        make_flat_button(
            btn_row,
            text="Close",
            command=self._on_viewer_closing,
            bg=d_bg,
            active_bg=d_active,
        ).pack(side=tk.RIGHT)

        tree_outer = tk.Frame(self._frm, bg=THEME["bg"])
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
            style="Launcher.Treeview",
        )
        self._tree.heading("idx", text="#")
        self._tree.heading("tracking_number", text="Tracking number")
        self._tree.heading("carrier", text="Carrier")
        self._tree.heading("quick_status", text="POD status")
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
            foreground=THEME["fg"],
        )
        self._tree.tag_configure(
            _ROW_POD_RESOLVED,
            background="#991b1b",
            foreground="#fff1f2",
        )

        scroll = launcher_scrollbar(tree_outer, tk.VERTICAL, self._tree.yview)
        self._tree.configure(yscrollcommand=scroll.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._pod_poll_after_id = None
        self._apply_pod_completion_layout(initial=True)

        key = None
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
        self._schedule_hands_free_state_sync()

        self._tree.bind("<Double-1>", self._on_double)
        self._tree.bind("<Return>", lambda _e: self._activate_row_open_or_capture())
        self._bind_tree_context_menu()

        if not self._row_infos:
            tk.Label(
                self._frm,
                text="No tracking numbers in this file.",
                fg=THEME["muted"],
                bg=THEME["bg"],
                font=theme_font("body"),
            ).pack()

        self._root.protocol("WM_DELETE_WINDOW", self._on_viewer_closing)
        if self._auto_start_pod_capture and self._row_infos:
            try:
                self._tree.selection_set("0")
                self._tree.focus("0")
                self._tree.see("0")
                self._hands_free_state_syncing = True
                try:
                    self._hands_free_pdf_var.set(1)
                finally:
                    self._hands_free_state_syncing = False
                write_hands_free_capture_enabled(True)
                self._set_batch_capture_active(True)
                self._root.after(450, self._start_or_continue_batch_capture)
                self._root.lift()
                self._root.attributes("-topmost", True)
                def clear_topmost() -> None:
                    try:
                        self._root.attributes("-topmost", False)
                    except tk.TclError:
                        pass

                self._root.after(1200, clear_topmost)
            except tk.TclError:
                self._hands_free_state_syncing = False
                pass

    def _on_viewer_closing(self) -> None:
        try:
            write_hands_free_capture_enabled(False)
        except OSError:
            pass
        self._set_batch_capture_active(False)
        self._stop_capture_controller()
        self._destroy_review_rescan_decision_window()
        self._cancel_pod_poll()
        if self._hands_free_sync_after_id is not None:
            try:
                self._root.after_cancel(self._hands_free_sync_after_id)
            except tk.TclError:
                pass
            self._hands_free_sync_after_id = None
        try:
            self._root.destroy()
        except tk.TclError:
            pass

    def _load_row_infos_with_progress(
        self,
        numbers: list[str],
        loading_var: tk.DoubleVar,
    ) -> list[dict[str, object]]:
        work_q: queue.Queue[tuple[str, object]] = queue.Queue()

        def progress(done: int, total: int) -> None:
            work_q.put(("progress", (done, total)))

        def worker() -> None:
            try:
                work_q.put(("done", self._build_row_infos(numbers, progress_callback=progress)))
            except Exception as exc:
                work_q.put(("error", exc))

        threading.Thread(target=worker, daemon=True).start()
        rows: list[dict[str, object]] | None = None
        error: Exception | None = None
        while rows is None and error is None:
            try:
                while True:
                    kind, payload = work_q.get_nowait()
                    if kind == "progress":
                        done, total = payload  # type: ignore[misc]
                        total_int = max(int(total), 1)
                        done_int = max(0, min(int(done), total_int))
                        pct = int(100 * done_int / total_int)
                        loading_var.set(pct)
                        self._load_status_var.set(
                            f"Scanned {done_int} of {total_int} order records..."
                        )
                    elif kind == "done":
                        rows = payload  # type: ignore[assignment]
                    elif kind == "error":
                        error = payload if isinstance(payload, Exception) else Exception(str(payload))
            except queue.Empty:
                pass
            try:
                self._root.update()
            except tk.TclError:
                return []
            if rows is None and error is None:
                time.sleep(0.03)
        if error is not None:
            messagebox.showerror("Tracking", f"Could not load POD tracking rows:\n{error}")
            return []
        loading_var.set(100)
        self._load_status_var.set("Done.")
        try:
            self._root.update_idletasks()
        except tk.TclError:
            pass
        return rows or []

    def _build_row_infos(
        self,
        numbers: list[str],
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> list[dict[str, object]]:
        if self._hub_remaining_mode and self._project_root is not None:
            rows = pod_status_viewer_rows(self._project_root, progress_callback=progress_callback)
            out: list[dict[str, object]] = []
            for row in rows:
                num = str(row.get("tracking_number") or "").strip()
                if not num:
                    continue
                out.append(
                    {
                        "tracking_number": num,
                        "carrier": str(row.get("carrier") or "").strip() or carrier_display_for_number(num),
                        "company": str(row.get("company") or "").strip(),
                        "purchase_datetime": str(row.get("purchase_datetime") or "").strip(),
                        "order_number": str(row.get("order_number") or "").strip(),
                        "category": str(row.get("category") or "").strip(),
                        "quick_status": "Needs POD",
                        "pod_review_status": str(row.get("pod_review_status") or "").strip(),
                        "pod_review_resolved": bool(row.get("pod_review_resolved")),
                        "pod_review_skipped_eternally": bool(row.get("pod_review_skipped_eternally")),
                        "greyed_out": bool(row.get("greyed_out")),
                    }
                )
            return out

        out: list[dict[str, object]] = []
        company = str(self._context.get("company") or "").strip()
        purchase_datetime = str(self._context.get("purchase_datetime") or "").strip()
        order_number = str(self._context.get("order_number") or "").strip()
        category = str(self._context.get("email_category") or self._context.get("category") or "").strip()
        total = len(numbers)
        if progress_callback is not None:
            progress_callback(0, total)
        for idx, num in enumerate(numbers, start=1):
            s_num = str(num or "").strip()
            if not s_num:
                if progress_callback is not None:
                    progress_callback(idx, total)
                continue
            out.append(
                {
                    "tracking_number": s_num,
                    "carrier": carrier_display_for_number(s_num),
                    "company": company,
                    "purchase_datetime": purchase_datetime,
                    "order_number": order_number,
                    "category": category,
                    "quick_status": "Needs POD",
                    "greyed_out": tracking_is_greyed_out(s_num),
                }
            )
            if progress_callback is not None:
                progress_callback(idx, total)
        return out

    def _selected_index(self) -> int | None:
        sel = self._tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except ValueError:
            return None

    def _info_for_index(self, idx: int) -> dict[str, object]:
        return self._row_infos[idx]

    def _clipboard_set(self, text: str) -> None:
        try:
            self._root.clipboard_clear()
            self._root.clipboard_append(text)
            self._root.update_idletasks()
        except tk.TclError:
            pass

    def _copy_context_cell(self) -> None:
        item = getattr(self, "_menu_item", None)
        col = getattr(self, "_menu_col", None)
        if not item or not col:
            return
        self._clipboard_set(treeview_cell_text(self._tree, item, col))

    def _copy_context_row(self) -> None:
        item = getattr(self, "_menu_item", None)
        if not item:
            return
        self._clipboard_set(treeview_row_text_tsv(self._tree, item))

    def _context_menu_index(self) -> int | None:
        item = getattr(self, "_menu_item", None)
        if not item:
            return None
        try:
            idx = int(item)
        except ValueError:
            return None
        if 0 <= idx < len(self._row_infos):
            return idx
        return None

    def _context_menu_can_delete(self) -> bool:
        idx = self._context_menu_index()
        if idx is None:
            return False
        return self._is_processed(self._info_for_index(idx))

    def _context_menu_review_pdf(self) -> Path | None:
        idx = self._context_menu_index()
        if idx is None:
            return None
        return self._review_pdf_path_for_info(self._info_for_index(idx))

    def _bind_tree_context_menu(self) -> None:
        self._menu_item = None
        self._menu_col = None
        self._tree_menu = tk.Menu(self._root, tearoff=0)

        def on_button(event: tk.Event) -> None:
            region = self._tree.identify_region(event.x, event.y)
            if region not in ("cell", "tree"):
                return
            item = self._tree.identify_row(event.y)
            col = self._tree.identify_column(event.x)
            if not item or not col:
                return
            try:
                self._tree.selection_set(item)
                self._tree.focus(item)
            except tk.TclError:
                pass
            self._menu_item = item
            self._menu_col = col

            self._tree_menu.delete(0, tk.END)
            review_pdf = self._context_menu_review_pdf()
            if review_pdf is not None:
                self._tree_menu.add_command(label="Open review PDF", command=self._open_selected_review_pdf)
            self._tree_menu.add_command(label="Open tracking page", command=self._open_carrier_in_browser_only)
            if self._context_menu_can_delete():
                self._tree_menu.add_command(label="Delete", command=self._delete_selected_processed_row)
            self._tree_menu.add_separator()
            self._tree_menu.add_command(label="Copy cell", command=self._copy_context_cell)
            self._tree_menu.add_command(label="Copy row", command=self._copy_context_row)
            try:
                self._tree_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._tree_menu.grab_release()

        self._tree.bind("<Button-3>", on_button)
        if sys.platform == "darwin":
            self._tree.bind("<Button-2>", on_button)
            self._tree.bind("<Control-Button-1>", on_button)

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

    def _load_tracking_pdf_audit_entries(self) -> list[dict[str, object]]:
        try:
            path = audit_path()
        except Exception:
            return []
        try:
            stat = path.stat()
        except OSError:
            self._audit_entries_cache = []
            self._audit_entries_mtime_ns = None
            return []
        if self._audit_entries_mtime_ns == stat.st_mtime_ns:
            return self._audit_entries_cache
        entries = load_tracking_pdf_audit_entries()
        self._audit_entries_cache = [entry for entry in entries if isinstance(entry, dict)]
        self._audit_entries_mtime_ns = stat.st_mtime_ns
        return self._audit_entries_cache

    def _audit_entry_matches_info(self, info: dict[str, object], entry: dict[str, object]) -> bool:
        track_num = _normalize_tracking_number_text(info.get("tracking_number"))
        if track_num:
            entry_track = _normalize_tracking_number_text(entry.get("tracking_number"))
            if entry_track:
                return entry_track == track_num

        company = _normalize_ctx_text(info.get("company"))
        order_last4 = _order_last4_text(info.get("order_number"))
        category = _normalize_ctx_text(info.get("category"))
        entry_company = _normalize_ctx_text(entry.get("company"))
        entry_last4 = _order_last4_text(entry.get("order_number") or entry.get("order_last4"))
        entry_category = _normalize_ctx_text(entry.get("category"))
        return (
            bool(company)
            and company == entry_company
            and order_last4 == entry_last4
            and category == entry_category
        )

    def _audited_pdf_path_for_info(self, info: dict[str, object]) -> Path | None:
        for entry in self._load_tracking_pdf_audit_entries():
            if not self._audit_entry_matches_info(info, entry):
                continue
            if entry.get("latest_tracking_info_visible") is False:
                continue
            raw_path = str(entry.get("path") or "").strip()
            if not raw_path:
                continue
            path = Path(raw_path).expanduser().resolve()
            if path.is_file():
                return path
        return None

    def _review_pdf_path_for_info(self, info: dict[str, object]) -> Path | None:
        for entry in self._load_tracking_pdf_audit_entries():
            if not self._audit_entry_matches_info(info, entry):
                continue
            if entry.get("latest_tracking_info_visible") is not False:
                continue
            raw_path = str(entry.get("path") or "").strip()
            if not raw_path:
                continue
            try:
                path = Path(raw_path).expanduser().resolve()
            except OSError:
                continue
            if path.is_file():
                return path
        return None

    def _processed_pdf_path_for_info(self, info: dict[str, object]) -> Path | None:
        if self._has_grey_status(info):
            return None
        expected_pdf = self._expected_pdf_path_for_info(info)
        if expected_pdf is not None and expected_pdf.is_file():
            return expected_pdf
        if self._project_root is not None:
            num = str(info.get("tracking_number") or "").strip()
            if num:
                existing_pdf = first_existing_pod_pdf_path(
                    self._project_root,
                    str(info.get("company") or "").strip() or "Unknown",
                    str(info.get("purchase_datetime") or "").strip(),
                    num,
                    str(info.get("carrier") or "").strip() or carrier_display_for_number(num),
                )
                if existing_pdf is not None:
                    return existing_pdf
        return self._audited_pdf_path_for_info(info)

    def _is_processed(self, info: dict[str, object]) -> bool:
        return self._processed_pdf_path_for_info(info) is not None

    def _is_resolved_review(self, info: dict[str, object]) -> bool:
        return bool(info.get("pod_review_resolved")) or (
            str(info.get("pod_review_status") or "").strip().lower() == "resolved"
        )

    def _has_grey_status(self, info: dict[str, object]) -> bool:
        if self._is_resolved_review(info):
            return False
        return self._review_pdf_path_for_info(info) is not None

    def _is_grey_row(self, info: dict[str, object]) -> bool:
        return (not self._is_processed(info)) and self._has_grey_status(info)

    def _row_sort_key(self, pair: tuple[int, dict[str, object]]) -> tuple[int, int]:
        idx, info = pair
        if self._is_processed(info):
            bucket = 1
        elif self._is_resolved_review(info):
            bucket = 2
        elif self._has_grey_status(info):
            bucket = 3
        else:
            bucket = 0
        return (bucket, idx)

    def _pod_layout_enabled(self) -> bool:
        return self._project_root is not None

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
        enumerated.sort(key=self._row_sort_key)
        new_infos = [pair[1] for pair in enumerated]
        new_order = tuple(str(x.get("tracking_number") or "") for x in new_infos)
        order_changed = old_order != new_order

        self._row_infos = new_infos
        self._numbers = [str(info["tracking_number"]) for info in self._row_infos]

        for row in self._row_infos:
            row["pod_completed_row"] = bool(self._is_processed(row))
            row["greyed_out"] = bool(self._has_grey_status(row))

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

    def _render_row(self, idx: int, info: dict[str, object]) -> None:
        num = str(info.get("tracking_number") or "").strip()
        carrier = str(info.get("carrier") or "").strip() or carrier_display_for_number(num)
        if self._is_processed(info):
            info["quick_status"] = "POD saved"
        elif self._is_resolved_review(info):
            info["quick_status"] = "Resolved"
        elif self._has_grey_status(info):
            info["quick_status"] = "Needs manual review"
        else:
            info["quick_status"] = str(info.get("quick_status") or "Needs POD")
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
        elif self._is_resolved_review(info):
            tags = (_ROW_POD_RESOLVED,)
        elif self._is_grey_row(info):
            tags = (_ROW_GREYED,)
        else:
            tags = ()
        if self._tree.exists(str(idx)):
            self._tree.item(str(idx), values=values, tags=tags)
        else:
            self._tree.insert("", tk.END, iid=str(idx), values=values, tags=tags)

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

    def _open_selected_review_pdf(self) -> None:
        idx = self._selected_index()
        if idx is None or idx < 0 or idx >= len(self._row_infos):
            return
        path = self._review_pdf_path_for_info(self._info_for_index(idx))
        if path is None:
            self._set_capture_status("No review PDF is saved for this row.")
            return
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception:
            webbrowser.open(path.as_uri())

    def _delete_selected_processed_row(self) -> None:
        idx = self._selected_index()
        if idx is None or idx < 0 or idx >= len(self._row_infos):
            return
        if self._project_root is None:
            messagebox.showerror("Delete", "Project data directory is unavailable.")
            return
        info = self._info_for_index(idx)
        if not self._is_processed(info):
            return

        processed_path = self._processed_pdf_path_for_info(info)
        result = delete_processed_tracking_artifacts(
            self._project_root,
            tracking_number=info.get("tracking_number"),
            company=info.get("company"),
            purchase_datetime=info.get("purchase_datetime"),
            carrier_display=info.get("carrier"),
            order_number=info.get("order_number"),
            category=info.get("category"),
            processed_pdf_path=processed_path,
        )
        self._audit_entries_cache = []
        self._audit_entries_mtime_ns = None
        self._apply_pod_completion_layout()
        deleted_pdfs = int(result.get("deleted_pdfs", 0))
        removed_refs = int(result.get("removed_pod_records", 0)) + int(
            result.get("removed_audit_entries", 0)
        )
        if deleted_pdfs > 0 or removed_refs > 0:
            self._set_capture_status("Deleted. Row is now unprocessed.")
        else:
            self._set_capture_status("Nothing was deleted for this row.")

    def _record_for_hands_free_capture(self, info: dict[str, object]) -> dict:
        """Build the record shape expected by the convention filename builder."""
        record = dict(self._context)
        num = str(info.get("tracking_number") or "").strip()
        carrier = str(info.get("carrier") or "").strip() or carrier_display_for_number(num)
        category = str(
            info.get("category")
            or record.get("email_category")
            or record.get("category")
            or "Unknown"
        ).strip()
        record.update(
            {
                "company": str(info.get("company") or record.get("company") or "Unknown").strip(),
                "order_number": str(info.get("order_number") or record.get("order_number") or "").strip(),
                "purchase_datetime": str(
                    info.get("purchase_datetime") or record.get("purchase_datetime") or ""
                ).strip(),
                "email_category": category or "Unknown",
                "category": category or "Unknown",
                "tracking_number": num,
                "tracking_numbers": [num] if num else [],
                "carrier": carrier,
            }
        )
        return record

    def _on_hands_free_pdf_toggle(self, *_args: object) -> None:
        """Persist this switch so both status windows share the same assisted-capture mode."""
        if self._hands_free_state_syncing:
            return
        enabled = bool(self._hands_free_pdf_var.get())
        try:
            write_hands_free_capture_enabled(enabled)
        except OSError:
            pass
        if not enabled:
            self._set_batch_capture_active(False)
            self._stop_capture_controller()
            self._set_capture_status("")
        else:
            self._set_batch_capture_active(True)
            self._set_capture_status("Opening the next eligible POD row...")
            try:
                self._root.after(0, self._start_or_continue_batch_capture)
            except tk.TclError:
                pass
        self._update_capture_controls()

    def _sync_hands_free_state_from_disk(self) -> None:
        enabled = read_hands_free_capture_enabled(default=bool(self._hands_free_pdf_var.get()))
        if enabled == bool(self._hands_free_pdf_var.get()):
            return
        self._hands_free_state_syncing = True
        try:
            self._hands_free_pdf_var.set(1 if enabled else 0)
        finally:
            self._hands_free_state_syncing = False
        if not enabled:
            self._set_batch_capture_active(False)
            self._stop_capture_controller()
            self._set_capture_status("")
        self._update_capture_controls()

    def _schedule_hands_free_state_sync(self) -> None:
        if self._hands_free_sync_after_id is not None:
            try:
                self._root.after_cancel(self._hands_free_sync_after_id)
            except tk.TclError:
                pass
            self._hands_free_sync_after_id = None

        def tick() -> None:
            self._sync_hands_free_state_from_disk()
            self._hands_free_sync_after_id = self._root.after(1200, tick)

        self._hands_free_sync_after_id = self._root.after(1200, tick)

    def _update_capture_controls(self) -> None:
        return

    def _set_capture_status(self, message: str) -> None:
        var = getattr(self, "_capture_status_var", None)
        if var is None:
            return
        clean = " ".join(str(message or "").split())
        try:
            var.set(clean[:180])
        except tk.TclError:
            pass

    def _place_capture_popup(self, win: tk.Toplevel) -> None:
        try:
            win.update_idletasks()
            self._root.update_idletasks()
            x = self._root.winfo_rootx() + (self._root.winfo_width() // 2) - (win.winfo_width() // 2)
            y = self._root.winfo_rooty() + 72
            win.geometry(f"+{max(0, x)}+{max(0, y)}")
            win.lift()
            win.attributes("-topmost", True)
            def clear_topmost() -> None:
                try:
                    win.attributes("-topmost", False)
                except tk.TclError:
                    pass

            win.after(900, clear_topmost)
        except tk.TclError:
            pass

    def _destroy_capture_instruction_window(self) -> None:
        win = self._capture_instruction_window
        self._capture_instruction_window = None
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass

    def _destroy_capture_progress_window(self) -> None:
        bar = self._capture_progress_bar
        self._capture_progress_bar = None
        if bar is not None:
            try:
                bar.stop()
            except tk.TclError:
                pass
        win = self._capture_progress_window
        self._capture_progress_window = None
        self._capture_progress_status_var = None
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass

    def _destroy_selector_window(self) -> None:
        win = self._selector_window
        self._selector_window = None
        self._selector_status_var = None
        self._selector_time_var = None
        self._selector_detail_var = None
        self._selector_target_id = None
        self._selector_pdf_saved = False
        self._selector_capture_btn = None
        self._selector_next_pod_btn = None
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass

    def _destroy_review_rescan_decision_window(self) -> None:
        win = self._review_rescan_decision_window
        self._review_rescan_decision_window = None
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass

    def _destroy_capture_popups(self, *, keep_progress: bool = False) -> None:
        self._destroy_capture_instruction_window()
        if not keep_progress:
            self._destroy_capture_progress_window()

    def _keep_batch_progress_window(self) -> bool:
        try:
            hands_free = bool(self._hands_free_pdf_var.get())
        except (AttributeError, tk.TclError):
            hands_free = False
        return (
            self._batch_capture_active
            and hands_free
            and not self._review_rescan_mode
            and not pod_readiness_debug_enabled()
        )

    def _skip_review_from_capture(self) -> None:
        if self._project_root is None:
            self._set_capture_status("Project data directory is unavailable.")
            return
        record_id = str(self._context.get("fix_flagged_record_identity") or "").strip()
        request_id = str(self._context.get("fix_flagged_request_id") or "").strip()
        if not record_id:
            self._set_capture_status("This review item cannot be matched for skip.")
            return
        try:
            from shared.fix_flagged import skip_record_eternally

            self._set_batch_capture_active(False)
            self._stop_capture_controller()
            skip_record_eternally(
                self._project_root,
                record_id=record_id,
                request_id=request_id or None,
            )
            self._set_capture_status("Skipped eternally and removed from Fix Flagged.")
            try:
                self._root.after(650, self._on_viewer_closing)
            except tk.TclError:
                pass
        except Exception as exc:
            self._set_capture_status(f"Could not skip this review item: {exc}")

    def _add_review_rescan_skip_prompt(self, parent: tk.Misc) -> None:
        if not self._review_rescan_mode:
            return
        prompt = tk.Frame(parent, bg=THEME["surface"], padx=12, pady=10)
        prompt.pack(fill=tk.X, pady=(12, 0))
        tk.Label(
            prompt,
            text="This POD capture needs review. If this email/tracking number is bad, skip it eternally.",
            wraplength=420,
            justify=tk.LEFT,
            fg=THEME["fg"],
            bg=THEME["surface"],
            font=theme_font("body"),
            anchor=tk.W,
        ).pack(fill=tk.X, anchor=tk.W)
        row = tk.Frame(prompt, bg=THEME["surface"])
        row.pack(fill=tk.X, pady=(10, 0))
        make_flat_button(
            row,
            text="Skip Eternally",
            command=self._skip_review_from_capture,
            bg=THEME["stop_fg"],
            active_bg="#da3633",
            width=14,
        ).pack(side=tk.LEFT, padx=(0, 8))
        make_flat_button(
            row,
            text="Keep Trying",
            command=self._retry_review_rescan,
            bg=THEME["track"],
            active_bg=THEME["border"],
            fg=THEME["fg"],
            active_fg=THEME["fg"],
            width=12,
        ).pack(side=tk.LEFT)

    def _review_rescan_record_id(self) -> str:
        return str(self._context.get("fix_flagged_record_identity") or "").strip()

    def _review_rescan_request_id(self) -> str:
        return str(self._context.get("fix_flagged_request_id") or "").strip()

    def _open_review_rescan_pdf(self, path: Path | None) -> None:
        if path is None or not path.is_file():
            self._set_capture_status("The new scan PDF is not available yet.")
            return
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception:
            webbrowser.open(path.as_uri())

    def _show_review_rescan_decision(self, scanned_pdf: Path | None, *, reason: str = "") -> None:
        if not self._review_rescan_mode:
            return
        path = scanned_pdf
        if path is not None:
            try:
                path = path.expanduser().resolve()
            except OSError:
                pass
        self._active_capture_output_pdf = path
        self._destroy_review_rescan_decision_window()
        status_reason = " ".join(str(reason or "").split())
        self._set_capture_status("Review the new scan, then replace the current POD or rescan again.")
        try:
            win = tk.Toplevel(self._root)
            win.title("Review Manual POD Rescan")
            win.configure(bg=THEME["bg"])
            win.resizable(False, False)
            win.transient(self._root)
            win.protocol("WM_DELETE_WINDOW", lambda: None)
            frame = tk.Frame(win, bg=THEME["bg"], padx=18, pady=16)
            frame.pack(fill=tk.BOTH, expand=True)
            tk.Label(
                frame,
                text="Use this scan?",
                fg=THEME["fg"],
                bg=THEME["bg"],
                font=theme_font("title"),
                anchor=tk.W,
            ).pack(fill=tk.X, anchor=tk.W)
            detail = "A new POD scan was saved. Replace the current reviewed POD only if this one is correct."
            if status_reason:
                detail += f"\n\nAudit note: {status_reason}"
            if path is not None:
                detail += f"\n\nNew scan: {path.name}"
            tk.Label(
                frame,
                text=detail,
                wraplength=520,
                justify=tk.LEFT,
                fg=THEME["fg"],
                bg=THEME["bg"],
                font=theme_font("body"),
                anchor=tk.W,
            ).pack(fill=tk.X, anchor=tk.W, pady=(8, 12))
            actions = tk.Frame(frame, bg=THEME["bg"])
            actions.pack(fill=tk.X)
            make_flat_button(
                actions,
                text="Replace Current POD",
                command=self._accept_review_rescan,
                bg=THEME["excel_accent"],
                active_bg=THEME["excel_accent_dim"],
                width=20,
            ).pack(side=tk.LEFT, padx=(0, 8))
            make_flat_button(
                actions,
                text="Open New Scan",
                command=lambda p=path: self._open_review_rescan_pdf(p),
                bg=THEME["surface"],
                active_bg=THEME["track"],
                fg=THEME["fg"],
                active_fg=THEME["fg"],
                width=14,
            ).pack(side=tk.LEFT, padx=(0, 8))
            make_flat_button(
                actions,
                text="Rescan Again",
                command=self._retry_review_rescan,
                bg=THEME["run_accent"],
                active_bg=THEME["run_accent_dim"],
                width=14,
            ).pack(side=tk.LEFT, padx=(0, 8))
            make_flat_button(
                actions,
                text="Skip Eternally",
                command=self._skip_review_from_capture,
                bg=THEME["stop_fg"],
                active_bg="#da3633",
                width=14,
            ).pack(side=tk.RIGHT)
            self._review_rescan_decision_window = win
            self._place_capture_popup(win)
        except tk.TclError:
            self._review_rescan_decision_window = None

    def _accept_review_rescan(self) -> None:
        if self._project_root is None:
            self._set_capture_status("Project data directory is unavailable.")
            return
        record_id = self._review_rescan_record_id()
        if not record_id:
            self._set_capture_status("This review item cannot be matched for replacement.")
            return
        scan_path = self._active_capture_output_pdf
        if scan_path is None:
            self._set_capture_status("No manual rescan PDF is ready to replace the current POD.")
            return
        try:
            from shared.fix_flagged import resolve_pod_review_with_manual_scan

            result = resolve_pod_review_with_manual_scan(
                self._project_root,
                record_id=record_id,
                scanned_pdf_path=scan_path,
                request_id=self._review_rescan_request_id() or None,
            )
            self._destroy_review_rescan_decision_window()
            self._set_batch_capture_active(False)
            self._stop_capture_controller()
            self._set_capture_status(f"Replaced current POD with {Path(str(result.get('final_pdf') or scan_path)).name}.")
            try:
                self._root.after(900, self._on_viewer_closing)
            except tk.TclError:
                pass
        except Exception as exc:
            self._set_capture_status(f"Could not replace current POD: {exc}")
            messagebox.showerror("POD review rescan", str(exc), parent=self._root)

    def _retry_review_rescan(self) -> None:
        if not self._review_rescan_mode:
            return
        path = self._active_capture_output_pdf
        if path is not None:
            try:
                expected = self._active_capture_expected_pdf.resolve() if self._active_capture_expected_pdf else None
                if path.is_file() and (expected is None or path.resolve() != expected):
                    path.unlink()
            except OSError:
                pass
        self._destroy_review_rescan_decision_window()
        self._stop_capture_controller()
        self._hands_free_state_syncing = True
        try:
            self._hands_free_pdf_var.set(1)
        finally:
            self._hands_free_state_syncing = False
        self._set_batch_capture_active(True)
        self._set_capture_status("Opening this POD again for a fresh manual rescan...")
        try:
            self._root.after(250, self._start_or_continue_batch_capture)
        except tk.TclError:
            pass

    def _show_capture_instruction_window(self, info: dict[str, object]) -> None:
        self._destroy_capture_progress_window()
        self._destroy_capture_instruction_window()
        num = str(info.get("tracking_number") or "").strip()
        carrier = str(info.get("carrier") or "").strip() or carrier_display_for_number(num)
        debug_selector = pod_readiness_debug_enabled()
        title = "Starting POD selector debug" if debug_selector else "Starting assisted PDF capture"
        try:
            win = tk.Toplevel(self._root)
            win.title("Assisted PDF Capture")
            win.configure(bg=THEME["bg"])
            win.resizable(False, False)
            win.transient(self._root)
            win.protocol("WM_DELETE_WINDOW", lambda: None)
            frame = tk.Frame(win, bg=THEME["bg"], padx=18, pady=16)
            frame.pack(fill=tk.BOTH, expand=True)
            tk.Label(
                frame,
                text=title,
                fg=THEME["fg"],
                bg=THEME["bg"],
                font=theme_font("title"),
                anchor=tk.W,
            ).pack(fill=tk.X, anchor=tk.W)
            detail = f"{carrier} {num}".strip()
            if detail:
                tk.Label(
                    frame,
                    text=detail,
                    fg=THEME["muted"],
                    bg=THEME["bg"],
                    font=theme_font("body"),
                    anchor=tk.W,
                ).pack(fill=tk.X, anchor=tk.W, pady=(5, 0))
            tk.Label(
                frame,
                text=(
                    (
                        "Chrome will stay open. Use the POD Element Selector controls "
                        "to move the blue box and capture the selected element."
                    )
                    if debug_selector
                    else (
                        "The app will wait for the shipping information, then save the POD PDF automatically."
                    )
                ),
                wraplength=430,
                justify=tk.LEFT,
                fg=THEME["fg"],
                bg=THEME["bg"],
                font=theme_font("body"),
                anchor=tk.W,
            ).pack(fill=tk.X, anchor=tk.W, pady=(12, 0))
            self._add_review_rescan_skip_prompt(frame)
            self._capture_instruction_window = win
            self._place_capture_popup(win)
        except tk.TclError:
            self._capture_instruction_window = None

    def _show_capture_progress_window(self, message: str) -> None:
        self._destroy_capture_instruction_window()
        clean = " ".join(str(message or "").split()) or "Capture trigger received."
        if self._capture_progress_window is not None and self._capture_progress_status_var is not None:
            try:
                self._capture_progress_status_var.set(clean)
                self._capture_progress_window.lift()
            except tk.TclError:
                self._destroy_capture_progress_window()
            return
        try:
            win = tk.Toplevel(self._root)
            win.title("PDF Capture Progress")
            win.configure(bg=THEME["bg"])
            win.resizable(False, False)
            win.transient(self._root)
            win.protocol("WM_DELETE_WINDOW", lambda: None)
            frame = tk.Frame(win, bg=THEME["bg"], padx=18, pady=16)
            frame.pack(fill=tk.BOTH, expand=True)
            tk.Label(
                frame,
                text="Capture trigger received",
                fg=THEME["fg"],
                bg=THEME["bg"],
                font=theme_font("title"),
                anchor=tk.W,
            ).pack(fill=tk.X, anchor=tk.W)
            status_var = tk.StringVar(value=clean)
            tk.Label(
                frame,
                textvariable=status_var,
                wraplength=430,
                justify=tk.LEFT,
                fg=THEME["muted"],
                bg=THEME["bg"],
                font=theme_font("body"),
                anchor=tk.W,
            ).pack(fill=tk.X, anchor=tk.W, pady=(8, 12))
            bar = ttk.Progressbar(frame, mode="indeterminate", length=420)
            bar.pack(fill=tk.X)
            self._add_review_rescan_skip_prompt(frame)
            bar.start(12)
            self._capture_progress_window = win
            self._capture_progress_status_var = status_var
            self._capture_progress_bar = bar
            self._place_capture_popup(win)
        except tk.TclError:
            self._capture_progress_window = None
            self._capture_progress_status_var = None
            self._capture_progress_bar = None

    def _selector_text_from_status(self, status: dict | None) -> tuple[str, str, str]:
        if not isinstance(status, dict):
            return ("Waiting for selector...", "", "")
        current = status.get("current")
        if not isinstance(current, dict):
            current = {}
        readiness = status.get("readiness")
        if not isinstance(readiness, dict):
            readiness = {}
        total = int(status.get("total") or 0)
        index = int(status.get("index") or 0)
        selector = str(current.get("selector") or "").strip()
        tag = str(current.get("tag") or "").strip()
        text = " ".join(str(current.get("text") or "").split())
        head = f"Element {index + 1} of {total}" if total else "Best visible element"
        if tag:
            head += f"  |  {tag}"
        if status.get("pdf_saved"):
            head = "PDF saved. Use Next POD when ready."
        elif status.get("approved"):
            head = "Selection accepted. Saving PDF..."

        def seconds(raw: object) -> str | None:
            try:
                value = float(raw)
            except (TypeError, ValueError):
                return None
            if value < 0:
                return None
            return f"{value:.1f}s"

        def ms_seconds(raw: object) -> str | None:
            try:
                value = float(raw) / 1000.0
            except (TypeError, ValueError):
                return None
            if value < 0:
                return None
            return f"{value:.1f}s"

        timing_bits = []
        mode = str(readiness.get("mode") or "").strip()
        if mode:
            timing_bits.append(f"mode: {mode}")
        elapsed = seconds(readiness.get("elapsed_seconds"))
        if elapsed:
            timing_bits.append(f"ready wait: {elapsed}")
        quiet = seconds(readiness.get("quiet_seconds"))
        if quiet:
            timing_bits.append(f"quiet: {quiet}")
        first_seen = ms_seconds(current.get("first_seen_ms"))
        last_seen = ms_seconds(current.get("last_seen_ms"))
        if first_seen:
            timing_bits.append(f"element first seen: {first_seen}")
        if last_seen:
            timing_bits.append(f"last changed: {last_seen}")
        if not first_seen and not last_seen:
            page_elapsed = ms_seconds(current.get("page_elapsed_ms"))
            timing_bits.append(
                f"element timing: not observed directly{f' (page sampled at {page_elapsed})' if page_elapsed else ''}"
            )
        event_count = readiness.get("event_count")
        if event_count not in (None, ""):
            timing_bits.append(f"DOM/CSS events: {event_count}")
        timing = "  |  ".join(timing_bits)

        detail_bits = []
        if text:
            detail_bits.append(text[:220])
        if selector:
            detail_bits.append(selector[:220])
        return (head, timing, "\n".join(detail_bits))

    def _update_selector_window(self, status: dict | None = None) -> None:
        if status is None and self._selector_target_id and self._capture_controller is not None:
            status = self._capture_controller.debug_selector_status(self._selector_target_id)
        head, timing, detail = self._selector_text_from_status(status)
        if self._selector_status_var is not None:
            try:
                self._selector_status_var.set(head)
            except tk.TclError:
                pass
        if self._selector_time_var is not None:
            try:
                self._selector_time_var.set(timing)
            except tk.TclError:
                pass
        if self._selector_detail_var is not None:
            try:
                self._selector_detail_var.set(detail)
            except tk.TclError:
                pass
        saved = bool((status or {}).get("pdf_saved"))
        self._selector_pdf_saved = saved
        if self._selector_capture_btn is not None:
            try:
                self._selector_capture_btn.config(state=tk.DISABLED if saved or bool((status or {}).get("approved")) else tk.NORMAL)
            except tk.TclError:
                pass
        if self._selector_next_pod_btn is not None:
            try:
                self._selector_next_pod_btn.config(state=tk.NORMAL if saved else tk.DISABLED)
            except tk.TclError:
                pass

    def _move_selector(self, step: int) -> None:
        target_id = self._selector_target_id
        controller = self._capture_controller
        if not target_id or controller is None:
            return
        status = controller.select_debug_element(target_id, step)
        self._update_selector_window(status)

    def _approve_selector(self) -> None:
        target_id = self._selector_target_id
        controller = self._capture_controller
        if not target_id or controller is None:
            return
        if controller.approve_debug_element(target_id):
            self._update_selector_window(controller.debug_selector_status(target_id))

    def _selector_next_pod(self) -> None:
        target_id = self._selector_target_id
        controller = self._capture_controller
        if not target_id or controller is None or not self._selector_pdf_saved:
            return
        if self._review_rescan_mode:
            status = controller.debug_selector_status(target_id) or {}
            path = str(status.get("final_pdf") or "").strip()
            self._show_review_rescan_decision(Path(path) if path else self._active_capture_output_pdf)
            return
        controller.close_debug_target(target_id)
        self._destroy_selector_window()
        if self._batch_capture_active and self._hands_free_pdf_var.get():
            self._cancel_batch_after()
            self._batch_capture_after_id = self._root.after(250, self._start_or_continue_batch_capture)

    def _stop_selector_batch(self) -> None:
        self._set_batch_capture_active(False)
        try:
            self._hands_free_pdf_var.set(0)
        except tk.TclError:
            pass
        self._stop_capture_controller()
        self._set_capture_status("POD selector debug stopped.")

    def _on_selector_ready(self, target_id: str, status: dict) -> None:
        def show() -> None:
            self._destroy_capture_instruction_window()
            self._destroy_capture_progress_window()
            self._selector_target_id = target_id
            self._selector_pdf_saved = bool(status.get("pdf_saved"))
            if self._selector_window is None:
                try:
                    win = tk.Toplevel(self._root)
                    win.title("POD Element Selector")
                    win.configure(bg=THEME["bg"])
                    win.resizable(False, False)
                    win.transient(self._root)
                    win.protocol("WM_DELETE_WINDOW", lambda: None)
                    frame = tk.Frame(win, bg=THEME["bg"], padx=18, pady=16)
                    frame.pack(fill=tk.BOTH, expand=True)
                    tk.Label(
                        frame,
                        text="POD ready element",
                        fg=THEME["fg"],
                        bg=THEME["bg"],
                        font=theme_font("title"),
                        anchor=tk.W,
                    ).pack(fill=tk.X, anchor=tk.W)
                    self._selector_status_var = tk.StringVar(value="")
                    self._selector_time_var = tk.StringVar(value="")
                    self._selector_detail_var = tk.StringVar(value="")
                    tk.Label(
                        frame,
                        textvariable=self._selector_status_var,
                        height=2,
                        wraplength=520,
                        justify=tk.LEFT,
                        fg=THEME["fg"],
                        bg=THEME["bg"],
                        font=theme_font("body"),
                        anchor=tk.NW,
                    ).pack(fill=tk.X, anchor=tk.W, pady=(8, 0))
                    tk.Label(
                        frame,
                        textvariable=self._selector_time_var,
                        height=4,
                        wraplength=520,
                        justify=tk.LEFT,
                        fg=THEME["muted"],
                        bg=THEME["bg"],
                        font=theme_font("body"),
                        anchor=tk.NW,
                    ).pack(fill=tk.X, anchor=tk.W, pady=(6, 0))
                    tk.Label(
                        frame,
                        textvariable=self._selector_detail_var,
                        height=15,
                        width=72,
                        wraplength=520,
                        justify=tk.LEFT,
                        fg=THEME["muted"],
                        bg=THEME["bg"],
                        font=theme_font("body"),
                        anchor=tk.NW,
                    ).pack(fill=tk.X, anchor=tk.W, pady=(6, 12))

                    nav = tk.Frame(frame, bg=THEME["bg"])
                    nav.pack(fill=tk.X, pady=(0, 8))
                    make_flat_button(
                        nav,
                        text="Previous",
                        command=lambda: self._move_selector(-1),
                        bg=THEME["surface"],
                        active_bg=THEME["track"],
                        fg=THEME["fg"],
                        active_fg=THEME["fg"],
                        width=11,
                    ).pack(side=tk.LEFT, padx=(0, 8))
                    make_flat_button(
                        nav,
                        text="Next Element",
                        command=lambda: self._move_selector(1),
                        bg=THEME["surface"],
                        active_bg=THEME["track"],
                        fg=THEME["fg"],
                        active_fg=THEME["fg"],
                        width=13,
                    ).pack(side=tk.LEFT)

                    actions = tk.Frame(frame, bg=THEME["bg"])
                    actions.pack(fill=tk.X)
                    self._selector_capture_btn = make_flat_button(
                        actions,
                        text="Capture Selected",
                        command=self._approve_selector,
                        bg=THEME["excel_accent"],
                        active_bg=THEME["excel_accent_dim"],
                        width=16,
                    )
                    self._selector_capture_btn.pack(side=tk.LEFT, padx=(0, 8))
                    self._selector_next_pod_btn = make_flat_button(
                        actions,
                        text="Replace POD" if self._review_rescan_mode else "Next POD",
                        command=self._selector_next_pod,
                        bg=THEME["run_accent"],
                        active_bg=THEME["run_accent_dim"],
                        width=12,
                    )
                    self._selector_next_pod_btn.pack(side=tk.LEFT, padx=(0, 8))
                    d_bg, d_active = danger_colors()
                    make_flat_button(
                        actions,
                        text="Stop",
                        command=self._stop_selector_batch,
                        bg=d_bg,
                        active_bg=d_active,
                        width=10,
                    ).pack(side=tk.RIGHT)

                    for seq, step in (("<Left>", -1), ("<Up>", -1), ("<Right>", 1), ("<Down>", 1)):
                        win.bind(seq, lambda _e, s=step: self._move_selector(s))
                    win.bind("<Return>", lambda _e: self._approve_selector())
                    win.bind("n", lambda _e: self._selector_next_pod())
                    self._selector_window = win
                    self._place_capture_popup(win)
                    try:
                        win.focus_force()
                    except tk.TclError:
                        pass
                except tk.TclError:
                    self._selector_window = None
                    self._selector_status_var = None
                    self._selector_time_var = None
                    self._selector_detail_var = None
                    return
            self._update_selector_window(status)
            self._set_capture_status("Selector debug: choose the blue-boxed element, then capture.")

        try:
            self._root.after(0, show)
        except tk.TclError:
            pass

    def _capture_notify(self, level: str, message: str) -> None:
        audit_pause = level == "error" and "AI audit did not approve this capture" in str(message or "")
        chrome_closed = level == "info" and "The capture Chrome was closed." in str(message or "")

        def show() -> None:
            self._set_capture_status(message)
            if level == "progress":
                self._show_capture_progress_window(message)
                return
            if chrome_closed:
                self._destroy_capture_popups(keep_progress=self._keep_batch_progress_window())
                self._hands_free_capture_running = False
                if self._batch_capture_active and self._hands_free_pdf_var.get() and not self._review_rescan_mode:
                    self._cancel_batch_after()
                    self._batch_capture_after_id = self._root.after(
                        350, self._start_or_continue_batch_capture
                    )
                return
            if level == "error":
                self._destroy_capture_progress_window()
                if not audit_pause:
                    self._destroy_capture_instruction_window()
                    self._hands_free_capture_running = False
                    self._set_batch_capture_active(False)
                else:
                    if pod_readiness_debug_enabled():
                        self._update_selector_window()
                        self._set_capture_status(
                            "AI audit wants another look. Adjust the blue box, then capture again."
                        )
                    else:
                        self._set_capture_status(
                            "AI audit wants another look. Fix the still-open tab, then press "
                            f"{CAPTURE_HOTKEY_LABEL} again."
                        )
                messagebox.showerror("PDF capture", message)

        try:
            self._root.after(0, show)
        except tk.TclError:
            pass

    def _ensure_capture_controller(self) -> HtmlCaptureController | None:
        if self._capture_controller is None:
            self._capture_controller = HtmlCaptureController(
                on_notify=self._capture_notify,
                on_saved=lambda: self._root.after(0, self._on_assisted_capture_saved),
                on_review_saved=lambda payload: self._root.after(
                    0,
                    lambda: self._on_assisted_capture_review_saved(payload),
                ),
                on_selector_ready=self._on_selector_ready,
            )
        if not self._capture_controller.start():
            return None
        return self._capture_controller

    def _stop_capture_controller(self) -> None:
        controller = self._capture_controller
        self._capture_controller = None
        self._hands_free_capture_running = False
        self._destroy_capture_popups()
        self._destroy_selector_window()
        if controller is not None:
            controller.stop()

    def _cancel_batch_after(self) -> None:
        if self._batch_capture_after_id is not None:
            try:
                self._root.after_cancel(self._batch_capture_after_id)
            except tk.TclError:
                pass
            self._batch_capture_after_id = None

    def _set_batch_capture_active(self, active: bool) -> None:
        self._batch_capture_active = bool(active)
        if not active:
            self._cancel_batch_after()
        self._update_capture_controls()

    def _capture_eligible(self, info: dict[str, object]) -> bool:
        if self._review_rescan_decision_window is not None:
            return False
        if self._review_rescan_mode:
            return bool(str(info.get("tracking_number") or "").strip())
        return (not self._is_processed(info)) and (
            self._review_rescan_mode or not self._is_grey_row(info)
        )

    def _next_capture_index(self) -> int | None:
        if self._review_rescan_mode and self._row_infos:
            selected = self._selected_index()
            idx = selected if selected is not None and 0 <= selected < len(self._row_infos) else 0
            return idx if self._capture_eligible(self._row_infos[idx]) else None
        selected = self._selected_index()
        if selected is not None and 0 <= selected < len(self._row_infos):
            if self._capture_eligible(self._row_infos[selected]):
                return selected
        for idx, info in enumerate(self._row_infos):
            if self._capture_eligible(info):
                return idx
        return None

    def _start_or_continue_batch_capture(self) -> None:
        if not self._batch_capture_active:
            return
        if not self._hands_free_pdf_var.get():
            self._set_batch_capture_active(False)
            return
        if self._hands_free_capture_running:
            return
        idx = self._next_capture_index()
        if idx is None:
            self._set_batch_capture_active(False)
            done_message = (
                "Waiting for your manual rescan decision."
                if self._review_rescan_mode and self._review_rescan_decision_window is not None
                else "Done. No unprocessed, active POD rows remain."
            )
            self._set_capture_status(done_message)
            if self._capture_progress_window is not None:
                self._show_capture_progress_window(done_message)
                if self._capture_progress_bar is not None:
                    try:
                        self._capture_progress_bar.stop()
                    except tk.TclError:
                        pass
            return
        if not self._open_assisted_capture_for_index(idx):
            self._set_batch_capture_active(False)

    def _open_assisted_capture_for_index(self, idx: int) -> bool:
        if idx < 0 or idx >= len(self._row_infos):
            return False
        info = self._info_for_index(idx)
        if self._is_processed(info) and not self._review_rescan_mode:
            self._set_capture_status("That POD is already processed.")
            return False
        if self._is_grey_row(info) and not self._review_rescan_mode:
            messagebox.showinfo(
                "PDF capture",
                "This tracking number is temporarily unavailable for assisted capture.",
            )
            return False
        url = self._carrier_url_for_index(idx)
        if not url:
            messagebox.showerror("Tracking", "Could not build a carrier URL for this number.")
            return False
        expected_pdf = self._expected_pdf_path_for_info(info)
        if expected_pdf is None:
            messagebox.showerror("PDF capture", "Could not build the expected PDF filename for this row.")
            return False
        capture_pdf = expected_pdf
        if self._review_rescan_mode:
            capture_pdf = _unique_sibling_path(expected_pdf, "_manual_rescan")
            self._active_capture_info = dict(info)
            self._active_capture_expected_pdf = expected_pdf
            self._active_capture_output_pdf = capture_pdf
        controller = self._ensure_capture_controller()
        if controller is None:
            self._destroy_capture_instruction_window()
            messagebox.showerror(
                "PDF capture",
                f"Could not start assisted capture. {CAPTURE_HOTKEY_LABEL} capture requires Windows and Chrome.",
            )
            return False
        if self._hands_free_pdf_var.get():
            self._set_batch_capture_active(True)
        self._tree.selection_set(str(idx))
        self._tree.see(str(idx))
        num = str(info.get("tracking_number") or "").strip()
        if pod_readiness_debug_enabled():
            self._show_capture_instruction_window(info)
        else:
            carrier = str(info.get("carrier") or "").strip() or carrier_display_for_number(num)
            self._show_capture_progress_window(
                f"Starting {carrier} {num}. Waiting for page readiness, then saving automatically."
            )
        record = self._record_for_hands_free_capture(info)
        if not controller.enqueue_capture(
            url,
            capture_pdf,
            record=record,
            auto_print_pdf=True,
            visible_chrome=self._review_rescan_mode,
        ):
            self._destroy_capture_popups()
            return False
        self._hands_free_capture_running = True
        if pod_readiness_debug_enabled():
            self._set_capture_status(f"Starting {num}. Selector controls will appear when the page is ready.")
        else:
            self._set_capture_status(f"Starting {num}. Waiting for page readiness, then saving automatically.")
        return True

    def _on_assisted_capture_saved(self) -> None:
        """Refresh visible POD state and advance the batch after an approved capture."""
        self._hands_free_capture_running = False
        if self._review_rescan_mode:
            self._set_batch_capture_active(False)
            self._destroy_capture_popups()
            self._show_review_rescan_decision(self._active_capture_output_pdf)
            return
        self._destroy_capture_popups(keep_progress=self._keep_batch_progress_window())
        self._apply_pod_completion_layout()
        self._sync_open_workbook_after_pod_saved()
        if pod_readiness_debug_enabled() and self._selector_target_id:
            status = None
            if self._capture_controller is not None:
                status = self._capture_controller.debug_selector_status(self._selector_target_id)
            self._update_selector_window(status)
            self._set_capture_status("PDF saved. Use Next POD when ready.")
            return
        if self._batch_capture_active:
            next_message = "Saved. Opening next active POD row..."
            self._set_capture_status(next_message)
            self._show_capture_progress_window(next_message)
            self._cancel_batch_after()
            self._batch_capture_after_id = self._root.after(700, self._start_or_continue_batch_capture)
        else:
            self._set_capture_status("Saved. Ready for the next selected row.")

    def _on_assisted_capture_review_saved(self, payload: dict) -> None:
        """Flag a captured POD that needs manual review, then continue the batch."""
        self._hands_free_capture_running = False
        path = str(payload.get("path") or "").strip()
        record = payload.get("record") if isinstance(payload.get("record"), dict) else {}
        reason = str(payload.get("reason") or "").strip()
        if self._review_rescan_mode:
            self._set_batch_capture_active(False)
            self._destroy_capture_popups()
            self._show_review_rescan_decision(Path(path) if path else self._active_capture_output_pdf, reason=reason)
            return
        if self._project_root is not None and path:
            try:
                upsert_pod_review_record(
                    self._project_root,
                    pdf_path=path,
                    source_record=record,
                    reason=reason,
                )
                self._audit_entries_cache = []
                self._audit_entries_mtime_ns = None
            except Exception as exc:
                self._set_capture_status(f"POD review PDF saved, but flagging needs attention: {exc}")
        self._destroy_capture_popups(keep_progress=self._keep_batch_progress_window())
        self._apply_pod_completion_layout()
        self._sync_open_workbook_after_pod_saved()
        if pod_readiness_debug_enabled() and self._selector_target_id:
            status = None
            if self._capture_controller is not None:
                status = self._capture_controller.debug_selector_status(self._selector_target_id)
            self._update_selector_window(status)
            self._set_capture_status("POD saved for manual review. Use Next POD when ready.")
            return
        if self._batch_capture_active:
            next_message = "Flagged for review. Opening next active POD row..."
            self._set_capture_status(next_message)
            self._show_capture_progress_window(next_message)
            self._cancel_batch_after()
            self._batch_capture_after_id = self._root.after(700, self._start_or_continue_batch_capture)
        else:
            self._set_capture_status("POD saved for manual review and flagged.")

    def _sync_open_workbook_after_pod_saved(self) -> None:
        workbook_path = str(self._context.get("workbook_path") or "").strip()
        if not workbook_path or self._pod_excel_sync_running:
            return
        try:
            path = Path(workbook_path).expanduser().resolve()
        except OSError:
            return
        if not path.is_file():
            return
        self._pod_excel_sync_running = True

        def finish(message: str) -> None:
            self._pod_excel_sync_running = False
            if message:
                self._set_capture_status(message)

        def worker() -> None:
            try:
                from proofOfDelivery.pod_workflow import sync_open_workbook

                changed = sync_open_workbook(path)
                message = (
                    "POD row added to the open Excel workbook."
                    if changed
                    else "Saved. Excel workbook already had this POD row."
                )
            except Exception as exc:
                message = f"POD saved, but Excel sync needs attention: {exc}"
            try:
                self._root.after(0, lambda: finish(message))
            except tk.TclError:
                pass

        threading.Thread(target=worker, daemon=True).start()

    def _activate_row_open_or_capture(self) -> None:
        idx = self._selected_index()
        if idx is None or idx < 0 or idx >= len(self._row_infos):
            return
        info = self._info_for_index(idx)
        if self._review_pdf_path_for_info(info) is not None:
            self._open_selected_review_pdf()
            return
        url = self._carrier_url_for_index(idx)
        if not url:
            messagebox.showerror("Tracking", "Could not build a carrier URL for this number.")
            return
        if self._hands_free_pdf_var.get() == True:
            if self._open_assisted_capture_for_index(idx):
                return

        self._open_carrier_in_browser_only()

    def _force_check_selected_tracking_number(self) -> None:
        messagebox.showinfo(
            "Tracking",
            "Live tracking validation has been removed. The POD list keeps parser-found tracking numbers available for capture.",
        )

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

    parser = argparse.ArgumentParser(description="View tracking numbers and POD capture status.")
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
