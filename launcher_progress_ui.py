"""Shared dark-themed progress dialogs for email_sorter_launcher (Run / Excel)."""

from __future__ import annotations

import re
import sys
import tkinter as tk
from tkinter import messagebox
from typing import Callable

# Universal dark theme (Run = blue accent, Excel = green accent)
THEME = {
    "bg": "#0f1117",
    "surface": "#161b22",
    "fg": "#e6edf3",
    "muted": "#8b949e",
    "track": "#21262d",
    "border": "#30363d",
    "run_accent": "#3b82f6",
    "run_accent_dim": "#1d4ed8",
    "excel_accent": "#22c55e",
    "excel_accent_dim": "#15803d",
    "stop_fg": "#f85149",
    "font_title": ("Segoe UI", 13, "bold"),
    "font_body": ("Segoe UI", 10),
    "font_pct": ("Segoe UI", 22, "bold"),
}


_RUN_LINE = re.compile(
    r"^EMAIL_SORTER_RUN_PROGRESS\s+pct=(\d+)(?:\s+msg=(.*))?\s*$", re.I
)
_EXCEL_LINE = re.compile(
    r"^EMAIL_SORTER_EXCEL_PROGRESS\s+pct=(\d+)(?:\s+msg=(.*))?\s*$", re.I
)


def parse_run_progress_line(line: str) -> tuple[int, str] | None:
    m = _RUN_LINE.match(line.strip())
    if not m:
        return None
    pct = max(0, min(100, int(m.group(1))))
    msg = (m.group(2) or "").strip()
    return pct, msg


def parse_excel_progress_line(line: str) -> tuple[int, str] | None:
    m = _EXCEL_LINE.match(line.strip())
    if not m:
        return None
    pct = max(0, min(100, int(m.group(1))))
    msg = (m.group(2) or "").strip()
    return pct, msg


class PipelineProgressWindow(tk.Toplevel):
    """
    Dark modal with canvas progress bar + percentage + status line + Stop.
    accent: "run" (blue) or "excel" (green).
    show_log: if True, adds a scrollable log pane for debug output.
    on_skip_17track: when set, shows "Skip 17Track" (Excel rebuild / 17TRACK prefetch flows).
    """

    def __init__(
        self,
        parent: tk.Tk,
        *,
        title: str,
        headline: str,
        accent: str,
        on_stop: Callable[[], None],
        on_skip_17track: Callable[[], None] | None = None,
        bar_width: int = 420,
        bar_height: int = 14,
        show_log: bool = False,
    ) -> None:
        super().__init__(parent)
        self._accent = accent
        self._on_stop = on_stop
        self._on_skip_17track = on_skip_17track
        self._skip_17track_btn: tk.Button | None = None
        self._skip_17track_done = False
        self._bar_w = bar_width
        self._bar_h = bar_height
        self._pct = 0
        self._stopped = False
        self._show_log = show_log

        self.title(title)
        self.configure(bg=THEME["bg"])
        # Avoid transient()+grab: on Windows that yields a slim title bar (no minimize) and
        # keeps the dialog above unrelated apps. User should be able to minimize and cover
        # this window with other programs while the pipeline runs.
        self.resizable(True, True)
        if sys.platform == "win32":
            try:
                self.attributes("-toolwindow", False)
            except tk.TclError:
                pass
        self.protocol("WM_DELETE_WINDOW", self._block_close_attempt)

        if accent == "excel":
            fill = THEME["excel_accent"]
            glow = THEME["excel_accent_dim"]
        else:
            fill = THEME["run_accent"]
            glow = THEME["run_accent_dim"]

        self._fill = fill
        self._glow = glow

        pad = {"padx": 22, "pady": 8}
        outer = tk.Frame(self, bg=THEME["bg"])
        outer.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            outer,
            text=headline,
            font=THEME["font_title"],
            fg=THEME["fg"],
            bg=THEME["bg"],
            justify=tk.LEFT,
        ).pack(anchor=tk.W, **pad)

        self._status = tk.Label(
            outer,
            text="Starting…",
            font=THEME["font_body"],
            fg=THEME["muted"],
            bg=THEME["bg"],
            anchor=tk.W,
            wraplength=bar_width + 40,
            justify=tk.LEFT,
        )
        self._status.pack(anchor=tk.W, padx=22, pady=(0, 6))

        pct_row = tk.Frame(outer, bg=THEME["bg"])
        pct_row.pack(fill=tk.X, padx=22, pady=(4, 10))
        self._pct_lbl = tk.Label(
            pct_row,
            text="0%",
            font=THEME["font_pct"],
            fg=THEME["fg"],
            bg=THEME["bg"],
        )
        self._pct_lbl.pack(side=tk.LEFT)

        canvas_frame = tk.Frame(
            outer,
            bg=THEME["surface"],
            highlightthickness=1,
            highlightbackground=THEME["border"],
        )
        canvas_frame.pack(fill=tk.X, padx=22, pady=(0, 14))
        self._cv = tk.Canvas(
            canvas_frame,
            width=bar_width,
            height=bar_height + 8,
            bg=THEME["surface"],
            highlightthickness=0,
        )
        self._cv.pack(padx=8, pady=8)
        self._draw_bar(0)

        # --- Log pane (debug mode) ---
        if show_log:
            log_outer = tk.Frame(outer, bg=THEME["bg"])
            log_outer.pack(fill=tk.BOTH, expand=True, padx=22, pady=(0, 8))

            tk.Label(
                log_outer,
                text="Debug output",
                font=("Segoe UI", 9),
                fg=THEME["muted"],
                bg=THEME["bg"],
                anchor=tk.W,
            ).pack(anchor=tk.W, pady=(0, 2))

            log_frame = tk.Frame(
                log_outer,
                bg=THEME["surface"],
                highlightthickness=1,
                highlightbackground=THEME["border"],
            )
            log_frame.pack(fill=tk.BOTH, expand=True)

            self._log_text = tk.Text(
                log_frame,
                bg=THEME["surface"],
                fg=THEME["muted"],
                font=("Consolas", 9),
                relief=tk.FLAT,
                state=tk.DISABLED,
                wrap=tk.NONE,
                height=10,
                highlightthickness=0,
            )
            vsb = tk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self._log_text.yview)
            hsb = tk.Scrollbar(log_frame, orient=tk.HORIZONTAL, command=self._log_text.xview)
            self._log_text.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            vsb.pack(side=tk.RIGHT, fill=tk.Y)
            hsb.pack(side=tk.BOTTOM, fill=tk.X)
            self._log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        else:
            self._log_text = None  # type: ignore[assignment]

        # --- Stop / Skip 17Track row ---
        btn_row = tk.Frame(outer, bg=THEME["bg"])
        btn_row.pack(fill=tk.X, padx=22, pady=(4, 18))
        if on_skip_17track is not None:
            self._skip_17track_btn = tk.Button(
                btn_row,
                text="Skip 17Track",
                command=self._skip_17track_clicked,
                font=("Segoe UI", 10, "bold"),
                fg=THEME["excel_accent"],
                bg=THEME["surface"],
                activeforeground=THEME["excel_accent_dim"],
                activebackground=THEME["track"],
                relief=tk.FLAT,
                padx=18,
                pady=6,
                cursor="hand2",
            )
            self._skip_17track_btn.pack(side=tk.RIGHT, padx=(0, 10))
        self._stop_btn = tk.Button(
            btn_row,
            text="Stop",
            command=self._stop_clicked,
            font=("Segoe UI", 10, "bold"),
            fg=THEME["stop_fg"],
            bg=THEME["surface"],
            activeforeground=THEME["stop_fg"],
            activebackground=THEME["track"],
            relief=tk.FLAT,
            padx=18,
            pady=6,
            cursor="hand2",
        )
        self._stop_btn.pack(side=tk.RIGHT)

        # Size and position
        self.update_idletasks()
        req_w = max(outer.winfo_reqwidth() + 8, 480)
        req_h = outer.winfo_reqheight() + 8
        if show_log:
            req_h = max(req_h, 520)
        else:
            req_h = max(req_h, 280)
        self.geometry(f"{req_w}x{req_h}")
        self._center_on_parent(parent)

    def _center_on_parent(self, parent: tk.Tk) -> None:
        try:
            self.update_idletasks()
            px = parent.winfo_rootx() + (parent.winfo_width() // 2) - (self.winfo_width() // 2)
            py = parent.winfo_rooty() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
            self.geometry(f"+{max(0, px)}+{max(0, py)}")
        except tk.TclError:
            pass

    def _block_close_attempt(self) -> None:
        """Do not allow closing via the window X — only Stop can end the flow."""
        try:
            messagebox.showinfo(
                self.title(),
                "This window cannot be closed.\n\nUse Stop if you need to cancel.",
                parent=self,
            )
        except tk.TclError:
            pass

    def _skip_17track_clicked(self) -> None:
        if self._skip_17track_done or self._stopped:
            return
        self._skip_17track_done = True
        if self._skip_17track_btn is not None:
            try:
                self._skip_17track_btn.config(state=tk.DISABLED, text="Skipping…")
            except tk.TclError:
                pass
        if self._on_skip_17track:
            self._on_skip_17track()

    def _stop_clicked(self) -> None:
        if self._stopped:
            return
        try:
            if not messagebox.askyesno(
                self.title(),
                "Are you sure you want to stop?",
                parent=self,
                icon=messagebox.WARNING,
                default=messagebox.NO,
            ):
                return
        except tk.TclError:
            return
        self._stopped = True
        self._stop_btn.config(state=tk.DISABLED, text="Stopping…")
        if self._skip_17track_btn is not None:
            try:
                self._skip_17track_btn.config(state=tk.DISABLED)
            except tk.TclError:
                pass
        self._on_stop()

    @property
    def stop_requested(self) -> bool:
        return self._stopped

    def _draw_bar(self, pct: int) -> None:
        self._cv.delete("all")
        w, h = self._bar_w, self._bar_h
        y0 = 4
        # Track
        self._cv.create_rectangle(0, y0, w, y0 + h, fill=THEME["track"], outline="", width=0)
        fill_w = max(0, int(round(w * (pct / 100.0))))
        if fill_w > 2:
            # subtle glow strip under fill
            self._cv.create_rectangle(
                0, y0 + h - 2, fill_w, y0 + h + 2, fill=self._glow, outline="", width=0
            )
            self._cv.create_rectangle(
                0, y0, fill_w, y0 + h, fill=self._fill, outline="", width=0
            )

    def set_progress(self, pct: int, message: str | None = None) -> None:
        pct = max(0, min(100, int(pct)))
        self._pct = pct
        self._pct_lbl.config(text=f"{pct}%")
        if message:
            self._status.config(text=message, fg=THEME["fg"])
        self._draw_bar(pct)
        self.update_idletasks()

    def append_log(self, text: str) -> None:
        """Append a line to the debug log pane (only visible when show_log=True)."""
        if self._log_text is None:
            return
        try:
            self._log_text.configure(state=tk.NORMAL)
            self._log_text.insert(tk.END, text if text.endswith("\n") else text + "\n")
            self._log_text.see(tk.END)
            self._log_text.configure(state=tk.DISABLED)
        except tk.TclError:
            pass

    def close_window(self) -> None:
        try:
            self.destroy()
        except tk.TclError:
            pass
