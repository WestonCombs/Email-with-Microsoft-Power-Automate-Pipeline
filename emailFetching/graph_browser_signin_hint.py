"""Dark-mode hint while Microsoft Graph interactive / device-code sign-in runs."""

from __future__ import annotations

import threading
import time
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import messagebox

# Duo-inspired dark layout (not identical assets)
_BG = "#1a1a1e"
_SURFACE = "#25252b"
_FG = "#f3f3f8"
_MUTED = "#a8a8b3"
_ACCENT = "#569cd6"
_BORDER = "#3e3e48"


def run_blocking_task_with_browser_signin_hint(
    fn: Callable[[], dict],
    *,
    cancel_check,
    cancel_message: str,
    timeout_seconds: float | None,
    base_dir: Path | None,
    allow_retry: bool = False,
    on_retry: Callable[[], None] | None = None,
) -> dict:
    """
    Run *fn* on a worker thread (MSAL interactive / device flow) while showing a
    non-blocking hint window on the main thread. Closes automatically when *fn*
    finishes. Silent token acquisition must be attempted before calling this.

    When *allow_retry* is true, the Retry button starts a new worker and makes
    any older worker result stale. Python cannot kill the old MSAL browser flow,
    but the app will ignore it and will not save its token cache.
    """
    root = tk.Tk()
    root.title("Microsoft sign-in")
    root.configure(bg=_BG)
    root.resizable(False, False)
    root.minsize(420, 420)

    dismissed = {"v": False}
    state_lock = threading.Lock()
    state: dict[str, object] = {
        "attempt": 0,
        "done": False,
        "value": None,
        "error": None,
    }
    status_var = tk.StringVar(value="Waiting for Microsoft sign-in...")

    def _start_attempt() -> None:
        with state_lock:
            attempt = int(state["attempt"]) + 1
            state.update(
                {
                    "attempt": attempt,
                    "done": False,
                    "value": None,
                    "error": None,
                }
            )

        def _worker(attempt_no: int) -> None:
            value: dict | None = None
            error: BaseException | None = None
            try:
                value = fn()
            except BaseException as e:
                error = e
            with state_lock:
                if dismissed["v"] or attempt_no != state["attempt"]:
                    return
                state["value"] = value
                state["error"] = error
                state["done"] = True

        threading.Thread(target=_worker, args=(attempt,), daemon=True).start()

    def _current_done() -> bool:
        with state_lock:
            return bool(state["done"])

    def _current_result() -> tuple[dict | None, BaseException | None]:
        with state_lock:
            value = state.get("value")
            error = state.get("error")
        return (
            value if isinstance(value, dict) else None,
            error if isinstance(error, BaseException) else None,
        )

    def _need_help() -> None:
        msg = (
            "1. Open the main Email Sorter runner window.\n"
            "2. Click the X at the top-right to close the program.\n"
            "3. Reopen the program and try signing in again.\n\n"
            "Step 2: If the problem continues, contact your administrator."
        )
        try:
            messagebox.showinfo("Need help", msg, parent=root)
        except tk.TclError:
            pass

    def _cancel() -> None:
        if dismissed["v"]:
            return
        dismissed["v"] = True
        if base_dir is not None:
            try:
                from shared.cancel_control import request_cancel

                request_cancel(base_dir)
            except Exception:
                pass
        try:
            root.destroy()
        except tk.TclError:
            pass

    def _retry() -> None:
        if dismissed["v"]:
            return
        if on_retry is not None:
            try:
                on_retry()
            except Exception:
                pass
        status_var.set("Starting a new Microsoft sign-in...")
        _start_attempt()

    outer = tk.Frame(root, bg=_BG, padx=28, pady=28)
    outer.pack(fill=tk.BOTH, expand=True)

    tk.Label(
        outer,
        text="Please sign in to the email in your browser.",
        font=("Segoe UI", 14, "bold"),
        fg=_FG,
        bg=_BG,
        wraplength=360,
        justify=tk.CENTER,
    ).pack(pady=(0, 12))

    tk.Label(
        outer,
        text=(
            "Verify your account credentials by signing in through the log in screen "
            "on your browser. It should have opened."
        ),
        font=("Segoe UI", 10),
        fg=_MUTED,
        bg=_BG,
        wraplength=360,
        justify=tk.CENTER,
    ).pack(pady=(0, 20))

    box = tk.Frame(
        outer,
        bg=_SURFACE,
        highlightthickness=1,
        highlightbackground=_BORDER,
        height=140,
    )
    box.pack(fill=tk.X, pady=(0, 28))
    box.pack_propagate(False)

    tk.Label(
        box,
        textvariable=status_var,
        font=("Segoe UI", 10),
        fg=_MUTED,
        bg=_SURFACE,
        wraplength=340,
        justify=tk.CENTER,
    ).pack(expand=True)

    btn_row = tk.Frame(outer, bg=_BG)
    btn_row.pack(pady=(0, 16))

    if allow_retry:
        tk.Button(
            btn_row,
            text="Retry",
            font=("Segoe UI", 10, "bold"),
            fg=_FG,
            bg=_ACCENT,
            activeforeground=_FG,
            activebackground="#3577a8",
            relief=tk.FLAT,
            padx=28,
            pady=10,
            cursor="hand2",
            command=_retry,
        ).pack(side=tk.LEFT, padx=(0, 10))

    tk.Button(
        btn_row,
        text="Close",
        font=("Segoe UI", 10, "bold"),
        fg=_FG,
        bg=_SURFACE,
        activeforeground=_FG,
        activebackground=_BORDER,
        relief=tk.FLAT,
        padx=28,
        pady=10,
        cursor="hand2",
        command=_cancel,
    ).pack(side=tk.LEFT)

    help_lbl = tk.Label(
        outer,
        text="Need help",
        font=("Segoe UI", 10, "underline"),
        fg=_ACCENT,
        bg=_BG,
        cursor="hand2",
    )
    help_lbl.pack(anchor=tk.W)
    help_lbl.bind("<Button-1>", lambda _e: _need_help())

    def _on_close() -> None:
        _cancel()

    root.protocol("WM_DELETE_WINDOW", _on_close)

    try:
        root.update_idletasks()
        w, h = root.winfo_reqwidth() + 40, root.winfo_reqheight() + 40
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        root.geometry(f"{max(420, w)}x{max(420, h)}+{(sw - w) // 2}+{(sh - h) // 2}")
    except tk.TclError:
        pass

    _start_attempt()
    started = time.monotonic()
    while not _current_done():
        if cancel_check and cancel_check():
            _cancel()
            raise RuntimeError(cancel_message)
        if timeout_seconds is not None and (time.monotonic() - started) > timeout_seconds:
            try:
                root.destroy()
            except tk.TclError:
                pass
            raise RuntimeError(
                f"Microsoft sign-in timed out after {timeout_seconds:.0f}s."
            )
        try:
            if not root.winfo_exists():
                break
        except tk.TclError:
            break
        try:
            root.update()
        except tk.TclError:
            break
        time.sleep(0.02)

    try:
        if root.winfo_exists():
            root.destroy()
    except tk.TclError:
        pass

    if dismissed["v"]:
        raise RuntimeError(cancel_message)

    while not _current_done():
        if cancel_check and cancel_check():
            raise RuntimeError(cancel_message)
        if timeout_seconds is not None and (time.monotonic() - started) > timeout_seconds:
            raise RuntimeError(
                f"Microsoft sign-in timed out after {timeout_seconds:.0f}s."
            )
        time.sleep(0.2)

    value, error = _current_result()
    if error is not None:
        raise error
    return value or {}
