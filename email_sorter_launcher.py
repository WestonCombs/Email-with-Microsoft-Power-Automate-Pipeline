"""Tkinter launcher — main screen for Email Sorter (Run, Excel, Update, Settings, Exit)."""

from __future__ import annotations

from collections import deque
import os
import queue
import re
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import tkinter.font as tkfont
import webbrowser
from pathlib import Path
from tkinter import messagebox

from launcher_progress_ui import (
    THEME,
    PipelineProgressWindow,
    parse_excel_progress_line,
    parse_run_progress_line,
)
from shared.project_paths import ensure_base_dir_in_environ
from shared.settings_store import (
    apply_runtime_settings_from_json,
    read_settings_for_write_merge,
    read_settings_json,
    write_settings_json,
)

_PYTHON_FILES_DIR = Path(__file__).resolve().parent
_BASE_DIR = _PYTHON_FILES_DIR.parent
_LAUNCHER_CANCEL_FILE = ".email_sorter_cancel"
_TITLEBAR_BG = "#1f1f23"
_TITLEBAR_ACTIVE_BG = "#2a2a2f"
_TITLEBAR_HEIGHT = 39
_TITLEBAR_MIN_HOVER_BG = THEME["run_accent_dim"]
_TITLEBAR_MAX_HOVER_BG = "#d97706"
_DANGER_BG = THEME["stop_fg"]
_DANGER_ACTIVE_BG = "#da3633"
_UPDATE_BG = "#f59e0b"
_UPDATE_ACTIVE_BG = "#d97706"


def _font(name: str) -> tuple[str, int] | tuple[str, int, str]:
    fonts = {
        "titlebar": ("Segoe UI", 9),
        "title": ("Segoe UI", 14, "bold"),
        "body": ("Segoe UI", 10),
        "button": ("Segoe UI", 10, "bold"),
        "input": ("Segoe UI", 10),
    }
    return fonts[name]


def _blend_hex_color(hex_color: str, target: str = "#ffffff", amount: float = 0.14) -> str:
    try:
        base = hex_color.strip().lstrip("#")
        dest = target.strip().lstrip("#")
        if len(base) != 6 or len(dest) != 6:
            return hex_color
        b = tuple(int(base[i : i + 2], 16) for i in (0, 2, 4))
        d = tuple(int(dest[i : i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return hex_color
    mixed = tuple(max(0, min(255, round(c + (t - c) * amount))) for c, t in zip(b, d))
    return "#" + "".join(f"{c:02x}" for c in mixed)


def _add_button_hover(btn: tk.Button, *, normal_bg: str, hover_bg: str | None = None) -> tk.Button:
    hover = hover_bg or _blend_hex_color(normal_bg)

    def on_enter(_event: tk.Event) -> None:
        try:
            if str(btn.cget("state")) != tk.DISABLED:
                btn.configure(bg=hover)
        except tk.TclError:
            pass

    def on_leave(_event: tk.Event) -> None:
        try:
            btn.configure(bg=normal_bg)
        except tk.TclError:
            pass

    btn.bind("<Enter>", on_enter, add="+")
    btn.bind("<Leave>", on_leave, add="+")
    return btn


def _attach_tooltip(widget: tk.Misc, text: str) -> None:
    tip: dict[str, tk.Toplevel | None] = {"win": None}

    def show(event: tk.Event) -> None:
        if tip["win"] is not None:
            return
        try:
            win = tk.Toplevel(widget)
            win.overrideredirect(True)
            win.configure(bg=THEME["border"])
            tk.Label(
                win,
                text=text,
                font=("Segoe UI", 9),
                fg=THEME["fg"],
                bg=THEME["surface"],
                padx=8,
                pady=4,
            ).pack(padx=1, pady=1)
            win.geometry(f"+{event.x_root + 12}+{event.y_root + 12}")
            tip["win"] = win
        except tk.TclError:
            tip["win"] = None

    def hide(_event: tk.Event) -> None:
        win = tip["win"]
        tip["win"] = None
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass

    widget.bind("<Enter>", show, add="+")
    widget.bind("<Leave>", hide, add="+")
    widget.bind("<ButtonPress>", hide, add="+")


def _make_file_explorer_icon(master: tk.Misc) -> tk.PhotoImage:
    icon = tk.PhotoImage(master=master, width=36, height=28)
    try:
        for y in range(28):
            for x in range(36):
                icon.transparency_set(x, y, True)
    except tk.TclError:
        icon.put(THEME["excel_accent"], to=(0, 0, 36, 28))

    icon.put("#92400e", to=(4, 9, 30, 23))
    icon.put("#fbbf24", to=(5, 6, 15, 12))
    icon.put("#f59e0b", to=(6, 9, 31, 22))
    icon.put("#fde68a", to=(7, 12, 30, 20))
    icon.put("#d97706", to=(6, 21, 31, 23))
    icon.put("#2563eb", to=(18, 14, 27, 20))
    icon.put("#60a5fa", to=(19, 15, 27, 18))
    icon.put("#1d4ed8", to=(18, 19, 27, 21))
    return icon


def _make_button(
    parent: tk.Misc,
    *,
    text: str,
    command,
    bg: str,
    fg: str = "#ffffff",
    active_bg: str | None = None,
    active_fg: str | None = None,
    padx: int = 18,
    pady: int = 7,
    width: int | None = None,
) -> tk.Button:
    opts: dict[str, object] = {
        "text": text,
        "command": command,
        "font": _font("button"),
        "bg": bg,
        "fg": fg,
        "activebackground": active_bg or bg,
        "activeforeground": active_fg or fg,
        "relief": tk.FLAT,
        "bd": 0,
        "highlightthickness": 0,
        "cursor": "hand2",
        "padx": padx,
        "pady": pady,
    }
    if width is not None:
        opts["width"] = width
    btn = tk.Button(parent, **opts)
    return _add_button_hover(btn, normal_bg=bg, hover_bg=active_bg)


def _set_app_icon(win: tk.Toplevel | tk.Tk) -> None:
    try:
        icon = tk.PhotoImage(width=32, height=32)
        icon.put("#0f1117", to=(0, 0, 32, 32))
        icon.put("#3b82f6", to=(5, 6, 27, 13))
        icon.put("#22c55e", to=(5, 15, 27, 22))
        icon.put("#f59e0b", to=(5, 24, 27, 28))
        win.iconphoto(True, icon)
        win._email_sorter_icon = icon  # type: ignore[attr-defined]
    except tk.TclError:
        pass


def _center_window(win: tk.Toplevel | tk.Tk, parent: tk.Misc | None = None) -> None:
    try:
        win.update_idletasks()
        if parent is None:
            sw = win.winfo_screenwidth()
            sh = win.winfo_screenheight()
            x = (sw // 2) - (win.winfo_width() // 2)
            y = (sh // 2) - (win.winfo_height() // 2)
        else:
            parent.update_idletasks()
            x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (win.winfo_width() // 2)
            y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (win.winfo_height() // 2)
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
    except tk.TclError:
        pass


def _keep_window_above_parent(win: tk.Toplevel, parent: tk.Misc) -> None:
    try:
        parent_top = parent.winfo_toplevel()
    except tk.TclError:
        return
    bindings: list[tuple[tk.Misc, str, str]] = []

    def lift_child(_event: tk.Event | None = None) -> None:
        try:
            if win.winfo_exists():
                win.lift(parent_top)
        except tk.TclError:
            pass

    def cleanup(_event: tk.Event | None = None) -> None:
        for widget, sequence, funcid in bindings:
            try:
                widget.unbind(sequence, funcid)
            except tk.TclError:
                pass
        bindings.clear()

    for widget, sequence in (
        (parent_top, "<FocusIn>"),
        (parent_top, "<Map>"),
        (win, "<FocusIn>"),
        (win, "<Map>"),
    ):
        try:
            funcid = widget.bind(sequence, lift_child, add="+")
            bindings.append((widget, sequence, funcid))
        except tk.TclError:
            pass
    try:
        win.bind("<Destroy>", cleanup, add="+")
    except tk.TclError:
        pass
    lift_child()


def _show_dialog(
    win: tk.Toplevel,
    *,
    parent: tk.Misc,
    focus_widget: tk.Misc | None = None,
    modal: bool = True,
) -> None:
    try:
        win.transient(parent.winfo_toplevel())
    except tk.TclError:
        pass
    _center_window(win, parent)
    try:
        win.deiconify()
        win.lift(parent.winfo_toplevel())
        win.attributes("-topmost", True)
        win.after(0, lambda: _clear_topmost(win))
    except tk.TclError:
        pass
    _keep_window_above_parent(win, parent)
    try:
        if focus_widget is not None:
            focus_widget.focus_set()
        else:
            win.focus_force()
    except tk.TclError:
        pass
    if modal:
        try:
            win.grab_set()
        except tk.TclError:
            pass


def _clear_topmost(win: tk.Toplevel) -> None:
    try:
        win.attributes("-topmost", False)
    except tk.TclError:
        pass


def _make_titlebar_button(
    parent: tk.Misc,
    *,
    text: str,
    command,
    hover_bg: str,
    font: tuple[str, int] | tuple[str, int, str],
) -> tk.Button:
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        font=font,
        fg=THEME["fg"],
        bg=_TITLEBAR_BG,
        activeforeground="#ffffff",
        activebackground=hover_bg,
        relief=tk.FLAT,
        bd=0,
        width=4,
        padx=3,
        pady=2,
        cursor="hand2",
    )

    def on_enter(_event: tk.Event) -> None:
        try:
            btn.configure(bg=hover_bg, fg="#ffffff")
        except tk.TclError:
            pass

    def on_leave(_event: tk.Event) -> None:
        try:
            btn.configure(bg=_TITLEBAR_BG, fg=THEME["fg"])
        except tk.TclError:
            pass

    btn.bind("<Enter>", on_enter)
    btn.bind("<Leave>", on_leave)
    return btn


def _apply_frameless_window(
    win: tk.Toplevel | tk.Tk,
    *,
    title: str,
    on_close,
    show_minimize: bool = True,
) -> tk.Frame:
    win.overrideredirect(True)
    win.configure(bg=_TITLEBAR_BG)

    shell = tk.Frame(win, bg=_TITLEBAR_BG, highlightthickness=0)
    shell.pack(fill=tk.BOTH, expand=True)

    titlebar = tk.Frame(shell, bg=_TITLEBAR_BG, height=_TITLEBAR_HEIGHT)
    titlebar.pack(fill=tk.X, side=tk.TOP)
    titlebar.pack_propagate(False)

    title_label = tk.Label(
        titlebar,
        text=title,
        font=_font("titlebar"),
        fg=THEME["fg"],
        bg=_TITLEBAR_BG,
        anchor=tk.W,
    )
    title_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

    drag = {"x": 0, "y": 0}

    def start_move(event: tk.Event) -> None:
        drag["x"] = event.x
        drag["y"] = event.y

    def do_move(event: tk.Event) -> None:
        try:
            win.geometry(f"+{event.x_root - drag['x']}+{event.y_root - drag['y']}")
        except tk.TclError:
            pass

    def minimize() -> None:
        try:
            win.overrideredirect(False)
            win.iconify()
        except tk.TclError:
            return

    def toggle_maximize() -> None:
        try:
            win.overrideredirect(False)
            win.state("normal" if str(win.state()) == "zoomed" else "zoomed")
            win.after(50, lambda: win.overrideredirect(True))
        except tk.TclError:
            pass

    def restore_frameless(_event: tk.Event | None = None) -> None:
        try:
            if str(win.state()) == "normal":
                win.overrideredirect(True)
        except tk.TclError:
            pass

    win.bind("<Map>", restore_frameless, add="+")
    for widget in (titlebar, title_label):
        widget.bind("<ButtonPress-1>", start_move)
        widget.bind("<B1-Motion>", do_move)

    close_btn = _make_titlebar_button(
        titlebar,
        text="X",
        command=on_close,
        hover_bg=_DANGER_ACTIVE_BG,
        font=("Segoe UI Symbol", 12, "bold"),
    )
    close_btn.pack(side=tk.RIGHT, fill=tk.Y)

    if show_minimize:
        max_btn = _make_titlebar_button(
            titlebar,
            text="□",
            command=toggle_maximize,
            hover_bg=_TITLEBAR_MAX_HOVER_BG,
            font=("Segoe UI Symbol", 11),
        )
        max_btn.pack(side=tk.RIGHT, fill=tk.Y)

        min_btn = _make_titlebar_button(
            titlebar,
            text="—",
            command=minimize,
            hover_bg=_TITLEBAR_MIN_HOVER_BG,
            font=("Segoe UI Symbol", 12, "bold"),
        )
        min_btn.pack(side=tk.RIGHT, fill=tk.Y)

    content = tk.Frame(shell, bg=THEME["bg"], highlightthickness=0)
    content.pack(fill=tk.BOTH, expand=True, padx=1, pady=(0, 1))
    win.protocol("WM_DELETE_WINDOW", on_close)
    return content


def _themed_message(
    parent: tk.Misc,
    *,
    title: str,
    message: str,
    kind: str = "info",
) -> None:
    win = tk.Toplevel(parent)
    win.title(title)
    win.withdraw()
    win.resizable(False, False)

    def close() -> None:
        win.destroy()

    content = _apply_frameless_window(win, title=title, on_close=close, show_minimize=False)
    outer = tk.Frame(content, bg=THEME["bg"], padx=22, pady=22)
    outer.pack(fill=tk.BOTH, expand=True)
    color = THEME["fg"] if kind != "error" else _DANGER_BG
    tk.Label(
        outer,
        text=title,
        font=_font("title"),
        fg=color,
        bg=THEME["bg"],
        anchor=tk.W,
    ).pack(anchor=tk.W, pady=(0, 10))
    tk.Label(
        outer,
        text=message,
        font=_font("body"),
        fg=THEME["fg"],
        bg=THEME["bg"],
        justify=tk.LEFT,
        wraplength=520,
    ).pack(anchor=tk.W, fill=tk.X, pady=(0, 18))
    row = tk.Frame(outer, bg=THEME["bg"])
    row.pack(fill=tk.X)
    _make_button(row, text="Close", command=close, bg=_DANGER_BG, active_bg=_DANGER_ACTIVE_BG).pack(
        side=tk.RIGHT
    )
    win.update_idletasks()
    win.geometry(f"{max(420, outer.winfo_reqwidth() + 4)}x{outer.winfo_reqheight() + _TITLEBAR_HEIGHT + 8}")
    _show_dialog(win, parent=parent)
    win.wait_window()


def _themed_ask_yes_no(parent: tk.Misc, *, title: str, message: str) -> bool:
    answer = [False]
    win = tk.Toplevel(parent)
    win.title(title)
    win.withdraw()
    win.resizable(False, False)

    def finish(value: bool) -> None:
        answer[0] = value
        win.destroy()

    content = _apply_frameless_window(
        win,
        title=title,
        on_close=lambda: finish(False),
        show_minimize=False,
    )
    outer = tk.Frame(content, bg=THEME["bg"], padx=22, pady=22)
    outer.pack(fill=tk.BOTH, expand=True)
    tk.Label(
        outer,
        text=title,
        font=_font("title"),
        fg=THEME["fg"],
        bg=THEME["bg"],
        anchor=tk.W,
    ).pack(anchor=tk.W, pady=(0, 10))
    tk.Label(
        outer,
        text=message,
        font=_font("body"),
        fg=THEME["fg"],
        bg=THEME["bg"],
        justify=tk.LEFT,
        wraplength=520,
    ).pack(anchor=tk.W, fill=tk.X, pady=(0, 18))
    row = tk.Frame(outer, bg=THEME["bg"])
    row.pack(fill=tk.X)
    _make_button(
        row,
        text="No",
        command=lambda: finish(False),
        bg=_DANGER_BG,
        active_bg=_DANGER_ACTIVE_BG,
        width=10,
    ).pack(side=tk.RIGHT, padx=(8, 0))
    _make_button(
        row,
        text="Yes",
        command=lambda: finish(True),
        bg=THEME["excel_accent"],
        active_bg=THEME["excel_accent_dim"],
        width=10,
    ).pack(side=tk.RIGHT)
    win.update_idletasks()
    win.geometry(f"{max(420, outer.winfo_reqwidth() + 4)}x{outer.winfo_reqheight() + _TITLEBAR_HEIGHT + 8}")
    _show_dialog(win, parent=parent)
    win.wait_window()
    return answer[0]


def _optional_path(env_name: str, default: Path) -> Path:
    raw = os.getenv(env_name)
    if raw:
        return Path(raw).expanduser().resolve()
    return default


def _env_debug_enabled() -> bool:
    """True when ``runLogger.is_debug`` reports DEBUG mode enabled."""
    try:
        from shared import runLogger as _RL

        return _RL.is_debug()
    except Exception:
        return os.getenv("DEBUG_MODE", "0").strip().lower() in ("1", "true", "yes")


def _record_launcher_subprocess_error(
    *,
    component: str,
    exit_code: int | None,
    err_msg: str | None,
    log_tail: str | None = None,
) -> None:
    """Append to ``BASE_DIR/logs/program_errors.txt`` (independent of ``DEBUG_MODE``)."""
    apply_runtime_settings_from_json()
    parts: list[str] = []
    if err_msg:
        parts.append(err_msg)
    if exit_code is not None:
        parts.append(f"exit_code={exit_code}")
    summary = " — ".join(parts) if parts else "Subprocess failed"
    detail: str | None = None
    if log_tail and log_tail.strip():
        t = log_tail.strip()
        if len(t) > 8000:
            t = t[-8000:]
        detail = f"Recent output:\n{t}"
    ec = 1 if exit_code is None else int(exit_code)
    try:
        from shared import runLogger as _RL

        _RL.record_program_error_exit(
            exit_code=ec,
            summary=summary,
            detail=detail,
            source=f"email_sorter_launcher.{component}",
        )
    except Exception:
        pass


def _run_cancel_request_file() -> Path | None:
    apply_runtime_settings_from_json()
    base_raw = (os.getenv("BASE_DIR") or "").strip()
    if not base_raw:
        return None
    return Path(base_raw).expanduser().resolve() / "logs" / _LAUNCHER_CANCEL_FILE


def _request_pipeline_run_cancel() -> None:
    p = _run_cancel_request_file()
    if p is None:
        return
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text("1\n", encoding="utf-8")
    except OSError:
        pass


def _subprocess_creationflags(*, show_console: bool) -> int:
    if sys.platform != "win32":
        return 0
    if show_console and hasattr(subprocess, "CREATE_NEW_CONSOLE"):
        return subprocess.CREATE_NEW_CONSOLE  # type: ignore[attr-defined]
    if not show_console and hasattr(subprocess, "CREATE_NO_WINDOW"):
        return subprocess.CREATE_NO_WINDOW  # type: ignore[attr-defined]
    return 0


def resolve_orders_workbook_path() -> Path | None:
    """Match createExcelDocument output path (template -> .xlsm when applicable)."""
    apply_runtime_settings_from_json()
    ensure_base_dir_in_environ()
    base_raw = (os.getenv("BASE_DIR") or "").strip()
    if not base_raw:
        return None
    project_root = Path(base_raw).expanduser().resolve()

    template_path = _optional_path(
        "EXCEL_TEMPLATE_PATH",
        project_root / "email_contents" / "orders_template.xlsm",
    )
    using_template = template_path.is_file()

    excel_path = Path(
        _optional_path(
            "EXCEL_OUTPUT_PATH",
            project_root / "email_contents" / "orders.xlsm",
        )
    )

    p = excel_path
    if using_template:
        if p.suffix.lower() != ".xlsm":
            return p.with_suffix(".xlsm")
        return p
    if p.suffix.lower() == ".xlsm":
        return p.with_suffix(".xlsx")
    return p


def resolve_pdf_folder_path() -> Path:
    apply_runtime_settings_from_json()
    return ensure_base_dir_in_environ() / "email_contents" / "pdf"


def open_pdf_folder(parent: tk.Misc) -> None:
    folder = resolve_pdf_folder_path()
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        messagebox.showerror(
            "PDF folder",
            f"Could not create the PDF folder:\n{folder}\n\n{e}",
            parent=parent,
        )
        return

    try:
        if sys.platform == "win32":
            os.startfile(str(folder.resolve()))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(folder.resolve())])
        else:
            subprocess.Popen(["xdg-open", str(folder.resolve())])
    except Exception as e:
        messagebox.showerror(
            "PDF folder",
            f"Could not open the PDF folder:\n{folder}\n\n{e}",
            parent=parent,
        )


def _ask_debug_run_mode(parent: tk.Tk) -> str | None:
    """
    When DEBUG_MODE is on: choose full pipeline vs Excel-only rebuild.

    Returns ``\"full\"``, ``\"excel\"``, or ``None`` if the user closes the window with **X** (no run).
    """
    choice: list[str | None] = [None]

    win = tk.Toplevel(parent)
    win.title("Run")
    win.transient(parent)
    win.resizable(False, False)
    win.grab_set()

    def finish(value: str | None) -> None:
        choice[0] = value
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", lambda: finish(None))

    frm = tk.Frame(win, padx=16, pady=16)
    frm.pack(fill=tk.BOTH, expand=True)

    tk.Label(
        frm,
        text=(
            "DEBUG_MODE is on.\n\n"
            "Fetch new emails from your Microsoft mailbox?\n\n"
            "yes - Full run from Microsoft mailbox.\n"
            "no - Skip processing from Microsoft mailbox."
        ),
        justify=tk.LEFT,
        wraplength=480,
    ).pack(anchor=tk.W, pady=(0, 12))

    btn_row = tk.Frame(frm)
    btn_row.pack(anchor=tk.E)

    tk.Button(btn_row, text="Yes", width=10, command=lambda: finish("full")).pack(
        side=tk.LEFT, padx=(0, 8)
    )
    tk.Button(btn_row, text="No", width=10, command=lambda: finish("excel")).pack(side=tk.LEFT)

    parent.wait_window(win)
    return choice[0]


def _ask_debug_custom_html_import(parent: tk.Tk, custom_dir: Path) -> bool | None:
    """
    After the user chose **No** (Excel-only) on the first debug prompt: offer a
    full pipeline run from ``BASE_DIR/custom_import_html_files/*.html``.

    Returns ``True`` to run ``mainRunner.py --custom-import-html``, ``False`` to
    rebuild Excel from ``results.json`` only, or ``None`` if the window is closed.
    """
    choice: list[bool | None] = [None]

    win = tk.Toplevel(parent)
    win.title("Run")
    win.transient(parent)
    win.resizable(False, False)
    win.grab_set()

    def finish(value: bool | None) -> None:
        choice[0] = value
        win.destroy()

    win.protocol("WM_DELETE_WINDOW", lambda: finish(None))

    frm = tk.Frame(win, padx=16, pady=16)
    frm.pack(fill=tk.BOTH, expand=True)

    tk.Label(
        frm,
        text=(
            "Run the full pipeline using HTML files from your debug import folder?\n\n"
            f"Folder (under project):\n{custom_dir.name}/\n"
            f"Full path:\n{custom_dir}\n\n"
            "• Yes — Graph skipped: each *.html is processed like a mailbox message "
            "(extract → sort JSON → Excel). Sender/subject use debug placeholders "
            '("customImportHTML") because there is no live email metadata.\n'
            "• No — Excel only: rebuild the workbook from existing results.json "
            "(no mail, no OpenAI, no re-sort).\n\n"
            "If the folder is missing or has no .html files, the Yes run will stop "
            "after Step 2 with a short message."
        ),
        justify=tk.LEFT,
        wraplength=520,
    ).pack(anchor=tk.W, pady=(0, 12))

    btn_row = tk.Frame(frm)
    btn_row.pack(anchor=tk.E)

    tk.Button(btn_row, text="Yes", width=10, command=lambda: finish(True)).pack(
        side=tk.LEFT, padx=(0, 8)
    )
    tk.Button(btn_row, text="No", width=10, command=lambda: finish(False)).pack(side=tk.LEFT)

    parent.wait_window(win)
    return choice[0]


def _find_workbook_by_path(excel_app, path: str):
    want = str(Path(path).resolve())
    try:
        n = int(excel_app.Workbooks.Count)
    except Exception:
        return None
    for i in range(1, n + 1):
        try:
            wb = excel_app.Workbooks(i)
            cur = str(Path(str(wb.FullName)).resolve())
        except Exception:
            continue
        if cur.lower() == want.lower():
            return wb
    return None


def orders_workbook_open_in_excel(target: Path) -> bool:
    """True if Excel has this workbook open (Windows + pywin32)."""
    if sys.platform != "win32":
        return False
    try:
        import win32com.client

        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return False
    return _find_workbook_by_path(excel, str(target.resolve())) is not None


def focus_or_open_orders_workbook() -> None:
    target = resolve_orders_workbook_path()
    if target is None:
        messagebox.showerror(
            "Excel",
            "Could not resolve the project data folder. Restart the app or check Settings / email_sorter_settings.json.",
        )
        return
    if not target.is_file():
        messagebox.showwarning(
            "Excel",
            f"File does not exist yet:\n{target}\n\nRun the pipeline first to create it.",
        )
        return
    path_str = str(target.resolve())

    if sys.platform == "win32":
        try:
            import win32com.client

            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            os.startfile(path_str)
            return
        wb = _find_workbook_by_path(excel, path_str)
        if wb is not None:
            try:
                excel.Visible = True
                wb.Activate()
                try:
                    win = wb.Windows(1)
                    win.Activate()
                except Exception:
                    pass
            except Exception:
                messagebox.showinfo("Excel", "Excel already open.")
            return
        try:
            os.startfile(path_str)
        except OSError as e:
            messagebox.showerror("Excel", f"Could not open file:\n{e}")
        return

    try:
        os.startfile(path_str)
    except OSError as e:
        messagebox.showerror("Excel", f"Could not open file:\n{e}")


def prompt_update(parent: tk.Misc) -> None:
    if not _themed_ask_yes_no(
        parent,
        title="Update",
        message=(
            "Are you sure you want to force update from GitHub?\n\n"
            "This overwrites local tracked code changes."
        ),
    ):
        return

    updater = _PYTHON_FILES_DIR / "tools" / "git" / "pull_latest.py"
    if not updater.is_file():
        _themed_message(parent, title="Update", message=f"Update script not found:\n{updater}", kind="error")
        return

    try:
        result = subprocess.run(
            [sys.executable, str(updater)],
            cwd=str(_PYTHON_FILES_DIR),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=_subprocess_creationflags(show_console=False),
        )
    except Exception as e:
        _themed_message(parent, title="Update", message=f"Could not start update:\n{e}", kind="error")
        return

    details = "\n".join(
        part.strip() for part in (result.stdout, result.stderr) if part and part.strip()
    )
    if len(details) > 3000:
        details = details[-3000:]

    if result.returncode == 0:
        msg = "Force update completed from GitHub.\nRestart the app to load latest code."
        if details:
            msg += f"\n\n{details}"
        _themed_message(parent, title="Update", message=msg)
        return

    msg = "Force update failed."
    if details:
        msg += f"\n\n{details}"
    else:
        msg += f"\n\nProcess exited with code {result.returncode}."
    _themed_message(parent, title="Update", message=msg, kind="error")


def _missing_run_config_message(*, require_mail_and_azure: bool) -> str | None:
    """Return an error message if required Settings/.env values are missing."""
    apply_runtime_settings_from_json()
    missing: list[str] = []
    if require_mail_and_azure:
        if not (os.getenv("GRAPH_MAIL_FOLDER") or "").strip():
            missing.append("Mailbox folder (GRAPH_MAIL_FOLDER)")
        if not (os.getenv("AZURE_CLIENT_ID") or "").strip():
            missing.append("Azure application client ID (AZURE_CLIENT_ID)")
    if not (os.getenv("OPENAI_API_KEY") or "").strip():
        missing.append("OpenAI API key (OPENAI_API_KEY)")
    if not (os.getenv("SEVENTEEN_TRACK_API_KEY") or "").strip():
        missing.append("17TRACK API key (SEVENTEEN_TRACK_API_KEY)")
    if not missing:
        return None
    return (
        "Required configuration is missing:\n\n"
        + "\n".join(f"  - {m}" for m in missing)
        + "\n\nFill these in Settings or in python_files/.env as fallback values."
    )


def _missing_excel_menu_config_message() -> str | None:
    """Excel workbook refresh (no mailbox fetch): OpenAI + 17TRACK only."""
    return _missing_run_config_message(require_mail_and_azure=False)


def _settings_truthy(raw: str | None) -> bool:
    return (raw or "").strip().lower() in ("1", "true", "yes")


class _SettingsSwitch(tk.Frame):
    """Compact on/off control (IntVar 0/1) for Settings rows."""

    def __init__(self, parent: tk.Misc, variable: tk.IntVar) -> None:
        try:
            bg = parent.cget("bg")
        except tk.TclError:
            bg = "SystemButtonFace"
        super().__init__(parent, bg=bg)
        self._var = variable
        self._cv = tk.Canvas(
            self,
            width=48,
            height=26,
            highlightthickness=0,
            bg=bg,
            cursor="hand2",
        )
        self._cv.pack(side=tk.LEFT)
        self._var.trace_add("write", lambda *_: self._draw())
        self._cv.bind("<Button-1>", self._toggle)
        self._draw()

    def _toggle(self, _evt: object | None = None) -> None:
        self._var.set(1 - (1 if self._var.get() else 0))

    def _draw(self) -> None:
        self._cv.delete("all")
        on = bool(self._var.get())
        color = THEME["excel_accent"] if on else THEME["track"]
        self._cv.create_rectangle(2, 6, 46, 20, fill=color, outline="", width=0)
        cx = 33 if on else 15
        self._cv.create_oval(cx - 9, 4, cx + 9, 22, fill="#ffffff", outline=THEME["border"], width=1)


class SettingsDialog:
    """Edit required app values; Save writes ``python_files/email_sorter_settings.json``."""

    def __init__(self, parent: tk.Tk) -> None:
        self._saved_ok = False
        self._win = tk.Toplevel(parent)
        self._win.title("Settings")
        self._win.withdraw()
        self._win.minsize(560, 520)
        self._win.configure(bg=THEME["bg"])
        self._win.protocol("WM_DELETE_WINDOW", self._cancel)

        apply_runtime_settings_from_json()
        cfg = read_settings_json()
        mail = (cfg.get("GRAPH_MAIL_FOLDER") or os.getenv("GRAPH_MAIL_FOLDER") or "").strip()
        azure = (cfg.get("AZURE_CLIENT_ID") or os.getenv("AZURE_CLIENT_ID") or "").strip()
        tenant = (cfg.get("AZURE_TENANT_ID") or os.getenv("AZURE_TENANT_ID") or "common").strip()
        oa = (cfg.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY") or "").strip()
        t17 = (cfg.get("SEVENTEEN_TRACK_API_KEY") or os.getenv("SEVENTEEN_TRACK_API_KEY") or "").strip()
        dbg = _settings_truthy(cfg.get("DEBUG_MODE") or os.getenv("DEBUG_MODE"))
        login_next = _settings_truthy(cfg.get("LOGIN_NEW_ACCOUNT_NEXT_RUN"))

        content = self._win
        outer = tk.Frame(content, padx=18, pady=18, bg=THEME["bg"])
        outer.pack(fill=tk.BOTH, expand=True)
        label_opts = {"font": _font("body"), "fg": THEME["fg"], "bg": THEME["bg"]}
        entry_opts = {
            "font": _font("input"),
            "fg": THEME["fg"],
            "bg": THEME["surface"],
            "insertbackground": THEME["fg"],
            "relief": tk.FLAT,
            "highlightthickness": 1,
            "highlightbackground": THEME["border"],
            "highlightcolor": THEME["run_accent"],
        }

        r = 0
        tk.Label(
            outer,
            text="Mailbox folder name (GRAPH_MAIL_FOLDER)",
            anchor=tk.W,
            **label_opts,
        ).grid(
            row=r, column=0, columnspan=2, sticky=tk.W, pady=(0, 4)
        )
        self._mail = tk.Entry(outer, width=64, **entry_opts)
        self._mail.grid(row=r + 1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 8))
        self._mail.insert(0, mail)
        r += 2

        self._debug = tk.IntVar(value=1 if dbg else 0)
        dbg_row = tk.Frame(outer, bg=THEME["bg"])
        dbg_row.grid(row=r, column=0, columnspan=2, sticky=tk.W, pady=(0, 8))
        _SettingsSwitch(dbg_row, self._debug).pack(side=tk.LEFT)
        tk.Label(
            dbg_row,
            text="Detailed debug logs and extra UI (DEBUG_MODE)",
            anchor=tk.W,
            **label_opts,
        ).pack(side=tk.LEFT, padx=(10, 0))
        r += 1

        self._login_next = tk.IntVar(value=1 if login_next else 0)
        login_row = tk.Frame(outer, bg=THEME["bg"])
        login_row.grid(row=r, column=0, columnspan=2, sticky=tk.W, pady=(0, 8))
        _SettingsSwitch(login_row, self._login_next).pack(side=tk.LEFT)
        tk.Label(
            login_row,
            text="Login to new account on next run (clears saved Microsoft sign-in)",
            anchor=tk.W,
            **label_opts,
        ).pack(side=tk.LEFT, padx=(10, 0))
        r += 1

        tk.Label(
            outer,
            text="Azure application (client) ID (AZURE_CLIENT_ID)",
            anchor=tk.W,
            **label_opts,
        ).grid(row=r, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))
        self._azure = tk.Entry(outer, width=64, **entry_opts)
        self._azure.grid(row=r + 1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 8))
        self._azure.insert(0, azure)
        r += 2

        tk.Label(
            outer,
            text="Microsoft login tenant (AZURE_TENANT_ID)",
            anchor=tk.W,
            **label_opts,
        ).grid(row=r, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))
        self._tenant = tk.Entry(outer, width=64, **entry_opts)
        self._tenant.grid(row=r + 1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 4))
        self._tenant.insert(0, tenant)
        tk.Label(
            outer,
            text="Use consumers for Outlook.com; use common if you need both personal and work/school accounts.",
            anchor=tk.W,
            wraplength=520,
            justify=tk.LEFT,
            fg=THEME["muted"],
            bg=THEME["bg"],
            font=_font("body"),
        ).grid(row=r + 2, column=0, columnspan=2, sticky=tk.W, pady=(0, 8))
        r += 3

        tk.Label(
            outer,
            text="OpenAI API key (OPENAI_API_KEY)",
            anchor=tk.W,
            **label_opts,
        ).grid(row=r, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))
        self._openai = tk.Entry(outer, width=64, **entry_opts)
        self._openai.grid(row=r + 1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 8))
        self._openai.insert(0, oa)
        r += 2

        tk.Label(
            outer,
            text="17TRACK API key (SEVENTEEN_TRACK_API_KEY)",
            anchor=tk.W,
            **label_opts,
        ).grid(row=r, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))
        self._track = tk.Entry(outer, width=64, **entry_opts)
        self._track.grid(row=r + 1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 8))
        self._track.insert(0, t17)
        r += 2

        btn_row = tk.Frame(outer, bg=THEME["bg"])
        btn_row.grid(row=r, column=0, columnspan=2, sticky=tk.E)
        _make_button(
            btn_row,
            text="Cancel",
            command=self._cancel,
            bg=_DANGER_BG,
            active_bg=_DANGER_ACTIVE_BG,
        ).pack(side=tk.RIGHT, padx=(8, 0))
        _make_button(
            btn_row,
            text="Save",
            command=self._save,
            bg=THEME["excel_accent"],
            active_bg=THEME["excel_accent_dim"],
        ).pack(side=tk.RIGHT)

        outer.grid_columnconfigure(0, weight=1)
        self._win.update_idletasks()
        self._win.geometry(
            f"{max(590, outer.winfo_reqwidth() + 4)}x{outer.winfo_reqheight() + 8}"
        )
        try:
            self._win.transient(parent.winfo_toplevel())
        except tk.TclError:
            pass
        _center_window(self._win, parent)
        def focus_first_field() -> None:
            try:
                if self._win.winfo_exists():
                    self._mail.focus_force()
                    self._mail.icursor(tk.END)
            except tk.TclError:
                pass

        try:
            self._win.deiconify()
            self._win.lift(parent.winfo_toplevel())
            self._win.after(50, focus_first_field)
        except tk.TclError:
            pass

    def _cancel(self) -> None:
        self._win.destroy()

    def _save(self) -> None:
        mail = self._mail.get().strip()
        azure = self._azure.get().strip()
        tenant = self._tenant.get().strip()
        oa = self._openai.get().strip()
        t17 = self._track.get().strip()

        updates = {
            "GRAPH_MAIL_FOLDER": mail,
            "AZURE_CLIENT_ID": azure,
            "AZURE_TENANT_ID": tenant,
            "OPENAI_API_KEY": oa,
            "SEVENTEEN_TRACK_API_KEY": t17,
            "DEBUG_MODE": "1" if self._debug.get() else "0",
            "LOGIN_NEW_ACCOUNT_NEXT_RUN": "1" if self._login_next.get() else "0",
        }
        try:
            write_settings_json(updates)
            apply_runtime_settings_from_json()
        except OSError as e:
            _themed_message(self._win, title="Settings", message=f"Could not save settings:\n{e}", kind="error")
            return
        self._saved_ok = True
        _themed_message(self._win, title="Settings", message="Saved.")
        self._win.destroy()


def main() -> None:
    root = tk.Tk()
    root.title("Email Sorter")
    _set_app_icon(root)
    root.configure(bg=THEME["bg"])
    apply_runtime_settings_from_json()
    root.minsize(360, 535)
    root.geometry("420x595")
    if sys.platform == "win32":
        # Windows taskbar integration is unreliable for overrideredirect root
        # windows. Keep the main menu native so Explorer always has a real app
        # button.
        app = tk.Frame(root, bg=THEME["bg"], highlightthickness=0)
        app.pack(fill=tk.BOTH, expand=True)
        root.protocol("WM_DELETE_WINDOW", root.destroy)
    else:
        app = _apply_frameless_window(root, title="Email Sorter", on_close=root.destroy)

    title_font = tkfont.Font(family="Segoe UI", size=18, weight="bold")
    btn_font = tkfont.Font(family="Segoe UI", size=14, weight="bold")

    pad = {"padx": 24, "pady": 14}
    common_btn = {
        "font": btn_font,
        "width": 18,
        "height": 2,
        "cursor": "hand2",
        "relief": tk.FLAT,
        "bd": 0,
    }

    tk.Label(
        app,
        text="Welcome to Email Sorter",
        font=title_font,
        fg=THEME["fg"],
        bg=THEME["bg"],
        pady=20,
    ).pack()

    inner = tk.Frame(app, padx=28, pady=8, bg=THEME["bg"])
    inner.pack(fill=tk.BOTH, expand=True)

    run_in_progress = False

    def set_pipeline_ui_busy(busy: bool) -> None:
        st = tk.DISABLED if busy else tk.NORMAL
        run_btn.config(state=st)
        excel_btn.config(state=st)

    def on_excel() -> None:
        if run_in_progress:
            messagebox.showinfo(
                "Excel",
                "Wait until Run finishes before opening the workbook in Excel.",
            )
            return
        ex_err = _missing_excel_menu_config_message()
        if ex_err:
            messagebox.showerror(
                "Excel — configuration required",
                ex_err + "\n\nUse Settings to save the required values.",
                parent=root,
            )
            return
        target = resolve_orders_workbook_path()
        if target is None:
            messagebox.showerror(
                "Excel",
                "Could not resolve the project data folder. Restart the app or check Settings / email_sorter_settings.json.",
            )
            return
        if orders_workbook_open_in_excel(target):
            messagebox.showwarning(
                "Excel",
                "Close the orders workbook in Excel first.\n\n"
                "The file must be closed so it can be updated from results.json before opening.",
            )
            return

        rebuild_helper = _PYTHON_FILES_DIR / "launcher_rebuild_excel.py"
        if not rebuild_helper.is_file():
            messagebox.showerror(
                "Excel",
                f"Missing helper script:\n{rebuild_helper}",
            )
            return

        excel_btn.config(state=tk.DISABLED)
        apply_runtime_settings_from_json()
        show_console = _env_debug_enabled()

        proc_holder: list[subprocess.Popen | None] = [None]

        def on_stop_excel() -> None:
            p = proc_holder[0]
            if p is not None and p.poll() is None:
                p.terminate()

        fd_skip, skip17_path = tempfile.mkstemp(prefix="email_sorter_skip17_", suffix=".flag")
        os.close(fd_skip)
        skip17_flag_path = str(Path(skip17_path).resolve())

        def on_skip_17track_excel() -> None:
            try:
                Path(skip17_flag_path).write_text("1", encoding="utf-8")
            except OSError:
                pass

        win = PipelineProgressWindow(
            root,
            title="Excel",
            headline="Updating workbook and 17TRACK data…",
            accent="excel",
            on_stop=on_stop_excel,
            on_skip_17track=on_skip_17track_excel,
            show_log=show_console,
        )

        line_q: queue.Queue[tuple[str, object]] = queue.Queue()
        done_flag = [False]
        excel_log_tail: deque[str] = deque(maxlen=40)

        def read_thread() -> None:
            env = os.environ.copy()
            env["PYTHONUNBUFFERED"] = "1"
            env["PYTHONIOENCODING"] = "utf-8"
            env["EXCEL_LAUNCHER_PROGRESS"] = "1"
            env["EMAIL_SORTER_17TRACK_QUOTA_SESSION"] = "1"
            env["EMAIL_SORTER_17TRACK_SKIP_FLAG"] = skip17_flag_path
            try:
                kw: dict = {
                    "args": [sys.executable, str(rebuild_helper), str(target.resolve())],
                    "cwd": str(_PYTHON_FILES_DIR),
                    "env": env,
                    "stdout": subprocess.PIPE,
                    "stderr": subprocess.STDOUT,
                    "stdin": subprocess.DEVNULL,
                    "text": True,
                    "encoding": "utf-8",
                    "errors": "replace",
                }
                cf = _subprocess_creationflags(show_console=False)
                if cf:
                    kw["creationflags"] = cf
                p = subprocess.Popen(**kw)
                proc_holder[0] = p
                if p.stdout:
                    for line in p.stdout:
                        line_q.put(("line", line))
                code = p.wait()
                line_q.put(("code", code))
            except Exception as e:
                line_q.put(("err", str(e)))

        def finish_excel(
            stopped: bool,
            err_msg: str | None,
            code: int | None,
            *,
            log_tail: list[str] | None = None,
        ) -> None:
            try:
                Path(skip17_flag_path).unlink(missing_ok=True)
            except OSError:
                pass

            def release_excel_btn() -> None:
                excel_btn.config(state=tk.NORMAL)

            if stopped:
                release_excel_btn()
                win.close_window()
                return

            if err_msg is not None:
                _record_launcher_subprocess_error(
                    component="excel",
                    exit_code=None,
                    err_msg=err_msg,
                    log_tail="\n".join(log_tail) if log_tail else None,
                )
                detail = "Could not update the orders workbook before opening.\n\n" + err_msg
                if log_tail:
                    tail = "\n".join(log_tail).strip()
                    if len(tail) > 4000:
                        tail = tail[-4000:]
                    detail = f"{detail}\n\nRecent output:\n{tail}"
                if show_console:
                    win.prepare_error_review_mode(
                        status_message=(
                            "Excel rebuild failed before finishing.\n\n"
                            "Review the debug log above, then click Close."
                        ),
                        on_dismiss=lambda: (
                            release_excel_btn(),
                            messagebox.showerror("Excel", detail, parent=root),
                        ),
                    )
                    return
                release_excel_btn()
                win.close_window()
                messagebox.showerror("Excel", detail, parent=root)
                return

            if code not in (0, None):
                _record_launcher_subprocess_error(
                    component="excel",
                    exit_code=int(code),
                    err_msg=f"Excel rebuild subprocess failed (exit code {code})",
                    log_tail="\n".join(log_tail) if log_tail else None,
                )
                brief = f"Rebuild failed (exit code {code})."
                if log_tail:
                    tail = "\n".join(log_tail).strip()
                    if len(tail) > 4000:
                        tail = tail[-4000:]
                    brief = f"{brief}\n\nRecent output:\n{tail}"
                if show_console:
                    win.prepare_error_review_mode(
                        status_message=(
                            f"{brief}\n\nReview the debug log above, then click Close."
                        ),
                        on_dismiss=lambda: (
                            release_excel_btn(),
                            messagebox.showerror("Excel", brief, parent=root),
                        ),
                    )
                    return
                release_excel_btn()
                win.close_window()
                messagebox.showerror("Excel", brief, parent=root)
                return

            release_excel_btn()
            win.close_window()
            focus_or_open_orders_workbook()

        def pump_excel() -> None:
            if done_flag[0]:
                return
            try:
                while True:
                    kind, payload = line_q.get_nowait()
                    if kind == "line":
                        line = str(payload)
                        parsed = parse_excel_progress_line(line)
                        if parsed:
                            pct, msg = parsed
                            win.set_progress(pct, msg or None)
                        else:
                            one = line.rstrip("\n\r")
                            if one.strip():
                                excel_log_tail.append(one[-800:])
                        if show_console:
                            win.append_log(line)
                    elif kind == "code":
                        done_flag[0] = True
                        stopped = win.stop_requested
                        c = int(payload)
                        tail = list(excel_log_tail)
                        root.after(
                            0,
                            lambda s=stopped, c=c, t=tail: finish_excel(
                                s,
                                None,
                                c,
                                log_tail=t,
                            ),
                        )
                        return
                    elif kind == "err":
                        done_flag[0] = True
                        stopped = win.stop_requested
                        msg = str(payload)
                        tail = list(excel_log_tail)
                        root.after(
                            0,
                            lambda s=stopped, m=msg, t=tail: finish_excel(
                                s,
                                m,
                                None,
                                log_tail=t,
                            ),
                        )
                        return
            except queue.Empty:
                pass
            root.after(60, pump_excel)

        threading.Thread(target=read_thread, daemon=True).start()
        root.after(30, pump_excel)

    def on_run() -> None:
        nonlocal run_in_progress
        if run_in_progress:
            return
        script = _PYTHON_FILES_DIR / "mainRunner.py"
        if not script.is_file():
            messagebox.showerror("Run", f"Missing {script.name}")
            return
        target = resolve_orders_workbook_path()
        if target is not None and target.is_file() and orders_workbook_open_in_excel(target):
            messagebox.showwarning(
                "Run",
                "Close the orders workbook in Excel before running.\n\n"
                "The pipeline needs the file closed for a stable run.",
            )
            return

        apply_runtime_settings_from_json()
        rebuild_helper = _PYTHON_FILES_DIR / "launcher_rebuild_excel.py"
        is_excel_rebuild_only = False

        if _env_debug_enabled():
            mode = _ask_debug_run_mode(root)
            if mode is None:
                return
            if mode == "full":
                run_args = [sys.executable, str(script)]
            else:
                ensure_base_dir_in_environ()
                base_raw = (os.getenv("BASE_DIR") or "").strip()
                custom_dir = Path(base_raw).expanduser().resolve() / "custom_import_html_files"
                sub = _ask_debug_custom_html_import(root, custom_dir)
                if sub is None:
                    return
                if sub:
                    run_args = [sys.executable, str(script), "--custom-import-html"]
                else:
                    if not rebuild_helper.is_file():
                        messagebox.showerror(
                            "Run",
                            f"Missing helper script:\n{rebuild_helper}",
                        )
                        return
                    target_wb = resolve_orders_workbook_path()
                    if target_wb is None:
                        messagebox.showerror(
                            "Run",
                            "Could not resolve the orders workbook path. Restart the app or check Settings / email_sorter_settings.json.",
                        )
                        return
                    is_excel_rebuild_only = True
                    run_args = [sys.executable, str(rebuild_helper), str(target_wb.resolve())]
        else:
            run_args = [sys.executable, str(script)]

        cfg_err = _missing_run_config_message(require_mail_and_azure=not is_excel_rebuild_only)
        if cfg_err:
            title = "Run — configuration required"
            if is_excel_rebuild_only:
                title = "Run (Excel-only) — configuration required"
            messagebox.showerror(
                title,
                cfg_err + "\n\nUse Settings to save the required values.",
                parent=root,
            )
            return

        apply_runtime_settings_from_json()
        if _settings_truthy(os.getenv("LOGIN_NEW_ACCOUNT_NEXT_RUN")):
            try:
                (_PYTHON_FILES_DIR / ".graph_token_cache.bin").unlink(missing_ok=True)
            except OSError:
                pass
            try:
                merged = read_settings_for_write_merge()
                merged["LOGIN_NEW_ACCOUNT_NEXT_RUN"] = "0"
                write_settings_json(merged)
            except (OSError, ValueError):
                os.environ["LOGIN_NEW_ACCOUNT_NEXT_RUN"] = "0"
            else:
                apply_runtime_settings_from_json()

        run_in_progress = True
        set_pipeline_ui_busy(True)
        show_console = _env_debug_enabled()

        skip17_flag_path_run: str | None = None
        if is_excel_rebuild_only:
            fd_sr, skip17_path_run = tempfile.mkstemp(
                prefix="email_sorter_skip17_", suffix=".flag"
            )
            os.close(fd_sr)
            skip17_flag_path_run = str(Path(skip17_path_run).resolve())

        def on_skip_17track_run() -> None:
            p = skip17_flag_path_run
            if not p:
                return
            try:
                Path(p).write_text("1", encoding="utf-8")
            except OSError:
                pass

        proc_holder: list[subprocess.Popen | None] = [None]

        def on_stop_run() -> None:
            if is_excel_rebuild_only:
                p = proc_holder[0]
                if p is not None and p.poll() is None:
                    p.terminate()
            else:
                _request_pipeline_run_cancel()

        win = PipelineProgressWindow(
            root,
            title="Excel rebuild" if is_excel_rebuild_only else "Run",
            headline=(
                "Rebuilding workbook (17TRACK prefetch)…"
                if is_excel_rebuild_only
                else "Running pipeline…"
            ),
            accent="excel" if is_excel_rebuild_only else "run",
            on_stop=on_stop_run,
            on_skip_17track=on_skip_17track_run if is_excel_rebuild_only else None,
            show_log=show_console,
        )

        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        env["EMAIL_SORTER_17TRACK_QUOTA_SESSION"] = "1"
        if is_excel_rebuild_only:
            env["EXCEL_LAUNCHER_PROGRESS"] = "1"
            if skip17_flag_path_run:
                env["EMAIL_SORTER_17TRACK_SKIP_FLAG"] = skip17_flag_path_run
        else:
            env["EMAIL_SORTER_LAUNCHER_PROGRESS"] = "1"

        line_q: queue.Queue[tuple[str, object]] = queue.Queue()
        done_flag = [False]
        run_log_tail: deque[str] = deque(maxlen=40)

        def read_thread() -> None:
            try:
                kw: dict = {
                    "args": run_args,
                    "cwd": str(_PYTHON_FILES_DIR),
                    "env": env,
                    "stdout": subprocess.PIPE,
                    "stderr": subprocess.STDOUT,
                    "stdin": subprocess.DEVNULL,
                    "text": True,
                    "encoding": "utf-8",
                    "errors": "replace",
                }
                cf = _subprocess_creationflags(show_console=False)
                if cf:
                    kw["creationflags"] = cf
                p = subprocess.Popen(**kw)
                proc_holder[0] = p
                if p.stdout:
                    for line in p.stdout:
                        line_q.put(("line", line))
                code = p.wait()
                line_q.put(("code", code))
            except Exception as e:
                line_q.put(("err", str(e)))

        def finish_run(
            stopped: bool,
            err_msg: str | None,
            code: int | None,
            *,
            log_tail: list[str] | None = None,
        ) -> None:
            nonlocal run_in_progress
            if skip17_flag_path_run:
                try:
                    Path(skip17_flag_path_run).unlink(missing_ok=True)
                except OSError:
                    pass

            def release_run_ui() -> None:
                nonlocal run_in_progress
                run_in_progress = False
                set_pipeline_ui_busy(False)

            if stopped and is_excel_rebuild_only:
                release_run_ui()
                win.close_window()
                return
            if stopped:
                release_run_ui()
                win.close_window()
                return

            if err_msg is not None:
                _record_launcher_subprocess_error(
                    component="run",
                    exit_code=None,
                    err_msg=err_msg,
                    log_tail="\n".join(log_tail) if log_tail else None,
                )
                if show_console:
                    win.prepare_error_review_mode(
                        status_message=(
                            "The pipeline subprocess failed before finishing.\n\n"
                            "Review the debug log above, then click Close."
                        ),
                        on_dismiss=lambda: (
                            release_run_ui(),
                            messagebox.showerror("Run", err_msg, parent=root),
                        ),
                    )
                    return
                release_run_ui()
                win.close_window()
                messagebox.showerror("Run", err_msg, parent=root)
                return

            if code == 3:
                tail_txt = None
                if log_tail:
                    tail_txt = "\n".join(log_tail).strip()
                    if len(tail_txt) > 8000:
                        tail_txt = tail_txt[-8000:]
                _record_launcher_subprocess_error(
                    component="run",
                    exit_code=3,
                    err_msg="OpenAI rate limit or fatal OpenAI error (exit code 3)",
                    log_tail=tail_txt,
                )
                openai_msg = (
                    "The pipeline stopped due to an OpenAI rate limit or another fatal "
                    "OpenAI error. Check logs and your OpenAI billing/settings."
                )
                if show_console:
                    win.prepare_error_review_mode(
                        status_message=(
                            "Pipeline exited with code 3 (OpenAI rate limit or fatal OpenAI error).\n\n"
                            "Review the debug log above, then click Close."
                        ),
                        on_dismiss=lambda: (
                            release_run_ui(),
                            messagebox.showerror("Run", openai_msg, parent=root),
                        ),
                    )
                    return
                release_run_ui()
                win.close_window()
                messagebox.showerror("Run", openai_msg, parent=root)
                return

            if code not in (0, None):
                tail_txt = None
                if log_tail:
                    tail_txt = "\n".join(log_tail).strip()
                    if len(tail_txt) > 8000:
                        tail_txt = tail_txt[-8000:]
                _record_launcher_subprocess_error(
                    component="run",
                    exit_code=int(code),
                    err_msg=f"Pipeline subprocess exited with code {code}",
                    log_tail=tail_txt,
                )
                msg = f"Pipeline exited with code {code}."
                if log_tail:
                    tail = "\n".join(log_tail).strip()
                    if len(tail) > 4000:
                        tail = tail[-4000:]
                    msg = f"{msg}\n\nRecent output:\n{tail}"
                if show_console:
                    win.prepare_error_review_mode(
                        status_message=(
                            f"Pipeline exited with code {code}.\n\n"
                            "Review the debug log above, then click Close."
                        ),
                        on_dismiss=lambda: (
                            release_run_ui(),
                            messagebox.showerror("Run", msg, parent=root),
                        ),
                    )
                    return
                release_run_ui()
                win.close_window()
                messagebox.showerror("Run", msg, parent=root)
                return

            release_run_ui()
            win.close_window()

        def pump_run() -> None:
            if done_flag[0]:
                return
            try:
                while True:
                    kind, payload = line_q.get_nowait()
                    if kind == "line":
                        line = str(payload)
                        if is_excel_rebuild_only:
                            parsed = parse_excel_progress_line(line)
                        else:
                            parsed = parse_run_progress_line(line)
                        if parsed:
                            pct, msg = parsed
                            win.set_progress(pct, msg or None)
                        else:
                            one = line.rstrip("\n\r")
                            if one.strip():
                                run_log_tail.append(one[-800:])
                        if show_console:
                            win.append_log(line)
                    elif kind == "code":
                        done_flag[0] = True
                        stopped = win.stop_requested
                        c = int(payload)
                        tail = list(run_log_tail)
                        root.after(
                            0,
                            lambda s=stopped, co=c, t=tail: finish_run(s, None, co, log_tail=t),
                        )
                        return
                    elif kind == "err":
                        done_flag[0] = True
                        stopped = win.stop_requested
                        msg = str(payload)
                        root.after(0, lambda s=stopped, m=msg: finish_run(s, m, None))
                        return
            except queue.Empty:
                pass
            root.after(60, pump_run)

        threading.Thread(target=read_thread, daemon=True).start()
        root.after(30, pump_run)

    run_btn = tk.Button(
        inner,
        text="Run",
        bg=THEME["run_accent"],
        fg="#ffffff",
        activebackground=THEME["run_accent_dim"],
        activeforeground="#ffffff",
        command=on_run,
        **common_btn,
    )
    _add_button_hover(run_btn, normal_bg=THEME["run_accent"], hover_bg="#60a5fa")
    run_btn.pack(fill=tk.X, **pad)

    excel_row = tk.Frame(inner, bg=THEME["bg"])
    excel_row.pack(fill=tk.X, **pad)

    excel_btn = tk.Button(
        excel_row,
        text="Excel",
        bg=THEME["excel_accent"],
        fg="#ffffff",
        activebackground=THEME["excel_accent_dim"],
        activeforeground="#ffffff",
        command=on_excel,
        **common_btn,
    )
    _add_button_hover(excel_btn, normal_bg=THEME["excel_accent"], hover_bg="#4ade80")
    excel_btn.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    tk.Frame(excel_row, bg=THEME["run_accent_dim"], width=2).pack(side=tk.LEFT, fill=tk.Y)

    pdf_folder_icon = _make_file_explorer_icon(root)
    pdf_folder_btn = tk.Button(
        excel_row,
        image=pdf_folder_icon,
        bg=THEME["excel_accent"],
        activebackground=THEME["excel_accent_dim"],
        relief=tk.FLAT,
        bd=0,
        highlightthickness=0,
        width=64,
        height=54,
        cursor="hand2",
        command=lambda: open_pdf_folder(root),
    )
    pdf_folder_btn.image = pdf_folder_icon  # type: ignore[attr-defined]
    _add_button_hover(pdf_folder_btn, normal_bg=THEME["excel_accent"], hover_bg="#4ade80")
    _attach_tooltip(pdf_folder_btn, "Open PDF folder")
    pdf_folder_btn.pack(side=tk.LEFT, fill=tk.Y)

    update_btn = tk.Button(
        inner,
        text="Update",
        bg=_UPDATE_BG,
        fg="#0f1117",
        activebackground=_UPDATE_ACTIVE_BG,
        activeforeground="#0f1117",
        command=lambda: prompt_update(root),
        **common_btn,
    )
    _add_button_hover(update_btn, normal_bg=_UPDATE_BG, hover_bg="#fbbf24")
    update_btn.pack(fill=tk.X, **pad)

    settings_btn = tk.Button(
        inner,
        text="Settings",
        bg=THEME["surface"],
        fg=THEME["fg"],
        activebackground=THEME["track"],
        activeforeground=THEME["fg"],
        highlightthickness=1,
        highlightbackground=THEME["border"],
        highlightcolor=THEME["border"],
        command=lambda: SettingsDialog(root),
        **common_btn,
    )
    _add_button_hover(settings_btn, normal_bg=THEME["surface"], hover_bg=THEME["track"])
    settings_btn.pack(fill=tk.X, **pad)

    exit_btn = tk.Button(
        inner,
        text="Exit",
        bg=THEME["stop_fg"],
        fg="#ffffff",
        activebackground="#da3633",
        activeforeground="#ffffff",
        command=root.destroy,
        **common_btn,
    )
    _add_button_hover(exit_btn, normal_bg=THEME["stop_fg"], hover_bg="#ff6b66")
    exit_btn.pack(fill=tk.X, **pad)

    root.mainloop()


if __name__ == "__main__":
    main()
