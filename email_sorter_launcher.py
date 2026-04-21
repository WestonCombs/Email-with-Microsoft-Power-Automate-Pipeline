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

from dotenv import load_dotenv

from launcher_progress_ui import (
    THEME,
    PipelineProgressWindow,
    parse_excel_progress_line,
    parse_run_progress_line,
)

_PYTHON_FILES_DIR = Path(__file__).resolve().parent
_LAUNCHER_CANCEL_FILE = ".email_sorter_cancel"
_ENV_PATH = _PYTHON_FILES_DIR / ".env"
_PDF_CAPTURE_REQUIREMENTS = _PYTHON_FILES_DIR / "pdfCaptureFromChrome" / "requirements_mitmproxy.txt"

_MITM_WIZARD_BG_OPACITY = 0.50  # 50/50 blend: gray base + image (image half visible)

_KEY_LINE = re.compile(r"^([A-Za-z_][A-Za-z0-9_]*)=(.*)$")


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
    load_dotenv(_ENV_PATH, override=True)
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
    load_dotenv(_ENV_PATH, override=True)
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
    load_dotenv(_ENV_PATH, override=True)
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


def merge_env_keys(env_path: Path, updates: dict[str, str]) -> None:
    """Insert or replace KEY=value lines; leave all other lines unchanged."""
    if not updates:
        return
    text = env_path.read_text(encoding="utf-8") if env_path.is_file() else ""
    lines = text.splitlines(keepends=False)
    seen: set[str] = set()
    out: list[str] = []
    for line in lines:
        m = _KEY_LINE.match(line.strip())
        if m:
            key = m.group(1)
            if key in updates:
                out.append(f"{key}={updates[key]}")
                seen.add(key)
                continue
        out.append(line)
    for key, val in updates.items():
        if key not in seen:
            out.append(f"{key}={val}")
    env_path.write_text("\n".join(out) + ("\n" if out else ""), encoding="utf-8")


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
            'BASE_DIR is not set.\nOpen Settings, set "Project folder on disk", and click Save.',
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


def prompt_update() -> None:
    if not messagebox.askyesno(
        "Update",
        "Are you sure you want to force update from GitHub?\n\n"
        "This overwrites local tracked code changes.",
    ):
        return

    updater = _PYTHON_FILES_DIR / "tools" / "git" / "pull_latest.py"
    if not updater.is_file():
        messagebox.showerror("Update", f"Update script not found:\n{updater}")
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
        messagebox.showerror("Update", f"Could not start update:\n{e}")
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
        messagebox.showinfo("Update", msg)
        return

    msg = "Force update failed."
    if details:
        msg += f"\n\n{details}"
    else:
        msg += f"\n\nProcess exited with code {result.returncode}."
    messagebox.showerror("Update", msg)


class MitmInitializeWindow:
    """Step-through MITM education; last step launches capture Chrome + mitmdump."""

    _SLIDES: tuple[dict, ...] = (
        {
            "title": "PDF capture and MITM",
            "body": (
                "Email Sorter can capture carrier PDFs by running a local HTTPS proxy (mitmproxy) "
                "and a dedicated Chrome profile. That proxy must decrypt TLS briefly so it can see "
                "responses — which requires installing mitmproxy’s certificate authority (CA) "
                "where your browser trusts it.\n\n"
                "Only proceed if you accept the risk: any software that trusts that CA could be "
                "fooled by a forged certificate if the CA or its private key were ever misused. "
                "Use the → arrow (or Next) to read short background articles, then install the CA "
                "on the last step using the same isolated Chrome this tool uses."
            ),
        },
        {
            "title": "How mitmproxy works (official)",
            "body": (
                "mitmproxy sits between your browser and the internet. For HTTPS it presents "
                "its own certificate chain signed by its local CA. That is why you must install "
                "that CA as trusted — otherwise the browser will warn or block the connection."
            ),
            "article_url": "https://docs.mitmproxy.org/stable/concepts/how-mitmproxy-works",
        },
        {
            "title": "MITM and SSL/TLS hijacking",
            "body": (
                "A man-in-the-middle can decrypt traffic if the client trusts the attacker’s CA. "
                "Debugging tools do this on purpose on your machine; attackers try to trick you "
                "into trusting a malicious CA. Understand the parallel before you add a new root."
            ),
            "article_url": "https://www.invicti.com/learn/mitm-ssl-hijacking",
        },
        {
            "title": "Trusted root certificates",
            "body": (
                "Operating systems and browsers ship with a set of trusted root CAs. Adding "
                "another root increases what you implicitly trust. Remove or avoid installing "
                "debugging CAs on machines where you do not need them."
            ),
            "article_url": "https://www.threatdown.com/blog/when-you-shouldnt-trust-a-trusted-root-certificate",
        },
        {
            "title": "Before you install",
            "body": (
                "1. Install mitmproxy so mitmdump is on your PATH (e.g. pip install mitmproxy).\n"
                "2. The next step opens https://mitm.it/ inside the same isolated Chrome profile "
                "used for PDF capture, with mitmdump running — install the mitmproxy CA there so "
                "HTTPS works for capture.\n"
                "3. On Windows, you may also install the CA into the system Trusted Root store "
                "if you need other tools to trust it; the wizard’s last step focuses on "
                "capture Chrome."
            ),
        },
        {
            "title": "Install the CA (capture Chrome + mitmdump)",
            "body": (
                "Click the button below. Your Python environment will install/update PDF capture "
                "dependencies (mitmproxy, PyMuPDF, Pillow) from requirements_mitmproxy.txt first; "
                "then mitmdump starts and isolated Chrome opens https://mitm.it/ — install the "
                "mitmproxy CA in that browser profile. Close Chrome when you are done; capture "
                "runs again from the Shipping status window."
            ),
            "final_install": True,
        },
    )

    def __init__(self, parent: tk.Toplevel | tk.Tk) -> None:
        self._dlg = tk.Toplevel(parent)
        self._dlg.title("Initialize MITM (PDF capture)")
        self._dlg.transient(parent)
        self._dlg.grab_set()
        self._dlg.minsize(560, 420)
        self._dlg.geometry("600x480")

        self._idx = 0
        self._mitm_bg_photo: object | None = None
        self._mitm_pil_original = None
        self._outer_win_id: int | None = None

        load_dotenv(_ENV_PATH, override=True)
        base_raw = (os.getenv("BASE_DIR") or "").strip()
        bg_candidates: list[Path] = []
        if base_raw:
            bg_candidates.append(
                Path(base_raw).expanduser().resolve() / "assets" / "images" / "shrek.png"
            )
        bg_candidates.append(_PYTHON_FILES_DIR / "assets" / "images" / "shrek.png")
        self._mitm_pil_original = None
        try:
            from PIL import Image as _PILImage

            for cand in bg_candidates:
                if cand.is_file():
                    self._mitm_pil_original = _PILImage.open(cand).convert("RGB")
                    break
        except Exception:
            self._mitm_pil_original = None

        self._mitm_canvas = tk.Canvas(self._dlg, highlightthickness=0, bg="#ececec")
        self._mitm_canvas.pack(fill=tk.BOTH, expand=True)

        outer = tk.Frame(self._mitm_canvas, padx=14, pady=12, bg="#ececec")
        self._outer = outer

        self._counter_lbl = tk.Label(outer, text="", fg="#555", anchor=tk.W, bg="#ececec")
        self._counter_lbl.pack(fill=tk.X, pady=(0, 4))

        self._title_lbl = tk.Label(
            outer, text="", font=("", 12, "bold"), anchor=tk.W, justify=tk.LEFT, bg="#ececec"
        )
        self._title_lbl.pack(fill=tk.X, pady=(0, 8))

        body_frame = tk.Frame(outer, bg="#ececec")
        body_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self._body_text = tk.Text(
            body_frame,
            wrap=tk.WORD,
            height=14,
            width=68,
            padx=4,
            pady=4,
            relief=tk.FLAT,
            highlightthickness=0,
            cursor="arrow",
            bg="#fafafa",
            insertbackground="#222",
        )
        self._body_text.pack(fill=tk.BOTH, expand=True)

        self._links_frame = tk.Frame(outer, bg="#ececec")

        self._install_frame = tk.Frame(outer, bg="#ececec")
        self._install_font = tkfont.Font(size=10, weight="bold", underline=True)
        self._install_btn = tk.Button(
            self._install_frame,
            text="start mitm.it installation 😊",
            command=self._launch_pdf_capture_helper,
            cursor="hand2",
            font=self._install_font,
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief=tk.RAISED,
            bd=2,
            padx=14,
            pady=8,
        )
        self._install_btn.pack(anchor=tk.CENTER)

        self._nav = tk.Frame(outer, bg="#ececec")
        self._nav.pack(fill=tk.X, pady=(8, 0))
        self._btn_prev = tk.Button(self._nav, text="← Previous", command=self._prev, cursor="hand2")
        self._btn_prev.pack(side=tk.LEFT)
        self._btn_next = tk.Button(self._nav, text="Next →", command=self._next, cursor="hand2")
        self._btn_next.pack(side=tk.RIGHT)

        self._hint = tk.Frame(outer, bg="#ececec")
        self._hint.pack(fill=tk.X, pady=(6, 0))
        self._btn_close = tk.Button(self._hint, text="Close", command=self._dlg.destroy, cursor="hand2")
        self._btn_close.pack(side=tk.RIGHT)

        for w in (
            self._dlg,
            self._mitm_canvas,
            outer,
            body_frame,
            self._body_text,
            self._title_lbl,
            self._counter_lbl,
            self._nav,
            self._hint,
            self._links_frame,
            self._install_frame,
            self._btn_prev,
            self._btn_next,
            self._install_btn,
            self._btn_close,
        ):
            self._apply_arrow_bindings(w)

        self._mitm_canvas.bind("<Configure>", self._mitm_on_canvas_configure)

        self._render()
        self._dlg.after(50, lambda: self._body_text.focus_set())

    def _mitm_on_canvas_configure(self, event: tk.Event) -> None:
        w, h = max(event.width, 2), max(event.height, 2)

        self._mitm_canvas.delete("mitm_bg")

        if self._mitm_pil_original is not None and w >= 32 and h >= 32:
            try:
                from PIL import Image, ImageTk

                base = Image.new("RGB", (w, h), (236, 236, 236))
                try:
                    resample = Image.Resampling.LANCZOS
                except AttributeError:
                    resample = Image.LANCZOS  # type: ignore[attr-defined]
                fg = self._mitm_pil_original.resize((w, h), resample).convert("RGB")
                blended = Image.blend(base, fg, _MITM_WIZARD_BG_OPACITY)
                self._mitm_bg_photo = ImageTk.PhotoImage(blended)
                self._mitm_canvas.create_image(
                    0, 0, anchor=tk.NW, image=self._mitm_bg_photo, tags="mitm_bg"
                )
                self._mitm_canvas.tag_lower("mitm_bg")
            except Exception:
                self._mitm_bg_photo = None

        # Inset so the blended image stays visible around the panel (Tk frames are opaque).
        margin = 0.10
        iw = max(int(w * (1 - 2 * margin)), 260)
        ih = max(int(h * (1 - 2 * margin)), 200)
        cx, cy = w // 2, h // 2

        if self._outer_win_id is None:
            self._outer_win_id = self._mitm_canvas.create_window(
                cx,
                cy,
                window=self._outer,
                anchor=tk.CENTER,
                width=iw,
                height=ih,
            )
        else:
            self._mitm_canvas.coords(self._outer_win_id, cx, cy)
            self._mitm_canvas.itemconfig(self._outer_win_id, width=iw, height=ih)

        if self._outer_win_id is not None:
            self._mitm_canvas.tag_raise(self._outer_win_id)

    def _arrow_left(self, _e: tk.Event) -> str | None:
        self._prev()
        return "break"

    def _arrow_right(self, _e: tk.Event) -> str | None:
        self._next()
        return "break"

    def _apply_arrow_bindings(self, w: tk.Misc) -> None:
        w.bind("<Left>", self._arrow_left)
        w.bind("<Right>", self._arrow_right)

    def _spawn_pdf_capture_process(self) -> None:
        script = _PYTHON_FILES_DIR / "pdfCaptureFromChrome" / "run_pdf_capture.py"
        cwd = _PYTHON_FILES_DIR / "pdfCaptureFromChrome"
        if not script.is_file():
            messagebox.showerror("PDF capture", f"Missing script:\n{script}")
            return
        try:
            subprocess.Popen(
                [sys.executable, str(script)],
                cwd=str(cwd),
            )
        except OSError as e:
            messagebox.showerror("PDF capture", f"Could not start:\n{e}")

    def _launch_pdf_capture_helper(self) -> None:
        script = _PYTHON_FILES_DIR / "pdfCaptureFromChrome" / "run_pdf_capture.py"
        if not script.is_file():
            messagebox.showerror("PDF capture", f"Missing script:\n{script}")
            return
        if not _PDF_CAPTURE_REQUIREMENTS.is_file():
            messagebox.showerror(
                "PDF capture",
                f"Requirements file not found:\n{_PDF_CAPTURE_REQUIREMENTS}",
            )
            return

        busy = tk.Toplevel(self._dlg)
        busy.title("Installing dependencies")
        busy.transient(self._dlg)
        busy.resizable(False, False)
        tk.Label(
            busy,
            text=(
                "Installing PDF capture dependencies from\n"
                "pdfCaptureFromChrome/requirements_mitmproxy.txt\n\n"
                "(mitmproxy, PyMuPDF, Pillow)\n\n"
                "Please wait — mitm.it opens after this finishes."
            ),
            justify=tk.CENTER,
            padx=20,
            pady=20,
        ).pack()
        busy.update_idletasks()
        x = self._dlg.winfo_rootx() + (self._dlg.winfo_width() // 2) - (busy.winfo_reqwidth() // 2)
        y = self._dlg.winfo_rooty() + (self._dlg.winfo_height() // 2) - (busy.winfo_reqheight() // 2)
        busy.geometry(f"+{x}+{y}")

        self._install_btn.config(state=tk.DISABLED)

        def worker() -> None:
            err: str | None = None
            try:
                run_kw: dict = {
                    "args": [
                        sys.executable,
                        "-m",
                        "pip",
                        "install",
                        "-r",
                        str(_PDF_CAPTURE_REQUIREMENTS),
                    ],
                    "cwd": str(_PYTHON_FILES_DIR),
                    "capture_output": True,
                    "text": True,
                    "timeout": 900,
                }
                if sys.platform == "win32" and hasattr(subprocess, "CREATE_NO_WINDOW"):
                    run_kw["creationflags"] = subprocess.CREATE_NO_WINDOW
                proc = subprocess.run(**run_kw)
                if proc.returncode != 0:
                    tail = (proc.stderr or proc.stdout or "").strip()
                    err = tail[-6000:] if tail else f"pip exited with code {proc.returncode}"
            except subprocess.TimeoutExpired:
                err = "pip install timed out (over 15 minutes)."
            except OSError as e:
                err = str(e)

            def finish() -> None:
                try:
                    busy.destroy()
                except tk.TclError:
                    pass
                self._install_btn.config(state=tk.NORMAL)
                if err is not None:
                    messagebox.showerror(
                        "Dependency install failed",
                        "Could not install PDF capture dependencies.\n\n" + err,
                    )
                    return
                self._spawn_pdf_capture_process()

            self._dlg.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def _render(self) -> None:
        slide = self._SLIDES[self._idx]
        n = len(self._SLIDES)
        self._counter_lbl.config(text=f"Step {self._idx + 1} of {n}")
        self._title_lbl.config(text=slide["title"])

        self._body_text.configure(state=tk.NORMAL)
        self._body_text.delete("1.0", tk.END)
        self._body_text.insert(tk.END, slide["body"])
        self._body_text.configure(state=tk.DISABLED)

        for c in self._links_frame.winfo_children():
            c.destroy()

        article_url = (slide.get("article_url") or "").strip()
        if article_url:
            self._links_frame.pack(fill=tk.X, pady=(0, 10), before=self._nav)
            row = tk.Frame(self._links_frame, bg="#ececec")
            row.pack(anchor=tk.W, pady=(4, 0))
            tk.Label(
                row,
                text="click here to read an article about it: ",
                anchor=tk.W,
                fg="#333",
                bg="#ececec",
            ).pack(side=tk.LEFT)
            url_lbl = tk.Label(
                row,
                text=article_url,
                fg="#0066cc",
                cursor="hand2",
                anchor=tk.W,
                bg="#ececec",
            )
            url_lbl.pack(side=tk.LEFT)
            url_lbl.bind("<Button-1>", lambda e, u=article_url: webbrowser.open(u))
            self._apply_arrow_bindings(row)
            self._apply_arrow_bindings(url_lbl)
        else:
            self._links_frame.pack_forget()

        if slide.get("final_install"):
            self._install_frame.pack(fill=tk.X, pady=(0, 12), before=self._nav)
        else:
            self._install_frame.pack_forget()

        self._btn_prev.config(state=tk.NORMAL if self._idx > 0 else tk.DISABLED)
        self._btn_next.config(state=tk.NORMAL if self._idx < n - 1 else tk.DISABLED)

    def _prev(self) -> None:
        if self._idx > 0:
            self._idx -= 1
            self._render()

    def _next(self) -> None:
        if self._idx < len(self._SLIDES) - 1:
            self._idx += 1
            self._render()


class SettingsDialog:
    def __init__(self, parent: tk.Tk) -> None:
        self._win = tk.Toplevel(parent)
        self._win.title("Settings")
        self._win.transient(parent)
        self._win.grab_set()
        self._win.minsize(480, 280)

        load_dotenv(_ENV_PATH, override=True)
        mail = (os.getenv("GRAPH_MAIL_FOLDER") or "").strip()
        base = (os.getenv("BASE_DIR") or "").strip()

        outer = tk.Frame(self._win, padx=16, pady=16)
        outer.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            outer,
            text="Mailbox folder name (GRAPH_MAIL_FOLDER)",
            anchor=tk.W,
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 4))
        self._mail = tk.Entry(outer, width=64)
        self._mail.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 12))
        self._mail.insert(0, mail)

        tk.Label(
            outer,
            text="Project folder on disk (BASE_DIR)",
            anchor=tk.W,
        ).grid(row=2, column=0, sticky=tk.W, pady=(0, 4))
        self._base = tk.Entry(outer, width=64)
        self._base.grid(row=3, column=0, columnspan=2, sticky=tk.EW, pady=(0, 12))
        self._base.insert(0, base)

        tk.Label(
            outer,
            text="Initialize MITM (for capturing tracking PDF documents from (UPS, FEDEX, etc.))",
            anchor=tk.W,
            wraplength=520,
            justify=tk.LEFT,
        ).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(0, 4))

        tk.Button(
            outer,
            text="Begin",
            command=lambda: MitmInitializeWindow(self._win),
            cursor="hand2",
        ).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(0, 12))

        btn_row = tk.Frame(outer)
        btn_row.grid(row=6, column=0, columnspan=2, sticky=tk.E)
        tk.Button(btn_row, text="Cancel", command=self._win.destroy).pack(
            side=tk.RIGHT, padx=(8, 0)
        )
        tk.Button(btn_row, text="Save", command=self._save).pack(side=tk.RIGHT)

        outer.grid_columnconfigure(0, weight=1)

    def _save(self) -> None:
        mail_new = self._mail.get().strip()
        base_new = self._base.get().strip()
        updates: dict[str, str] = {}
        if mail_new:
            updates["GRAPH_MAIL_FOLDER"] = mail_new
        if base_new:
            updates["BASE_DIR"] = base_new
        if not updates:
            messagebox.showinfo("Settings", "No changes to save (both fields were blank).")
            self._win.destroy()
            return
        try:
            if not _ENV_PATH.is_file():
                _ENV_PATH.write_text("", encoding="utf-8")
            merge_env_keys(_ENV_PATH, updates)
            load_dotenv(_ENV_PATH, override=True)
        except OSError as e:
            messagebox.showerror("Settings", f"Could not write .env:\n{e}")
            return
        messagebox.showinfo("Settings", "Saved.")
        self._win.destroy()


def main() -> None:
    root = tk.Tk()
    root.title("Email Sorter")
    root.configure(bg=THEME["bg"])
    root.minsize(360, 500)
    root.geometry("420x560")

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
        root,
        text="Welcome to Email Sorter",
        font=title_font,
        fg=THEME["fg"],
        bg=THEME["bg"],
        pady=20,
    ).pack()

    inner = tk.Frame(root, padx=28, pady=8, bg=THEME["bg"])
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
        target = resolve_orders_workbook_path()
        if target is None:
            messagebox.showerror(
                "Excel",
                'BASE_DIR is not set.\nOpen Settings, set "Project folder on disk", and click Save.',
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
        load_dotenv(_ENV_PATH, override=True)
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

        def read_thread() -> None:
            env = os.environ.copy()
            env["PYTHONUNBUFFERED"] = "1"
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

        def finish_excel(stopped: bool, err_msg: str | None, code: int | None) -> None:
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
                )
                detail = "Could not update the orders workbook before opening.\n\n" + err_msg
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
                )
                brief = f"Rebuild failed (exit code {code})."
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
                        if show_console:
                            win.append_log(line)
                    elif kind == "code":
                        done_flag[0] = True
                        stopped = win.stop_requested
                        c = int(payload)
                        root.after(0, lambda s=stopped, c=c: finish_excel(s, None, c))
                        return
                    elif kind == "err":
                        done_flag[0] = True
                        stopped = win.stop_requested
                        msg = str(payload)
                        root.after(0, lambda s=stopped, m=msg: finish_excel(s, m, None))
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

        load_dotenv(_ENV_PATH, override=True)
        rebuild_helper = _PYTHON_FILES_DIR / "launcher_rebuild_excel.py"
        is_excel_rebuild_only = False

        if _env_debug_enabled():
            mode = _ask_debug_run_mode(root)
            if mode is None:
                return
            if mode == "full":
                run_args = [sys.executable, str(script)]
            else:
                base_raw = (os.getenv("BASE_DIR") or "").strip()
                custom_dir = (
                    Path(base_raw).expanduser().resolve() / "custom_import_html_files"
                    if base_raw
                    else Path("(set BASE_DIR in Settings)") / "custom_import_html_files"
                )
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
                            'BASE_DIR is not set.\nOpen Settings, set "Project folder on disk", and click Save.',
                        )
                        return
                    is_excel_rebuild_only = True
                    run_args = [sys.executable, str(rebuild_helper), str(target_wb.resolve())]
        else:
            run_args = [sys.executable, str(script)]

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
    run_btn.pack(fill=tk.X, **pad)

    excel_btn = tk.Button(
        inner,
        text="Excel",
        bg=THEME["excel_accent"],
        fg="#ffffff",
        activebackground=THEME["excel_accent_dim"],
        activeforeground="#ffffff",
        command=on_excel,
        **common_btn,
    )
    excel_btn.pack(fill=tk.X, **pad)

    tk.Button(
        inner,
        text="Update",
        bg="#f59e0b",
        fg="#0f1117",
        activebackground="#d97706",
        activeforeground="#0f1117",
        command=prompt_update,
        **common_btn,
    ).pack(fill=tk.X, **pad)

    tk.Button(
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
    ).pack(fill=tk.X, **pad)

    tk.Button(
        inner,
        text="Exit",
        bg=THEME["stop_fg"],
        fg="#ffffff",
        activebackground="#da3633",
        activeforeground="#ffffff",
        command=root.destroy,
        **common_btn,
    ).pack(fill=tk.X, **pad)

    root.mainloop()


if __name__ == "__main__":
    main()
