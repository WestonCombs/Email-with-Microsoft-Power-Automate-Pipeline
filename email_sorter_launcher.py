"""Tkinter launcher — main screen for Email Sorter (Run, Excel, Update, Settings, Exit)."""

from __future__ import annotations

import os
import re
import subprocess
import sys
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import messagebox

from dotenv import load_dotenv

_PYTHON_FILES_DIR = Path(__file__).resolve().parent
_ENV_PATH = _PYTHON_FILES_DIR / ".env"

_KEY_LINE = re.compile(r"^([A-Za-z_][A-Za-z0-9_]*)=(.*)$")


def _optional_path(env_name: str, default: Path) -> Path:
    raw = os.getenv(env_name)
    if raw:
        return Path(raw).expanduser().resolve()
    return default


def resolve_orders_workbook_path() -> Path | None:
    """Match createExcelDocument output path (template -> .xlsm when applicable)."""
    load_dotenv(_ENV_PATH, override=True)
    base_raw = (os.getenv("BASE_DIR") or "").strip()
    if not base_raw:
        return None
    project_root = Path(base_raw).expanduser().resolve()

    template_path = _optional_path(
        "EXCEL_TEMPLATE_PATH", _PYTHON_FILES_DIR / "orders_template.xlsm"
    )
    using_template = template_path.is_file()

    excel_path = Path(
        _optional_path(
            "EXCEL_OUTPUT_PATH",
            project_root / "email_contents" / "orders.xlsx",
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


def focus_or_open_orders_workbook() -> None:
    target = resolve_orders_workbook_path()
    if target is None:
        messagebox.showerror(
            "Excel",
            "BASE_DIR is not set in .env.\nSet it under Settings or edit python_files/.env.",
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


def start_main_runner() -> None:
    script = _PYTHON_FILES_DIR / "mainRunner.py"
    if not script.is_file():
        messagebox.showerror("Run", f"Missing {script.name}")
        return
    kwargs: dict = {
        "args": [sys.executable, str(script)],
        "cwd": str(_PYTHON_FILES_DIR),
    }
    if sys.platform == "win32" and hasattr(subprocess, "CREATE_NEW_CONSOLE"):
        kwargs["creationflags"] = subprocess.CREATE_NEW_CONSOLE  # type: ignore[assignment]
    try:
        subprocess.Popen(**kwargs)
    except OSError as e:
        messagebox.showerror("Run", str(e))


def prompt_update() -> None:
    messagebox.askyesno(
        "Update",
        "Are you sure you want to update?",
    )


class SettingsDialog:
    def __init__(self, parent: tk.Tk) -> None:
        self._win = tk.Toplevel(parent)
        self._win.title("Settings")
        self._win.transient(parent)
        self._win.grab_set()
        self._win.minsize(480, 160)

        load_dotenv(_ENV_PATH, override=True)
        mail = (os.getenv("GRAPH_MAIL_FOLDER") or "").strip()
        base = (os.getenv("BASE_DIR") or "").strip()

        frm = tk.Frame(self._win, padx=16, pady=16)
        frm.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            frm,
            text="Mailbox folder name (GRAPH_MAIL_FOLDER)",
            anchor=tk.W,
        ).grid(row=0, column=0, sticky=tk.W, pady=(0, 4))
        self._mail = tk.Entry(frm, width=64)
        self._mail.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 12))
        self._mail.insert(0, mail)

        tk.Label(
            frm,
            text="Project folder on disk (BASE_DIR)",
            anchor=tk.W,
        ).grid(row=2, column=0, sticky=tk.W, pady=(0, 4))
        self._base = tk.Entry(frm, width=64)
        self._base.grid(row=3, column=0, columnspan=2, sticky=tk.EW, pady=(0, 12))
        self._base.insert(0, base)

        tk.Label(
            frm,
            text="Leave a field blank and click Save to keep the current value in .env.",
            fg="#555",
            anchor=tk.W,
            wraplength=520,
            justify=tk.LEFT,
        ).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(0, 12))

        btn_row = tk.Frame(frm)
        btn_row.grid(row=5, column=0, columnspan=2, sticky=tk.E)
        tk.Button(btn_row, text="Cancel", command=self._win.destroy).pack(
            side=tk.RIGHT, padx=(8, 0)
        )
        tk.Button(btn_row, text="Save", command=self._save).pack(side=tk.RIGHT)

        frm.grid_columnconfigure(0, weight=1)

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
    root.minsize(360, 500)
    root.geometry("420x560")

    title_font = tkfont.Font(size=18, weight="bold")
    btn_font = tkfont.Font(size=14, weight="bold")

    pad = {"padx": 24, "pady": 14}
    common_btn = {
        "font": btn_font,
        "width": 18,
        "height": 2,
        "cursor": "hand2",
        "relief": tk.RAISED,
        "bd": 2,
    }

    tk.Label(
        root,
        text="Welcome to Email Sorter",
        font=title_font,
        pady=20,
    ).pack()

    inner = tk.Frame(root, padx=28, pady=8)
    inner.pack(fill=tk.BOTH, expand=True)

    tk.Button(
        inner,
        text="Run",
        bg="#1e88e5",
        fg="white",
        activebackground="#1565c0",
        activeforeground="white",
        command=start_main_runner,
        **common_btn,
    ).pack(fill=tk.X, **pad)

    tk.Button(
        inner,
        text="Excel",
        bg="#43a047",
        fg="white",
        activebackground="#2e7d32",
        activeforeground="white",
        command=focus_or_open_orders_workbook,
        **common_btn,
    ).pack(fill=tk.X, **pad)

    tk.Button(
        inner,
        text="Update",
        bg="#fb8c00",
        fg="white",
        activebackground="#ef6c00",
        activeforeground="white",
        command=prompt_update,
        **common_btn,
    ).pack(fill=tk.X, **pad)

    tk.Button(
        inner,
        text="Settings",
        bg="#546e7a",
        fg="white",
        activebackground="#37474f",
        activeforeground="white",
        command=lambda: SettingsDialog(root),
        **common_btn,
    ).pack(fill=tk.X, **pad)

    tk.Button(
        inner,
        text="Exit",
        bg="#e53935",
        fg="white",
        activebackground="#c62828",
        activeforeground="white",
        command=root.destroy,
        **common_btn,
    ).pack(fill=tk.X, **pad)

    root.mainloop()


if __name__ == "__main__":
    main()

