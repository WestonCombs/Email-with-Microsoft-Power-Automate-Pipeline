from __future__ import annotations

import importlib.util
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

if sys.platform != "win32":
    print("This helper requires Windows + Excel.", file=sys.stderr)
    sys.exit(1)

_PYTHON_FILES = Path(__file__).resolve().parent.parent
if str(_PYTHON_FILES) not in sys.path:
    sys.path.insert(0, str(_PYTHON_FILES))

from shared.settings_store import apply_runtime_settings_from_json

apply_runtime_settings_from_json()

from tkinter import BOTH, LEFT, RIGHT, W, Button, Frame, Label, Tk, Toplevel, messagebox

from shared.gui_aux_singleton import detach_console_win32, register_current_aux_gui
from giftcardInvoiceLink.excel_link_sync import find_workbook_by_path
from proofOfDelivery.pod_data import (
    AUTOMATION_HUB_CATEGORY,
    AUTOMATION_HUB_COMPANY_LABEL,
    AUTOMATION_HUB_ORDER_LABEL,
    POD_HUB_MODE,
    missing_proof_of_delivery_records,
    project_root_from_env,
    sync_proof_of_delivery_records,
)
from shared.ui_dark_theme import UI_BG, UI_FG, UI_FG_DIM, UI_BTN, UI_BTN_ACTIVE

PROJECT_ROOT = project_root_from_env()


def _create_excel_document_module():
    mod_path = _PYTHON_FILES / "createExcelDocument" / "createExcelDocument.py"
    spec = importlib.util.spec_from_file_location("_email_sorter_ced_live_sync", mod_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load {mod_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _get_excel_app():
    import win32com.client

    return win32com.client.GetActiveObject("Excel.Application")


def _copy_orders_sheet_from_temp(target_workbook, temp_workbook_path: Path) -> None:
    excel = target_workbook.Application
    old_alerts = excel.DisplayAlerts
    old_screen_updating = excel.ScreenUpdating
    temp_wb = None
    try:
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        temp_wb = excel.Workbooks.Open(str(temp_workbook_path.resolve()), UpdateLinks=False, ReadOnly=True)
        src_ws = temp_wb.Worksheets("Orders")
        try:
            old_ws = target_workbook.Worksheets("Orders")
        except Exception:
            old_ws = None
        if old_ws is not None:
            src_ws.Copy(Before=old_ws)
            new_ws = excel.ActiveSheet
            old_ws.Delete()
        else:
            src_ws.Copy(After=target_workbook.Worksheets(target_workbook.Worksheets.Count))
            new_ws = excel.ActiveSheet
        new_ws.Name = "Orders"
        target_workbook.Save()
    finally:
        if temp_wb is not None:
            try:
                temp_wb.Close(SaveChanges=False)
            except Exception:
                pass
        excel.ScreenUpdating = old_screen_updating
        excel.DisplayAlerts = old_alerts


def sync_open_workbook(workbook_path: str | Path) -> bool:
    workbook_path = str(Path(workbook_path).resolve())
    _records, changed = sync_proof_of_delivery_records(PROJECT_ROOT)

    excel = _get_excel_app()
    wb = find_workbook_by_path(excel, workbook_path)
    if wb is None:
        raise RuntimeError("Open this workbook in Excel and keep it as the active file.")
    if not changed:
        return False

    ced = _create_excel_document_module()
    suffix = Path(workbook_path).suffix.lower() or ".xlsx"
    temp_dir = Path(tempfile.mkdtemp(prefix="email_sorter_pod_sync_"))
    temp_workbook = temp_dir / f"orders_sync{suffix}"
    try:
        ced.rebuild_orders_workbook(temp_workbook)
        _copy_orders_sheet_from_temp(wb, temp_workbook)
    finally:
        try:
            if temp_workbook.is_file():
                temp_workbook.unlink()
        except OSError:
            pass
        try:
            temp_dir.rmdir()
        except OSError:
            pass
    return True


def _ask_sync_prompt(missing_count: int) -> bool:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    choice = {"sync_now": False}

    dlg = Toplevel(root)
    dlg.title("Sync POD Rows")
    dlg.transient(root)
    dlg.attributes("-topmost", True)
    dlg.resizable(False, False)
    dlg.grab_set()
    dlg.configure(bg=UI_BG)

    def finish(sync_now: bool) -> None:
        choice["sync_now"] = sync_now
        dlg.destroy()

    dlg.protocol("WM_DELETE_WINDOW", lambda: finish(False))

    outer = Frame(dlg, padx=16, pady=14, bg=UI_BG)
    outer.pack(fill=BOTH, expand=True)

    Label(
        outer,
        text=(
            "You have proof of delivery (POD) PDF files that are not in the Excel document yet.\n\n"
            f"Missing POD row(s): {missing_count}\n\n"
            "Would you like to populate them now?"
        ),
        justify=LEFT,
        wraplength=420,
        anchor=W,
        bg=UI_BG,
        fg=UI_FG,
    ).pack(fill=BOTH, expand=True)

    btn_row = Frame(outer, bg=UI_BG)
    btn_row.pack(fill=BOTH, pady=(12, 0))
    Button(
        btn_row,
        text="Yes",
        width=12,
        command=lambda: finish(True),
        bg=UI_BTN,
        fg=UI_FG,
        activebackground=UI_BTN_ACTIVE,
        activeforeground=UI_FG,
        relief="flat",
        bd=0,
    ).pack(side=LEFT, padx=(0, 8))
    Button(
        btn_row,
        text="Maybe later",
        width=12,
        command=lambda: finish(False),
        bg="#475569",
        fg=UI_FG,
        activebackground="#334155",
        activeforeground=UI_FG,
        relief="flat",
        bd=0,
    ).pack(side=RIGHT)

    root.wait_window(dlg)
    root.destroy()
    return choice["sync_now"]


def _show_status_dialog(title: str, body: str) -> None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    dlg = Toplevel(root)
    dlg.title(title)
    dlg.transient(root)
    dlg.attributes("-topmost", True)
    dlg.resizable(False, False)
    dlg.grab_set()
    dlg.configure(bg=UI_BG)

    outer = Frame(dlg, padx=16, pady=14, bg=UI_BG)
    outer.pack(fill=BOTH, expand=True)
    Label(
        outer,
        text=body,
        justify=LEFT,
        wraplength=440,
        anchor=W,
        bg=UI_BG,
        fg=UI_FG,
    ).pack(fill=BOTH, expand=True)
    Label(outer, text="", bg=UI_BG, fg=UI_FG_DIM).pack()
    Button(
        outer,
        text="Close",
        width=12,
        command=dlg.destroy,
        bg=UI_BTN,
        fg=UI_FG,
        activebackground=UI_BTN_ACTIVE,
        activeforeground=UI_FG,
        relief="flat",
        bd=0,
    ).pack(anchor="e")

    root.wait_window(dlg)
    root.destroy()


def watch_workbook(workbook_path: str | Path) -> int:
    workbook_path = str(Path(workbook_path).resolve())
    deferred_until_close = False
    last_selection_signature: tuple[str, str] | None = None

    while True:
        try:
            excel = _get_excel_app()
        except Exception:
            return 0
        wb = find_workbook_by_path(excel, workbook_path)
        if wb is None:
            return 0

        try:
            active_full_name = str(Path(str(excel.ActiveWorkbook.FullName)).resolve())
        except Exception:
            active_full_name = ""
        if active_full_name.lower() != workbook_path.lower():
            time.sleep(0.40)
            continue

        try:
            sel = excel.Selection
            selection_signature = (str(sel.Worksheet.Name), str(sel.Address))
        except Exception:
            selection_signature = None

        if selection_signature and selection_signature != last_selection_signature:
            last_selection_signature = selection_signature
            if not deferred_until_close:
                missing = missing_proof_of_delivery_records(PROJECT_ROOT)
                if missing:
                    if _ask_sync_prompt(len(missing)):
                        try:
                            changed = sync_open_workbook(workbook_path)
                            if changed:
                                _show_status_dialog(
                                    "POD Sync",
                                    "Proof of delivery rows were added to the workbook.",
                                )
                        except Exception as exc:
                            _show_status_dialog(
                                "POD Sync",
                                f"Could not sync POD rows right now:\n{exc}",
                            )
                    else:
                        deferred_until_close = True

        time.sleep(0.40)


def launch_remaining_pod_viewer(workbook_path: str | Path, row_number: str | int | None = None) -> int:
    temp_dir = Path(tempfile.mkdtemp(prefix="email_sorter_remaining_pod_"))
    txt_path = temp_dir / "remaining_pod.txt"
    ctx_path = txt_path.with_suffix(".ctx.tsv")
    txt_path.write_text("", encoding="utf-8")
    lines = [
        f"company\t{AUTOMATION_HUB_COMPANY_LABEL}\n",
        f"order_number\t{AUTOMATION_HUB_ORDER_LABEL}\n",
        f"category\t{AUTOMATION_HUB_CATEGORY}\n",
        f"pod_mode\t{POD_HUB_MODE}\n",
        f"workbook_path\t{Path(workbook_path).resolve()}\n",
    ]
    if row_number is not None:
        lines.append(f"row_number\t{row_number}\n")
    ctx_path.write_text("".join(lines), encoding="utf-8")

    viewer = _PYTHON_FILES / "trackingNumbersViewer" / "tracking_status_viewer.py"
    subprocess.Popen([sys.executable, str(viewer), str(txt_path)], cwd=str(_PYTHON_FILES))
    return 0


def main() -> int:
    detach_console_win32()
    register_current_aux_gui()

    if len(sys.argv) < 3:
        print(
            "Usage: python pod_workflow.py <watch|sync|launch-remaining> <workbook> [row]",
            file=sys.stderr,
        )
        return 1

    command = sys.argv[1].strip().lower()
    workbook_path = sys.argv[2]

    if command == "watch":
        return watch_workbook(workbook_path)
    if command == "sync":
        try:
            changed = sync_open_workbook(workbook_path)
            return 0 if changed or True else 0
        except Exception as exc:
            print(str(exc), file=sys.stderr)
            return 1
    if command == "launch-remaining":
        row_number = sys.argv[3] if len(sys.argv) > 3 else None
        return launch_remaining_pod_viewer(workbook_path, row_number)

    print(f"Unknown command: {command}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
