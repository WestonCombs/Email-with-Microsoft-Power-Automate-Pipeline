"""
Passive PDF capture via mitmproxy + isolated Chrome.

Stops after **one** saved PDF: closes Chrome and mitmdump, then shows a confirmation
window with an optional first-page image preview (PyMuPDF + Pillow). Auto-closes
after 10 seconds.

Optional **last positional** ``0`` = quiet (no mitmdump log files, minimal console),
``1`` or omitted = debug (writes ``mitmdump.*.log``, verbose output).

    pip install -r requirements_mitmproxy.txt

    python run_pdf_capture.py
    python run_pdf_capture.py 0
    python run_pdf_capture.py "https://..." 1
    python run_pdf_capture.py "https://..." "C:\\dir" "file.pdf" 0

Press Ctrl+C to cancel before capture.

When launched from the Shipping Status viewer, the console is detached — errors and step traces
are appended to ``<BASE_DIR>/logs/pdfCaptureFromChrome/pdf_capture_session.log``.
"""

from __future__ import annotations

import argparse
import ctypes
import json
import os
import shutil
import subprocess
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

from chrome_devtools import (  # noqa: E402
    export_page_pdf,
    inspect_page,
    list_page_targets,
    looks_like_print_preview,
    next_pdf_path,
    reserve_free_port,
    wait_for_debugger,
)
from paths import (  # noqa: E402
    PDF_CAPTURE_SESSION_LOG,
    PDF_CAPTURE_STDERR_LOG,
    PDF_CAPTURE_STDOUT_LOG,
    PDF_CAPTURE_DONE_FILE,
    PDF_CAPTURE_ROOT,
    PREVIEW_MAX_HEIGHT,
    PREVIEW_MAX_WIDTH,
    default_pdf_output_dir,
    ensure_import_path,
    is_mitm_it_install_url,
    normalize_start_url,
    split_debug_positional,
)

ensure_import_path()

SUCCESS_DIALOG_SECONDS = 10

# When the viewer detaches the console (FreeConsole), stderr is invisible — always append here too.
SESSION_LOG = PDF_CAPTURE_SESSION_LOG


def _log_session(message: str) -> None:
    line = f"{datetime.now().isoformat(timespec='seconds')} {message}"
    try:
        with open(SESSION_LOG, "a", encoding="utf-8", newline="\n") as f:
            f.write(line + "\n")
    except OSError:
        pass


def _log_mitmdump_file_tails(out_path: Path, err_path: Path) -> None:
    """After mitmdump dies, append last lines of its logs to the session log (debug mode)."""
    for label, path in (("stdout", out_path), ("stderr", err_path)):
        try:
            txt = path.read_text(encoding="utf-8", errors="replace").strip()
            if not txt:
                _log_session(f"mitmdump {label} log: (empty)")
                continue
            lines = txt.splitlines()
            tail = "\n".join(lines[-15:])
            _log_session(f"mitmdump {label} tail:\n{tail}")
        except OSError as e:
            _log_session(f"mitmdump {label} log read failed: {e}")


def _write_capture_done(saved_path: Path) -> None:
    payload = {
        "path": str(saved_path.resolve()),
        "filename": saved_path.name,
    }
    PDF_CAPTURE_DONE_FILE.write_text(json.dumps(payload), encoding="utf-8", newline="\n")


def _expected_manual_save_path(output_dir: Path | None, output_basename: str | None) -> Path | None:
    if output_dir is None or not output_basename:
        return None
    return output_dir.expanduser().resolve() / f"{output_basename}.pdf"


def _copy_text_to_clipboard_windows(text: str) -> bool:
    if sys.platform != "win32" or not text:
        return False

    GMEM_MOVEABLE = 0x0002
    CF_UNICODETEXT = 13

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    data = ctypes.create_unicode_buffer(text)
    size_bytes = ctypes.sizeof(data)
    handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, size_bytes)
    if not handle:
        return False

    locked = kernel32.GlobalLock(handle)
    if not locked:
        kernel32.GlobalFree(handle)
        return False

    try:
        ctypes.memmove(locked, ctypes.addressof(data), size_bytes)
    finally:
        kernel32.GlobalUnlock(handle)

    if not user32.OpenClipboard(None):
        kernel32.GlobalFree(handle)
        return False

    try:
        user32.EmptyClipboard()
        if not user32.SetClipboardData(CF_UNICODETEXT, handle):
            kernel32.GlobalFree(handle)
            return False
        handle = None
        return True
    finally:
        user32.CloseClipboard()
        if handle:
            kernel32.GlobalFree(handle)


def _prime_manual_save_clipboard(expected_path: Path | None, *, verbose: bool) -> None:
    if expected_path is None:
        return
    if _copy_text_to_clipboard_windows(str(expected_path)):
        _log_session(f"Copied manual Save As path to clipboard: {expected_path}")
        if verbose:
            print(f"Save As path copied to clipboard:\n{expected_path}\n")
    else:
        _log_session(f"Could not copy manual Save As path to clipboard: {expected_path}")


class _PrintPreviewFallback:
    """Fallback exporter for browser-generated print previews that never hit mitm."""

    def __init__(self, debug_port: int, output_dir: Path, basename: str, *, verbose: bool) -> None:
        self.debug_port = debug_port
        self.output_dir = output_dir
        self.basename = basename
        self.verbose = verbose
        self.ready = wait_for_debugger(debug_port, timeout=10.0)
        self.started_at = time.time()
        self._last_attempt: dict[str, float] = {}
        if self.ready:
            try:
                targets = list_page_targets(debug_port)
                _log_session(
                    f"DevTools ready on port={debug_port}; initial page targets={len(targets)}"
                )
            except Exception as e:
                _log_session(f"DevTools target listing failed after ready: {e}")
        else:
            _log_session(f"DevTools not ready on port={debug_port}; print-preview fallback disabled")

    def poll(self) -> tuple[str, Path | None]:
        """
        Returns:
            ("none", None)   - keep waiting
            ("saved", Path)  - fallback PDF saved
            ("closed", None) - browser window(s) gone before save
        """
        if not self.ready:
            return "none", None

        try:
            targets = list_page_targets(self.debug_port)
        except Exception as e:
            _log_session(f"DevTools poll failed on port={self.debug_port}: {e}")
            return "closed", None

        if not targets and (time.time() - self.started_at) > 2.0:
            _log_session("No Chrome page targets remain; treating capture as cancelled by browser close")
            return "closed", None

        now = time.time()
        for target in targets:
            target_id = str(target.get("id") or "")
            if not target_id:
                continue
            if (now - self._last_attempt.get(target_id, 0.0)) < 1.0:
                continue
            self._last_attempt[target_id] = now

            ws_url = str(target.get("webSocketDebuggerUrl") or "")
            if not ws_url:
                continue

            page_info = inspect_page(ws_url)
            if not looks_like_print_preview(target, page_info):
                continue

            desc = str((page_info or {}).get("title") or target.get("title") or "")
            href = str((page_info or {}).get("href") or target.get("url") or "")
            _log_session(
                f"Print-preview fallback candidate target_id={target_id} title={desc!r} href={href!r}"
            )
            try:
                pdf_bytes = export_page_pdf(ws_url)
                out_path = next_pdf_path(self.output_dir, self.basename)
                out_path.write_bytes(pdf_bytes)
                _write_capture_done(out_path)
                _log_session(f"Print-preview fallback saved PDF to {out_path}")
                return "saved", out_path
            except Exception as e:
                _log_session(f"Print-preview fallback failed for target_id={target_id}: {e}")

        return "none", None


def _basename_from_filename(filename: str) -> str:
    """Stem for PDF_INTERCEPTOR_BASENAME (saved as ``{stem}.pdf``)."""
    p = Path(filename.strip().strip('"'))
    if p.suffix.lower() == ".pdf":
        return p.stem
    return p.name or "captured"


def _mitm_setup(
    port: int,
    *,
    output_dir: Path | None = None,
    output_basename: str | None = None,
) -> tuple[list[str], Path, Path, Path]:
    """Configure env for the PDF addon; return mitmdump cmd, captured dir, log paths."""
    if output_dir is not None:
        captured = output_dir.expanduser().resolve()
        captured.mkdir(parents=True, exist_ok=True)
        os.environ["PDF_INTERCEPTOR_OUTPUT_DIR"] = str(captured)
    else:
        captured = default_pdf_output_dir()
        os.environ["PDF_INTERCEPTOR_OUTPUT_DIR"] = str(captured)

    if output_basename is not None:
        os.environ["PDF_INTERCEPTOR_BASENAME"] = output_basename
    else:
        os.environ.setdefault("PDF_INTERCEPTOR_BASENAME", "captured")

    os.environ["PDF_INTERCEPTOR_MAX_PDFS"] = "1"
    os.environ.setdefault(
        "PDF_INTERCEPTOR_URL_KEYWORDS",
        "proof,delivery,pod,label,tracking,shipment,invoice,fedex,ups",
    )
    os.environ["PDF_CAPTURE_DONE_FILE"] = str(PDF_CAPTURE_DONE_FILE)

    script = PDF_CAPTURE_ROOT / "mitm_pdf_interceptor" / "mitm_pdf_addon.py"
    out_path = PDF_CAPTURE_STDOUT_LOG
    err_path = PDF_CAPTURE_STDERR_LOG
    cmd = [
        "mitmdump",
        "-s",
        str(script),
        "--listen-port",
        str(port),
    ]
    return cmd, captured, out_path, err_path


def _terminate_chrome_process(proc: subprocess.Popen | None) -> None:
    if proc is None or proc.poll() is not None:
        return
    if sys.platform == "win32":
        kwargs: dict = {"capture_output": True}
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        subprocess.run(
            ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
            **kwargs,
        )
    else:
        proc.terminate()
        try:
            proc.wait(timeout=8)
        except subprocess.TimeoutExpired:
            proc.kill()


def _terminate_mitmdump(proc: subprocess.Popen) -> None:
    if proc.poll() is not None:
        return
    proc.terminate()
    try:
        proc.wait(timeout=12)
    except subprocess.TimeoutExpired:
        proc.kill()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            pass


def _stop_chrome_and_mitm(chrome_proc: subprocess.Popen | None, mitm_proc: subprocess.Popen) -> None:
    _terminate_chrome_process(chrome_proc)
    _terminate_mitmdump(mitm_proc)


def _pdf_first_page_preview_photoimage(full_path: Path, *, master, verbose: bool) -> object | None:
    """Render first PDF page for Tk preview, or ``None`` if unavailable."""
    try:
        import fitz  # PyMuPDF
        from io import BytesIO

        from PIL import Image, ImageTk
    except ImportError:
        if verbose:
            print(
                "[pdf preview] Missing pymupdf or Pillow. Run: pip install pymupdf Pillow",
                file=sys.stderr,
            )
        return None

    path = full_path.expanduser().resolve()
    if not path.is_file():
        if verbose:
            print(f"[pdf preview] Not a file: {path}", file=sys.stderr)
        return None

    try:
        doc = fitz.open(path)
        if len(doc) < 1:
            doc.close()
            return None
        page = doc[0]
        mat = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        doc.close()
        img = Image.open(BytesIO(pix.tobytes("png"))).convert("RGB")
        w, h = img.size
        ratio = min(PREVIEW_MAX_WIDTH / w, PREVIEW_MAX_HEIGHT / h, 1.0)
        if ratio < 1.0:
            img = img.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(img, master=master)
    except Exception as e:
        if verbose:
            print(f"[pdf preview] Could not render PDF: {e}", file=sys.stderr)
        return None


def _show_success_dialog(full_path: Path, filename: str, *, verbose: bool) -> None:
    title = "PDF captured"
    text_fallback = (
        "PDF successfully grabbed.\n\n"
        f"File name:\n{filename}\n\n"
        f"Full path:\n{full_path}"
    )
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception:
        print(text_fallback, file=sys.stdout)
        return

    root = tk.Tk()
    root.title(title)
    try:
        root.attributes("-topmost", True)
    except tk.TclError:
        pass

    preview = _pdf_first_page_preview_photoimage(full_path, master=root, verbose=verbose)

    frm = ttk.Frame(root, padding=20)
    frm.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frm, text="PDF successfully grabbed.", font=("", 11, "bold")).pack(anchor=tk.W)

    if preview is not None:
        pic = tk.Label(frm, image=preview, bd=0)
        pic.pack(anchor=tk.CENTER, pady=(10, 12))
        pic.image = preview

    ttk.Label(frm, text=f"File name:\n{filename}\n\nFull path:\n{full_path}", justify=tk.LEFT).pack(
        anchor=tk.W, pady=(4, 0)
    )

    root.after(SUCCESS_DIALOG_SECONDS * 1000, root.destroy)
    root.mainloop()


def run_capture_session(
    port: int,
    chrome_path: Path | None,
    start_url: str,
    *,
    output_dir: Path | None = None,
    output_basename: str | None = None,
    debug: bool = True,
) -> int:
    """Start mitmdump + Chrome; exit after first PDF or Ctrl+C.

    ``debug``: when False, mitmdump stdout/stderr go to ``DEVNULL`` (no ``mitmdump.*.log``),
    and console/Chrome launcher output is reduced.
    """
    from launch_mitm_chrome import launch_isolated_chrome

    _log_session(
        f"run_capture_session begin port={port} debug={debug} start_url={start_url!r} "
        f"output_dir={output_dir!r} basename={output_basename!r} "
        f"mitmdump={shutil.which('mitmdump')!r} chrome_path={chrome_path!r}"
    )

    PDF_CAPTURE_DONE_FILE.unlink(missing_ok=True)
    expected_manual_save_path = _expected_manual_save_path(output_dir, output_basename)

    cmd, captured, out_path, err_path = _mitm_setup(
        port,
        output_dir=output_dir,
        output_basename=output_basename,
    )
    ts = datetime.now().isoformat(timespec="seconds")
    banner = f"=== mitmdump session started {ts} (cwd={PDF_CAPTURE_ROOT}) ===\n"

    if debug:
        print("PDF capture (mitmdump + isolated Chrome) — stops after 1 PDF is saved.")
        print("Starting:", " ".join(cmd))
        print(f"Stdout log: {out_path}")
        print(f"Stderr log: {err_path}")
        print(f"PDF output:  {captured}")
        if output_basename:
            print(f"PDF basename: {output_basename} → {output_basename}.pdf")
        print(f"Chrome opens: {start_url}")
        print(f"Proxy (this Chrome only): 127.0.0.1:{port}")
        if is_mitm_it_install_url(start_url):
            print("First-time HTTPS: install the mitm CA from this page.")
        else:
            print(
                "If you see HTTPS errors, run with no arguments first, open http://mitm.it, "
                "install the CA, then try your tracking URL again.",
            )
        print("Press Ctrl+C to cancel before capture.\n")
    else:
        print("PDF capture (quiet). Ctrl+C to cancel.\n")

    _prime_manual_save_clipboard(expected_manual_save_path, verbose=debug)

    out_f = None
    err_f = None
    proc: subprocess.Popen | None = None
    chrome_proc: subprocess.Popen | None = None
    debug_port = reserve_free_port()
    print_preview_fallback: _PrintPreviewFallback | None = None
    try:
        if debug:
            out_f = open(out_path, "w", encoding="utf-8", newline="\n")
            err_f = open(err_path, "w", encoding="utf-8", newline="\n")
            out_f.write(banner)
            err_f.write(banner)
            out_f.flush()
            err_f.flush()
            stdout_dest: int | object = out_f
            stderr_dest: int | object = err_f
        else:
            stdout_dest = subprocess.DEVNULL
            stderr_dest = subprocess.DEVNULL

        proc = subprocess.Popen(
            cmd,
            cwd=PDF_CAPTURE_ROOT,
            stdout=stdout_dest,
            stderr=stderr_dest,
        )
        _log_session(f"mitmdump Popen pid={proc.pid} cmd={' '.join(cmd)}")
        time.sleep(1.0)
        if proc.poll() is not None:
            rc = proc.returncode or 1
            _log_session(f"mitmdump exited before Chrome (immediate) rc={rc}")
            if debug:
                try:
                    if out_f is not None:
                        out_f.close()
                        out_f = None
                    if err_f is not None:
                        err_f.close()
                        err_f = None
                except OSError:
                    pass
                _log_mitmdump_file_tails(out_path, err_path)
            else:
                _log_session(
                    "Quiet mode hides mitmdump output. Set DEBUG_MODE=1 for full logs, or "
                    "if port 8080 is in use set PDF_CAPTURE_MITM_PORT in .env to a free port (e.g. 8082)."
                )
            if debug:
                print(
                    "mitmdump exited immediately; check mitmdump.stderr.log / mitmdump.stdout.log.",
                    file=sys.stderr,
                )
            else:
                print("mitmdump exited immediately.", file=sys.stderr)
            return rc

        chrome_proc = launch_isolated_chrome(
            port,
            chrome_path=chrome_path,
            start_url=start_url,
            remote_debugging_port=debug_port,
            verbose=debug,
        )
        if chrome_proc is None:
            _log_session("launch_isolated_chrome returned None — chrome.exe missing or Popen failed (see stderr above if console)")
            _terminate_mitmdump(proc)
            return 1
        _log_session(f"Chrome Popen pid={chrome_proc.pid}")
        print_preview_fallback = _PrintPreviewFallback(
            debug_port,
            captured,
            output_basename or os.environ.get("PDF_INTERCEPTOR_BASENAME", "captured"),
            verbose=debug,
        )

        try:
            while proc.poll() is None:
                if PDF_CAPTURE_DONE_FILE.is_file():
                    break
                if expected_manual_save_path is not None and expected_manual_save_path.is_file():
                    _write_capture_done(expected_manual_save_path)
                    _log_session(
                        f"Detected manually saved PDF on disk: {expected_manual_save_path}"
                    )
                    break
                if chrome_proc.poll() is not None:
                    _log_session(
                        f"Chrome process exited before capture completion rc={chrome_proc.returncode}"
                    )
                    _terminate_mitmdump(proc)
                    return 130
                if print_preview_fallback is not None:
                    state, _saved_path = print_preview_fallback.poll()
                    if state == "saved":
                        break
                    if state == "closed":
                        _stop_chrome_and_mitm(chrome_proc, proc)
                        return 130
                time.sleep(0.25)
        except KeyboardInterrupt:
            if debug:
                print("\n[cancelled]")
            _stop_chrome_and_mitm(chrome_proc, proc)
            return 130

        if proc.poll() is not None and not PDF_CAPTURE_DONE_FILE.is_file():
            _log_session(f"mitmdump died during wait rc={proc.returncode}")
            _terminate_chrome_process(chrome_proc)
            return proc.returncode or 1

        if not PDF_CAPTURE_DONE_FILE.is_file():
            _stop_chrome_and_mitm(chrome_proc, proc)
            return 1

        try:
            data = json.loads(PDF_CAPTURE_DONE_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as e:
            print(f"Could not read capture done file: {e}", file=sys.stderr)
            _stop_chrome_and_mitm(chrome_proc, proc)
            return 1

        full_path = Path(data.get("path", ""))
        filename = data.get("filename", full_path.name)

        _stop_chrome_and_mitm(chrome_proc, proc)

        PDF_CAPTURE_DONE_FILE.unlink(missing_ok=True)

        _show_success_dialog(full_path, filename, verbose=debug)
        return 0
    finally:
        if out_f is not None:
            out_f.close()
        if err_f is not None:
            err_f.close()


def main(argv: list[str] | None = None) -> int:
    try:
        from dotenv import load_dotenv

        load_dotenv(PDF_CAPTURE_ROOT.parent / ".env", override=False)
    except ImportError:
        pass

    default_port = 8080
    raw_port = os.environ.get("PDF_CAPTURE_MITM_PORT", "").strip()
    if raw_port:
        try:
            default_port = int(raw_port)
        except ValueError:
            pass

    _log_session(
        f"main argv={argv if argv is not None else sys.argv[1:]!r} cwd={os.getcwd()} "
        f"exe={sys.executable} default_mitm_port_from_env={default_port}"
    )
    p = argparse.ArgumentParser(
        description="Capture carrier PDFs: mitmdump plus isolated Chrome proxied through it.",
    )
    p.add_argument(
        "positional",
        nargs="*",
        metavar="ARG",
        help=(
            "Optional last argument 0 = quiet (no mitmdump log files), 1 = debug (default if omitted). "
            "Otherwise: 0 args → defaults; 1 arg → URL; 3 args → URL, output dir, filename."
        ),
    )
    p.add_argument(
        "--port",
        type=int,
        default=default_port,
        help=f"mitmproxy listen port (default {default_port}; override with PDF_CAPTURE_MITM_PORT in .env)",
    )
    p.add_argument(
        "--chrome-path",
        type=Path,
        default=None,
        help="Path to chrome.exe if not found automatically",
    )
    try:
        args = p.parse_args(argv)
        rest, debug = split_debug_positional(list(args.positional))
        _log_session(f"parsed positional (stripped 0/1) rest={rest!r} verbose_mitmdump={debug}")

        if len(rest) not in (0, 1, 3):
            p.error(
                "After removing optional trailing 0/1, provide 0, 1, or 3 arguments "
                "(URL; or URL, directory, filename). "
                f"Got {len(rest)} argument(s)."
            )

        if len(rest) == 3:
            return run_capture_session(
                args.port,
                args.chrome_path,
                normalize_start_url(rest[0]),
                output_dir=Path(rest[1].strip().strip('"')),
                output_basename=_basename_from_filename(rest[2]),
                debug=debug,
            )

        start_url = normalize_start_url(rest[0] if len(rest) == 1 else None)
        return run_capture_session(args.port, args.chrome_path, start_url, debug=debug)
    except SystemExit as e:
        code = e.code
        if isinstance(code, int) and code != 0:
            _log_session(f"argparse/SystemExit code={code!r}")
        raise
    except Exception:
        _log_session(f"main exception:\n{traceback.format_exc()}")
        raise


if __name__ == "__main__":
    raise SystemExit(main())
