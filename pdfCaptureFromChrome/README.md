# PDF capture (mitmproxy)

Intercepts HTTPS responses in **Google Chrome** (isolated profile) and saves **carrier PDFs** (e.g. FedEx proof-of-delivery) to disk using **mitmproxy**. Stops after **one** PDF per run, closes Chrome, and shows a short confirmation window (optional first-page preview).

In this repository the tool lives under **`pdfCaptureFromChrome/`** (run commands from that directory).

## Requirements

- Python 3.10+
- **Google Chrome** (installed normally on Windows)
- **mitmproxy** and dependencies:

```bash
pip install -r requirements_mitmproxy.txt
```

`requirements_mitmproxy.txt` includes `mitmproxy`, `pymupdf`, and `Pillow` (for the success-dialog preview). Preview is skipped if those are missing.

## First-time HTTPS (mitm CA)

mitmproxy decrypts HTTPS; the browser must trust its certificate once per Windows user:

1. Run with no arguments (opens `http://mitm.it` through the proxy).
2. Download/install the **Windows** certificate into **Trusted Root Certification Authorities** (see mitmproxy docs).
3. On later runs you can pass a tracking URL directly.

## Usage

Run from this directory:

```bash
python run_pdf_capture.py
```

### Positional arguments

After stripping an **optional last token** `0` or `1`:

| Remaining args | Meaning |
|----------------|---------|
| *(none)* | Open `http://mitm.it`; save under `captured_pdfs/captured.pdf` |
| `https://...` | Open that URL first; same default folder/name |
| `URL` `DIR` `file.pdf` | Open `URL`; save as `DIR\<stem>.pdf` (e.g. `file.pdf` → stem `file`) |

**Trailing debug flag (optional):**

| Last token | Effect |
|------------|--------|
| *(omitted)* | **Debug** (default): writes `mitmdump.stdout.log` / `mitmdump.stderr.log`, verbose console + Chrome launcher output, `[pdf preview]` hints on stderr if preview fails |
| `1` | Same as omitted (explicit debug) |
| `0` | **Quiet**: mitmdump stdout/stderr discarded (those log files are **not** written this run); minimal console; no Chrome launch spam; preview failure messages suppressed |

Examples:

```bash
python run_pdf_capture.py "https://www.fedex.com/wtrk/track/?tracknumbers=YOURNUMBER"
python run_pdf_capture.py "https://..." "C:\output" "pod.pdf"
python run_pdf_capture.py 0
python run_pdf_capture.py "https://..." 0
python run_pdf_capture.py "https://..." "C:\output" "pod.pdf" 1
```

Optional flags: `--port 8080`, `--chrome-path "C:\Path\chrome.exe"`.

Press **Ctrl+C** in the terminal to cancel before a PDF is captured (in debug mode a `[cancelled]` line is printed; quiet mode stays silent).

## Outputs

| Path | Purpose |
|------|---------|
| `captured_pdfs/` | Default PDF output (or custom `DIR` from 3-arg form) |
| `mitmdump.stdout.log` / `mitmdump.stderr.log` | mitmdump logs (**debug mode only**; quiet mode does not write these) |
| `chrome_user_data_mitm/` | Isolated Chrome profile (delete to reset) |
| `.pdfCaptureFromChrome_done.json` | Internal signal file (removed after success) |

## Embedding in another app

Import paths from `paths.py` (or the package `__init__.py`):

- `PDF_CAPTURE_ROOT` — package directory (use as `cwd` for subprocesses)
- `normalize_start_url`, `is_mitm_it_install_url`, `split_debug_positional`, `PDF_CAPTURE_DONE_FILE`

Or spawn:

`python run_pdf_capture.py ...` with `cwd=PDF_CAPTURE_ROOT`.

## Scripts

- **`run_pdf_capture.py`** — main entry: mitmdump + Chrome + success UI.
- **`launch_mitm_chrome.py`** — Chrome only (for manual mitmdump runs).
- **`mitm_pdf_interceptor/mitm_pdf_addon.py`** — mitmproxy addon (PDF detection + save).

## Troubleshooting

- **“Connection not private”** — Install the mitm CA (see above) or run with no args and use `http://mitm.it` first.
- **No preview in the dialog** — Install `pymupdf` and `Pillow`; check stderr for `[pdf preview]` messages.
- **FedEx / site blocks** — Some sites block automated or proxied access; this tool only captures what the browser loads through the proxy.
