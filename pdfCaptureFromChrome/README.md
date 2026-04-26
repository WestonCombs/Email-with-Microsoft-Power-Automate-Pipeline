# PDF capture (under `python_files/`)

## Primary workflow: proof-of-delivery (no MITM)

The **Shipping status (17TRACK)** window uses **Assisted PDF Capture**, **Play/Pause**, and **Ctrl+Shift+P** in the isolated capture Chrome. Implementation: **`html_capture/`** (Chrome DevTools `Page.printToPDF`).

## MITM (archival / optional)

The mitmproxy-based entrypoint and addon were moved out of this folder to keep the `python_files` tree smaller:

- **`../mitm_pdf_capture/run_pdf_capture.py`**
- **`../mitm_pdf_capture/mitm_pdf_interceptor/`**
- **`../mitm_pdf_capture/requirements_mitmproxy.txt`**

See `../mitm_pdf_capture/README.md` for CLI and wizard use.

## Still in this package

- **`chrome_devtools.py`**, **`paths.py`**, **`launch_mitm_chrome.py`** — shared by `html_capture` and by `../mitm_pdf_capture/run_pdf_capture.py`
- Runtime logs and Chrome profile: `<BASE_DIR>/logs/pdfCaptureFromChrome/` (see `paths.py`)
