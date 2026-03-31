# Email Sorter (Python)

Fetches shopping emails via IMAP, extracts structured order data with the OpenAI API, stores results as JSON, sorts them, and builds an Excel workbook. Everything runs from a single entry point — `mainRunner.py`.

## Requirements

- **Python 3** (3.10+ recommended)
- Install dependencies:

```bash
pip install beautifulsoup4 python-dotenv openpyxl openai
```

## Configuration

Copy **`.env.example`** to **`.env`** (same folder as this README) and fill in the values:

| Variable | Purpose |
|----------|---------|
| `BASE_DIR` | Absolute path to your project data root (folders like `email_contents/` live under here). |
| `OPENAI_API_KEY` | Your OpenAI API key (used by the email extraction script). |
| `IMAP_SERVER` | IMAP server hostname (e.g. `imap.gmail.com`). |
| `IMAP_PORT` | IMAP port (default `993`). |
| `IMAP_USE_SSL` | `1` to use SSL (default), `0` for plain. |
| `IMAP_USERNAME` | Email account username. |
| `IMAP_PASSWORD` | Email account password (use an App Password for Gmail). |
| `IMAP_MAIL_FOLDER` | Mailbox folder to fetch from (e.g. `INBOX`, `Test1`). |

Optional: `DEMO_MODE=1` for demo behavior in the extractor; `openai_max_chars_per_prompting` to cap HTML size sent to the model (see `htmlHandler/convertHTMLToPlaintext.py`).

## Running

```bash
python mainRunner.py
```

This single command runs the full pipeline:

1. **Environment initialization** — creates required folders under `BASE_DIR` (`email_contents/attachments`, `email_contents/html`, `email_contents/json`, etc.).
2. **Email fetching** — connects to the IMAP server, fetches all emails from the configured folder, and saves attachments to `BASE_DIR/email_contents/attachments`.
3. **Extraction** — for each email, writes the HTML body to disk and runs the OpenAI extraction pipeline to produce structured JSON in `email_contents/json/results.json`.
4. **Sort** — sorts `results.json` by order number and purchase date.
5. **Excel export** — builds `email_contents/orders.xlsx` from the JSON and resets duplicate flags.

An OpenAI usage log (`usageN.txt`) is created per run in `BASE_DIR/email_contents/openai usage/`. A summary with total tokens, cost, and average time per email is printed at the end.

Logs and admin traces go to paths under `BASE_DIR` (e.g. `programFileOutput.txt`, `adminLog/`).

## Project layout

- `mainRunner.py` — main entry point; orchestrates the full pipeline
- `emailFetching/` — IMAP email fetching module
- `EnvironmentInitialization/` — folder/file verification for first run
- `grabbingImportantEmailContent/` — HTML → JSON (OpenAI)
- `sortJSONByOrderNumber/` — sort `results.json`
- `createExcelDocument/` — JSON → Excel
- `htmlHandler/` — HTML cleanup and helpers used by the extractor
