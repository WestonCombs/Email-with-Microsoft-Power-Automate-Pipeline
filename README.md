# Email Sorter (Python)

Scripts that read HTML shopping emails, extract structured order data with the OpenAI API, store results as JSON, optionally sort them, and build an Excel workbook. Designed to be called from **Microsoft Power Automate** (or run manually).

## Requirements

- **Python 3** (3.10+ recommended)
- Install dependencies:

```bash
pip install beautifulsoup4 python-dotenv openpyxl openai
```

## Configuration

Create or edit **`python_files/.env`** (same folder as this README):

| Variable | Purpose |
|----------|---------|
| `BASE_DIR` | Absolute path to your project data root (folders like `python_files/` live under here). |
| `OPENAI_API_KEY` | Your OpenAI API key (used by the email extraction script). |

Optional: `DEMO_MODE=1` for demo behavior in the extractor; `openai_max_chars_per_prompting` to cap HTML size sent to the model (see `htmlHandler/convertHTMLToPlaintext.py`).

## Typical run order

1. **Create expected folders** (from `EnvironmentInitialization/`):

   ```bash
   cd EnvironmentInitialization
   python runner.py
   ```

2. **Place HTML** under `BASE_DIR/email_contents/html/` (or pass `--file`).

3. **Extract data to JSON** (from `grabbingImportantEmailContent/`):

   ```bash
   cd grabbingImportantEmailContent
   python grabbingImportantEmailContent.py --file yourfile.html --subject "..." --sender-name "..." --email "..."
   ```

   Omit `--file` to process all HTML files in `email_contents/html/`. Output is appended to `email_contents/json/results.json` under `BASE_DIR`.

4. **Sort JSON by order** (optional, from `sortJSONByOrderNumber/`):

   ```bash
   cd sortJSONByOrderNumber
   python sortJSONByOrderNumber.py
   ```

5. **Build Excel** (from `createExcelDocument/`):

   ```bash
   cd createExcelDocument
   python createExcelDocument.py
   ```

   Reads `results.json`, writes `email_contents/orders.xlsx`, and resets duplicate flags in the JSON.

Logs and admin traces may go to paths under `BASE_DIR` (e.g. `programFileOutput.txt`, `adminLog/`).

## Project layout

- `EnvironmentInitialization/` — folder checks for Power Automate / first run  
- `grabbingImportantEmailContent/` — HTML → JSON (OpenAI)  
- `sortJSONByOrderNumber/` — sort `results.json`  
- `createExcelDocument/` — JSON → Excel  
- `htmlHandler/` — HTML cleanup and helpers used by the extractor  
