# Email Sorter (Python)

Fetches shopping emails via **Microsoft Graph** (OAuth 2.0), extracts structured order data with the OpenAI API, stores results as JSON, sorts them, and builds an Excel workbook. Everything runs from a single entry point — `mainRunner.py`.

## Requirements

- **Python 3** (3.10+ recommended)
- Install dependencies:

```bash
pip install -r requirements.txt
```

(or: `pip install beautifulsoup4 python-dotenv openpyxl openai msal`)

## Azure app registration (one-time)

1. In [Azure Portal](https://portal.azure.com/) go to **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name the app, choose **Accounts in any organizational directory and personal Microsoft accounts** (or the option that matches who will sign in), then register.
3. Under **Authentication** → **Platform configurations** → **Add a platform** → **Mobile and desktop applications**. Add the redirect URI **`http://localhost`** (MSAL uses a localhost redirect for interactive login). Enable **Allow public client flows** (under **Advanced settings**) if you use the device-code flow without a client secret.
4. Under **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated** → add **`Mail.Read`**. Use **Grant admin consent** if your tenant requires it.
5. Copy the **Application (client) ID** into `.env` as `AZURE_CLIENT_ID`.

The first run opens a browser (or prints a device-code link) to sign in; tokens are cached in `python_files/.graph_token_cache.bin` (gitignored) so later runs usually stay silent until refresh.

### Long-term support (what the Azure portal warnings mean)

- **“Upgrade to MSAL and Microsoft Graph”** — This project already uses **MSAL** and **Microsoft Graph** for mail. You do not need ADAL or the legacy Azure AD Graph API.
- **“Applications … not contained within any directory” / directory-less apps deprecated** — That refers to **where** the app registration lives in Azure, not to this repo’s code. Personal sign-ins sometimes show app registrations that are not tied to a proper **Microsoft Entra ID (tenant) directory**. Microsoft wants new apps registered **inside a directory** (a tenant).

**What you should do for a stable setup**

1. Ensure you are working **inside a directory (tenant)** in the Azure Portal: use the **directory + subscription** control (top right) and pick a tenant, or create one.
2. If you have no tenant yet, pick one path:
   - **[Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)** — free sandbox tenant (good for development and learning), or  
   - **[Create an Azure free account](https://azure.microsoft.com/free/)** — creates a subscription and an **Entra ID tenant** you can use for app registrations, or  
   - **Work or school account** — register the app in your organization’s tenant (with admin consent for `Mail.Read` if required).
3. With that tenant active, go to **Microsoft Entra ID** → **App registrations** → **New registration** and complete the steps above. Use **`AZURE_TENANT_ID`** set to that tenant’s ID (Directory ID) when you want to restrict sign-in to that tenant; keep **`common`** only if you intentionally support multiple account types and your app registration allows it.

The mailbox you read with Graph can still be a **personal Microsoft account** (e.g. Outlook.com) in many setups, as long as the **app registration** lives in a supported directory and the supported account types on the app include personal accounts.

## Configuration

Copy **`.env.example`** to **`.env`** (same folder as this README) and fill in the values:

| Variable | Purpose |
|----------|---------|
| `BASE_DIR` | Absolute path to your project data root (folders like `email_contents/` live under here). |
| `OPENAI_API_KEY` | Your OpenAI API key (used by the email extraction script). |
| `AZURE_CLIENT_ID` | Azure app registration client ID (see above). |
| `AZURE_TENANT_ID` | Tenant ID, or `common` for mixed personal/work sign-in (default). |
| `GRAPH_AUTH_FLOW` | `interactive` (default) or `device_code`. |
| `GRAPH_MAIL_FOLDER` | Mailbox folder to read: `INBOX`, or the **exact display name** as in Outlook (e.g. `Shopping`). Well-known names like Inbox/Sent/Drafts are matched case-insensitively. |
| `GRAPH_TOKEN_CACHE_PATH` | Optional full path for the token cache file. |

If `GRAPH_MAIL_FOLDER` is unset, `IMAP_MAIL_FOLDER` is still read so older `.env` files keep working for the folder name only.

Optional: `DEMO_MODE=1` turns on demo behavior in the extractor **and** forces a **full Microsoft Graph browser sign-in every run** (no silent token): MSAL uses `prompt=login` so you can switch accounts and step through MFA. Set `DEMO_MODE=0` for normal runs that reuse the cached token. Also optional: `openai_max_chars_per_prompting` to cap HTML size sent to the model (see `htmlHandler/convertHTMLToPlaintext.py`).

## Running

```bash
python mainRunner.py
```

This single command runs the full pipeline:

1. **Environment initialization** — creates required folders under `BASE_DIR` (`email_contents/attachments`, `email_contents/html`, `email_contents/json`, etc.).
2. **Email fetching** — signs in with Microsoft Graph (if needed), lists all messages in the configured folder, downloads bodies and file attachments to `BASE_DIR/email_contents/attachments` where applicable.
3. **Extraction** — for each email, writes the HTML body to disk and runs the OpenAI extraction pipeline to produce structured JSON in `email_contents/json/results.json`.
4. **Sort** — sorts `results.json` by order number and purchase date.
5. **Excel export** — builds `email_contents/orders.xlsx` from the JSON and resets duplicate flags.

An OpenAI usage log (`usageN.txt`) is created per run in `BASE_DIR/email_contents/openai usage/`. A summary with total tokens, cost, and average time per email is printed at the end.

Logs and admin traces go to paths under `BASE_DIR` (e.g. `programFileOutput.txt`, `adminLog/`).

## Project layout

- `mainRunner.py` — main entry point; orchestrates the full pipeline
- `emailFetching/` — Microsoft Graph mail fetch (`ms_graph_fetcher.py`, shared models in `emailFetcher.py`)
- `EnvironmentInitialization/` — folder/file verification for first run
- `grabbingImportantEmailContent/` — HTML → JSON (OpenAI)
- `sortJSONByOrderNumber/` — sort `results.json`
- `createExcelDocument/` — JSON → Excel
- `htmlHandler/` — HTML cleanup and helpers used by the extractor
