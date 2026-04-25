"""Apply Email Sorter settings with a small optional ``.env`` fallback.

Runtime order is intentionally simple:
1. built-in defaults for non-secret runtime switches;
2. ``python_files/.env`` or parent-process environment values for Settings fields;
3. nonempty values from ``email_sorter_settings.json`` written by the Settings dialog.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

from shared.project_paths import ensure_base_dir_in_environ, python_files_dir

_SETTINGS_FILENAME = "email_sorter_settings.json"
_ENV_FILENAME = ".env"

# Keys written by Settings.
STORED_SETTING_KEYS = frozenset(
    {
        "GRAPH_MAIL_FOLDER",
        "AZURE_CLIENT_ID",
        "AZURE_TENANT_ID",
        "OPENAI_API_KEY",
        "SEVENTEEN_TRACK_API_KEY",
        "DEBUG_MODE",
        "LOGIN_NEW_ACCOUNT_NEXT_RUN",
    }
)

# Only these Settings values may fall back to .env when the JSON value is empty.
ENV_FALLBACK_SETTING_KEYS = frozenset(
    {
        "GRAPH_MAIL_FOLDER",
        "AZURE_CLIENT_ID",
        "AZURE_TENANT_ID",
        "OPENAI_API_KEY",
        "SEVENTEEN_TRACK_API_KEY",
        "DEBUG_MODE",
    }
)

_PROCESS_ENV_FALLBACKS = {
    key: (os.environ.get(key) or "").strip() for key in ENV_FALLBACK_SETTING_KEYS
}

_DEFAULT_ENV: dict[str, str] = {
    "AZURE_TENANT_ID": "common",
    "GRAPH_AUTH_FLOW": "interactive",
    "DEMO_MODE": "0",
    "EMAIL_LINK_DEBUG": "0",
}


def settings_json_path() -> Path:
    return python_files_dir() / _SETTINGS_FILENAME


def env_path() -> Path:
    return python_files_dir() / _ENV_FILENAME


def _clean(raw: object) -> str:
    return str(raw).strip()


def _unquote_env_value(raw: str) -> str:
    value = raw.strip()
    if not value:
        return ""
    if value[0] in ("'", '"'):
        quote = value[0]
        end = value.rfind(quote)
        if end > 0:
            return value[1:end].strip()
    if " #" in value:
        value = value.split(" #", 1)[0].rstrip()
    return value.strip()


def _read_env_file_settings() -> dict[str, str]:
    p = env_path()
    if not p.is_file():
        return {}
    try:
        lines = p.read_text(encoding="utf-8-sig").splitlines()
    except (OSError, UnicodeError):
        return {}

    values: dict[str, str] = {}
    for raw_line in lines:
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.lower().startswith("export "):
            line = line[7:].lstrip()
        if "=" not in line:
            continue
        key, raw_value = line.split("=", 1)
        key = key.strip()
        if key not in ENV_FALLBACK_SETTING_KEYS:
            continue
        value = _unquote_env_value(raw_value)
        if value:
            values[key] = value
    return values


def read_env_fallback_settings() -> dict[str, str]:
    """Return allowed Settings fallback values from process env, then ``.env``."""
    values = {k: v for k, v in _PROCESS_ENV_FALLBACKS.items() if v}
    values.update(_read_env_file_settings())
    return values


def read_settings_json() -> dict[str, str]:
    p = settings_json_path()
    if not p.is_file():
        return {}
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return {}
    if not isinstance(raw, dict):
        return {}
    out: dict[str, str] = {}
    for k, v in raw.items():
        if not isinstance(k, str) or k not in STORED_SETTING_KEYS:
            continue
        if v is None:
            continue
        s = _clean(v)
        if s:
            out[k] = s
    return out


def read_settings_for_write_merge() -> dict[str, str]:
    """Return all Settings keys from disk, preserving blanks and avoiding env fallback."""
    p = settings_json_path()
    disk: dict[str, object] = {}
    if p.is_file():
        try:
            raw = json.loads(p.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                disk = {str(k): v for k, v in raw.items()}
        except (OSError, UnicodeError, json.JSONDecodeError):
            pass
    out: dict[str, str] = {k: "" for k in STORED_SETTING_KEYS}
    for k in STORED_SETTING_KEYS:
        if k in disk and disk[k] is not None:
            out[k] = _clean(disk[k])
    return out


def write_settings_json(updates: dict[str, str]) -> None:
    """Persist *updates* (must include every key in ``STORED_SETTING_KEYS``)."""
    missing = STORED_SETTING_KEYS - set(updates)
    if missing:
        raise ValueError(f"write_settings_json: missing keys {sorted(missing)}")
    data = {k: str(updates[k]).strip() for k in STORED_SETTING_KEYS}
    p = settings_json_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")
    tmp.replace(p)


def apply_runtime_settings_from_json() -> None:
    """Apply defaults, allowed ``.env`` fallbacks, then nonempty saved Settings."""
    ensure_base_dir_in_environ()
    for k, v in _DEFAULT_ENV.items():
        os.environ.setdefault(k, v)

    env_data = read_env_fallback_settings()
    json_data = read_settings_json()

    for key in ENV_FALLBACK_SETTING_KEYS:
        value = (json_data.get(key) or env_data.get(key) or "").strip()
        if value:
            os.environ[key] = value
        elif key in _DEFAULT_ENV:
            os.environ[key] = _DEFAULT_ENV[key]
        else:
            os.environ.pop(key, None)

    login_next = (json_data.get("LOGIN_NEW_ACCOUNT_NEXT_RUN") or "").strip()
    if login_next:
        os.environ["LOGIN_NEW_ACCOUNT_NEXT_RUN"] = login_next
    else:
        os.environ.pop("LOGIN_NEW_ACCOUNT_NEXT_RUN", None)
