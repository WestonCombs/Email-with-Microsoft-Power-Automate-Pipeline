"""Entry point for Power Automate: load BASE_DIR from .env, then run the verification checklist below."""

from __future__ import annotations

import os
import sys
from pathlib import Path

_PYTHON_FILES_DIR = Path(__file__).resolve().parent.parent
_ENV_PATH = _PYTHON_FILES_DIR / ".env"
BASE_DIR_ENV = "BASE_DIR"


def _run_verifications(root: Path) -> None:
    """Everything here is relative to BASE_DIR — edit this list to match your flow."""
    from FolderVerification import folder_verification
    from fileVerification import file_verification

    # Folders (repeat calls are safe; same path = idempotent)
    folder_verification(root, "email_contents/attachments", clear_if_exists=True)
    folder_verification(root, "email_contents/html")
    folder_verification(root, "email_contents/json")
    folder_verification(root, "email_contents/openai usage")
    folder_verification(root, "adminLog")

    # Files: parent dirs are created automatically; use overwrite=True only if you must truncate
    # file_verification(root, "email_contents/html/incoming.html")
    # file_verification(root, "email_contents/json/results.json")
    # file_verification(root, "programFileOutput.txt")


def main() -> int:
    from dotenv import load_dotenv

    load_dotenv(_ENV_PATH)

    base_raw = os.getenv(BASE_DIR_ENV)
    if not base_raw:
        print(
            f"ERROR: {BASE_DIR_ENV} is not set. Add it to {_ENV_PATH}",
            file=sys.stderr,
        )
        return 1

    root = Path(base_raw).expanduser().resolve()

    try:
        _run_verifications(root)
    except (OSError, ValueError, NotADirectoryError, IsADirectoryError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1

    print(f"Environment checks OK under BASE_DIR={root}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
