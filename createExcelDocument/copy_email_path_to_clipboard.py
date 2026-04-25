"""
Decode a file URI or Windows path, then open File Explorer with that file selected.

Usage:
  python copy_email_path_to_clipboard.py "file:///C:/path/to/file.pdf"

Used from Excel VBA when the user clicks an "Open File Location" hyperlink.
"""

from __future__ import annotations

import os
import subprocess
import sys
import urllib.parse


def file_uri_to_windows_path(uri: str) -> str:
    raw = uri.strip()
    if not raw.lower().startswith("file:"):
        return os.path.normpath(raw)

    parsed = urllib.parse.urlparse(raw)
    path = urllib.parse.unquote(parsed.path or "")
    if not path:
        raise ValueError("Empty path in file URI")

    # file:///C:/... -> urlparse path is often "/C:/..."
    if len(path) >= 3 and path[0] == "/" and path[2] == ":" and path[1].isalpha():
        path = path[1:]

    # UNC: file://server/share/... -> netloc + path
    if parsed.netloc and not (len(path) >= 2 and path[1] == ":"):
        path = "//" + parsed.netloc + path

    return os.path.normpath(path)


def main() -> int:
    if sys.platform != "win32":
        print("Open file location is only supported on Windows.", file=sys.stderr)
        return 1
    if len(sys.argv) < 2:
        print(
            "Usage: python copy_email_path_to_clipboard.py <file_uri_or_path>",
            file=sys.stderr,
        )
        return 1

    try:
        path = file_uri_to_windows_path(sys.argv[1])
    except Exception as e:
        print(str(e), file=sys.stderr)
        return 1

    if not os.path.isfile(path):
        print(f"File not found (cannot select in Explorer): {path}", file=sys.stderr)
        return 1

    windir = os.environ.get("WINDIR", r"C:\Windows")
    explorer = os.path.join(windir, "explorer.exe")
    subprocess.run([explorer, f"/select,{path}"], check=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
