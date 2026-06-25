from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from createExcelDocument.copy_email_path_to_clipboard import (  # noqa: E402
    explorer_select_command,
    file_uri_to_windows_path,
)


class FileLocationHelperTests(unittest.TestCase):
    def test_explorer_select_command_quotes_selected_file_path(self) -> None:
        command = explorer_select_command(
            r"C:\Windows\explorer.exe",
            r"C:\Users\Weston\Documents\Projects\Email Sorter\DOC Store 2026-05-05.pdf",
        )
        self.assertEqual(
            command,
            r'"C:\Windows\explorer.exe" /select,"C:\Users\Weston\Documents\Projects\Email Sorter\DOC Store 2026-05-05.pdf"',
        )

    def test_file_uri_to_windows_path_handles_spaces(self) -> None:
        self.assertEqual(
            file_uri_to_windows_path(
                "file:///C:/Users/Weston/Documents/Projects/Email%20Sorter/DOC%20Store.pdf"
            ),
            r"C:\Users\Weston\Documents\Projects\Email Sorter\DOC Store.pdf",
        )


if __name__ == "__main__":
    unittest.main()
