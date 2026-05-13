"""Focused checks for purchase-date accuracy fallbacks.

These tests avoid live email/OpenAI calls and exercise the consolidation and
filename behavior that can otherwise turn mailbox timing into an inaccurate
purchase date.
"""

from __future__ import annotations

import os
import sys
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ.setdefault("OPENAI_API_KEY", "test-key")

from grabbingImportantEmailContent import grabbingImportantEmailContent as extractor  # noqa: E402

extractor.RL.log = lambda *args, **kwargs: None

build_convention_filename = extractor.build_convention_filename
unify_purchase_dates_by_order = extractor.unify_purchase_dates_by_order


class PurchaseDateAccuracyTests(unittest.TestCase):
    def test_missing_purchase_date_does_not_use_email_datetime(self) -> None:
        rows = [
            {
                "order_number": "A1",
                "purchase_datetime": "",
                "email_datetime": "2026-05-13T12:00:00Z",
                "company": "Example Store",
                "email_category": "Invoice",
            }
        ]

        with redirect_stdout(StringIO()):
            unify_purchase_dates_by_order(rows)

        self.assertEqual(rows[0]["purchase_datetime"], "")
        filename = build_convention_filename(rows[0])
        self.assertNotIn("2026-05-13", filename)
        self.assertEqual(filename, "DOC Example Store 0001 INVOICE.pdf")

    def test_verified_order_date_from_same_order_is_shared(self) -> None:
        rows = [
            {
                "order_number": "B2",
                "purchase_datetime": "",
                "email_datetime": "2026-05-13T12:00:00Z",
                "company": "Example Store",
                "email_category": "Shipped",
            },
            {
                "order_number": "B2",
                "purchase_datetime": "2026-04-28",
                "email_datetime": "2026-05-14T12:00:00Z",
                "company": "Example Store",
                "email_category": "Invoice",
            },
        ]

        with redirect_stdout(StringIO()):
            unify_purchase_dates_by_order(rows)

        self.assertEqual(rows[0]["purchase_datetime"], "2026-04-28")
        self.assertEqual(rows[1]["purchase_datetime"], "2026-04-28")
        self.assertEqual(
            rows[0]["purchase_datetime_source"],
            "order_consensus:source=explicit_order_date;first_allowed_in_order:index=1",
        )
        self.assertEqual(rows[0]["purchase_datetime_confidence"], "medium")
        self.assertIn("2026-04-28", build_convention_filename(rows[0]))


if __name__ == "__main__":
    unittest.main()
