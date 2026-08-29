"""Focused tests for the simulated email header added during PDF printing."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from emailFetching.emailFetcher import (  # noqa: E402
    EmailMessage,
    _build_outlook_header_fragment,
    prepend_outlook_style_header,
)


class EmailPdfHeaderTests(unittest.TestCase):
    def test_all_expected_fields_are_always_rendered(self) -> None:
        fragment = _build_outlook_header_fragment(
            EmailMessage(from_raw="", subject="", body_html="")
        )

        for label in ("From:", "Sent:", "Received:", "To:", "Subject:"):
            self.assertIn(label, fragment)
        self.assertEqual(fragment.count("(not available)"), 5)

    def test_values_are_escaped_and_rows_do_not_use_table_elements(self) -> None:
        fragment = _build_outlook_header_fragment(
            EmailMessage(
                from_raw='Store <orders@example.com>',
                subject="Order <123>",
                body_html="",
                to_line="Client <client@example.com>",
                sent_line="Friday, August 28, 2026 1:00 PM",
                received_line="Friday, August 28, 2026 1:01 PM",
                header_title="Client Name",
            )
        )

        self.assertNotIn("<table", fragment.lower())
        self.assertNotIn("<tr", fragment.lower())
        self.assertNotIn("<td", fragment.lower())
        self.assertIn("Order &lt;123&gt;", fragment)
        self.assertIn("Store &lt;orders@example.com&gt;", fragment)
        self.assertIn("Received:", fragment)

    def test_hostile_email_css_cannot_target_header_by_table_tag(self) -> None:
        source = """<!doctype html><html><head><style>
            table, tr, td { display:none !important; color:white !important; font-size:0 !important; }
            div, span { visibility:hidden !important; }
        </style></head><body><p>Email body</p></body></html>"""
        rendered = prepend_outlook_style_header(
            source,
            EmailMessage(
                from_raw="Sender <sender@example.com>",
                subject="Visible subject",
                body_html=source,
            ),
        )

        self.assertIn("email-sorter-meta-header", rendered)
        self.assertIn("display:block !important", rendered)
        self.assertIn("visibility:visible !important", rendered)
        self.assertIn("display:table-row !important", rendered)
        self.assertIn("Visible subject", rendered)


if __name__ == "__main__":
    unittest.main()
