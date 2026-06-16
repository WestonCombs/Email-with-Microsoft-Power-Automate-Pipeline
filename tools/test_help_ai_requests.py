from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from shared.help_ai import (  # noqa: E402
    active_requests,
    load_requests,
    save_record_updates,
    sync_help_ai_requests,
)
from shared.excel_user_edits import record_identity  # noqa: E402


class HelpAIRequestTests(unittest.TestCase):
    def test_invoice_total_mismatch_creates_request_and_resolve_sticks(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            project_root = Path(td)
            json_dir = project_root / "email_contents" / "json"
            json_dir.mkdir(parents=True)
            record = {
                "source_file_link": "file:///C:/tmp/order.pdf",
                "email_category": "Invoice",
                "company": "Example",
                "order_number": "1234",
                "subtotal_amount": 20.0,
                "tax_paid": 2.0,
                "gift_card_amount": 0.0,
                "total_amount_paid": 30.0,
            }
            (json_dir / "results.json").write_text(
                json.dumps([record], indent=2),
                encoding="utf-8",
            )

            requests = sync_help_ai_requests(project_root)
            active = active_requests(requests)
            self.assertEqual(len(active), 1)
            self.assertEqual(active[0]["type"], "invoice_totals_ambiguous")

            save_record_updates(
                project_root,
                record_id=record_identity(record),
                updates={"total_amount_paid": "22.00"},
                request_id=active[0]["id"],
                status="resolved",
            )
            requests_after = load_requests(project_root)
            self.assertEqual(active_requests(requests_after), [])

            rows = json.loads((json_dir / "results.json").read_text(encoding="utf-8"))
            self.assertEqual(rows[0]["total_amount_paid"], 22.0)


if __name__ == "__main__":
    unittest.main()
