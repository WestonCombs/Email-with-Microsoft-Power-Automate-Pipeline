"""Second-pass OpenAI check: is this invoice email a gift card purchase vs items only.

Runs only when the first extraction pass returns a confident Invoice category.
Import the main module at call time so this works when grabbingImportantEmailContent.py
is run as a script (__main__) or as a package submodule.
"""

from __future__ import annotations

import importlib
import sys
from typing import Any, Literal

# True = gift card purchase receipt; False = items purchase; UNKNOWN = cannot decide
UNKNOWN: Literal[2] = 2


def _gie() -> Any:
    """Return the grabbingImportantEmailContent main module (avoids circular imports)."""
    name = "grabbingImportantEmailContent.grabbingImportantEmailContent"
    if name in sys.modules:
        return sys.modules[name]
    main = sys.modules.get("__main__")
    if main is not None:
        path = getattr(main, "__file__", "") or ""
        if path.endswith("grabbingImportantEmailContent.py"):
            return main
    return importlib.import_module(name)


def should_run_is_gift_card(extracted: dict) -> bool:
    """True when the first pass is a confident Invoice (only then we can check gift card vs items)."""
    if extracted.get("email_category", "Unknown") != "Invoice":
        return False
    try:
        conf = float(extracted.get("email_category_confidence", 0))
    except (TypeError, ValueError):
        return False
    return conf >= _gie().CATEGORY_CONFIDENCE_THRESHOLD


def is_gift_card(
    text_only: str,
    subject: str | None = None,
) -> bool | Literal[2]:
    """Classify invoice email: gift card purchase vs items.

    Returns:
        True — gift card purchase receipt.
        False — invoice for other items (including when a gift card is only used as payment).
        UNKNOWN (2) — model cannot decide confidently.
    """
    gie = _gie()
    text_only = gie._sanitize_for_api(text_only)
    if subject:
        subject = gie._sanitize_for_api(subject)
    subject_section = f"\nEMAIL SUBJECT: {subject}" if subject else ""

    prompt = f"""You are a specialist classifier. Your ONLY task is to decide:

- Set is_gift_card to TRUE if the email is primarily about buying, funding, loading, or receiving a gift card as the thing purchased or delivered (e.g. e-gift card code delivery, gift card reload), not about paying for separate merchandise with a gift card.

- Set is_gift_card to FALSE if the email is an order confirmation, receipt, or invoice for goods or services where a gift card appears ONLY as a payment method, payment option, applied balance, store credit, or "paid with gift card" toward other items.

- Set is_unknown to TRUE only if the email truly does not allow a confident choice between the two above; then is_gift_card should be false.

Use ONLY the subject and body below. Return structured JSON only.
{subject_section}

EMAIL TEXT:
{text_only}""".strip()

    api_kwargs = dict(
        model=gie.MODEL,
        messages=[
            {
                "role": "developer",
                "content": (
                    "Return whether an invoice email is a gift card purchase (is_gift_card) "
                    "or items (not), or unknown. JSON only."
                ),
            },
            {"role": "user", "content": prompt},
        ],
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "is_gift_card_result",
                "schema": {
                    "type": "object",
                    "properties": {
                        "is_gift_card": {
                            "type": "boolean",
                            "description": "True if this receipt is for purchasing/receiving a gift card as the product.",
                        },
                        "is_unknown": {
                            "type": "boolean",
                            "description": "True if the model cannot confidently choose gift card vs items.",
                        },
                        "confidence": {
                            "type": "number",
                            "description": "Confidence 0-100 in the chosen outcome.",
                        },
                    },
                    "required": ["is_gift_card", "is_unknown", "confidence"],
                    "additionalProperties": False,
                },
            },
        },
        temperature=0,
    )

    data = gie._chat_completion_json_parsed(api_kwargs)

    if data.get("is_unknown"):
        return UNKNOWN
    return bool(data.get("is_gift_card"))
