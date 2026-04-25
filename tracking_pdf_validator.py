"""LLM validation for captured package tracking PDFs."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any


VALIDATION_PROMPT = """You are analyzing text extracted from a captured package tracking PDF.

Your job is to determine whether the PDF appears to show the latest tracking information that was available on the page at capture time.

Return ONLY a JSON object with this format:

{
  "latest_tracking_info_visible": true/false,
  "confidence": 0-100,
  "status_found": "Delivered/Shipped/In Transit/Out for Delivery/Label Created/Exception/Delayed/Unknown",
  "latest_update_found": "short summary of the most recent visible tracking update, or null",
  "reason": "short explanation"
}

Criteria for true:
- Concrete shipment status present
- At least one tracking update visible
- Page is clearly fully loaded

Criteria for false:
- Loading screen
- Generic tracking landing page
- Login/CAPTCHA/error
- No actual tracking updates"""

_STATUSES = {
    "Delivered",
    "Shipped",
    "In Transit",
    "Out for Delivery",
    "Label Created",
    "Exception",
    "Delayed",
    "Unknown",
}


def _fallback_validation(reason: str) -> dict[str, Any]:
    return {
        "latest_tracking_info_visible": False,
        "confidence": 0,
        "status_found": "Unknown",
        "latest_update_found": None,
        "reason": reason,
    }


def _extract_pdf_text(pdf_path: str) -> str:
    """Extract all text from a PDF using PyPDF2."""
    try:
        from PyPDF2 import PdfReader
    except Exception as exc:  # pragma: no cover - depends on local install
        raise RuntimeError("PyPDF2 is required for tracking PDF validation") from exc

    reader = PdfReader(str(Path(pdf_path).expanduser().resolve()))
    chunks: list[str] = []
    for page in reader.pages:
        try:
            chunks.append(page.extract_text() or "")
        except Exception:
            continue
    return "\n".join(chunks).strip()


def _normalize_validation(raw: object) -> dict[str, Any]:
    """Coerce LLM output into the audit schema."""
    data = raw if isinstance(raw, dict) else {}
    visible = bool(data.get("latest_tracking_info_visible"))
    try:
        confidence = int(float(data.get("confidence", 0)))
    except (TypeError, ValueError):
        confidence = 0
    confidence = max(0, min(100, confidence))

    status = str(data.get("status_found") or "Unknown").strip()
    if status not in _STATUSES:
        status = "Unknown"

    latest_update = data.get("latest_update_found")
    if latest_update is not None:
        latest_update = str(latest_update).strip() or None

    return {
        "latest_tracking_info_visible": visible,
        "confidence": confidence,
        "status_found": status,
        "latest_update_found": latest_update,
        "reason": str(data.get("reason") or "").strip(),
    }


def validate_pdf_with_llm(pdf_path: str) -> dict:
    """Validate that a captured tracking PDF shows loaded, concrete tracking updates."""
    text = _extract_pdf_text(pdf_path)
    if not text:
        return _fallback_validation("No text could be extracted from the PDF.")

    try:
        from shared.settings_store import apply_runtime_settings_from_json

        apply_runtime_settings_from_json()
    except Exception:
        pass

    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        return _fallback_validation("OPENAI_API_KEY is not configured.")

    try:
        from openai import OpenAI
    except Exception as exc:  # pragma: no cover - depends on local install
        raise RuntimeError("openai is required for tracking PDF validation") from exc

    client = OpenAI(api_key=api_key)
    response = client.chat.completions.create(
        model=os.getenv("OPENAI_TRACKING_PDF_MODEL", "gpt-4o-mini"),
        messages=[
            {"role": "developer", "content": VALIDATION_PROMPT},
            {"role": "user", "content": f"PDF TEXT:\n{text[:60000]}"},
        ],
        response_format={"type": "json_object"},
        temperature=0,
    )
    content = response.choices[0].message.content or "{}"
    try:
        parsed = json.loads(content)
    except json.JSONDecodeError:
        return _fallback_validation("The LLM did not return valid JSON.")
    return _normalize_validation(parsed)
