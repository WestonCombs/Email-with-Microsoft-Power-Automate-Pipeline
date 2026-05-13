"""Automated saved-email HTML to PDF test using capture Chrome plus CDP."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from pdfCaptureFromChrome.email_html_pdf import convert_email_html_to_pdf  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("html", type=Path)
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=ROOT / "visible_email_pdf_tests",
    )
    parser.add_argument("--wait-seconds", type=float, default=35.0)
    args = parser.parse_args()

    html_path = args.html.resolve()
    if not html_path.is_file():
        raise SystemExit(f"HTML file not found: {html_path}")
    args.out_dir.mkdir(parents=True, exist_ok=True)

    stem = _safe_stem(html_path.stem)
    pdf_path = args.out_dir / f"{stem}.cdp-capture.pdf"
    preview_path = pdf_path.with_suffix(".preview.png")
    report_path = args.out_dir / f"{stem}.cdp-capture-report.json"

    result = convert_email_html_to_pdf(
        html_path,
        pdf_path,
        wait_timeout=args.wait_seconds,
    )
    preview = _render_pdf_preview(pdf_path, preview_path)

    report_path.write_text(
        json.dumps(
            {
                "source_html": str(html_path),
                "pdf": str(pdf_path),
                "preview": str(preview) if preview else None,
                "ready_state": result.ready_state,
                "image_count": result.image_count,
                "images_complete": result.images_complete,
                "images_loaded": result.images_loaded,
                "images_failed": result.images_failed,
                "elapsed_seconds": result.elapsed_seconds,
            },
            indent=2,
        ),
        encoding="utf-8",
    )

    print(f"images_loaded={result.images_loaded}/{result.image_count}")
    print(f"images_complete={result.images_complete}/{result.image_count}")
    print(f"images_failed={result.images_failed}")
    print(f"pdf={pdf_path}")
    print(f"preview={preview}")
    print(f"report={report_path}")
    return 0


def _render_pdf_preview(pdf_path: Path, preview_path: Path) -> Path | None:
    try:
        import fitz
    except Exception:
        return None
    doc = fitz.open(pdf_path)
    try:
        page_index = 0
        for i, page in enumerate(doc):
            text = page.get_text("text") or ""
            if "In This Package" in text or "Items in Your Order" in text:
                page_index = i
                break
        page = doc[page_index]
        rect = page.rect
        crop = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * 0.72)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), clip=crop, alpha=False)
        out = preview_path.with_name(f"{preview_path.stem}-page{page_index + 1}{preview_path.suffix}")
        pix.save(out)
        return out
    finally:
        doc.close()


def _safe_stem(value: str) -> str:
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value or "")
    safe = re.sub(r"\s+", " ", safe).strip(" ._")
    return safe[:120] or "email"


if __name__ == "__main__":
    raise SystemExit(main())
