"""Sandbox remote-image preservation for saved email HTML files.

This script does not modify the live workflow or the source HTML. It writes
test artifacts under the chosen output directory.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from dataclasses import asdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from emailFetching.remote_image_preserver import preserve_remote_images


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("html", type=Path, help="HTML file to test")
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=ROOT / "remote_image_preservation_tests",
        help="Directory for test artifacts",
    )
    parser.add_argument(
        "--pdf",
        action="store_true",
        help="Also render the preserved HTML to PDF using Chrome or Edge",
    )
    parser.add_argument(
        "--referer",
        default=None,
        help="Optional Referer header to use while fetching images",
    )
    args = parser.parse_args()

    html_path = args.html.resolve()
    if not html_path.is_file():
        raise SystemExit(f"HTML file not found: {html_path}")

    args.out_dir.mkdir(parents=True, exist_ok=True)
    source_html = html_path.read_text(encoding="utf-8", errors="replace")
    result = preserve_remote_images(source_html, referer=args.referer)

    stem = _safe_stem(html_path.stem)
    preserved_html = args.out_dir / f"{stem}.remote-images-preserved.html"
    report_json = args.out_dir / f"{stem}.remote-images-report.json"
    preserved_html.write_text(result.html, encoding="utf-8")
    report_json.write_text(
        json.dumps(
            {
                "source_html": str(html_path),
                "preserved_html": str(preserved_html),
                "summary": {
                    "img_tags": result.img_tags,
                    "remote_src": result.remote_src,
                    "replaced_src": result.replaced_src,
                    "failed_src": result.failed_src,
                    "skipped_src": result.skipped_src,
                },
                "attempts": [asdict(attempt) for attempt in result.attempts],
            },
            indent=2,
        ),
        encoding="utf-8",
    )

    print(result.to_log_line())
    print(f"preserved_html={preserved_html}")
    print(f"report_json={report_json}")

    if args.pdf:
        pdf_path = args.out_dir / f"{stem}.remote-images-preserved.pdf"
        browser = _find_browser()
        if not browser:
            print("pdf=skipped (Chrome/Edge not found)")
        else:
            _render_pdf(browser, preserved_html, pdf_path)
            print(f"pdf={pdf_path}")

    return 0


def _find_browser() -> Path | None:
    candidates = [
        Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
        Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
        Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
        Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
    ]
    return next((path for path in candidates if path.is_file()), None)


def _render_pdf(browser: Path, html_path: Path, pdf_path: Path) -> None:
    proc = subprocess.run(
        [
            str(browser),
            "--headless",
            "--disable-gpu",
            "--allow-file-access-from-files",
            "--blink-settings=imagesEnabled=true",
            "--run-all-compositor-stages-before-draw",
            "--virtual-time-budget=30000",
            "--no-pdf-header-footer",
            f"--print-to-pdf={pdf_path}",
            html_path.resolve().as_uri(),
        ],
        capture_output=True,
        timeout=90,
        text=True,
    )
    if not pdf_path.exists() or pdf_path.stat().st_size <= 0:
        stderr = (proc.stderr or "").strip()
        raise RuntimeError(f"PDF render failed: {stderr or proc.returncode}")


def _safe_stem(value: str) -> str:
    safe = "".join(ch if ch.isalnum() or ch in "._- " else "_" for ch in value)
    safe = " ".join(safe.split()).strip(" ._")
    return safe[:120] or "email"


if __name__ == "__main__":
    raise SystemExit(main())
