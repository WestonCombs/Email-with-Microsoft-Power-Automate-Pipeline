"""Preserve remote email images by inlining fetchable HTTP(S) image URLs."""

from __future__ import annotations

import base64
import dataclasses
import mimetypes
import re
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Callable

from bs4 import BeautifulSoup


DEFAULT_IMAGE_TIMEOUT_S = 20.0
DEFAULT_MAX_IMAGE_BYTES = 8 * 1024 * 1024

_REMOTE_URL_RE = re.compile(r"^https?://", re.IGNORECASE)
_IMAGE_LIKE_EXTENSIONS = {
    ".apng",
    ".avif",
    ".bmp",
    ".gif",
    ".jpeg",
    ".jpg",
    ".png",
    ".svg",
    ".webp",
}
_FALLBACK_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/125.0.0.0 Safari/537.36"
)


@dataclasses.dataclass(frozen=True)
class RemoteImageAttempt:
    url: str
    status: str
    content_type: str = ""
    byte_count: int = 0
    error: str = ""


@dataclasses.dataclass(frozen=True)
class RemoteImagePreserveResult:
    html: str
    img_tags: int
    remote_src: int
    replaced_src: int
    failed_src: int
    skipped_src: int
    attempts: tuple[RemoteImageAttempt, ...]

    def to_log_line(self) -> str:
        return (
            f"img_tags={self.img_tags} remote_src={self.remote_src} "
            f"replaced_src={self.replaced_src} failed_src={self.failed_src} "
            f"skipped_src={self.skipped_src}"
        )


def preserve_remote_images(
    html: str,
    *,
    timeout_s: float = DEFAULT_IMAGE_TIMEOUT_S,
    max_image_bytes: int = DEFAULT_MAX_IMAGE_BYTES,
    user_agent: str = _FALLBACK_USER_AGENT,
    referer: str | None = None,
    fetcher: Callable[[str], tuple[str, bytes]] | None = None,
    image_cache: dict[str, str | None] | None = None,
) -> RemoteImagePreserveResult:
    """Inline fetchable remote ``<img src="https://...">`` images as data URIs.

    This intentionally leaves failed URLs in place. Email-image hosts sometimes
    block non-browser fetches; keeping the original URL makes failures visible
    and avoids replacing a potentially recoverable image with a placeholder.
    """
    if not html:
        return RemoteImagePreserveResult("", 0, 0, 0, 0, 0, ())

    soup = BeautifulSoup(html, "html.parser")
    img_tags = soup.find_all("img")
    attempts: list[RemoteImageAttempt] = []
    cache: dict[str, str | None] = image_cache if image_cache is not None else {}
    remote_src = 0
    replaced_src = 0
    failed_src = 0
    skipped_src = 0

    def _fetch(url: str) -> tuple[str, bytes]:
        if fetcher is not None:
            return fetcher(url)
        return fetch_remote_image(
            url,
            timeout_s=timeout_s,
            max_image_bytes=max_image_bytes,
            user_agent=user_agent,
            referer=referer,
        )

    for img in img_tags:
        raw_src = (img.get("src") or "").strip()
        if not raw_src or not _REMOTE_URL_RE.match(raw_src):
            skipped_src += 1
            continue

        remote_src += 1
        url = raw_src
        data_uri = cache.get(url)
        if url in cache:
            if data_uri:
                attempts.append(
                    RemoteImageAttempt(
                        url=url,
                        status="cached_ok",
                    )
                )
                img["src"] = data_uri
                replaced_src += 1
            else:
                attempts.append(
                    RemoteImageAttempt(
                        url=url,
                        status="cached_failed",
                        error="cached failure",
                    )
                )
                failed_src += 1
            continue

        try:
            content_type, payload = _fetch(url)
            data_uri = _to_data_uri(content_type, payload)
            cache[url] = data_uri
            attempts.append(
                RemoteImageAttempt(
                    url=url,
                    status="ok",
                    content_type=content_type,
                    byte_count=len(payload),
                )
            )
            img["src"] = data_uri
            replaced_src += 1
        except Exception as exc:
            cache[url] = None
            attempts.append(
                RemoteImageAttempt(
                    url=url,
                    status="failed",
                    error=_short_error(exc),
                )
            )
            failed_src += 1

    return RemoteImagePreserveResult(
        html=str(soup),
        img_tags=len(img_tags),
        remote_src=remote_src,
        replaced_src=replaced_src,
        failed_src=failed_src,
        skipped_src=skipped_src,
        attempts=tuple(attempts),
    )


def fetch_remote_image(
    url: str,
    *,
    timeout_s: float = DEFAULT_IMAGE_TIMEOUT_S,
    max_image_bytes: int = DEFAULT_MAX_IMAGE_BYTES,
    user_agent: str = _FALLBACK_USER_AGENT,
    referer: str | None = None,
) -> tuple[str, bytes]:
    headers = {
        "User-Agent": user_agent,
        "Accept": "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8",
    }
    if referer:
        headers["Referer"] = referer

    request = urllib.request.Request(url, headers=headers, method="GET")
    with urllib.request.urlopen(request, timeout=timeout_s) as response:
        content_type = (response.headers.get("Content-Type") or "").split(";", 1)[0].strip()
        payload = _read_limited(response, max_image_bytes)

    if not payload:
        raise ValueError("empty image response")
    if not _is_image_response(url, content_type):
        raise ValueError(f"response is not an image: {content_type or 'unknown content type'}")
    if not content_type:
        content_type = _guess_content_type(url)
    return content_type, payload


def _read_limited(response: object, max_bytes: int) -> bytes:
    data = response.read(max_bytes + 1)
    if len(data) > max_bytes:
        raise ValueError(f"image is larger than {max_bytes} bytes")
    return data


def _to_data_uri(content_type: str, payload: bytes) -> str:
    encoded = base64.b64encode(payload).decode("ascii")
    return f"data:{content_type};base64,{encoded}"


def _is_image_response(url: str, content_type: str) -> bool:
    if content_type.lower().startswith("image/"):
        return True
    return _guess_content_type(url).lower().startswith("image/")


def _guess_content_type(url: str) -> str:
    path = urllib.parse.urlparse(url).path
    guessed, _ = mimetypes.guess_type(path)
    if guessed:
        return guessed
    suffix = "." + path.rsplit(".", 1)[-1].lower() if "." in path else ""
    if suffix in _IMAGE_LIKE_EXTENSIONS:
        return f"image/{suffix.lstrip('.')}"
    return "application/octet-stream"


def _short_error(exc: BaseException) -> str:
    if isinstance(exc, urllib.error.HTTPError):
        return f"HTTP {exc.code}: {exc.reason}"
    if isinstance(exc, urllib.error.URLError):
        return str(exc.reason)
    return str(exc)
