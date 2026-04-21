"""Passive PDF capture via mitmproxy (no browser automation)."""

from __future__ import annotations

from importlib import import_module
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from mitm_pdf_interceptor.mitm_pdf_addon import PdfInterceptor

__all__ = ["PdfInterceptor"]


def __getattr__(name: str):
    if name == "PdfInterceptor":
        return import_module("mitm_pdf_interceptor.mitm_pdf_addon").PdfInterceptor
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
