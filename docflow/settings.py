"""Environment-driven runtime settings shared by backend modules."""

from __future__ import annotations

import os


VALID_PDF_MODES = {"accurate", "fast"}


def env_flag(name: str, default: bool = False) -> bool:
    """Read a boolean-like environment variable."""

    value = str(os.getenv(name, "")).strip().lower()
    if not value:
        return default
    return value in {"1", "true", "yes", "on"}


def env_int(name: str, default: int) -> int:
    """Read a positive integer environment variable."""

    try:
        value = int(str(os.getenv(name, "")).strip())
        return value if value > 0 else default
    except Exception:
        return default


def get_default_pdf_mode() -> str:
    """Return the configured PDF parsing mode with a safe fallback."""

    value = str(os.getenv("DOCFLOW_DEFAULT_PDF_MODE", "fast")).strip().lower()
    return value if value in VALID_PDF_MODES else "fast"


DEFAULT_PDF_MODE = get_default_pdf_mode()


def normalize_pdf_mode(value: str, fallback: str = DEFAULT_PDF_MODE) -> str:
    """Normalize a user supplied PDF mode to a known value."""

    normalized = str(value or "").strip().lower()
    return normalized if normalized in VALID_PDF_MODES else fallback

