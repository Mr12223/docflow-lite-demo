"""Project path definitions used across the DocFlow codebase.

This module keeps filesystem locations in one place so routes, scripts, and
storage code do not each redefine their own root-relative paths.
"""

from __future__ import annotations

from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
FRONTEND_DIR = PROJECT_ROOT / "frontend"
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
SAMPLE_DATA_DIR = PROJECT_ROOT / "sample_data"
REPORTS_DIR = PROJECT_ROOT / "reports"
UPLOADS_DIR = PROJECT_ROOT / "uploads_temp"
IMAGE_OCR_CACHE_DIR = UPLOADS_DIR / "ocr_cache"
DATA_DIR = PROJECT_ROOT / "data"
INVOICE_DB_PATH = DATA_DIR / "invoices.db"


def ensure_runtime_directories() -> None:
    """Create runtime directories that are safe to initialize eagerly."""

    for path in (REPORTS_DIR, UPLOADS_DIR, IMAGE_OCR_CACHE_DIR, DATA_DIR):
        path.mkdir(parents=True, exist_ok=True)

