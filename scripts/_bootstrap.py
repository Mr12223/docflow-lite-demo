"""Shared bootstrap helpers for repository scripts.

Directly running a file under `scripts/` only adds that directory to
`sys.path`. This module ensures the repository root is available before
importing application packages such as `docflow`.
"""

from __future__ import annotations

import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent


def ensure_project_root_on_path() -> Path:
    """Add the repository root to `sys.path` and return it."""

    project_root_str = str(PROJECT_ROOT)
    if project_root_str not in sys.path:
        sys.path.insert(0, project_root_str)
    return PROJECT_ROOT
