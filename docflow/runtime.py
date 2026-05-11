"""Runtime bootstrap helpers shared by CLI and web entrypoints."""

from __future__ import annotations

import io
import logging
import os
import sys


THREAD_LIMIT_ENV = {
    "OMP_NUM_THREADS": "1",
    "OMP_THREAD_LIMIT": "1",
    "OPENBLAS_NUM_THREADS": "1",
    "MKL_NUM_THREADS": "1",
    "NUMEXPR_NUM_THREADS": "1",
    "VECLIB_MAXIMUM_THREADS": "1",
    "BLIS_NUM_THREADS": "1",
    "GOTO_NUM_THREADS": "1",
    "OMP_WAIT_POLICY": "PASSIVE",
}


def configure_native_thread_env() -> None:
    """Limit native numeric libraries when running in constrained environments."""

    cloud_markers = (
        "RENDER",
        "RENDER_SERVICE_ID",
        "RENDER_INSTANCE_ID",
        "DOCFLOW_CLOUD_DEPLOYMENT",
    )
    cloud_like = any(str(os.getenv(name, "")).strip() for name in cloud_markers)
    limit_requested = str(os.getenv("DOCFLOW_LIMIT_OCR_THREADS", "")).strip().lower()
    should_limit = cloud_like or limit_requested in {"1", "true", "yes", "on"}
    if not should_limit:
        return

    for env_name, env_value in THREAD_LIMIT_ENV.items():
        os.environ.setdefault(env_name, env_value)


def configure_windows_utf8_stdio() -> None:
    """Force UTF-8 stdout/stderr wrappers on Windows consoles."""

    if sys.platform != "win32":
        return

    if hasattr(sys.stdout, "buffer") and not isinstance(sys.stdout, io.TextIOWrapper):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    if hasattr(sys.stderr, "buffer") and not isinstance(sys.stderr, io.TextIOWrapper):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")


def configure_logging(logger_name: str = "DocFlow") -> logging.Logger:
    """Set up a consistent logger shared by app entrypoints."""

    if not logging.getLogger().handlers:
        logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    logger = logging.getLogger(logger_name)
    logger.setLevel(logging.INFO)
    return logger

