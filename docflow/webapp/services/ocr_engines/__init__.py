"""OCR engine adapters, preprocessing, cache, and runtime bootstrap."""

import base64
import hashlib
import json
import os
import shutil
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Optional

from docflow.paths import IMAGE_OCR_CACHE_DIR
from docflow.settings import env_flag as _env_flag
from docflow.settings import env_int as _env_int
from docflow.core import import_with_base_fallback
from docflow.support import collect_dependency_status, prepare_pytesseract

from ...core import UPLOAD_FOLDER, logger

_GOOGLE_VISION_READY_CACHE = None

_EASYOCR_READER_CACHE = None
_EASYOCR_READER_ERROR = None

_PADDLEOCR_READER_CACHE = None
_PADDLEOCR_READER_ERROR = None

_RAPIDOCR_READER_CACHE = None
_RAPIDOCR_READER_ERROR = None
_RAPIDOCR_READER_LOCK = threading.Lock()

_RAPIDOCR_WARMUP_LOCK = threading.Lock()
_RAPIDOCR_WARMUP_STARTED = False
_RAPIDOCR_WARMUP_FINISHED = False

IMAGE_OCR_CACHE_VERSION = 5
IMAGE_OCR_MEMORY_CACHE: dict[str, dict] = {}
IMAGE_OCR_CACHE_LOCK = threading.Lock()
_OCR_VARIANT_PARALLEL_PROVIDERS = {"tesseract", "easyocr", "rapidocr"}


def clear_image_ocr_cache() -> dict:
    """Clear image OCR disk and in-memory caches."""

    cache_dir = IMAGE_OCR_CACHE_DIR
    cache_dir.mkdir(parents=True, exist_ok=True)
    target = cache_dir.resolve()
    expected = IMAGE_OCR_CACHE_DIR.resolve()
    if target != expected:
        raise RuntimeError(f"Unsafe OCR cache path: {target}")

    file_count = 0
    byte_count = 0
    for item in target.rglob("*"):
        try:
            if item.is_file() or item.is_symlink():
                file_count += 1
                byte_count += item.stat().st_size
        except OSError:
            continue

    with IMAGE_OCR_CACHE_LOCK:
        memory_count = len(IMAGE_OCR_MEMORY_CACHE)
        IMAGE_OCR_MEMORY_CACHE.clear()

    shutil.rmtree(target)
    cache_dir.mkdir(parents=True, exist_ok=True)
    return {
        "cache_dir": str(cache_dir),
        "deleted_files": file_count,
        "deleted_bytes": byte_count,
        "cleared_memory_items": memory_count,
    }


def _is_cloud_runtime() -> bool:
    """Detect whether the app is running in a constrained cloud deployment."""

    markers = (
        "RENDER",
        "RENDER_SERVICE_ID",
        "RENDER_INSTANCE_ID",
        "DOCFLOW_CLOUD_DEPLOYMENT",
    )
    return any(str(os.getenv(name, "")).strip() for name in markers)


def _get_image_ocr_order() -> list[str]:
    """Return the configured OCR engine fallback chain for image processing."""

    default_order = "rapidocr,tesseract"
    raw = str(os.getenv("DOCFLOW_IMAGE_OCR_ORDER", default_order)).strip().lower()
    engines = [item.strip() for item in raw.split(",") if item.strip()]
    valid = [item for item in engines if item in {"rapidocr", "paddleocr", "tesseract", "easyocr"}]
    return valid or default_order.split(",")


def _is_ascii_only_path(path_value: str) -> bool:
    try:
        str(path_value or "").encode("ascii")
        return True
    except UnicodeEncodeError:
        return False


def _describe_ocr_engine(engine_name: str) -> str:
    return {
        "rapidocr": "RapidOCR",
        "paddleocr": "PaddleOCR",
        "tesseract": "Tesseract",
        "easyocr": "EasyOCR",
    }.get(str(engine_name or "").strip().lower(), str(engine_name or "").strip() or "未知引擎")


def _get_next_ocr_engine(engine_order: list[str], current_name: str) -> str:
    normalized = [str(item or "").strip().lower() for item in engine_order]
    current = str(current_name or "").strip().lower()
    try:
        index = normalized.index(current)
    except ValueError:
        return ""
    for candidate in normalized[index + 1 :]:
        if candidate:
            return candidate
    return ""


def _format_ocr_engine_chain(engine_names: list[str]) -> str:
    labels = [_describe_ocr_engine(item) for item in engine_names if str(item or "").strip()]
    return " -> ".join(labels)


def _get_image_ocr_resize_config() -> dict:
    """Build a normalized image resize policy shared by OCR engines."""

    config = {
        "max_long_edge": _env_int("DOCFLOW_IMAGE_OCR_MAX_LONG_EDGE", 1600),
        "target_long_edge": _env_int("DOCFLOW_IMAGE_OCR_TARGET_LONG_EDGE", 1200),
        "fast_edge_trigger": _env_int("DOCFLOW_IMAGE_OCR_FAST_EDGE_TRIGGER", 2400),
        "huge_long_edge": _env_int("DOCFLOW_IMAGE_OCR_HUGE_LONG_EDGE", 1280),
        "grayscale": _env_flag("DOCFLOW_IMAGE_OCR_GRAYSCALE", True),
        "autocontrast": _env_flag("DOCFLOW_IMAGE_OCR_AUTOCONTRAST", True),
        "cloud_downsample_enabled": _env_flag("DOCFLOW_CLOUD_IMAGE_OCR_DOWNSAMPLE", _is_cloud_runtime()),
    }

    if config["cloud_downsample_enabled"]:
        config["max_long_edge"] = min(
            config["max_long_edge"],
            _env_int("DOCFLOW_CLOUD_IMAGE_OCR_MAX_LONG_EDGE", 960),
        )
        config["target_long_edge"] = min(
            config["target_long_edge"],
            _env_int("DOCFLOW_CLOUD_IMAGE_OCR_TARGET_LONG_EDGE", 900),
        )
        config["fast_edge_trigger"] = min(
            config["fast_edge_trigger"],
            _env_int("DOCFLOW_CLOUD_IMAGE_OCR_FAST_EDGE_TRIGGER", 1500),
        )
        config["huge_long_edge"] = min(
            config["huge_long_edge"],
            _env_int("DOCFLOW_CLOUD_IMAGE_OCR_HUGE_LONG_EDGE", 768),
        )

    config["huge_long_edge"] = min(config["huge_long_edge"], config["max_long_edge"])
    config["fast_edge_trigger"] = max(config["fast_edge_trigger"], config["max_long_edge"])
    config["target_long_edge"] = min(config["target_long_edge"], config["max_long_edge"])
    return config


def _split_csv_tokens(raw_value: str) -> list[str]:
    return [item.strip().lower() for item in str(raw_value or "").split(",") if item.strip()]


def _parse_binary_thresholds(raw_value: str, default_values: tuple[int, ...]) -> tuple[int, ...]:
    values: list[int] = []
    for token in _split_csv_tokens(raw_value):
        try:
            threshold = int(token)
        except Exception:
            continue
        if 0 < threshold < 255 and threshold not in values:
            values.append(threshold)
    if values:
        return tuple(values)
    return tuple(default_values or ())


def _get_image_ocr_speed_profile() -> str:
    profile = str(os.getenv("DOCFLOW_IMAGE_OCR_SPEED_PROFILE", "fast")).strip().lower()
    if profile in {"accurate", "full", "quality"}:
        return "accurate"
    if profile in {"balanced", "normal"}:
        return "balanced"
    return "fast"


def _get_image_ocr_variant_config(provider: str = "") -> dict:
    normalized_provider = str(provider or "").strip().lower()
    env_prefix = normalized_provider.upper()
    speed_profile = _get_image_ocr_speed_profile()
    default_variants_by_profile = {
        "fast": {
            "rapidocr": "gray,detail",
            "paddleocr": "gray,detail",
            "tesseract": "gray",
            "easyocr": "gray",
        },
        "balanced": {
            "rapidocr": "gray,detail,contrast",
            "paddleocr": "gray,detail,contrast",
            "tesseract": "gray,contrast",
            "easyocr": "gray,contrast",
        },
        "accurate": {
            "rapidocr": "rgb,gray,detail,contrast,binary",
            "paddleocr": "rgb,gray,detail,contrast,binary",
            "tesseract": "gray,detail,contrast,binary",
            "easyocr": "rgb,gray,detail,contrast",
        },
    }
    default_variants_map = default_variants_by_profile[speed_profile]
    default_thresholds_map = {
        "rapidocr": (170, 190),
        "paddleocr": (170, 190),
        "tesseract": (170, 190),
        "easyocr": (170,),
    }

    provider_variant_env = f"DOCFLOW_{env_prefix}_IMAGE_OCR_VARIANTS" if env_prefix else ""
    provider_threshold_env = f"DOCFLOW_{env_prefix}_IMAGE_OCR_BINARY_THRESHOLDS" if env_prefix else ""
    raw_variants = str(
        os.getenv(provider_variant_env, "") if provider_variant_env else ""
    ).strip() or str(
        os.getenv(
            "DOCFLOW_IMAGE_OCR_VARIANTS",
            default_variants_map.get(normalized_provider, "rgb,gray,contrast,binary"),
        )
    ).strip()

    enabled: list[str] = []
    inline_thresholds: list[int] = []
    for token in _split_csv_tokens(raw_variants):
        if token in {"rgb", "gray", "detail", "contrast", "binary"}:
            if token not in enabled:
                enabled.append(token)
            continue
        if token.startswith("binary_"):
            if "binary" not in enabled:
                enabled.append("binary")
            try:
                threshold = int(token.split("_", 1)[1])
            except Exception:
                continue
            if 0 < threshold < 255 and threshold not in inline_thresholds:
                inline_thresholds.append(threshold)

    if not enabled:
        enabled = _split_csv_tokens(default_variants_map.get(normalized_provider, "rgb,gray,contrast,binary"))

    default_thresholds = tuple(inline_thresholds) or default_thresholds_map.get(normalized_provider, (170, 190))
    raw_thresholds = str(
        os.getenv(provider_threshold_env, "") if provider_threshold_env else ""
    ).strip() or str(
        os.getenv(
            "DOCFLOW_IMAGE_OCR_BINARY_THRESHOLDS",
            ",".join(str(item) for item in default_thresholds),
        )
    ).strip()
    binary_thresholds = _parse_binary_thresholds(raw_thresholds, tuple(default_thresholds))

    variant_names: list[str] = []
    if "rgb" in enabled:
        variant_names.append("rgb")
    if "gray" in enabled:
        variant_names.append("gray_autocontrast")
    if "detail" in enabled:
        variant_names.append("detail_boost")
    if "contrast" in enabled:
        variant_names.append("contrast_sharp")
    if "binary" in enabled:
        for threshold in binary_thresholds:
            variant_names.append(f"binary_{threshold}")

    return {
        "enabled": tuple(enabled),
        "binary_thresholds": tuple(binary_thresholds),
        "variant_names": tuple(variant_names),
    }


def _get_ocr_variant_worker_count(provider: str, variant_count: int) -> int:
    if variant_count <= 1:
        return 1
    provider_name = str(provider or "").strip().lower()
    default_workers = min(4, variant_count) if provider_name in _OCR_VARIANT_PARALLEL_PROVIDERS else 1
    configured_workers = _env_int("DOCFLOW_IMAGE_OCR_VARIANT_WORKERS", default_workers)
    provider_env = f"DOCFLOW_{provider_name.upper()}_IMAGE_OCR_VARIANT_WORKERS" if provider_name else ""
    if provider_env:
        configured_workers = _env_int(provider_env, configured_workers)
    return max(1, min(int(configured_workers or 1), variant_count))


def _run_ocr_variant_tasks(provider: str, variant_entries: list[tuple[str, object, dict]], worker, close_image: bool = True) -> list[dict]:
    if not variant_entries:
        return []

    worker_count = _get_ocr_variant_worker_count(provider, len(variant_entries))
    if worker_count <= 1:
        results: list[dict] = []
        last_error: Optional[Exception] = None
        for variant_index, entry in enumerate(variant_entries):
            try:
                item = worker(variant_index, entry)
                if item:
                    results.append(item)
            except Exception as exc:
                last_error = exc
            finally:
                if close_image:
                    try:
                        entry[1].close()
                    except Exception:
                        pass
        if results:
            return sorted(results, key=lambda item: int((item.get("meta") or {}).get("variant_index") or 0))
        if last_error is not None:
            raise last_error
        return []

    futures = {}
    last_error: Optional[Exception] = None
    with ThreadPoolExecutor(max_workers=worker_count, thread_name_prefix=f"{provider}-ocr") as executor:
        for variant_index, entry in enumerate(variant_entries):
            futures[executor.submit(worker, variant_index, entry)] = entry
        results = []
        for future in as_completed(futures):
            entry = futures[future]
            try:
                item = future.result()
                if item:
                    results.append(item)
            except Exception as exc:
                last_error = exc
            finally:
                if close_image:
                    try:
                        entry[1].close()
                    except Exception:
                        pass

    if results:
        return sorted(results, key=lambda item: int((item.get("meta") or {}).get("variant_index") or 0))
    if last_error is not None:
        raise last_error
    return []


def _get_provider_resize_policy(provider: str) -> dict:
    resize_config = _get_image_ocr_resize_config()
    normalized_provider = str(provider or "").strip().lower()
    env_prefix = normalized_provider.upper()

    if not env_prefix:
        return dict(resize_config)

    max_long_edge = min(
        resize_config["max_long_edge"],
        _env_int(f"DOCFLOW_{env_prefix}_MAX_LONG_EDGE", resize_config["max_long_edge"]),
    )
    target_long_edge = min(
        max_long_edge,
        _env_int(f"DOCFLOW_{env_prefix}_TARGET_LONG_EDGE", resize_config["target_long_edge"]),
    )
    fast_edge_trigger = min(
        resize_config["fast_edge_trigger"],
        _env_int(f"DOCFLOW_{env_prefix}_FAST_EDGE_TRIGGER", resize_config["fast_edge_trigger"]),
    )
    huge_long_edge = min(
        resize_config["huge_long_edge"],
        _env_int(f"DOCFLOW_{env_prefix}_HUGE_LONG_EDGE", resize_config["huge_long_edge"]),
    )
    huge_long_edge = min(huge_long_edge, max_long_edge)
    fast_edge_trigger = max(fast_edge_trigger, max_long_edge)

    return {
        **resize_config,
        "max_long_edge": max_long_edge,
        "target_long_edge": target_long_edge,
        "fast_edge_trigger": fast_edge_trigger,
        "huge_long_edge": huge_long_edge,
    }


def _fit_image_for_ocr(image, *, max_long_edge: int, target_long_edge: int):
    from PIL import Image

    width, height = image.size
    long_edge = max(width, height)
    if long_edge > max_long_edge:
        scale = max_long_edge / max(long_edge, 1)
        resized = image.resize((max(1, int(width * scale)), max(1, int(height * scale))), Image.LANCZOS)
        return resized, "downscaled"
    if long_edge >= target_long_edge:
        return image.copy(), "original"

    scale = target_long_edge / max(long_edge, 1)
    resized = image.resize((max(1, int(width * scale)), max(1, int(height * scale))), Image.LANCZOS)
    return resized, "upscaled"


def _load_base_image_for_ocr(image_path: str, provider: str):
    from PIL import Image, ImageOps

    resize_policy = _get_provider_resize_policy(provider)
    with Image.open(image_path) as source:
        image = ImageOps.exif_transpose(source)
        if image.mode not in ("RGB", "RGBA", "L"):
            image = image.convert("RGB")

        original_size = image.size
        long_edge = max(image.size)
        effective_max_long_edge = (
            resize_policy["huge_long_edge"]
            if long_edge >= resize_policy["fast_edge_trigger"]
            else resize_policy["max_long_edge"]
        )
        prepared, resize_action = _fit_image_for_ocr(
            image,
            max_long_edge=effective_max_long_edge,
            target_long_edge=resize_policy["target_long_edge"],
        )

    return prepared, {
        "original_size": original_size,
        "prepared_size": prepared.size,
        "long_edge_target": effective_max_long_edge,
        "target_long_edge": resize_policy["target_long_edge"],
        "provider": provider,
        "resize_action": resize_action,
        "upscaled": resize_action == "upscaled",
        "downsampled": resize_action == "downscaled",
        "cloud_downsample_enabled": resize_policy["cloud_downsample_enabled"],
    }


def _build_image_variant_entries(base_image, provider: str) -> list[tuple[str, object, dict]]:
    from PIL import ImageEnhance, ImageFilter, ImageOps

    variant_config = _get_image_ocr_variant_config(provider)
    enabled = set(variant_config.get("enabled") or ())
    binary_thresholds = tuple(variant_config.get("binary_thresholds") or ())
    results: list[tuple[str, object, dict]] = []

    if "rgb" in enabled:
        results.append(
            (
                "rgb",
                base_image.convert("RGB"),
                {"preprocess_variant": "rgb", "grayscale": False},
            )
        )

    gray = None
    detail_equalized = None
    detail_boosted = None
    contrast = None
    denoised = None
    sharpened = None
    try:
        if {"gray", "detail", "contrast", "binary"} & enabled:
            gray = ImageOps.grayscale(base_image)
            gray = ImageOps.autocontrast(gray)

        if gray is not None and "gray" in enabled:
            results.append(
                (
                    "gray_autocontrast",
                    gray.copy(),
                    {
                        "preprocess_variant": "gray_autocontrast",
                        "grayscale": True,
                        "autocontrast": True,
                    },
                )
            )

        if gray is not None and ("detail" in enabled or {"contrast", "binary"} & enabled):
            detail_equalized = ImageOps.equalize(gray)
            detail_boosted = ImageEnhance.Contrast(detail_equalized).enhance(1.35)
            detail_boosted = detail_boosted.filter(ImageFilter.MedianFilter(size=3))
            detail_boosted = detail_boosted.filter(
                ImageFilter.UnsharpMask(radius=1.6, percent=180, threshold=2)
            )

        if detail_boosted is not None and "detail" in enabled:
            results.append(
                (
                    "detail_boost",
                    detail_boosted.copy(),
                    {
                        "preprocess_variant": "detail_boost",
                        "grayscale": True,
                        "autocontrast": True,
                        "equalize": True,
                        "contrast_boost": 1.35,
                        "median_filter": 3,
                        "unsharp_radius": 1.6,
                        "unsharp_percent": 180,
                        "unsharp_threshold": 2,
                    },
                )
            )

        if gray is not None and ({"contrast", "binary"} & enabled):
            contrast_source = detail_boosted if detail_boosted is not None else gray
            contrast = ImageEnhance.Contrast(contrast_source).enhance(1.8)
            denoised = contrast.filter(ImageFilter.MedianFilter(size=3))
            sharpened = ImageEnhance.Sharpness(denoised).enhance(2.4)

        if sharpened is not None and "contrast" in enabled:
            results.append(
                (
                    "contrast_sharp",
                    sharpened.copy(),
                    {
                        "preprocess_variant": "contrast_sharp",
                        "grayscale": True,
                        "autocontrast": True,
                        "detail_boost": detail_boosted is not None,
                        "equalize": detail_boosted is not None,
                        "contrast_boost": 1.8,
                        "median_filter": 3,
                        "sharpness_boost": 2.4,
                    },
                )
            )

        if sharpened is not None and "binary" in enabled:
            for threshold in binary_thresholds:
                binary = sharpened.point(
                    lambda value, limit=int(threshold): 255 if value > limit else 0,
                    mode="1",
                ).convert("L")
                results.append(
                    (
                        f"binary_{threshold}",
                        binary,
                        {
                            "preprocess_variant": f"binary_{threshold}",
                            "grayscale": True,
                            "autocontrast": True,
                            "detail_boost": detail_boosted is not None,
                            "equalize": detail_boosted is not None,
                            "contrast_boost": 1.8,
                            "median_filter": 3,
                            "sharpness_boost": 2.4,
                            "binary_threshold": int(threshold),
                        },
                    )
                )
    finally:
        for image_obj in (gray, detail_equalized, detail_boosted, contrast, denoised, sharpened):
            if image_obj is not None:
                image_obj.close()

    return results


def _run_path_based_ocr_variants(image_path: str, provider: str, ocr_runner) -> list[dict]:
    import tempfile
    base_image, base_meta = _load_base_image_for_ocr(image_path, provider)
    try:
        variant_entries = _build_image_variant_entries(base_image, provider)
    finally:
        base_image.close()

    def run_entry(variant_index: int, entry: tuple[str, object, dict]) -> dict:
        variant_name, variant_image, variant_meta = entry
        temp_path = ""
        try:
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png", dir=UPLOAD_FOLDER)
            temp_path = temp_file.name
            temp_file.close()
            variant_image.save(temp_path, format="PNG", optimize=True)
            candidate_text, extra_meta = ocr_runner(temp_path)
            normalized_text = str(candidate_text or "").strip()
            if not normalized_text:
                return {}
            return {
                "text": normalized_text,
                "meta": {
                    **dict(base_meta or {}),
                    **dict(variant_meta or {}),
                    **dict(extra_meta or {}),
                    "provider": provider,
                    "variant_name": variant_name,
                    "variant_index": variant_index,
                    "transport_format": "png",
                    "path_sanitized": True,
                },
            }
        finally:
            if temp_path:
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

    return _run_ocr_variant_tasks(provider, variant_entries, run_entry)


def _run_pil_ocr_variants(image_path: str, provider: str, ocr_runner) -> list[dict]:
    base_image, base_meta = _load_base_image_for_ocr(image_path, provider)
    try:
        variant_entries = _build_image_variant_entries(base_image, provider)
    finally:
        base_image.close()

    def run_entry(variant_index: int, entry: tuple[str, object, dict]) -> dict:
        variant_name, variant_image, variant_meta = entry
        candidate_text, extra_meta = ocr_runner(variant_image)
        normalized_text = str(candidate_text or "").strip()
        if not normalized_text:
            return {}
        return {
            "text": normalized_text,
            "meta": {
                **dict(base_meta or {}),
                **dict(variant_meta or {}),
                **dict(extra_meta or {}),
                "provider": provider,
                "variant_name": variant_name,
                "variant_index": variant_index,
            },
        }

    return _run_ocr_variant_tasks(provider, variant_entries, run_entry)


def _select_best_ocr_variant_result(results: list[dict]) -> tuple[str, dict]:
    if not results:
        return "", {}

    best = max(
        results,
        key=lambda item: (
            len(str(item.get("text") or "").splitlines()),
            len(str(item.get("text") or "")),
            -int((item.get("meta") or {}).get("variant_index") or 0),
        ),
    )
    return str(best.get("text") or ""), dict(best.get("meta") or {})


def _get_google_vision_api_key() -> str:
    """Read the first available Google Vision API key from known env names."""

    for env_name in (
        "DOCFLOW_GOOGLE_VISION_API_KEY",
        "GOOGLE_VISION_API_KEY",
        "GOOGLE_API_KEY",
    ):
        value = str(os.getenv(env_name, "")).strip()
        if value:
            return value
    return ""


def _get_google_vision_language_hints() -> list[str]:
    """Read optional Google Vision language hints from the environment."""

    raw = str(os.getenv("DOCFLOW_GOOGLE_VISION_LANGUAGE_HINTS", "")).strip()
    return [item.strip() for item in raw.split(",") if item.strip()]


def _get_google_vision_feature() -> str:
    """Normalize the requested Google Vision feature type."""

    value = str(os.getenv("DOCFLOW_GOOGLE_VISION_FEATURE", "DOCUMENT_TEXT_DETECTION")).strip().upper()
    return value if value in {"DOCUMENT_TEXT_DETECTION", "TEXT_DETECTION"} else "DOCUMENT_TEXT_DETECTION"


def _get_google_vision_endpoint() -> str:
    return str(
        os.getenv("DOCFLOW_GOOGLE_VISION_API_ENDPOINT", "https://vision.googleapis.com/v1/images:annotate")
    ).strip()


def _is_google_vision_enabled() -> bool:
    default_enabled = _is_cloud_runtime() or bool(_get_google_vision_api_key())
    return _env_flag("DOCFLOW_ENABLE_GOOGLE_VISION_OCR", default_enabled)


def _check_google_vision_ready() -> tuple[bool, str]:
    global _GOOGLE_VISION_READY_CACHE
    if _GOOGLE_VISION_READY_CACHE is not None:
        return _GOOGLE_VISION_READY_CACHE

    if not _is_google_vision_enabled():
        _GOOGLE_VISION_READY_CACHE = (False, "Google Vision OCR 已禁用")
        return _GOOGLE_VISION_READY_CACHE

    endpoint = _get_google_vision_endpoint()
    if not endpoint.startswith("https://"):
        _GOOGLE_VISION_READY_CACHE = (False, "Google Vision OCR 接口地址无效")
        return _GOOGLE_VISION_READY_CACHE

    api_key = _get_google_vision_api_key()
    if not api_key:
        _GOOGLE_VISION_READY_CACHE = (False, "未配置 Google Vision API Key")
        return _GOOGLE_VISION_READY_CACHE

    _GOOGLE_VISION_READY_CACHE = (True, "")
    return _GOOGLE_VISION_READY_CACHE


def _get_easyocr_reader():
    global _EASYOCR_READER_CACHE, _EASYOCR_READER_ERROR
    if _EASYOCR_READER_CACHE is not None:
        return _EASYOCR_READER_CACHE
    if _EASYOCR_READER_ERROR is not None:
        raise _EASYOCR_READER_ERROR

    try:
        easyocr = import_with_base_fallback("easyocr")
        _EASYOCR_READER_CACHE = easyocr.Reader(["ch_sim", "en"], verbose=False)
        return _EASYOCR_READER_CACHE
    except Exception as exc:
        _EASYOCR_READER_ERROR = exc
        raise


def _prepare_image_for_google_vision(image_path: str) -> tuple[bytes, dict]:
    import io
    from PIL import Image, ImageOps

    resize_config = _get_image_ocr_resize_config()
    max_long_edge = min(
        resize_config["max_long_edge"],
        _env_int("DOCFLOW_GOOGLE_VISION_MAX_LONG_EDGE", resize_config["max_long_edge"]),
    )
    huge_trigger = min(
        resize_config["fast_edge_trigger"],
        _env_int("DOCFLOW_GOOGLE_VISION_FAST_EDGE_TRIGGER", resize_config["fast_edge_trigger"]),
    )
    huge_long_edge = min(
        resize_config["huge_long_edge"],
        _env_int("DOCFLOW_GOOGLE_VISION_HUGE_LONG_EDGE", resize_config["huge_long_edge"]),
    )
    huge_long_edge = min(huge_long_edge, max_long_edge)
    huge_trigger = max(huge_trigger, max_long_edge)

    with Image.open(image_path) as source:
        image = ImageOps.exif_transpose(source)
        if image.mode not in ("RGB", "RGBA", "L"):
            image = image.convert("RGB")

        original_size = image.size
        width, height = image.size
        long_edge = max(width, height)
        target_long_edge = huge_long_edge if long_edge >= huge_trigger else max_long_edge

        if long_edge > target_long_edge:
            scale = target_long_edge / max(long_edge, 1)
            image = image.resize(
                (max(1, int(width * scale)), max(1, int(height * scale))),
                Image.LANCZOS,
            )

        if image.mode not in ("RGB", "RGBA", "L"):
            image = image.convert("RGB")

        payload = io.BytesIO()
        image.save(payload, format="PNG", optimize=True)
        image_bytes = payload.getvalue()
        prepared_size = image.size

    return image_bytes, {
        "original_size": original_size,
        "prepared_size": prepared_size,
        "long_edge_target": target_long_edge,
        "provider": "google_vision",
        "transport_format": "png",
        "cloud_downsample_enabled": resize_config["cloud_downsample_enabled"],
    }


def _run_google_vision_ocr(image_path: str) -> tuple[str, dict]:
    from urllib import error, parse, request

    is_ready, reason = _check_google_vision_ready()
    if not is_ready:
        raise RuntimeError(reason)

    api_key = _get_google_vision_api_key()
    endpoint = _get_google_vision_endpoint()
    feature_type = _get_google_vision_feature()
    language_hints = _get_google_vision_language_hints()
    timeout_sec = _env_int("DOCFLOW_GOOGLE_VISION_TIMEOUT_SEC", 20)
    image_bytes, prepared_meta = _prepare_image_for_google_vision(image_path)

    request_payload = {
        "requests": [
            {
                "image": {"content": base64.b64encode(image_bytes).decode("ascii")},
                "features": [{"type": feature_type}],
            }
        ]
    }
    if language_hints:
        request_payload["requests"][0]["imageContext"] = {"languageHints": language_hints}

    request_url = f"{endpoint}?key={parse.quote(api_key)}"
    req = request.Request(
        request_url,
        data=json.dumps(request_payload, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )

    try:
        with request.urlopen(req, timeout=timeout_sec) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Google Vision HTTP {exc.code}: {detail[:240] or exc.reason}") from exc
    except error.URLError as exc:
        raise RuntimeError(f"Google Vision 网络请求失败: {exc.reason}") from exc
    except Exception as exc:
        raise RuntimeError(f"Google Vision 请求异常: {exc}") from exc

    responses = payload.get("responses") or []
    if not responses:
        raise RuntimeError("Google Vision 未返回识别结果")

    first = responses[0] or {}
    error_info = first.get("error") or {}
    if error_info.get("message"):
        raise RuntimeError(str(error_info.get("message")))

    text = ((first.get("fullTextAnnotation") or {}).get("text")) or ""
    if not text:
        text_annotations = first.get("textAnnotations") or []
        if text_annotations:
            text = str(text_annotations[0].get("description") or "")

    prepared_meta.update(
        {
            "feature": feature_type,
            "language_hints": language_hints,
            "endpoint": endpoint,
            "payload_bytes": len(image_bytes),
        }
    )
    return text, prepared_meta


def _get_paddleocr_reader():
    global _PADDLEOCR_READER_CACHE, _PADDLEOCR_READER_ERROR
    if not _env_flag("DOCFLOW_ENABLE_PADDLEOCR", True):
        raise RuntimeError("PaddleOCR 已禁用")
    if _PADDLEOCR_READER_CACHE is not None:
        return _PADDLEOCR_READER_CACHE
    if _PADDLEOCR_READER_ERROR is not None:
        raise _PADDLEOCR_READER_ERROR

    try:
        import inspect

        paddleocr_module = import_with_base_fallback("paddleocr")
        paddleocr_cls = getattr(paddleocr_module, "PaddleOCR", None)
        if paddleocr_cls is None:
            raise RuntimeError("未找到 PaddleOCR 类")

        signature = inspect.signature(paddleocr_cls)
        kwargs = {}
        if "lang" in signature.parameters:
            kwargs["lang"] = str(os.getenv("DOCFLOW_PADDLEOCR_LANG", "ch")).strip() or "ch"
        if "show_log" in signature.parameters:
            kwargs["show_log"] = _env_flag("DOCFLOW_PADDLEOCR_SHOW_LOG", False)
        if "use_angle_cls" in signature.parameters:
            kwargs["use_angle_cls"] = _env_flag("DOCFLOW_PADDLEOCR_USE_ANGLE_CLS", False)
        if "use_gpu" in signature.parameters:
            kwargs["use_gpu"] = _env_flag("DOCFLOW_PADDLEOCR_USE_GPU", False)

        _PADDLEOCR_READER_CACHE = paddleocr_cls(**kwargs)
        return _PADDLEOCR_READER_CACHE
    except Exception as exc:
        _PADDLEOCR_READER_ERROR = exc
        raise


def _get_rapidocr_reader():
    global _RAPIDOCR_READER_CACHE, _RAPIDOCR_READER_ERROR
    if not _env_flag("DOCFLOW_ENABLE_RAPIDOCR", True):
        raise RuntimeError("RapidOCR 已禁用")
    if _RAPIDOCR_READER_CACHE is not None:
        return _RAPIDOCR_READER_CACHE
    if _RAPIDOCR_READER_ERROR is not None:
        raise _RAPIDOCR_READER_ERROR

    try:
        rapidocr_module = import_with_base_fallback("rapidocr")
        rapidocr_cls = getattr(rapidocr_module, "RapidOCR", None)
        if rapidocr_cls is None:
            raise RuntimeError("未找到 RapidOCR 类")
        try:
            _RAPIDOCR_READER_CACHE = rapidocr_cls()
        except TypeError:
            _RAPIDOCR_READER_CACHE = rapidocr_cls(params={})
        return _RAPIDOCR_READER_CACHE
    except Exception as exc:
        _RAPIDOCR_READER_ERROR = exc
        raise


def _get_rapidocr_reader():
    global _RAPIDOCR_READER_CACHE, _RAPIDOCR_READER_ERROR
    if not _env_flag("DOCFLOW_ENABLE_RAPIDOCR", True):
        raise RuntimeError("RapidOCR is disabled")
    if _RAPIDOCR_READER_CACHE is not None:
        return _RAPIDOCR_READER_CACHE
    if _RAPIDOCR_READER_ERROR is not None:
        raise _RAPIDOCR_READER_ERROR

    with _RAPIDOCR_READER_LOCK:
        if _RAPIDOCR_READER_CACHE is not None:
            return _RAPIDOCR_READER_CACHE
        if _RAPIDOCR_READER_ERROR is not None:
            raise _RAPIDOCR_READER_ERROR

        try:
            rapidocr_module = import_with_base_fallback("rapidocr")
            rapidocr_cls = getattr(rapidocr_module, "RapidOCR", None)
            if rapidocr_cls is None:
                raise RuntimeError("RapidOCR class not found")
            try:
                _RAPIDOCR_READER_CACHE = rapidocr_cls()
            except TypeError:
                _RAPIDOCR_READER_CACHE = rapidocr_cls(params={})
            return _RAPIDOCR_READER_CACHE
        except Exception as exc:
            _RAPIDOCR_READER_ERROR = exc
            raise


def _prepare_image_for_rapidocr(image_path: str) -> tuple[str, dict, callable]:
    import tempfile
    from PIL import Image, ImageOps

    resize_config = _get_image_ocr_resize_config()
    max_long_edge = min(
        resize_config["max_long_edge"],
        _env_int("DOCFLOW_RAPIDOCR_MAX_LONG_EDGE", resize_config["max_long_edge"]),
    )
    huge_trigger = min(
        resize_config["fast_edge_trigger"],
        _env_int("DOCFLOW_RAPIDOCR_FAST_EDGE_TRIGGER", resize_config["fast_edge_trigger"]),
    )
    huge_long_edge = min(
        resize_config["huge_long_edge"],
        _env_int("DOCFLOW_RAPIDOCR_HUGE_LONG_EDGE", resize_config["huge_long_edge"]),
    )
    huge_long_edge = min(huge_long_edge, max_long_edge)
    huge_trigger = max(huge_trigger, max_long_edge)

    def create_temp_copy(pil_image, *, downsampled: bool) -> tuple[str, dict, callable]:
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png", dir=UPLOAD_FOLDER)
        temp_path = temp_file.name
        temp_file.close()
        pil_image.save(temp_path, format="PNG", optimize=True)

        def cleanup() -> None:
            try:
                os.remove(temp_path)
            except Exception:
                pass

        return temp_path, {
            "transport_format": "png",
            "path_sanitized": True,
            "downsampled": downsampled,
        }, cleanup

    with Image.open(image_path) as source:
        image = ImageOps.exif_transpose(source)
        if image.mode not in ("RGB", "RGBA", "L"):
            image = image.convert("RGB")

        original_size = image.size
        width, height = image.size
        long_edge = max(width, height)
        target_long_edge = huge_long_edge if long_edge >= huge_trigger else max_long_edge

        if long_edge <= target_long_edge:
            path_meta = {
                "original_size": original_size,
                "prepared_size": original_size,
                "long_edge_target": target_long_edge,
                "provider": "rapidocr",
                "cloud_downsample_enabled": resize_config["cloud_downsample_enabled"],
            }
            if _is_ascii_only_path(image_path):
                return image_path, path_meta, (lambda: None)
            temp_path, temp_meta, cleanup = create_temp_copy(image, downsampled=False)
            return temp_path, {**path_meta, **temp_meta}, cleanup

        scale = target_long_edge / max(long_edge, 1)
        resized = image.resize(
            (max(1, int(width * scale)), max(1, int(height * scale))),
            Image.LANCZOS,
        )
        temp_path, temp_meta, cleanup = create_temp_copy(resized, downsampled=True)
        prepared_size = resized.size

    return temp_path, {
        "original_size": original_size,
        "prepared_size": prepared_size,
        "long_edge_target": target_long_edge,
        "provider": "rapidocr",
        "cloud_downsample_enabled": resize_config["cloud_downsample_enabled"],
        **temp_meta,
    }, cleanup


def _extract_text_from_rapidocr_result(payload) -> str:
    lines: list[str] = []

    def add_line(value) -> None:
        text = str(value or "").strip()
        if text:
            lines.append(text)

    def walk(node) -> None:
        if node is None:
            return
        txts = getattr(node, "txts", None)
        if isinstance(txts, (list, tuple)):
            for item in txts:
                add_line(item)
            return
        if isinstance(node, dict):
            for key in ("txts", "texts", "rec_texts"):
                value = node.get(key)
                if isinstance(value, (list, tuple)):
                    for item in value:
                        add_line(item)
                    return
            if isinstance(node.get("text"), str):
                add_line(node.get("text"))
                return
            for key in ("res", "result", "data", "ocr_result"):
                if key in node:
                    walk(node.get(key))
            return
        if isinstance(node, (list, tuple)):
            if len(node) >= 3 and isinstance(node[1], str):
                add_line(node[1])
                return
            if len(node) >= 2 and isinstance(node[1], (list, tuple)):
                candidate = node[1]
                if candidate and isinstance(candidate[0], str):
                    add_line(candidate[0])
                    return
            for item in node:
                walk(item)

    walk(payload)
    merged = []
    last_line = None
    for line in lines:
        if line != last_line:
            merged.append(line)
            last_line = line
    return "\n".join(merged)


def _run_rapidocr_variants(image_path: str) -> list[dict]:
    reader = _get_rapidocr_reader()

    def run_variant(prepared_path: str) -> tuple[str, dict]:
        raw_result = reader(prepared_path)
        return _extract_text_from_rapidocr_result(raw_result), {
            "api_variant": "__call__",
            "engine_class": reader.__class__.__name__,
            "elapsed": getattr(raw_result, "elapse", None),
        }

    return _run_path_based_ocr_variants(image_path, "rapidocr", run_variant)


def _run_rapidocr(image_path: str) -> tuple[str, dict]:
    return _select_best_ocr_variant_result(_run_rapidocr_variants(image_path))


def _should_prewarm_rapidocr() -> bool:
    return _env_flag("DOCFLOW_RAPIDOCR_PREWARM", True)


def _run_rapidocr_warmup() -> None:
    global _RAPIDOCR_WARMUP_FINISHED

    start = time.time()
    warmup_path = None
    try:
        import tempfile
        from PIL import Image, ImageDraw, ImageFont

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png", dir=UPLOAD_FOLDER)
        warmup_path = temp_file.name
        temp_file.close()

        image = Image.new("RGB", (640, 160), "white")
        draw = ImageDraw.Draw(image)
        font = ImageFont.load_default()
        draw.text((18, 24), "DocFlow OCR Warmup 123", fill="black", font=font)
        draw.text((18, 78), "Render CPU warmup", fill="black", font=font)
        image.save(warmup_path, format="PNG", optimize=True)
        image.close()

        text, meta = _run_rapidocr(warmup_path)
        elapsed_ms = (time.time() - start) * 1000
        logger.info(
            "RapidOCR warmup complete: elapsed=%.0fms prepared=%s chars=%s",
            elapsed_ms,
            meta.get("prepared_size"),
            len(str(text or "")),
        )
    except Exception as exc:
        logger.warning("RapidOCR warmup failed: %s", exc)
    finally:
        _RAPIDOCR_WARMUP_FINISHED = True
        if warmup_path:
            try:
                os.remove(warmup_path)
            except Exception:
                pass


def _schedule_rapidocr_warmup() -> None:
    global _RAPIDOCR_WARMUP_STARTED

    if not _should_prewarm_rapidocr():
        return
    if not _env_flag("DOCFLOW_ENABLE_RAPIDOCR", True):
        return
    if "rapidocr" not in _get_image_ocr_order():
        return

    with _RAPIDOCR_WARMUP_LOCK:
        if _RAPIDOCR_WARMUP_STARTED:
            return
        _RAPIDOCR_WARMUP_STARTED = True

    logger.info("RapidOCR warmup scheduled")
    threading.Thread(target=_run_rapidocr_warmup, name="rapidocr-warmup", daemon=True).start()


def _prepare_image_for_paddleocr(image_path: str) -> tuple[str, dict, callable]:
    import tempfile
    from PIL import Image, ImageOps

    resize_config = _get_image_ocr_resize_config()
    max_long_edge = min(
        resize_config["max_long_edge"],
        _env_int("DOCFLOW_PADDLEOCR_MAX_LONG_EDGE", resize_config["max_long_edge"]),
    )
    huge_trigger = min(
        resize_config["fast_edge_trigger"],
        _env_int("DOCFLOW_PADDLEOCR_FAST_EDGE_TRIGGER", resize_config["fast_edge_trigger"]),
    )
    huge_long_edge = min(
        resize_config["huge_long_edge"],
        _env_int("DOCFLOW_PADDLEOCR_HUGE_LONG_EDGE", resize_config["huge_long_edge"]),
    )
    huge_long_edge = min(huge_long_edge, max_long_edge)
    huge_trigger = max(huge_trigger, max_long_edge)

    def create_temp_copy(pil_image, *, downsampled: bool) -> tuple[str, dict, callable]:
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png", dir=UPLOAD_FOLDER)
        temp_path = temp_file.name
        temp_file.close()
        pil_image.save(temp_path, format="PNG", optimize=True)

        def cleanup() -> None:
            try:
                os.remove(temp_path)
            except Exception:
                pass

        return temp_path, {
            "transport_format": "png",
            "path_sanitized": True,
            "downsampled": downsampled,
        }, cleanup

    with Image.open(image_path) as source:
        image = ImageOps.exif_transpose(source)
        if image.mode not in ("RGB", "RGBA", "L"):
            image = image.convert("RGB")

        original_size = image.size
        width, height = image.size
        long_edge = max(width, height)
        target_long_edge = huge_long_edge if long_edge >= huge_trigger else max_long_edge

        if long_edge <= target_long_edge:
            path_meta = {
                "original_size": original_size,
                "prepared_size": original_size,
                "long_edge_target": target_long_edge,
                "provider": "paddleocr",
                "cloud_downsample_enabled": resize_config["cloud_downsample_enabled"],
            }
            if _is_ascii_only_path(image_path):
                return image_path, path_meta, (lambda: None)
            temp_path, temp_meta, cleanup = create_temp_copy(image, downsampled=False)
            return temp_path, {**path_meta, **temp_meta}, cleanup

        scale = target_long_edge / max(long_edge, 1)
        resized = image.resize(
            (max(1, int(width * scale)), max(1, int(height * scale))),
            Image.LANCZOS,
        )
        temp_path, temp_meta, cleanup = create_temp_copy(resized, downsampled=True)
        prepared_size = resized.size

    return temp_path, {
        "original_size": original_size,
        "prepared_size": prepared_size,
        "long_edge_target": target_long_edge,
        "provider": "paddleocr",
        "cloud_downsample_enabled": resize_config["cloud_downsample_enabled"],
        **temp_meta,
    }, cleanup


def _extract_text_from_paddleocr_result(payload) -> str:
    lines: list[str] = []

    def add_line(value) -> None:
        text = str(value or "").strip()
        if text:
            lines.append(text)

    def walk(node) -> None:
        if node is None:
            return
        if isinstance(node, dict):
            rec_texts = node.get("rec_texts")
            if isinstance(rec_texts, (list, tuple)):
                for item in rec_texts:
                    add_line(item)
                return
            if isinstance(node.get("text"), str):
                add_line(node.get("text"))
                return
            for key in ("res", "result", "ocr_result", "data"):
                if key in node:
                    walk(node.get(key))
            return
        if isinstance(node, (list, tuple)):
            if len(node) >= 2 and isinstance(node[1], (list, tuple)) and node[1]:
                candidate = node[1][0]
                if isinstance(candidate, str):
                    add_line(candidate)
                    return
            for item in node:
                walk(item)

    walk(payload)
    merged = []
    last_line = None
    for line in lines:
        if line != last_line:
            merged.append(line)
            last_line = line
    return "\n".join(merged)


def _run_paddleocr(image_path: str) -> tuple[str, dict]:
    reader = _get_paddleocr_reader()
    prepared_path, prepared_meta, cleanup = _prepare_image_for_paddleocr(image_path)
    try:
        if hasattr(reader, "predict"):
            raw_result = reader.predict(prepared_path)
            api_variant = "predict"
        elif hasattr(reader, "ocr"):
            raw_result = reader.ocr(
                prepared_path,
                cls=_env_flag("DOCFLOW_PADDLEOCR_USE_ANGLE_CLS", False),
            )
            api_variant = "ocr"
        else:
            raise RuntimeError("当前 PaddleOCR 版本缺少可调用的预测接口")
    finally:
        cleanup()

    text = _extract_text_from_paddleocr_result(raw_result)
    prepared_meta.update(
        {
            "api_variant": api_variant,
            "lang": str(os.getenv("DOCFLOW_PADDLEOCR_LANG", "ch")).strip() or "ch",
        }
    )
    return text, prepared_meta


def _run_paddleocr_variants(image_path: str) -> list[dict]:
    reader = _get_paddleocr_reader()

    def run_variant(prepared_path: str) -> tuple[str, dict]:
        if hasattr(reader, "predict"):
            raw_result = reader.predict(prepared_path)
            api_variant = "predict"
        elif hasattr(reader, "ocr"):
            raw_result = reader.ocr(
                prepared_path,
                cls=_env_flag("DOCFLOW_PADDLEOCR_USE_ANGLE_CLS", False),
            )
            api_variant = "ocr"
        else:
            raise RuntimeError("Current PaddleOCR runtime does not expose a supported inference API")

        return _extract_text_from_paddleocr_result(raw_result), {
            "api_variant": api_variant,
            "lang": str(os.getenv("DOCFLOW_PADDLEOCR_LANG", "ch")).strip() or "ch",
        }

    return _run_path_based_ocr_variants(image_path, "paddleocr", run_variant)


def _prepare_image_for_easyocr(image_path: str) -> tuple[str, dict, callable]:
    import tempfile
    from PIL import Image, ImageOps

    if _is_ascii_only_path(image_path):
        return image_path, {"provider": "easyocr", "path_sanitized": False}, (lambda: None)

    with Image.open(image_path) as source:
        image = ImageOps.exif_transpose(source)
        if image.mode not in ("RGB", "RGBA", "L"):
            image = image.convert("RGB")
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png", dir=UPLOAD_FOLDER)
        temp_path = temp_file.name
        temp_file.close()
        image.save(temp_path, format="PNG", optimize=True)

    def cleanup() -> None:
        try:
            os.remove(temp_path)
        except Exception:
            pass

    return temp_path, {"provider": "easyocr", "path_sanitized": True, "transport_format": "png"}, cleanup


def _run_easyocr_variants(image_path: str) -> list[dict]:
    reader = _get_easyocr_reader()

    def run_variant(prepared_path: str) -> tuple[str, dict]:
        results = reader.readtext(prepared_path, detail=0)
        return "\n".join(results), {
            "detail": 0,
            "result_count": len(results),
        }

    return _run_path_based_ocr_variants(image_path, "easyocr", run_variant)


def _prepare_image_for_tesseract(image_path: str):
    from PIL import Image, ImageOps

    resize_config = _get_image_ocr_resize_config()
    max_long_edge = resize_config["max_long_edge"]
    huge_trigger = resize_config["fast_edge_trigger"]
    huge_long_edge = resize_config["huge_long_edge"]
    apply_gray = resize_config["grayscale"]
    apply_autocontrast = resize_config["autocontrast"]

    with Image.open(image_path) as source:
        image = ImageOps.exif_transpose(source)
        if image.mode not in ("RGB", "L"):
            image = image.convert("RGB")

        original_size = image.size
        width, height = image.size
        long_edge = max(width, height)
        target_long_edge = huge_long_edge if long_edge >= huge_trigger else max_long_edge

        if long_edge > target_long_edge:
            scale = target_long_edge / max(long_edge, 1)
            image = image.resize(
                (max(1, int(width * scale)), max(1, int(height * scale))),
                Image.LANCZOS,
            )

        if apply_gray and image.mode != "L":
            image = ImageOps.grayscale(image)
        if apply_autocontrast:
            image = ImageOps.autocontrast(image)

        prepared = image.copy()

    return prepared, {
        "original_size": original_size,
        "prepared_size": prepared.size,
        "long_edge_target": target_long_edge,
        "grayscale": prepared.mode == "L",
        "cloud_downsample_enabled": resize_config["cloud_downsample_enabled"],
    }


def _build_image_tesseract_config(base_config: str = "") -> str:
    parts = [str(base_config or "").strip()]
    psm = _env_int("DOCFLOW_IMAGE_OCR_PSM", 6)
    oem = os.getenv("DOCFLOW_IMAGE_OCR_OEM", "").strip()
    if psm:
        parts.append(f"--psm {psm}")
    if oem:
        parts.append(f"--oem {oem}")
    return " ".join(part for part in parts if part).strip()


def _run_tesseract_variants(image_path: str) -> list[dict]:
    import pytesseract

    prepared = prepare_pytesseract()
    config = _build_image_tesseract_config(prepared.get("config", ""))
    lang = prepared.get("lang", "chi_sim+eng")

    def run_variant(image) -> tuple[str, dict]:
        return pytesseract.image_to_string(image, lang=lang, config=config), {
            "lang": lang,
            "config": config,
        }

    return _run_pil_ocr_variants(image_path, "tesseract", run_variant)


def _clone_json_payload(payload):
    return json.loads(json.dumps(payload, ensure_ascii=False))


def _is_image_ocr_cache_enabled() -> bool:
    return _env_flag("DOCFLOW_ENABLE_IMAGE_OCR_CACHE", True)


def _compute_file_sha256(file_path: str) -> str:
    digest = hashlib.sha256()
    with open(file_path, "rb") as file_obj:
        while True:
            chunk = file_obj.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def _get_image_ocr_profile() -> dict:
    resize_config = _get_image_ocr_resize_config()
    provider_variant_config = {
        provider: _get_image_ocr_variant_config(provider)
        for provider in ("rapidocr", "paddleocr", "tesseract", "easyocr")
    }
    provider_resize_policies = {
        provider: _get_provider_resize_policy(provider)
        for provider in ("rapidocr", "paddleocr", "tesseract", "easyocr")
    }
    result = {
        "version": IMAGE_OCR_CACHE_VERSION,
        "speed_profile": _get_image_ocr_speed_profile(),
        "order": _get_image_ocr_order(),
        "max_long_edge": resize_config["max_long_edge"],
        "target_long_edge": resize_config["target_long_edge"],
        "fast_edge_trigger": resize_config["fast_edge_trigger"],
        "huge_long_edge": resize_config["huge_long_edge"],
        "grayscale": resize_config["grayscale"],
        "autocontrast": resize_config["autocontrast"],
        "cloud_downsample_enabled": resize_config["cloud_downsample_enabled"],
        "rapidocr_enabled": _env_flag("DOCFLOW_ENABLE_RAPIDOCR", True),
        "rapidocr_max_long_edge": _env_int("DOCFLOW_RAPIDOCR_MAX_LONG_EDGE", resize_config["max_long_edge"]),
        "rapidocr_fast_edge_trigger": _env_int("DOCFLOW_RAPIDOCR_FAST_EDGE_TRIGGER", resize_config["fast_edge_trigger"]),
        "rapidocr_huge_long_edge": _env_int("DOCFLOW_RAPIDOCR_HUGE_LONG_EDGE", resize_config["huge_long_edge"]),
        "paddleocr_enabled": _env_flag("DOCFLOW_ENABLE_PADDLEOCR", True),
        "paddleocr_lang": str(os.getenv("DOCFLOW_PADDLEOCR_LANG", "ch")).strip() or "ch",
        "paddleocr_use_angle_cls": _env_flag("DOCFLOW_PADDLEOCR_USE_ANGLE_CLS", False),
        "psm": _env_int("DOCFLOW_IMAGE_OCR_PSM", 6),
        "oem": os.getenv("DOCFLOW_IMAGE_OCR_OEM", "").strip(),
        "rapidocr_variants": list(provider_variant_config["rapidocr"].get("variant_names") or ()),
        "paddleocr_variants": list(provider_variant_config["paddleocr"].get("variant_names") or ()),
        "tesseract_variants": list(provider_variant_config["tesseract"].get("variant_names") or ()),
        "easyocr_variants": list(provider_variant_config["easyocr"].get("variant_names") or ()),
        "provider_resize_policies": {
            provider: {
                "max_long_edge": int(policy.get("max_long_edge") or 0),
                "target_long_edge": int(policy.get("target_long_edge") or 0),
                "fast_edge_trigger": int(policy.get("fast_edge_trigger") or 0),
                "huge_long_edge": int(policy.get("huge_long_edge") or 0),
            }
            for provider, policy in provider_resize_policies.items()
        },
    }
    return result


def _build_image_ocr_cache_key(image_path: str) -> tuple[str, str, dict]:
    file_sha256 = _compute_file_sha256(image_path)
    profile = _get_image_ocr_profile()
    profile_json = json.dumps(profile, ensure_ascii=False, sort_keys=True)
    cache_key = hashlib.sha256(f"{file_sha256}|{profile_json}".encode("utf-8")).hexdigest()
    return cache_key, file_sha256, profile


def _get_image_ocr_cache_file(cache_key: str) -> Path:
    return IMAGE_OCR_CACHE_DIR / f"{cache_key}.json"


def _load_image_ocr_cache(cache_key: str) -> Optional[dict]:
    if not _is_image_ocr_cache_enabled():
        return None

    with IMAGE_OCR_CACHE_LOCK:
        cached_payload = IMAGE_OCR_MEMORY_CACHE.get(cache_key)
    if cached_payload:
        return _clone_json_payload(cached_payload)

    cache_file = _get_image_ocr_cache_file(cache_key)
    if not cache_file.exists():
        return None

    try:
        payload = json.loads(cache_file.read_text(encoding="utf-8"))
        if payload.get("version") != IMAGE_OCR_CACHE_VERSION:
            return None
        result = payload.get("result")
        if not isinstance(result, dict):
            return None
    except Exception:
        return None

    with IMAGE_OCR_CACHE_LOCK:
        IMAGE_OCR_MEMORY_CACHE[cache_key] = payload
    return _clone_json_payload(payload)


def _save_image_ocr_cache(cache_key: str, file_sha256: str, profile: dict, result: dict) -> None:
    if not _is_image_ocr_cache_enabled():
        return
    if not isinstance(result, dict):
        return
    if result.get("metadata", {}).get("engine") not in {"Tesseract", "EasyOCR", "PaddleOCR", "RapidOCR"}:
        return

    payload = {
        "version": IMAGE_OCR_CACHE_VERSION,
        "saved_at": time.time(),
        "file_sha256": file_sha256,
        "profile": profile,
        "result": _clone_json_payload(result),
    }
    cache_file = _get_image_ocr_cache_file(cache_key)
    try:
        cache_file.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        with IMAGE_OCR_CACHE_LOCK:
            IMAGE_OCR_MEMORY_CACHE[cache_key] = payload
    except Exception:
        return


def _format_image_ocr_output(
    filename: str,
    engine_used: str,
    char_count: int,
    elapsed_ms: float,
    text: str,
    *,
    cache_hit: bool = False,
    cache_original_processing_ms: Optional[float] = None,
) -> str:
    cache_lines = [f"缓存命中: {'是' if cache_hit else '否'}"]
    if cache_hit and cache_original_processing_ms is not None:
        cache_lines.append(f"原始OCR耗时: {cache_original_processing_ms:.0f}ms")

    return f"""[图片 OCR 结果] {filename}
{'━' * 40}
OCR 引擎: {engine_used}
识别字符数: {char_count}
处理耗时: {elapsed_ms:.0f}ms
{chr(10).join(cache_lines)}

识别内容:
{'━' * 20}
{text}
"""

# ── 创建 Flask 应用
def _format_image_ocr_output(
    filename: str,
    engine_used: str,
    char_count: int,
    elapsed_ms: float,
    text: str,
    *,
    cache_hit: bool = False,
    cache_original_processing_ms: Optional[float] = None,
    engine_order: Optional[list[str]] = None,
    attempted_engines: Optional[list[str]] = None,
    fallback_notes: Optional[list[str]] = None,
    selection_summary: Optional[str] = None,
) -> str:
    details = [f"缓存命中: {'是' if cache_hit else '否'}"]
    if cache_hit and cache_original_processing_ms is not None:
        details.append(f"原始OCR耗时: {cache_original_processing_ms:.0f}ms")
    if engine_order:
        details.append(f"引擎顺序: {_format_ocr_engine_chain(engine_order)}")
    if attempted_engines:
        details.append(f"尝试链路: {_format_ocr_engine_chain(attempted_engines)}")
    if selection_summary:
        details.append(f"候选打分: {selection_summary}")
    if fallback_notes:
        notes = " | ".join(str(item) for item in fallback_notes if str(item).strip())
        if notes:
            details.append(f"回退说明: {notes}")

    return f"""[图片 OCR 结果] {filename}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OCR 引擎: {engine_used}
识别字符数: {char_count}
处理耗时: {elapsed_ms:.0f}ms
{chr(10).join(details)}

识别内容:
━━━━━━━━━━━━━━━━━━━━
{text}
"""


def _build_image_ocr_cache_meta(
    *,
    cache_hit: bool,
    file_sha256: str = "",
    profile: Optional[dict] = None,
    saved_at: Optional[float] = None,
    original_processing_ms: Optional[float] = None,
    bypassed: bool = False,
) -> dict:
    payload = {
        "enabled": _is_image_ocr_cache_enabled(),
        "hit": bool(cache_hit),
        "bypassed": bool(bypassed),
        "file_sha256": file_sha256,
        "profile": profile or {},
    }
    if saved_at is not None:
        payload["saved_at"] = saved_at
    if original_processing_ms is not None:
        payload["original_processing_ms"] = round(float(original_processing_ms), 2)
    return payload


def _restore_cached_image_ocr_result(filename: str, cached_payload: dict, elapsed_ms: float) -> Optional[dict]:
    if not isinstance(cached_payload, dict):
        return None

    cached_result = cached_payload.get("result")
    if not isinstance(cached_result, dict):
        return None

    result = _clone_json_payload(cached_result)
    metadata = result.setdefault("metadata", {})
    statistics = result.setdefault("statistics", {})
    engine_used = str(metadata.get("engine") or "")
    text = str(result.get("text") or "")
    engine_order = metadata.get("ocr_engine_order") if isinstance(metadata.get("ocr_engine_order"), list) else []
    attempted_engines = metadata.get("ocr_attempted_engines") if isinstance(metadata.get("ocr_attempted_engines"), list) else []
    fallback_notes = metadata.get("ocr_fallback_notes") if isinstance(metadata.get("ocr_fallback_notes"), list) else []
    selection_summary = str(metadata.get("ocr_selection_summary") or "").strip()
    char_count = int(statistics.get("char_count") or len(text))
    original_processing_ms = result.get("processing_ms")
    if original_processing_ms is None:
        original_processing_ms = statistics.get("processing_ms")

    result["success"] = True
    result["file"] = filename
    result["error"] = ""
    result["processing_ms"] = elapsed_ms
    statistics["char_count"] = char_count
    statistics["paragraph_count"] = int(statistics.get("paragraph_count") or len(text.splitlines()))
    statistics["table_count"] = int(statistics.get("table_count") or 0)
    statistics["processing_ms"] = elapsed_ms
    metadata["file"] = filename
    metadata["image_ocr_cache"] = _build_image_ocr_cache_meta(
        cache_hit=True,
        file_sha256=str(cached_payload.get("file_sha256") or ""),
        profile=cached_payload.get("profile") if isinstance(cached_payload.get("profile"), dict) else {},
        saved_at=cached_payload.get("saved_at"),
        original_processing_ms=original_processing_ms,
    )
    result["formatted_output"] = _format_image_ocr_output(
        filename,
        engine_used,
        char_count,
        elapsed_ms,
        text,
        cache_hit=True,
        cache_original_processing_ms=original_processing_ms,
        engine_order=engine_order,
        attempted_engines=attempted_engines,
        fallback_notes=fallback_notes,
        selection_summary=selection_summary,
    )
    return result

def _log_ocr_runtime_status() -> None:
    try:
        dep_status = collect_dependency_status()
        items = {item.get("id"): item for item in dep_status.get("items", []) if isinstance(item, dict)}
        image_profile = dep_status.get("profiles", {}).get("image_ocr", {})
        logger.info(
            "OCR startup check: cloud=%s order=%s profile=%s reason=%s prewarm=%s threads=%s/%s",
            _is_cloud_runtime(),
            _format_ocr_engine_chain(_get_image_ocr_order()),
            image_profile.get("status", "unknown"),
            image_profile.get("reason", ""),
            _should_prewarm_rapidocr(),
            os.getenv("OMP_NUM_THREADS", "-"),
            os.getenv("OPENBLAS_NUM_THREADS", "-"),
        )
        for dep_id in ("rapidocr", "onnxruntime", "pytesseract"):
            item = items.get(dep_id)
            if not item:
                continue
            logger.info(
                "OCR dependency: %s installed=%s available=%s version=%s message=%s",
                dep_id,
                item.get("installed"),
                item.get("available", item.get("installed")),
                item.get("version", ""),
                item.get("message", ""),
            )
    except Exception as exc:
        logger.warning("OCR startup check failed: %s", exc)

_log_ocr_runtime_status()
_schedule_rapidocr_warmup()

IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.webp'}
