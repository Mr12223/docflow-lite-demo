"""OCR service facade for image processing and invoice extraction."""

from concurrent.futures import ThreadPoolExecutor, as_completed
import re
from typing import Optional

from docflow.core import DocFlowCancelledError
from docflow.settings import env_flag as _env_flag
from docflow.settings import env_int as _env_int

from ..core import logger
from .invoice_merge import (
    _OCR_TAX_ID_FIELDS,
    _apply_ocr_tax_id_box_references,
    _collect_rapidocr_tax_id_field_candidates,
    _collect_tesseract_tax_id_box_references,
    _collect_tesseract_tax_id_crop_candidates,
    _empty_invoice_fields_payload,
    _evaluate_ocr_invoice_text,
    _extract_ocr_tax_id_fields,
    _format_ocr_field_merge_summary,
    _format_ocr_selection_summary,
    _merge_ocr_invoice_fields,
    _ocr_candidate_sort_key,
    _summarize_ocr_candidates,
    _summarize_ocr_tax_id_crop_candidates,
    extract_invoice_fields,
)
from .ocr_engines import (
    IMAGE_EXTS,
    _build_image_ocr_cache_key,
    _build_image_ocr_cache_meta,
    _format_image_ocr_output,
    _format_ocr_engine_chain,
    _get_image_ocr_order,
    _is_image_ocr_cache_enabled,
    _load_image_ocr_cache,
    _restore_cached_image_ocr_result,
    _run_easyocr_variants,
    _run_paddleocr_variants,
    _run_rapidocr_variants,
    _run_tesseract_variants,
    _save_image_ocr_cache,
)


_INVOICE_REFINEMENT_CORE_TAX_ID_LENGTH_RANGE = range(15, 21)


def _get_image_ocr_engine_worker_count(engine_count: int) -> int:
    if engine_count <= 1:
        return 1
    if not _env_flag("DOCFLOW_IMAGE_OCR_PARALLEL_ENGINES", True):
        return 1
    return max(1, min(_env_int("DOCFLOW_IMAGE_OCR_ENGINE_WORKERS", min(2, engine_count)), engine_count))


def _normalize_invoice_tax_id_value(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", str(value or "").upper())


def _extract_invoice_field_value(fields: dict, field_name: str) -> str:
    source = fields.get(field_name) if isinstance(fields, dict) else None
    if isinstance(source, dict):
        return str(source.get("value") or "")
    return str(source or "")


def _extract_candidate_tax_id_value(fields: dict, tax_fields: dict, field_name: str) -> str:
    value = _extract_invoice_field_value(fields, field_name)
    if value:
        return value
    source = tax_fields.get(field_name) if isinstance(tax_fields, dict) else None
    if isinstance(source, dict):
        return str(source.get("value") or "")
    return str(source or "")


def _has_usable_candidate_tax_id(fields: dict, tax_fields: dict, field_name: str) -> bool:
    value = _normalize_invoice_tax_id_value(_extract_candidate_tax_id_value(fields, tax_fields, field_name))
    return len(value) in _INVOICE_REFINEMENT_CORE_TAX_ID_LENGTH_RANGE


def _should_run_invoice_refinement(selected_candidate: dict, selected_invoice_eval: dict) -> tuple[bool, str]:
    if not _env_flag("DOCFLOW_IMAGE_OCR_INVOICE_REFINEMENT", True):
        return False, "disabled"

    selected_fields = selected_invoice_eval.get("fields") or {}
    invoice_fields = selected_invoice_eval.get("invoice_fields") or {}
    expected_fields = set(invoice_fields.get("expected_fields") or [])
    if expected_fields and not any(field_name in expected_fields for field_name in _OCR_TAX_ID_FIELDS):
        return False, "tax_id_not_applicable"

    tax_fields = selected_candidate.get("ocr_tax_id_fields") or {}
    looks_like_invoice = bool(selected_invoice_eval.get("is_invoice")) or bool(tax_fields) or any(
        field_name in selected_fields for field_name in _OCR_TAX_ID_FIELDS
    )
    if not looks_like_invoice:
        return False, "not_invoice"

    if not _env_flag("DOCFLOW_IMAGE_OCR_ADAPTIVE_INVOICE_REFINEMENT", True):
        return True, "full"

    buyer_value = _normalize_invoice_tax_id_value(
        _extract_candidate_tax_id_value(selected_fields, tax_fields, "buyer_tax_id")
    )
    seller_value = _normalize_invoice_tax_id_value(
        _extract_candidate_tax_id_value(selected_fields, tax_fields, "seller_tax_id")
    )
    if not _has_usable_candidate_tax_id(selected_fields, tax_fields, "buyer_tax_id"):
        return True, "missing_buyer_tax_id"
    if not _has_usable_candidate_tax_id(selected_fields, tax_fields, "seller_tax_id"):
        return True, "missing_seller_tax_id"
    if buyer_value and seller_value and buyer_value == seller_value:
        if not _env_flag("DOCFLOW_IMAGE_OCR_DUPLICATE_TAX_ID_REFINEMENT", False):
            return False, "duplicate_tax_id_skipped"
        return True, "duplicate_tax_id"
    return False, "tax_id_complete"


def _get_invoice_refinement_target_fields(reason: str) -> tuple[str, ...]:
    normalized = str(reason or "").strip().lower()
    if normalized == "missing_buyer_tax_id":
        return ("buyer_tax_id",)
    if normalized == "missing_seller_tax_id":
        return ("seller_tax_id",)
    return _OCR_TAX_ID_FIELDS


def process_image_ocr(
    image_path: str,
    filename: str,
    progress_callback=None,
    cancel_callback=None,
    force_reprocess: bool = False,
    extract_invoice: bool = True,
) -> dict:
    import time
    start = time.time()

    def ensure_not_cancelled() -> None:
        if callable(cancel_callback) and cancel_callback():
            raise DocFlowCancelledError("Task cancelled")

    def emit(progress_pct: float, stage: str, message: str = "", **extra) -> None:
        ensure_not_cancelled()
        if callable(progress_callback):
            progress_callback(
                progress_pct=max(0.0, min(float(progress_pct), 100.0)),
                stage=stage,
                message=message,
                **extra,
            )

    emit(5, "image_prepare", f"Loading image: {filename}")
    text = ""
    engine_used = ""
    rapidocr_error = ""
    paddleocr_error = ""
    easyocr_error = ""
    tesseract_error = ""
    prepared_meta = {}
    engine_order = _get_image_ocr_order()
    attempted_engines: list[str] = []
    fallback_notes: list[str] = []
    cache_key = ""
    file_sha256 = ""
    cache_profile = {}
    candidates: list[dict] = []
    candidate_signatures: set[tuple[str, str]] = set()
    tax_id_crop_candidates: list[dict] = []
    rapid_tax_id_field_candidates: list[dict] = []
    tax_id_box_references: dict[str, str] = {}
    selected_invoice_eval = _evaluate_ocr_invoice_text("")
    selection_summary = ""
    merged_invoice_fields = _empty_invoice_fields_payload()
    field_merge_meta = {"applied": False, "fields": {}}
    field_merge_summary = ""
    invoice_refinement_applied = False
    invoice_refinement_reason = ""

    def engine_progress(index: int, phase: float) -> float:
        total = max(len(engine_order), 1)
        return 18 + ((index + phase) / total) * 62

    def register_candidate(
        engine_key: str,
        engine_label: str,
        candidate_text: str,
        candidate_meta: Optional[dict],
        engine_index: int,
    ) -> None:
        normalized_text = str(candidate_text or "").strip()
        if not normalized_text:
            return
        signature = (engine_key, normalized_text)
        if signature in candidate_signatures:
            return
        candidate_signatures.add(signature)
        candidates.append(
            {
                "engine": engine_label,
                "engine_key": engine_key,
                "engine_index": engine_index,
                "text": normalized_text,
                "meta": dict(candidate_meta or {}),
                "char_count": len(normalized_text),
                "invoice_eval": _evaluate_ocr_invoice_text(normalized_text),
                "ocr_tax_id_fields": _extract_ocr_tax_id_fields(normalized_text),
            }
        )

    def register_engine_candidates(
        engine_key: str,
        engine_label: str,
        engine_results: list[dict],
        engine_index: int,
    ) -> bool:
        before_count = len(candidates)
        for variant_offset, candidate in enumerate(engine_results):
            register_candidate(
                engine_key,
                engine_label,
                str(candidate.get("text") or ""),
                candidate.get("meta") if isinstance(candidate, dict) else {},
                engine_index * 100 + variant_offset,
            )
        return len(candidates) > before_count

    def run_engine(engine_name: str) -> dict:
        try:
            ensure_not_cancelled()
            if engine_name == "rapidocr":
                return {
                    "engine_name": engine_name,
                    "engine_key": "rapidocr",
                    "engine_label": "RapidOCR",
                    "results": _run_rapidocr_variants(image_path),
                    "error": "",
                }
            if engine_name == "paddleocr":
                return {
                    "engine_name": engine_name,
                    "engine_key": "paddleocr",
                    "engine_label": "PaddleOCR",
                    "results": _run_paddleocr_variants(image_path),
                    "error": "",
                }
            if engine_name == "tesseract":
                return {
                    "engine_name": engine_name,
                    "engine_key": "tesseract",
                    "engine_label": "Tesseract",
                    "results": _run_tesseract_variants(image_path),
                    "error": "",
                }
            if engine_name == "easyocr":
                return {
                    "engine_name": engine_name,
                    "engine_key": "easyocr",
                    "engine_label": "EasyOCR",
                    "results": _run_easyocr_variants(image_path),
                    "error": "",
                }
            return {
                "engine_name": engine_name,
                "engine_key": engine_name,
                "engine_label": engine_name,
                "results": [],
                "error": f"Unsupported OCR engine: {engine_name}",
            }
        except ImportError as exc:
            return {
                "engine_name": engine_name,
                "engine_key": engine_name,
                "engine_label": engine_name,
                "results": [],
                "error": str(exc),
                "unavailable": True,
            }
        except Exception as exc:
            return {
                "engine_name": engine_name,
                "engine_key": engine_name,
                "engine_label": engine_name,
                "results": [],
                "error": str(exc),
            }

    if _is_image_ocr_cache_enabled():
        try:
            emit(12, "image_cache_lookup", "Checking OCR cache")
            cache_key, file_sha256, cache_profile = _build_image_ocr_cache_key(image_path)
            cached_payload = None if force_reprocess else _load_image_ocr_cache(cache_key)
            if cached_payload:
                elapsed = (time.time() - start) * 1000
                ensure_not_cancelled()
                emit(90, "image_cache_hit", "OCR cache hit")
                cached_result = _restore_cached_image_ocr_result(filename, cached_payload, elapsed)
                if cached_result:
                    if not extract_invoice:
                        cached_stats = cached_result.setdefault("statistics", {})
                        cached_stats["invoice_fields"] = _empty_invoice_fields_payload()
                    return cached_result
            if force_reprocess:
                emit(13, "image_cache_bypass", "Skipping OCR cache")
        except Exception:
            cache_key = ""
            file_sha256 = ""
            cache_profile = {}

    import threading as _threading

    speculative_refinement_holder: list = [None]
    speculative_refinement_thread = None
    if extract_invoice and _env_flag("DOCFLOW_IMAGE_OCR_INVOICE_REFINEMENT", True) and _env_flag(
        "DOCFLOW_IMAGE_OCR_INVOICE_REFINEMENT_SPECULATIVE", False
    ):
        def _run_speculative_refinement() -> None:
            try:
                speculative_refinement_holder[0] = _collect_rapidocr_tax_id_field_candidates(image_path)
            except Exception:
                speculative_refinement_holder[0] = []

        speculative_refinement_thread = _threading.Thread(
            target=_run_speculative_refinement,
            name="image-ocr-refine",
            daemon=True,
        )
        speculative_refinement_thread.start()

    for engine_index, engine_name in enumerate(engine_order):
        attempted_engines.append(engine_name)
        if engine_name == "rapidocr":
            emit(engine_progress(engine_index, 0.05), "image_rapidocr_prepare", "Preparing RapidOCR")
        elif engine_name == "paddleocr":
            emit(engine_progress(engine_index, 0.05), "image_paddleocr_prepare", "Preparing PaddleOCR")
        elif engine_name == "tesseract":
            emit(engine_progress(engine_index, 0.05), "image_preprocess", "Preparing Tesseract image")
        elif engine_name == "easyocr":
            emit(engine_progress(engine_index, 0.05), "image_easyocr_prepare", "Preparing EasyOCR")

    engine_results_by_name: dict[str, dict] = {}
    engine_worker_count = _get_image_ocr_engine_worker_count(len(engine_order))
    if engine_worker_count <= 1:
        for engine_index, engine_name in enumerate(engine_order):
            emit(engine_progress(engine_index, 0.55), f"image_{engine_name}", f"Running {_format_ocr_engine_chain([engine_name])}")
            engine_results_by_name[engine_name] = run_engine(engine_name)
    else:
        emit(18, "image_ocr_parallel", f"Running OCR engines in parallel: {_format_ocr_engine_chain(engine_order)}")
        with ThreadPoolExecutor(max_workers=engine_worker_count, thread_name_prefix="image-ocr-engine") as executor:
            futures = {executor.submit(run_engine, engine_name): engine_name for engine_name in engine_order}
            for future in as_completed(futures):
                engine_name = futures[future]
                engine_results_by_name[engine_name] = future.result()

    for engine_index, engine_name in enumerate(engine_order):
        payload = engine_results_by_name.get(engine_name) or {}
        engine_results = payload.get("results") if isinstance(payload.get("results"), list) else []
        engine_key = str(payload.get("engine_key") or engine_name)
        engine_label = str(payload.get("engine_label") or _format_ocr_engine_chain([engine_name]) or engine_name)
        engine_error = str(payload.get("error") or "")
        unavailable = bool(payload.get("unavailable"))

        if engine_results and register_engine_candidates(engine_key, engine_label, engine_results, engine_index):
            emit(
                engine_progress(engine_index, 0.72),
                f"image_{engine_name}_select",
                f"{engine_label} variants ready: {len(engine_results)}",
            )
            continue

        if engine_name == "rapidocr":
            rapidocr_error = engine_error or "RapidOCR returned no usable text"
            fallback_notes.append(f"RapidOCR failed: {rapidocr_error}" if engine_error else rapidocr_error)
        elif engine_name == "paddleocr":
            paddleocr_error = engine_error or "PaddleOCR returned no usable text"
            fallback_notes.append(f"PaddleOCR failed: {paddleocr_error}" if engine_error and not unavailable else f"PaddleOCR unavailable: {paddleocr_error}" if unavailable else paddleocr_error)
        elif engine_name == "tesseract":
            tesseract_error = engine_error or "Tesseract returned no usable text"
            fallback_notes.append(f"Tesseract unavailable: {tesseract_error}" if unavailable else f"Tesseract failed: {tesseract_error}" if engine_error else tesseract_error)
        elif engine_name == "easyocr":
            easyocr_error = engine_error or "EasyOCR returned no usable text"
            fallback_notes.append(f"EasyOCR unavailable: {easyocr_error}" if unavailable else f"EasyOCR failed: {easyocr_error}" if engine_error else easyocr_error)

    ranked_candidates = sorted(candidates, key=_ocr_candidate_sort_key, reverse=True)
    if ranked_candidates:
        selected_candidate = ranked_candidates[0]
        for candidate in ranked_candidates:
            candidate["selected"] = candidate is selected_candidate
        text = str(selected_candidate.get("text") or "")
        engine_used = str(selected_candidate.get("engine") or "")
        prepared_meta = dict(selected_candidate.get("meta") or {})
        selected_invoice_eval = (
            selected_candidate.get("invoice_eval") or _evaluate_ocr_invoice_text(text)
            if extract_invoice
            else _evaluate_ocr_invoice_text("")
        )
        should_refine, invoice_refinement_reason = (
            _should_run_invoice_refinement(selected_candidate, selected_invoice_eval)
            if extract_invoice
            else (False, "disabled")
        )
        if should_refine:
            invoice_refinement_applied = True
            if speculative_refinement_thread is not None:
                speculative_refinement_thread.join()
                rapid_tax_id_field_candidates = speculative_refinement_holder[0] or []
            else:
                rapid_tax_id_field_candidates = _collect_rapidocr_tax_id_field_candidates(
                    image_path,
                    target_fields=_get_invoice_refinement_target_fields(invoice_refinement_reason),
                )
            if _env_flag("DOCFLOW_IMAGE_OCR_TESSERACT_TAX_ID_REFINEMENT", False):
                tax_id_crop_candidates = _collect_tesseract_tax_id_crop_candidates(image_path)
                tax_id_box_references = _collect_tesseract_tax_id_box_references(image_path)
        merge_candidates = [
            selected_candidate,
            *rapid_tax_id_field_candidates,
            *tax_id_crop_candidates,
            *[candidate for candidate in ranked_candidates if candidate is not selected_candidate],
        ]
        _apply_ocr_tax_id_box_references(merge_candidates, tax_id_box_references)
        merged_invoice_fields, field_merge_meta = _merge_ocr_invoice_fields(selected_candidate, merge_candidates)
        field_merge_summary = _format_ocr_field_merge_summary(field_merge_meta)
        selection_summary = _format_ocr_selection_summary(ranked_candidates, engine_used)
        emit(84, "image_ocr_select", f"Selected OCR engine: {engine_used}")

    if not engine_used:
        error_lines = []
        if rapidocr_error:
            error_lines.append(f"RapidOCR: {rapidocr_error}")
        if paddleocr_error:
            error_lines.append(f"PaddleOCR: {paddleocr_error}")
        if easyocr_error:
            error_lines.append(f"EasyOCR: {easyocr_error}")
        if tesseract_error:
            error_lines.append(f"Tesseract: {tesseract_error}")
        detail = "\n".join(error_lines)
        text = (
            "OCR unavailable or no usable text recognized.\n\n"
            "Check RapidOCR / PaddleOCR / pytesseract + Tesseract OCR / easyocr runtime setup."
        )
        if detail:
            text += f"\n\n{detail}"
        engine_used = "Unavailable"
        selected_invoice_eval = _evaluate_ocr_invoice_text("")
        merged_invoice_fields = _empty_invoice_fields_payload()
        field_merge_meta = {"applied": False, "fields": {}}
        field_merge_summary = ""
        tax_id_crop_candidates = []
        rapid_tax_id_field_candidates = []
        tax_id_box_references = {}

    elapsed = (time.time() - start) * 1000
    ensure_not_cancelled()
    emit(92, "image_finalize", "Formatting OCR result")
    char_count = len(text)

    result = {
        "success": True,
        "file": filename,
        "format": "image",
        "text": text,
        "tables": [],
        "metadata": {
            "engine": engine_used,
            "file": filename,
            "ocr_engine_order": engine_order,
            "ocr_attempted_engines": attempted_engines,
            "ocr_fallback_notes": fallback_notes,
            "fallback_reason": fallback_notes[0] if fallback_notes else "",
            "ocr_preprocess": prepared_meta,
            "ocr_candidates": _summarize_ocr_candidates(ranked_candidates),
            "ocr_selection_strategy": "multi_variant+invoice_field_score+rapidocr_field_crop+tesseract_box_ref",
            "ocr_selection_summary": selection_summary,
            "ocr_field_merge": field_merge_meta,
            "ocr_field_merge_summary": field_merge_summary,
            "ocr_invoice_refinement_applied": invoice_refinement_applied,
            "ocr_invoice_refinement_reason": invoice_refinement_reason,
            "ocr_tax_id_crop_candidates": _summarize_ocr_tax_id_crop_candidates(
                [*rapid_tax_id_field_candidates, *tax_id_crop_candidates]
            ),
            "ocr_tax_id_box_references": dict(tax_id_box_references),
            "image_ocr_cache": _build_image_ocr_cache_meta(
                cache_hit=False,
                file_sha256=file_sha256,
                profile=cache_profile,
                bypassed=force_reprocess,
            ),
        },
        "statistics": {
            "char_count": char_count,
            "paragraph_count": len(text.splitlines()),
            "table_count": 0,
            "keywords": [],
            "processing_ms": elapsed,
            "invoice_fields": (
                merged_invoice_fields or selected_invoice_eval.get("invoice_fields") or _empty_invoice_fields_payload()
                if extract_invoice
                else _empty_invoice_fields_payload()
            ),
        },
        "processing_ms": elapsed,
        "formatted_output": _format_image_ocr_output(
            filename,
            engine_used,
            char_count,
            elapsed,
            text,
            cache_hit=False,
            engine_order=engine_order,
            attempted_engines=attempted_engines,
            fallback_notes=fallback_notes,
            selection_summary=selection_summary,
        ),
        "error": "",
    }
    logger.info(
        "Image OCR complete: file=%s engine=%s strategy=%s summary=%s chain=%s notes=%s elapsed=%.0fms",
        filename,
        engine_used,
        result["metadata"].get("ocr_selection_strategy", ""),
        selection_summary or "-",
        _format_ocr_engine_chain(attempted_engines or engine_order),
        " | ".join(fallback_notes) if fallback_notes else "-",
        elapsed,
    )
    if cache_key and file_sha256 and cache_profile and engine_used in {"Tesseract", "EasyOCR", "PaddleOCR", "RapidOCR"}:
        _save_image_ocr_cache(cache_key, file_sha256, cache_profile, result)
    return result


__all__ = ["IMAGE_EXTS", "extract_invoice_fields", "process_image_ocr"]
