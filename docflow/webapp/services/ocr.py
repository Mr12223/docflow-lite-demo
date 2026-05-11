"""OCR service facade for image processing and invoice extraction."""

from typing import Optional

from docflow_core import DocFlowCancelledError

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


def process_image_ocr(image_path: str, filename: str, progress_callback=None, cancel_callback=None) -> dict:
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

    if _is_image_ocr_cache_enabled():
        try:
            emit(12, "image_cache_lookup", "Checking OCR cache")
            cache_key, file_sha256, cache_profile = _build_image_ocr_cache_key(image_path)
            cached_payload = _load_image_ocr_cache(cache_key)
            if cached_payload:
                elapsed = (time.time() - start) * 1000
                ensure_not_cancelled()
                emit(90, "image_cache_hit", "OCR cache hit")
                cached_result = _restore_cached_image_ocr_result(filename, cached_payload, elapsed)
                if cached_result:
                    return cached_result
        except Exception:
            cache_key = ""
            file_sha256 = ""
            cache_profile = {}

    for engine_index, engine_name in enumerate(engine_order):
        attempted_engines.append(engine_name)

        if engine_name == "rapidocr":
            try:
                ensure_not_cancelled()
                emit(engine_progress(engine_index, 0.05), "image_rapidocr_prepare", "Preparing RapidOCR")
                ensure_not_cancelled()
                emit(engine_progress(engine_index, 0.55), "image_rapidocr", "Running RapidOCR")
                engine_results = _run_rapidocr_variants(image_path)
                if register_engine_candidates("rapidocr", "RapidOCR", engine_results, engine_index):
                    emit(
                        engine_progress(engine_index, 0.72),
                        "image_rapidocr_select",
                        f"RapidOCR variants ready: {len(engine_results)}",
                    )
                else:
                    rapidocr_error = "RapidOCR returned no usable text"
                    fallback_notes.append(rapidocr_error)
            except Exception as e:
                rapidocr_error = str(e)
                fallback_notes.append(f"RapidOCR failed: {rapidocr_error}")

        elif engine_name == "paddleocr":
            try:
                ensure_not_cancelled()
                emit(engine_progress(engine_index, 0.05), "image_paddleocr_prepare", "Preparing PaddleOCR")
                ensure_not_cancelled()
                emit(engine_progress(engine_index, 0.55), "image_paddleocr", "Running PaddleOCR")
                engine_results = _run_paddleocr_variants(image_path)
                if register_engine_candidates("paddleocr", "PaddleOCR", engine_results, engine_index):
                    emit(
                        engine_progress(engine_index, 0.72),
                        "image_paddleocr_select",
                        f"PaddleOCR variants ready: {len(engine_results)}",
                    )
                else:
                    paddleocr_error = "PaddleOCR returned no usable text"
                    fallback_notes.append(paddleocr_error)
            except Exception as e:
                paddleocr_error = str(e)
                fallback_notes.append(f"PaddleOCR failed: {paddleocr_error}")

        elif engine_name == "tesseract":
            try:
                ensure_not_cancelled()
                emit(engine_progress(engine_index, 0.05), "image_preprocess", "Preparing Tesseract image")
                ensure_not_cancelled()
                emit(engine_progress(engine_index, 0.55), "image_tesseract", "Running Tesseract")
                engine_results = _run_tesseract_variants(image_path)
                if register_engine_candidates("tesseract", "Tesseract", engine_results, engine_index):
                    emit(
                        engine_progress(engine_index, 0.72),
                        "image_tesseract_select",
                        f"Tesseract variants ready: {len(engine_results)}",
                    )
                else:
                    tesseract_error = "Tesseract returned no usable text"
                    fallback_notes.append(tesseract_error)
            except ImportError as e:
                tesseract_error = str(e) or "pytesseract not installed"
                fallback_notes.append(f"Tesseract unavailable: {tesseract_error}")
            except Exception as e:
                tesseract_error = str(e)
                fallback_notes.append(f"Tesseract failed: {tesseract_error}")

        elif engine_name == "easyocr":
            try:
                ensure_not_cancelled()
                emit(engine_progress(engine_index, 0.55), "image_easyocr", "Running EasyOCR")
                engine_results = _run_easyocr_variants(image_path)
                if register_engine_candidates("easyocr", "EasyOCR", engine_results, engine_index):
                    emit(
                        engine_progress(engine_index, 0.72),
                        "image_easyocr_select",
                        f"EasyOCR variants ready: {len(engine_results)}",
                    )
                else:
                    easyocr_error = "EasyOCR returned no usable text"
                    fallback_notes.append(easyocr_error)
            except ImportError as e:
                easyocr_error = str(e) or "easyocr not installed"
                fallback_notes.append(f"EasyOCR unavailable: {easyocr_error}")
            except Exception as e:
                easyocr_error = str(e)
                fallback_notes.append(f"EasyOCR failed: {easyocr_error}")

    ranked_candidates = sorted(candidates, key=_ocr_candidate_sort_key, reverse=True)
    if ranked_candidates:
        selected_candidate = ranked_candidates[0]
        for candidate in ranked_candidates:
            candidate["selected"] = candidate is selected_candidate
        text = str(selected_candidate.get("text") or "")
        engine_used = str(selected_candidate.get("engine") or "")
        prepared_meta = dict(selected_candidate.get("meta") or {})
        selected_invoice_eval = selected_candidate.get("invoice_eval") or _evaluate_ocr_invoice_text(text)
        selected_fields = selected_invoice_eval.get("fields") or {}
        if selected_invoice_eval.get("is_invoice") or (
            selected_candidate.get("ocr_tax_id_fields")
            or any(field_name in selected_fields for field_name in _OCR_TAX_ID_FIELDS)
        ):
            rapid_tax_id_field_candidates = _collect_rapidocr_tax_id_field_candidates(image_path)
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
            "ocr_tax_id_crop_candidates": _summarize_ocr_tax_id_crop_candidates(
                [*rapid_tax_id_field_candidates, *tax_id_crop_candidates]
            ),
            "ocr_tax_id_box_references": dict(tax_id_box_references),
            "image_ocr_cache": _build_image_ocr_cache_meta(
                cache_hit=False,
                file_sha256=file_sha256,
                profile=cache_profile,
            ),
        },
        "statistics": {
            "char_count": char_count,
            "paragraph_count": len(text.splitlines()),
            "table_count": 0,
            "keywords": [],
            "processing_ms": elapsed,
            "invoice_fields": merged_invoice_fields or selected_invoice_eval.get("invoice_fields") or _empty_invoice_fields_payload(),
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
