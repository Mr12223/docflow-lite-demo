"""Invoice-specific OCR field extraction and cross-engine merge helpers."""

import os
import re
import statistics
from difflib import SequenceMatcher
from typing import Optional

from docflow_support import prepare_pytesseract
from invoice_extractor import FIELD_LABELS, InvoiceExtractor

from ..core import UPLOAD_FOLDER
from .ocr_engines import (
    _build_image_tesseract_config,
    _clone_json_payload,
    _extract_text_from_rapidocr_result,
    _get_rapidocr_reader,
)


_INVOICE_EXTRACTOR = InvoiceExtractor()
_OCR_INVOICE_FIELD_WEIGHTS = {
    "invoice_code": 30,
    "invoice_number": 30,
    "invoice_date": 14,
    "amount": 12,
    "tax": 8,
    "total": 14,
    "buyer_name": 10,
    "seller_name": 10,
    "buyer_tax_id": 12,
    "seller_tax_id": 12,
    "invoice_type": 10,
    "machine_number": 8,
    "check_code": 8,
}
_OCR_CONFIDENCE_RANK = {"none": 0, "low": 1, "medium": 2, "high": 3}
_OCR_TAX_ID_FIELDS = ("buyer_tax_id", "seller_tax_id")
_OCR_TAX_ID_LABELS = ("纳税人识别号", "税号")
_OCR_TAX_ID_PARTY_HINTS = {
    "buyer_tax_id": ("购买方信息", "购买方", "购方"),
    "seller_tax_id": ("销售方信息", "销售方", "销方"),
}
_OCR_TAX_ID_BLOCK_END_HINTS = (
    "货物名称",
    "项目",
    "规格",
    "数量",
    "税率",
    "金额",
    "价税合计",
    "开票人",
    "收款人",
    "备注",
)
_OCR_TAX_ID_INLINE_STOP_TERMS = _OCR_TAX_ID_LABELS + (
    "开户地址",
    "开户地址电话",
    "开户地址及电话",
    "开户行",
    "开户行及账号",
    "地址电话",
    "名称",
)


_OCR_TAX_ID_EXPECTED_LENGTH = 18
_OCR_TAX_ID_CROP_SPECS = (
    {"name": "tax_id_crop_top", "top": 0.12, "bottom": 0.30, "scale": 3, "psm": 6},
    {"name": "tax_id_crop_mid", "top": 0.18, "bottom": 0.42, "scale": 3, "psm": 6},
)
_OCR_TAX_ID_FIELD_CROP_SPECS = {
    "buyer_tax_id": {"left": 0.12, "right": 0.47, "top": 0.18, "bottom": 0.26},
    "seller_tax_id": {"left": 0.53, "right": 0.92, "top": 0.18, "bottom": 0.26},
}
_OCR_TAX_ID_FIELD_CROP_VARIANTS = (
    {"name": "gray", "scale": 4, "sharpen": False},
    {"name": "sharp", "scale": 4, "sharpen": True},
    {"name": "bin170", "scale": 4, "sharpen": False, "threshold": 170},
)
_OCR_TAX_ID_FUZZY_CHAR_MAP = {
    "0": "0",
    "O": "0",
    "Q": "0",
    "D": "0",
    "1": "1",
    "I": "1",
    "L": "1",
    "5": "5",
    "S": "5",
    "7": "7",
    "Z": "7",
    "8": "8",
    "B": "8",
}


def _get_invoice_extractor() -> InvoiceExtractor:
    return _INVOICE_EXTRACTOR


def _rank_confidence(level: str) -> int:
    return _OCR_CONFIDENCE_RANK.get(str(level or "").strip().lower(), 0)


def _evaluate_ocr_invoice_text(text: str) -> dict:
    normalized_text = str(text or "").strip()
    if not normalized_text:
        return {
            "score": 0,
            "is_invoice": False,
            "confidence": "none",
            "field_count": 0,
            "major_field_count": 0,
            "header_pair": False,
            "money_pair": False,
            "party_pair": False,
            "fields": {},
            "field_names": [],
            "invoice_fields": {
                "is_invoice": False,
                "confidence": "none",
                "field_count": 0,
                "fields": {},
            },
        }

    invoice_fields = _get_invoice_extractor().extract(normalized_text)
    fields = invoice_fields.get("fields") or {}
    score = 0
    for field_name, weight in _OCR_INVOICE_FIELD_WEIGHTS.items():
        if field_name not in fields:
            continue
        score += weight
        field_confidence = _rank_confidence((fields.get(field_name) or {}).get("confidence"))
        score += field_confidence * 2

    is_invoice = bool(invoice_fields.get("is_invoice"))
    confidence = str(invoice_fields.get("confidence") or "none")
    field_count = int(invoice_fields.get("field_count") or len(fields))
    major_field_count = sum(
        1
        for field_name in (
            "invoice_code",
            "invoice_number",
            "invoice_date",
            "amount",
            "total",
            "buyer_name",
            "seller_name",
            "invoice_type",
        )
        if field_name in fields
    )
    header_pair = "invoice_code" in fields and "invoice_number" in fields
    money_pair = "amount" in fields and "total" in fields
    party_pair = "buyer_name" in fields and "seller_name" in fields

    if is_invoice:
        score += 60
    score += _rank_confidence(confidence) * 10
    score += field_count * 3
    score += major_field_count * 4
    if header_pair:
        score += 18
    if money_pair:
        score += 14
    if party_pair:
        score += 8

    char_count = len(normalized_text)
    digit_count = len(re.findall(r"\d", normalized_text))
    if char_count >= 20:
        score += min(12, char_count // 40)
    if digit_count >= 8:
        score += min(8, digit_count // 8)

    return {
        "score": int(score),
        "is_invoice": is_invoice,
        "confidence": confidence,
        "field_count": field_count,
        "major_field_count": major_field_count,
        "header_pair": header_pair,
        "money_pair": money_pair,
        "party_pair": party_pair,
        "fields": fields,
        "field_names": sorted(fields.keys()),
        "invoice_fields": invoice_fields,
    }


def _empty_invoice_fields_payload() -> dict:
    return {
        "is_invoice": False,
        "confidence": "none",
        "field_count": 0,
        "fields": {},
    }


def _normalize_ocr_tax_id(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", str(value or "").upper())


def _normalize_ocr_tax_id_fuzzy_char(value: str) -> str:
    return _OCR_TAX_ID_FUZZY_CHAR_MAP.get(str(value or "").upper(), str(value or "").upper())


def _is_ocr_tax_id_fuzzy_equivalent(left: str, right: str) -> bool:
    if not left or not right:
        return False
    return _normalize_ocr_tax_id_fuzzy_char(left) == _normalize_ocr_tax_id_fuzzy_char(right)


def _extract_ocr_tax_id_repair_extras(base_segment: str, alt_segment: str, prev_char: str = "", next_char: str = "") -> str:
    extras: list[str] = []
    normalized_base = [_normalize_ocr_tax_id(char) for char in base_segment]
    for char in str(alt_segment or ""):
        upper_char = _normalize_ocr_tax_id(char)
        if not upper_char:
            continue
        if any(_is_ocr_tax_id_fuzzy_equivalent(upper_char, base_char) and upper_char != base_char for base_char in normalized_base):
            continue
        if prev_char and _is_ocr_tax_id_fuzzy_equivalent(upper_char, prev_char) and upper_char != prev_char:
            continue
        if next_char and _is_ocr_tax_id_fuzzy_equivalent(upper_char, next_char) and upper_char != next_char:
            continue
        extras.append(upper_char)
    return "".join(extras)


def _repair_ocr_tax_id_value(base_value: str, alt_value: str) -> str:
    base = _normalize_ocr_tax_id(base_value)
    alt = _normalize_ocr_tax_id(alt_value)
    if not base or not alt or base == alt:
        return base

    repaired_parts: list[str] = []
    for tag, i1, i2, j1, j2 in SequenceMatcher(None, base, alt).get_opcodes():
        base_segment = base[i1:i2]
        alt_segment = alt[j1:j2]
        prev_char = repaired_parts[-1][-1] if repaired_parts and repaired_parts[-1] else ""
        next_char = base[i1] if i1 < len(base) else ""

        if tag == "equal":
            repaired_parts.append(base_segment)
            continue

        if tag == "delete":
            repaired_parts.append(base_segment)
            continue

        if tag == "insert":
            extras = _extract_ocr_tax_id_repair_extras("", alt_segment, prev_char=prev_char, next_char=next_char)
            if extras:
                repaired_parts.append(extras)
            continue

        repaired_parts.append(base_segment)
        if len(alt_segment) > len(base_segment):
            extras = _extract_ocr_tax_id_repair_extras(base_segment, alt_segment, prev_char=prev_char, next_char=next_char)
            if extras:
                repaired_parts.append(extras)

    repaired = _normalize_ocr_tax_id("".join(repaired_parts))
    if not repaired:
        return base

    base_distance = abs(len(base) - _OCR_TAX_ID_EXPECTED_LENGTH)
    repaired_distance = abs(len(repaired) - _OCR_TAX_ID_EXPECTED_LENGTH)
    if repaired_distance > base_distance:
        return base
    if repaired_distance == base_distance and len(repaired) <= len(base):
        return base
    if len(repaired) > 20:
        return base
    return repaired


def _score_ocr_tax_id_candidate(value: str) -> tuple[int, int, int, int]:
    normalized = _normalize_ocr_tax_id(value)
    if not normalized:
        return (0, -99, 0, 0)
    digit_count = sum(char.isdigit() for char in normalized)
    alpha_count = sum(char.isalpha() for char in normalized)
    return (
        1 if digit_count and alpha_count else 0,
        -abs(len(normalized) - _OCR_TAX_ID_EXPECTED_LENGTH),
        len(normalized),
        digit_count,
    )


def _extract_ocr_tax_id_segment_values(segment: str) -> list[str]:
    cleaned_segment = str(segment or "").strip()
    if not cleaned_segment:
        return []

    cleaned_segment = re.sub(r"^[\s:：;；,，|丨\-—_]+", "", cleaned_segment)
    stop_expr = "|".join(map(re.escape, _OCR_TAX_ID_INLINE_STOP_TERMS))
    cleaned_segment = re.split(rf"(?:{stop_expr})", cleaned_segment, maxsplit=1, flags=re.IGNORECASE)[0].strip()
    if not cleaned_segment:
        return []

    values: list[str] = []
    seen: set[str] = set()

    def register(value: str) -> None:
        normalized = _normalize_ocr_tax_id(value)
        if 15 <= len(normalized) <= 20 and normalized not in seen:
            seen.add(normalized)
            values.append(normalized)

    chunks = [cleaned_segment]
    chunks.extend(chunk.strip() for chunk in re.split(r"\s{2,}", cleaned_segment) if chunk.strip())

    for chunk in chunks:
        register(chunk)
        for raw_value in re.findall(
            r"[A-Za-z0-9](?:[A-Za-z0-9]|\s(?!\s)){14,28}[A-Za-z0-9]",
            chunk,
            flags=re.IGNORECASE,
        ):
            register(raw_value)
        for raw_value in re.findall(r"[A-Za-z0-9]{15,20}", chunk, flags=re.IGNORECASE):
            register(raw_value)

    return sorted(values, key=_score_ocr_tax_id_candidate, reverse=True)


def _extract_single_ocr_tax_id_field(text: str, field_name: str, method: str) -> dict[str, dict]:
    structured = _extract_ocr_tax_id_fields(text)
    source = structured.get(field_name)
    if isinstance(source, dict):
        return {
            field_name: {
                "value": _normalize_ocr_tax_id(source.get("value", "")),
                "confidence": str(source.get("confidence") or "high"),
                "method": str(source.get("method") or method),
            }
        }

    for line in str(text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        values = _extract_ocr_tax_id_values_from_line(line)
        if not values:
            values = _extract_ocr_tax_id_segment_values(line)
        if not values:
            continue
        return {
            field_name: {
                "value": values[0],
                "confidence": "medium" if len(values[0]) >= _OCR_TAX_ID_EXPECTED_LENGTH else "low",
                "method": method,
            }
        }
    return {}


def _extract_ocr_tax_id_values_from_line(line: str) -> list[str]:
    if not line:
        return []

    label_expr = "|".join(map(re.escape, _OCR_TAX_ID_LABELS))
    values: list[str] = []
    seen: set[str] = set()

    label_matches = list(re.finditer(label_expr, line, flags=re.IGNORECASE))
    if label_matches:
        for index, match in enumerate(label_matches):
            segment_start = match.end()
            prefix_match = re.match(r"[\s:：;；,，|丨\-—_]*", line[segment_start:])
            if prefix_match:
                segment_start += prefix_match.end()
            segment_end = label_matches[index + 1].start() if index + 1 < len(label_matches) else len(line)
            segment_values = _extract_ocr_tax_id_segment_values(line[segment_start:segment_end])
            if len(label_matches) > 1 and segment_values:
                segment_values = [segment_values[0]]
            for normalized in segment_values:
                if normalized not in seen:
                    seen.add(normalized)
                    values.append(normalized)

    if values:
        return values

    for raw_value in re.findall(
        rf"(?:{label_expr})\s*[:：]?\s*([A-Za-z0-9](?:[A-Za-z0-9]|\s(?!\s)){{14,28}}[A-Za-z0-9])",
        line,
        flags=re.IGNORECASE,
    ):
        normalized = _normalize_ocr_tax_id(raw_value)
        if 15 <= len(normalized) <= 20 and normalized not in seen:
            seen.add(normalized)
            values.append(normalized)

    if values:
        return values

    if not re.search(label_expr, line, flags=re.IGNORECASE):
        return []

    for raw_value in re.findall(
        r"[A-Za-z0-9](?:[A-Za-z0-9]|\s(?!\s)){14,28}[A-Za-z0-9]",
        line,
        flags=re.IGNORECASE,
    ):
        normalized = _normalize_ocr_tax_id(raw_value)
        if 15 <= len(normalized) <= 20 and normalized not in seen:
            seen.add(normalized)
            values.append(normalized)
    return values


def _is_ocr_tax_id_block_end(line: str) -> bool:
    return any(term in line for term in _OCR_TAX_ID_BLOCK_END_HINTS)


def _contains_ocr_party_hint(line: str, field_name: str) -> bool:
    return any(term in line for term in _OCR_TAX_ID_PARTY_HINTS.get(field_name, ()))


def _infer_single_ocr_tax_id_field(
    lines: list[str],
    tax_line_index: int,
    name_line_indices: list[int],
    buyer_header_index: int,
    seller_header_index: int,
) -> tuple[str, str, str]:
    line = lines[tax_line_index]
    buyer_inline = _contains_ocr_party_hint(line, "buyer_tax_id")
    seller_inline = _contains_ocr_party_hint(line, "seller_tax_id")
    if buyer_inline and not seller_inline:
        return "buyer_tax_id", "high", "inline_party_hint"
    if seller_inline and not buyer_inline:
        return "seller_tax_id", "high", "inline_party_hint"

    if len(name_line_indices) >= 2:
        buyer_name_index = name_line_indices[0]
        seller_name_index = name_line_indices[1]
        nearest_name_index = min(
            (buyer_name_index, seller_name_index),
            key=lambda idx: (abs(tax_line_index - idx), -idx),
        )
        if nearest_name_index == buyer_name_index:
            return "buyer_tax_id", "medium", "nearest_name_line"
        return "seller_tax_id", "medium", "nearest_name_line"

    if len(name_line_indices) == 1 and name_line_indices[0] <= tax_line_index:
        return "buyer_tax_id", "medium", "single_name_line"

    if seller_header_index >= 0 and tax_line_index >= seller_header_index:
        return "seller_tax_id", "medium", "seller_header"
    if buyer_header_index >= 0 and tax_line_index >= buyer_header_index:
        return "buyer_tax_id", "medium", "buyer_header"
    return "", "none", ""


def _extract_ocr_tax_id_fields(text: str) -> dict[str, dict]:
    lines = [line.strip() for line in str(text or "").splitlines() if line.strip()]
    if not lines:
        return {}

    result: dict[str, dict] = {}
    name_line_indices: list[int] = []
    tax_line_entries: list[tuple[int, str]] = []
    buyer_header_index = -1
    seller_header_index = -1

    for idx, line in enumerate(lines[:20]):
        if idx >= 4 and tax_line_entries and _is_ocr_tax_id_block_end(line):
            break

        if buyer_header_index < 0 and _contains_ocr_party_hint(line, "buyer_tax_id"):
            buyer_header_index = idx
        if seller_header_index < 0 and _contains_ocr_party_hint(line, "seller_tax_id"):
            seller_header_index = idx
        if "名称" in line:
            name_line_indices.append(idx)

        tax_values = _extract_ocr_tax_id_values_from_line(line)
        if not tax_values:
            continue

        if len(tax_values) >= 2:
            result["buyer_tax_id"] = {
                "value": tax_values[0],
                "confidence": "high",
                "method": "same_line_pair",
                "line_index": idx,
            }
            result["seller_tax_id"] = {
                "value": tax_values[1],
                "confidence": "high",
                "method": "same_line_pair",
                "line_index": idx,
            }
            return result

        tax_line_entries.append((idx, tax_values[0]))

    if len(tax_line_entries) >= 2:
        confidence = "high" if name_line_indices or buyer_header_index >= 0 or seller_header_index >= 0 else "medium"
        buyer_idx, buyer_value = tax_line_entries[0]
        seller_idx, seller_value = tax_line_entries[1]
        result["buyer_tax_id"] = {
            "value": buyer_value,
            "confidence": confidence,
            "method": "ordered_pair",
            "line_index": buyer_idx,
        }
        result["seller_tax_id"] = {
            "value": seller_value,
            "confidence": confidence,
            "method": "ordered_pair",
            "line_index": seller_idx,
        }
        return result

    if len(tax_line_entries) == 1:
        line_index, value = tax_line_entries[0]
        field_name, confidence, method = _infer_single_ocr_tax_id_field(
            lines,
            line_index,
            name_line_indices,
            buyer_header_index,
            seller_header_index,
        )
        if field_name:
            result[field_name] = {
                "value": value,
                "confidence": confidence,
                "method": method,
                "line_index": line_index,
            }
    return result


def _collect_tesseract_tax_id_crop_candidates(image_path: str) -> list[dict]:
    try:
        import pytesseract
        from PIL import Image, ImageFilter, ImageOps
    except ImportError:
        return []

    prepared = prepare_pytesseract()
    base_config = str(prepared.get("config") or "").strip()
    lang = str(prepared.get("lang") or "chi_sim+eng").strip() or "chi_sim+eng"
    candidates: list[dict] = []
    seen_signatures: set[tuple[tuple[str, str], ...]] = set()

    with Image.open(image_path) as image:
        grayscale = image.convert("L")
        width, height = grayscale.size
        for index, spec in enumerate(_OCR_TAX_ID_CROP_SPECS):
            top = int(height * float(spec["top"]))
            bottom = int(height * float(spec["bottom"]))
            if bottom - top < 8:
                continue

            crop = grayscale.crop((0, top, width, bottom))
            try:
                scale = max(1, int(spec.get("scale", 1)))
                if scale > 1:
                    crop = crop.resize((crop.width * scale, crop.height * scale), Image.Resampling.LANCZOS)
                crop = ImageOps.autocontrast(crop)
                crop = crop.filter(ImageFilter.SHARPEN)
                config = _build_image_tesseract_config(f"{base_config} --psm {int(spec.get('psm', 6))}")
                candidate_text = str(
                    pytesseract.image_to_string(
                        crop,
                        lang=lang,
                        config=config,
                    )
                    or ""
                ).strip()
            finally:
                crop.close()

            if not candidate_text:
                continue

            tax_fields = _extract_ocr_tax_id_fields(candidate_text)
            if not tax_fields:
                continue

            signature = tuple(
                (field_name, _normalize_ocr_tax_id((tax_fields.get(field_name) or {}).get("value", "")))
                for field_name in _OCR_TAX_ID_FIELDS
                if _normalize_ocr_tax_id((tax_fields.get(field_name) or {}).get("value", ""))
            )
            if not signature or signature in seen_signatures:
                continue
            seen_signatures.add(signature)

            candidates.append(
                {
                    "engine": "TesseractTaxIDCrop",
                    "engine_key": "tesseract_tax_id_crop",
                    "engine_index": 100 + index,
                    "selected": False,
                    "text": candidate_text,
                    "char_count": len(candidate_text),
                    "ocr_tax_id_fields": tax_fields,
                    "meta": {
                        "crop_name": str(spec.get("name") or f"crop_{index}"),
                        "crop_top": float(spec["top"]),
                        "crop_bottom": float(spec["bottom"]),
                        "scale": scale,
                        "lang": lang,
                        "config": config,
                    },
                }
            )

    return candidates


def _collect_rapidocr_tax_id_field_candidates(image_path: str) -> list[dict]:
    try:
        import tempfile
        from PIL import Image, ImageFilter, ImageOps
    except ImportError:
        return []

    try:
        reader = _get_rapidocr_reader()
    except Exception:
        return []

    candidates: list[dict] = []
    seen_signatures: set[tuple[tuple[str, str], ...]] = set()

    with Image.open(image_path) as image:
        grayscale = ImageOps.autocontrast(image.convert("L"))
        width, height = grayscale.size

        for field_index, (field_name, spec) in enumerate(_OCR_TAX_ID_FIELD_CROP_SPECS.items()):
            left = int(width * float(spec["left"]))
            right = int(width * float(spec["right"]))
            top = int(height * float(spec["top"]))
            bottom = int(height * float(spec["bottom"]))
            if right - left < 8 or bottom - top < 8:
                continue

            base_crop = grayscale.crop((left, top, right, bottom))
            try:
                for variant_index, variant in enumerate(_OCR_TAX_ID_FIELD_CROP_VARIANTS):
                    crop = base_crop.copy()
                    try:
                        if variant.get("sharpen"):
                            crop = crop.filter(ImageFilter.SHARPEN)
                        threshold = variant.get("threshold")
                        if threshold is not None:
                            crop = crop.point(
                                lambda value, limit=int(threshold): 255 if value > limit else 0,
                                mode="1",
                            ).convert("L")
                        scale = max(1, int(variant.get("scale", 1)))
                        if scale > 1:
                            crop = crop.resize(
                                (crop.width * scale, crop.height * scale),
                                Image.Resampling.LANCZOS,
                            )

                        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png", dir=UPLOAD_FOLDER)
                        temp_path = temp_file.name
                        temp_file.close()
                        try:
                            crop.save(temp_path, format="PNG", optimize=True)
                            raw_result = reader(temp_path)
                            candidate_text = _extract_text_from_rapidocr_result(raw_result).strip()
                        finally:
                            try:
                                os.remove(temp_path)
                            except Exception:
                                pass

                        if not candidate_text:
                            continue

                        method = f"rapidocr_field_crop_{field_name}_{variant['name']}"
                        tax_fields = _extract_single_ocr_tax_id_field(candidate_text, field_name, method)
                        if not tax_fields:
                            continue

                        signature = tuple(
                            (name, _normalize_ocr_tax_id((tax_fields.get(name) or {}).get("value", "")))
                            for name in _OCR_TAX_ID_FIELDS
                            if _normalize_ocr_tax_id((tax_fields.get(name) or {}).get("value", ""))
                        )
                        if not signature or signature in seen_signatures:
                            continue
                        seen_signatures.add(signature)

                        candidates.append(
                            {
                                "engine": "RapidOCRTaxIDFieldCrop",
                                "engine_key": "rapidocr_tax_id_field_crop",
                                "engine_index": 120 + field_index * 10 + variant_index,
                                "selected": False,
                                "text": candidate_text,
                                "char_count": len(candidate_text),
                                "ocr_tax_id_fields": tax_fields,
                                "meta": {
                                    "crop_name": method,
                                    "field_name": field_name,
                                    "crop_left": float(spec["left"]),
                                    "crop_right": float(spec["right"]),
                                    "crop_top": float(spec["top"]),
                                    "crop_bottom": float(spec["bottom"]),
                                    "variant": str(variant["name"]),
                                    "scale": scale,
                                },
                            }
                        )
                    finally:
                        crop.close()
            finally:
                base_crop.close()

    return candidates


def _normalize_ocr_tax_id_patch(image) -> Optional["np.ndarray"]:
    try:
        import numpy as np
        from PIL import Image
    except ImportError:
        return None

    patch = image.convert("L")
    array = np.array(patch)
    mask = array < 200
    ys, xs = np.where(mask)
    if len(xs) == 0 or len(ys) == 0:
        return None

    x1, x2 = int(xs.min()), int(xs.max()) + 1
    y1, y2 = int(ys.min()), int(ys.max()) + 1
    normalized = (mask[y1:y2, x1:x2].astype(np.uint8) * 255)
    height, width = normalized.shape
    scale = min(28 / max(height, 1), 28 / max(width, 1))
    resized = Image.fromarray(normalized).resize(
        (max(1, int(round(width * scale))), max(1, int(round(height * scale)))),
        Image.Resampling.NEAREST,
    )
    canvas = Image.new("L", (32, 32), 0)
    canvas.paste(resized, ((32 - resized.width) // 2, (32 - resized.height) // 2))
    return np.array(canvas)


def _parse_tesseract_box_entries(boxes_text: str, image_height: int) -> list[dict]:
    entries: list[dict] = []
    for line in str(boxes_text or "").splitlines():
        parts = line.split()
        if len(parts) < 5:
            continue
        char = str(parts[0] or "").strip()
        if not char:
            continue
        try:
            x1, y1, x2, y2 = map(int, parts[1:5])
        except ValueError:
            continue
        entries.append(
            {
                "char": char.upper(),
                "box": (x1, image_height - y2, x2, image_height - y1),
                "height": max(1, (image_height - y1) - (image_height - y2)),
                "center_y": ((image_height - y2) + (image_height - y1)) / 2,
            }
        )
    return entries


def _extract_tesseract_tax_id_box_field_data(image, boxes_text: str) -> Optional[dict]:
    entries = _parse_tesseract_box_entries(boxes_text, image.height)
    if not entries:
        return None

    min_center_y = min(entry["center_y"] for entry in entries)
    median_height = statistics.median(entry["height"] for entry in entries) if entries else 0
    line_threshold = max(24, int(median_height * 0.75))
    line_entries = [
        entry
        for entry in entries
        if abs(entry["center_y"] - min_center_y) <= line_threshold
    ]
    if not line_entries:
        return None
    line_entries.sort(key=lambda entry: entry["box"][0])

    line_text = "".join(entry["char"] for entry in line_entries)
    colon_index = line_text.find(":")
    if colon_index < 0:
        colon_index = line_text.find("：")
    if colon_index >= 0:
        candidate_entries = line_entries[colon_index + 1 :]
    else:
        first_alnum = next((idx for idx, entry in enumerate(line_entries) if entry["char"].isalnum()), -1)
        candidate_entries = line_entries[first_alnum:] if first_alnum >= 0 else []
    candidate_entries = [entry for entry in candidate_entries if entry["char"].isalnum()]
    if not candidate_entries:
        return None

    return {
        "value": "".join(entry["char"] for entry in candidate_entries),
        "chars": [
            {
                "char": entry["char"],
                "patch": _normalize_ocr_tax_id_patch(image.crop(entry["box"])),
            }
            for entry in candidate_entries
        ],
    }


def _collect_tesseract_tax_id_box_references(image_path: str) -> dict[str, str]:
    try:
        from PIL import Image, ImageFilter, ImageOps
        import numpy as np
        import pytesseract
    except ImportError:
        return {}

    references: list[dict] = []
    field_data: dict[str, dict] = {}

    with Image.open(image_path) as source:
        grayscale = ImageOps.autocontrast(source.convert("L"))
        width, height = grayscale.size

        for field_name, spec in _OCR_TAX_ID_FIELD_CROP_SPECS.items():
            crop = grayscale.crop(
                (
                    int(width * float(spec["left"])),
                    int(height * float(spec["top"])),
                    int(width * float(spec["right"])),
                    int(height * float(spec["bottom"])),
                )
            )
            try:
                crop = crop.filter(ImageFilter.SHARPEN)
                crop = crop.resize((crop.width * 4, crop.height * 4), Image.Resampling.LANCZOS)
                boxes_text = pytesseract.image_to_boxes(crop, lang="chi_sim+eng", config="--psm 6")
                field_info = _extract_tesseract_tax_id_box_field_data(crop, boxes_text)
                if not field_info:
                    continue
                chars = field_info.get("chars") or []
                if not chars:
                    continue
                field_data[field_name] = field_info
                for index, char_info in enumerate(chars):
                    if char_info.get("char") in {"O", "0"} and char_info.get("patch") is not None:
                        references.append(
                            {
                                "field_name": field_name,
                                "index": index,
                                "char": char_info["char"],
                                "patch": char_info["patch"],
                            }
                        )
            finally:
                crop.close()

    refined_values: dict[str, str] = {}
    for field_name, info in field_data.items():
        chars = list(info.get("chars") or [])
        if not chars:
            continue

        refined_chars: list[str] = []
        for index, char_info in enumerate(chars):
            current_char = str(char_info.get("char") or "").upper()
            patch = char_info.get("patch")
            if current_char not in {"O", "0"} or patch is None:
                refined_chars.append(current_char)
                continue

            same_refs = [
                ref["patch"]
                for ref in references
                if ref["field_name"] != field_name or ref["index"] != index
                if ref["char"] == current_char
            ]
            other_char = "0" if current_char == "O" else "O"
            other_refs = [
                ref["patch"]
                for ref in references
                if ref["field_name"] != field_name or ref["index"] != index
                if ref["char"] == other_char
            ]
            if not same_refs or not other_refs:
                refined_chars.append(current_char)
                continue

            same_distance = min(float(np.abs(patch.astype(float) - reference.astype(float)).mean()) for reference in same_refs)
            other_distance = min(float(np.abs(patch.astype(float) - reference.astype(float)).mean()) for reference in other_refs)
            if other_distance + 4.0 < same_distance:
                refined_chars.append(other_char)
            else:
                refined_chars.append(current_char)

        refined_value = _normalize_ocr_tax_id("".join(refined_chars))
        if refined_value:
            refined_values[field_name] = refined_value

    return refined_values


def _apply_ocr_tax_id_reference_chars(candidate_value: str, reference_value: str) -> str:
    candidate = _normalize_ocr_tax_id(candidate_value)
    reference = _normalize_ocr_tax_id(reference_value)
    if not candidate or not reference:
        return candidate

    merged: list[str] = []
    for tag, i1, i2, j1, j2 in SequenceMatcher(None, candidate, reference).get_opcodes():
        candidate_segment = candidate[i1:i2]
        reference_segment = reference[j1:j2]

        if tag == "equal":
            merged.append(candidate_segment)
            continue

        if tag == "replace" and len(candidate_segment) == len(reference_segment):
            merged.append(
                "".join(
                    reference_char
                    if candidate_char == "O" and reference_char == "0"
                    else candidate_char
                    for candidate_char, reference_char in zip(candidate_segment, reference_segment)
                )
            )
            continue

        merged.append(candidate_segment)

    refined = _normalize_ocr_tax_id("".join(merged))
    return refined or candidate


def _apply_ocr_tax_id_box_references(candidates: list[dict], reference_values: dict[str, str]) -> None:
    if not reference_values:
        return

    for candidate in candidates:
        tax_fields = candidate.get("ocr_tax_id_fields")
        if not isinstance(tax_fields, dict):
            continue
        for field_name, source in list(tax_fields.items()):
            if not isinstance(source, dict):
                continue
            reference_value = reference_values.get(field_name)
            if not reference_value:
                continue
            current_value = _normalize_ocr_tax_id(source.get("value", ""))
            if not current_value:
                continue
            refined_value = _apply_ocr_tax_id_reference_chars(current_value, reference_value)
            if not refined_value or refined_value == current_value:
                continue
            source["value"] = refined_value
            source["method"] = f"{str(source.get('method') or '').strip('+')}+box_ref".strip("+")


def _summarize_ocr_tax_id_crop_candidates(candidates: list[dict]) -> list[dict]:
    summary: list[dict] = []
    for candidate in candidates:
        tax_fields = candidate.get("ocr_tax_id_fields") or {}
        meta = candidate.get("meta") or {}
        entry = {
            "engine": str(candidate.get("engine") or ""),
            "crop_name": str(meta.get("crop_name") or ""),
            "crop_top": meta.get("crop_top"),
            "crop_bottom": meta.get("crop_bottom"),
            "fields": {},
        }
        for field_name in _OCR_TAX_ID_FIELDS:
            value = _normalize_ocr_tax_id((tax_fields.get(field_name) or {}).get("value", ""))
            if value:
                entry["fields"][field_name] = value
        summary.append(entry)
    return summary


def _merge_ocr_invoice_fields(selected_candidate: Optional[dict], ranked_candidates: list[dict]) -> tuple[dict, dict]:
    if not isinstance(selected_candidate, dict):
        return _empty_invoice_fields_payload(), {"applied": False, "fields": {}}

    selected_invoice_eval = selected_candidate.get("invoice_eval") or {}
    base_invoice_fields = selected_invoice_eval.get("invoice_fields")
    if isinstance(base_invoice_fields, dict):
        merged_invoice_fields = _clone_json_payload(base_invoice_fields)
    else:
        merged_invoice_fields = _empty_invoice_fields_payload()

    fields = merged_invoice_fields.setdefault("fields", {})
    if not isinstance(fields, dict):
        fields = {}
        merged_invoice_fields["fields"] = fields

    merge_meta = {"applied": False, "fields": {}}
    selected_structured = selected_candidate.get("ocr_tax_id_fields")
    if not isinstance(selected_structured, dict):
        selected_structured = {}

    for field_name, source in selected_structured.items():
        normalized_value = _normalize_ocr_tax_id(source.get("value", ""))
        if not normalized_value:
            continue
        previous_value = _normalize_ocr_tax_id((fields.get(field_name) or {}).get("value", ""))
        fields[field_name] = {
            "value": normalized_value,
            "confidence": str(source.get("confidence") or "high"),
            "label": FIELD_LABELS.get(field_name, field_name),
        }
        if previous_value != normalized_value:
            merge_meta["applied"] = True
            merge_meta["fields"][field_name] = {
                "action": "selected_structured",
                "engine": str(selected_candidate.get("engine") or ""),
                "method": str(source.get("method") or ""),
                "value": normalized_value,
            }

    buyer_value = _normalize_ocr_tax_id((fields.get("buyer_tax_id") or {}).get("value", ""))
    seller_value = _normalize_ocr_tax_id((fields.get("seller_tax_id") or {}).get("value", ""))
    if buyer_value and seller_value and buyer_value == seller_value and selected_structured:
        for field_name in _OCR_TAX_ID_FIELDS:
            if field_name in selected_structured:
                continue
            if field_name in fields:
                fields.pop(field_name, None)
                merge_meta["applied"] = True
                merge_meta["fields"][field_name] = {
                    "action": "cleared_duplicate",
                    "engine": str(selected_candidate.get("engine") or ""),
                    "method": "duplicate_tax_id",
                    "value": buyer_value,
                }

    for field_name in _OCR_TAX_ID_FIELDS:
        current_value = _normalize_ocr_tax_id((fields.get(field_name) or {}).get("value", ""))
        if current_value:
            continue

        other_field = "seller_tax_id" if field_name == "buyer_tax_id" else "buyer_tax_id"
        other_value = _normalize_ocr_tax_id((fields.get(other_field) or {}).get("value", ""))

        for candidate in ranked_candidates:
            candidate_tax_fields = candidate.get("ocr_tax_id_fields")
            if not isinstance(candidate_tax_fields, dict):
                continue
            source = candidate_tax_fields.get(field_name)
            if not isinstance(source, dict):
                continue

            normalized_value = _normalize_ocr_tax_id(source.get("value", ""))
            if not normalized_value:
                continue
            if other_value and normalized_value == other_value:
                continue

            fields[field_name] = {
                "value": normalized_value,
                "confidence": str(source.get("confidence") or "high"),
                "label": FIELD_LABELS.get(field_name, field_name),
            }
            merge_meta["applied"] = True
            merge_meta["fields"][field_name] = {
                "action": "filled_from_other_engine",
                "engine": str(candidate.get("engine") or ""),
                "method": str(source.get("method") or ""),
                "value": normalized_value,
            }
            break

    for field_name in _OCR_TAX_ID_FIELDS:
        current_field = fields.get(field_name) or {}
        current_value = _normalize_ocr_tax_id(current_field.get("value", ""))
        if not current_value or len(current_value) >= _OCR_TAX_ID_EXPECTED_LENGTH:
            continue

        other_field = "seller_tax_id" if field_name == "buyer_tax_id" else "buyer_tax_id"
        other_value = _normalize_ocr_tax_id((fields.get(other_field) or {}).get("value", ""))
        best_value = current_value
        best_meta = None

        for candidate in ranked_candidates:
            if candidate is selected_candidate:
                continue
            candidate_tax_fields = candidate.get("ocr_tax_id_fields")
            if not isinstance(candidate_tax_fields, dict):
                continue

            source = candidate_tax_fields.get(field_name)
            if not isinstance(source, dict):
                continue

            normalized_value = _normalize_ocr_tax_id(source.get("value", ""))
            if not normalized_value:
                continue
            if other_value and normalized_value == other_value:
                continue

            repaired_value = _repair_ocr_tax_id_value(current_value, normalized_value)
            if repaired_value == current_value or repaired_value == other_value:
                continue
            if abs(len(repaired_value) - _OCR_TAX_ID_EXPECTED_LENGTH) > abs(len(best_value) - _OCR_TAX_ID_EXPECTED_LENGTH):
                continue
            if len(repaired_value) <= len(best_value) and abs(len(repaired_value) - _OCR_TAX_ID_EXPECTED_LENGTH) == abs(len(best_value) - _OCR_TAX_ID_EXPECTED_LENGTH):
                continue

            best_value = repaired_value
            best_meta = {
                "action": "repaired_from_other_engine",
                "engine": str(candidate.get("engine") or ""),
                "method": str(source.get("method") or ""),
                "value": repaired_value,
                "source_value": normalized_value,
            }

        if best_meta:
            fields[field_name] = {
                "value": best_value,
                "confidence": str(current_field.get("confidence") or "high"),
                "label": FIELD_LABELS.get(field_name, field_name),
            }
            merge_meta["applied"] = True
            merge_meta["fields"][field_name] = best_meta

    merged_invoice_fields["field_count"] = len(fields)
    merged_invoice_fields["is_invoice"] = bool(merged_invoice_fields.get("is_invoice") or fields)
    if merged_invoice_fields["is_invoice"] and str(merged_invoice_fields.get("confidence") or "none") == "none":
        merged_invoice_fields["confidence"] = "medium"
    return merged_invoice_fields, merge_meta


def _format_ocr_field_merge_summary(merge_meta: dict) -> str:
    if not isinstance(merge_meta, dict):
        return ""
    fields = merge_meta.get("fields")
    if not isinstance(fields, dict) or not fields:
        return ""

    parts = []
    for field_name in _OCR_TAX_ID_FIELDS:
        field_meta = fields.get(field_name)
        if not isinstance(field_meta, dict):
            continue
        action = str(field_meta.get("action") or "").strip()
        engine = str(field_meta.get("engine") or "").strip()
        method = str(field_meta.get("method") or "").strip()
        part = f"{field_name}:{action}"
        if engine:
            part += f"@{engine}"
        if method:
            part += f"/{method}"
        parts.append(part)
    return ", ".join(parts)


def _ocr_candidate_sort_key(candidate: dict) -> tuple:
    invoice_eval = candidate.get("invoice_eval") or {}
    return (
        1 if invoice_eval.get("is_invoice") else 0,
        int(invoice_eval.get("score") or 0),
        _rank_confidence(invoice_eval.get("confidence", "none")),
        int(invoice_eval.get("field_count") or 0),
        int(invoice_eval.get("major_field_count") or 0),
        int(candidate.get("char_count") or 0),
        -int(candidate.get("engine_index") or 0),
    )


def _format_ocr_candidate_label(candidate: dict) -> str:
    engine = str(candidate.get("engine") or "").strip()
    meta = candidate.get("meta") or {}
    variant = str(meta.get("preprocess_variant") or meta.get("variant_name") or "").strip()
    if engine and variant:
        return f"{engine}[{variant}]"
    return engine


def _summarize_ocr_candidates(candidates: list[dict]) -> list[dict]:
    summaries = []
    for candidate in candidates:
        invoice_eval = candidate.get("invoice_eval") or {}
        meta = candidate.get("meta") or {}
        summaries.append(
            {
                "engine": candidate.get("engine", ""),
                "engine_display": _format_ocr_candidate_label(candidate),
                "engine_key": candidate.get("engine_key", ""),
                "selected": bool(candidate.get("selected")),
                "char_count": int(candidate.get("char_count") or 0),
                "score": int(invoice_eval.get("score") or 0),
                "is_invoice": bool(invoice_eval.get("is_invoice")),
                "confidence": str(invoice_eval.get("confidence") or "none"),
                "field_count": int(invoice_eval.get("field_count") or 0),
                "field_names": list(invoice_eval.get("field_names") or []),
                "preprocess_variant": str(meta.get("preprocess_variant") or meta.get("variant_name") or ""),
            }
        )
    return summaries


def _format_ocr_selection_summary(candidates: list[dict], selected_engine: str) -> str:
    if not candidates:
        return ""
    parts = []
    for candidate in candidates:
        invoice_eval = candidate.get("invoice_eval") or {}
        parts.append(
            f"{_format_ocr_candidate_label(candidate)}={int(invoice_eval.get('score') or 0)}"
            f"/{int(invoice_eval.get('field_count') or 0)}"
        )
    summary = " > ".join(parts)
    if selected_engine:
        summary += f" | selected={selected_engine}"
    return summary

def extract_invoice_fields(text: str) -> dict:
    """? OCR ??????????"""
    return _get_invoice_extractor().extract(text)
