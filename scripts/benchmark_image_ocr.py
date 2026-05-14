"""Benchmark image OCR to break down where time is spent.

Instruments `_run_ocr_variant_tasks`, engine dispatch, and invoice refinement
to attribute every millisecond of `process_image_ocr` to a specific stage.
Runs on three sample invoice images, both cold (first call, includes warmup)
and warm (cache disabled, repeated calls).
"""

from __future__ import annotations

import os
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Optional

import _bootstrap  # noqa: F401

_bootstrap.ensure_project_root_on_path()

os.environ.setdefault("DOCFLOW_ENABLE_IMAGE_OCR_CACHE", "0")

from docflow.webapp.services import ocr as ocr_module  # noqa: E402
from docflow.webapp.services import ocr_engines as engines_module  # noqa: E402
from docflow.webapp.services import invoice_merge as invoice_merge_module  # noqa: E402


SAMPLE_DIR = Path("发票样本/invoices/images")
SAMPLE_FILES = [
    "invoice_electronic_003.png",
    "invoice_electronic_007.png",
    "invoice_electronic_011.png",
]


_timings: dict = {}


def _reset_timings() -> None:
    _timings.clear()
    _timings["engines"] = {}
    _timings["variants"] = []
    _timings["refinements"] = []
    _timings["image_load_total_ms"] = 0.0


def _install_instrumentation() -> None:
    original_variant_tasks = engines_module._run_ocr_variant_tasks
    original_load_base = engines_module._load_base_image_for_ocr
    original_rapid_collect = invoice_merge_module._collect_rapidocr_tax_id_field_candidates

    def timed_load_base(image_path: str, provider: str):
        start = time.perf_counter()
        try:
            return original_load_base(image_path, provider)
        finally:
            elapsed = (time.perf_counter() - start) * 1000
            _timings["image_load_total_ms"] += elapsed

    def timed_variant_tasks(provider, variant_entries, worker, close_image: bool = True):
        def wrapped_worker(variant_index, entry):
            variant_name = entry[0]
            start = time.perf_counter()
            try:
                return worker(variant_index, entry)
            finally:
                elapsed = (time.perf_counter() - start) * 1000
                _timings["variants"].append(
                    {
                        "provider": provider,
                        "variant_index": variant_index,
                        "variant_name": variant_name,
                        "elapsed_ms": elapsed,
                    }
                )

        return original_variant_tasks(provider, variant_entries, wrapped_worker, close_image=close_image)

    def timed_rapid_collect(image_path: str):
        start = time.perf_counter()
        try:
            return original_rapid_collect(image_path)
        finally:
            elapsed = (time.perf_counter() - start) * 1000
            _timings["refinements"].append(
                {"name": "rapidocr_tax_id_fields", "elapsed_ms": elapsed}
            )

    original_process = ocr_module.process_image_ocr

    def timed_process_image_ocr(image_path: str, filename: str, **kwargs):
        start = time.perf_counter()
        _reset_timings()
        try:
            return original_process(image_path, filename, **kwargs)
        finally:
            _timings["total_ms"] = (time.perf_counter() - start) * 1000

    engines_module._run_ocr_variant_tasks = timed_variant_tasks
    engines_module._load_base_image_for_ocr = timed_load_base
    invoice_merge_module._collect_rapidocr_tax_id_field_candidates = timed_rapid_collect
    # ocr.py imports this name at module load, so patch the local binding too.
    ocr_module._collect_rapidocr_tax_id_field_candidates = timed_rapid_collect

    original_tess_crop = invoice_merge_module._collect_tesseract_tax_id_crop_candidates
    original_tess_box = invoice_merge_module._collect_tesseract_tax_id_box_references

    def timed_tess_crop(image_path: str):
        start = time.perf_counter()
        try:
            return original_tess_crop(image_path)
        finally:
            _timings["refinements"].append(
                {"name": "tesseract_tax_id_crop", "elapsed_ms": (time.perf_counter() - start) * 1000}
            )

    def timed_tess_box(image_path: str):
        start = time.perf_counter()
        try:
            return original_tess_box(image_path)
        finally:
            _timings["refinements"].append(
                {"name": "tesseract_tax_id_box", "elapsed_ms": (time.perf_counter() - start) * 1000}
            )

    ocr_module._collect_tesseract_tax_id_crop_candidates = timed_tess_crop
    ocr_module._collect_tesseract_tax_id_box_references = timed_tess_box

    # Also time the candidate scoring / merge step to attribute the
    # remaining "overhead" in process_image_ocr.
    original_merge = invoice_merge_module._merge_ocr_invoice_fields

    def timed_merge(*args, **kwargs):
        start = time.perf_counter()
        try:
            return original_merge(*args, **kwargs)
        finally:
            _timings["refinements"].append(
                {"name": "merge_invoice_fields", "elapsed_ms": (time.perf_counter() - start) * 1000}
            )

    ocr_module._merge_ocr_invoice_fields = timed_merge

    ocr_module.process_image_ocr = timed_process_image_ocr


def _summarize_run(filename: str, result: dict) -> dict:
    metadata = result.get("metadata", {})
    candidates = metadata.get("ocr_candidates") or []
    engine_used = metadata.get("engine", "")
    refinement_applied = metadata.get("ocr_invoice_refinement_applied")
    refinement_reason = metadata.get("ocr_invoice_refinement_reason", "")

    by_provider: dict[str, dict] = {}
    for entry in _timings["variants"]:
        slot = by_provider.setdefault(
            entry["provider"],
            {"total_ms": 0.0, "variants": []},
        )
        slot["total_ms"] += entry["elapsed_ms"]
        slot["variants"].append(entry)

    refinement_total = sum(item["elapsed_ms"] for item in _timings["refinements"])

    return {
        "filename": filename,
        "total_ms": _timings["total_ms"],
        "engine_used": engine_used,
        "candidates": len(candidates),
        "image_load_total_ms": _timings["image_load_total_ms"],
        "by_provider": by_provider,
        "refinement_total_ms": refinement_total,
        "refinement_items": list(_timings["refinements"]),
        "refinement_applied": refinement_applied,
        "refinement_reason": refinement_reason,
    }


def _print_run(label: str, summary: dict) -> None:
    print(f"\n=== {label} :: {summary['filename']} ===")
    print(f"total: {summary['total_ms']:.0f} ms   engine_used: {summary['engine_used']}   candidates: {summary['candidates']}")
    print(f"image_load (sum of all engines): {summary['image_load_total_ms']:.0f} ms")
    for provider, slot in summary["by_provider"].items():
        print(f"  [{provider}] total {slot['total_ms']:.0f} ms across {len(slot['variants'])} variant(s)")
        for v in slot["variants"]:
            print(f"      - {v['variant_name']:<20} {v['elapsed_ms']:>7.0f} ms")
    if summary["refinement_total_ms"]:
        print(f"  refinement+merge: {summary['refinement_total_ms']:.0f} ms total  (applied={summary['refinement_applied']} reason={summary['refinement_reason']})")
        for item in summary.get("refinement_items", []):
            print(f"      - {item['name']:<25} {item['elapsed_ms']:>7.0f} ms")
    accounted = summary["image_load_total_ms"] + sum(
        s["total_ms"] for s in summary["by_provider"].values()
    ) + summary["refinement_total_ms"]
    other = summary["total_ms"] - accounted
    print(f"  unaccounted overhead: {other:.0f} ms (scoring, IO, scaffolding)")


def _run_one(image_path: Path, label: str) -> dict:
    result = ocr_module.process_image_ocr(
        str(image_path),
        image_path.name,
        force_reprocess=True,
    )
    summary = _summarize_run(image_path.name, result)
    _print_run(label, summary)
    return summary


def _run_parallel(paths: list[Path]) -> tuple[float, list[dict]]:
    """Run all images in parallel threads, measure wall time."""

    print("\n=== PARALLEL (3 threads) ===")
    start = time.perf_counter()
    results = []

    def worker(p: Path):
        # The instrumentation uses a single shared _timings dict, so
        # per-stage breakdown is unreliable here. We only care about
        # wall-clock total.
        return ocr_module.process_image_ocr(str(p), p.name, force_reprocess=True)

    with ThreadPoolExecutor(max_workers=len(paths)) as executor:
        futures = [executor.submit(worker, p) for p in paths]
        for f in futures:
            results.append(f.result())

    wall_ms = (time.perf_counter() - start) * 1000
    print(f"wall time for 3 parallel images: {wall_ms:.0f} ms")
    return wall_ms, results


def main() -> None:
    _install_instrumentation()

    paths = [SAMPLE_DIR / name for name in SAMPLE_FILES]
    for p in paths:
        if not p.exists():
            raise SystemExit(f"sample image not found: {p}")

    print("Engine order:", os.environ.get("DOCFLOW_IMAGE_OCR_ORDER", "(default: rapidocr,tesseract)"))
    print("Speed profile:", os.environ.get("DOCFLOW_IMAGE_OCR_SPEED_PROFILE", "(default: fast)"))
    print("Parallel engines:", os.environ.get("DOCFLOW_IMAGE_OCR_PARALLEL_ENGINES", "(default: 1)"))
    print("Cache:", os.environ.get("DOCFLOW_ENABLE_IMAGE_OCR_CACHE"))

    # COLD: first call includes RapidOCR model load.
    cold = _run_one(paths[0], "COLD (first call, includes model load)")

    # WARM: subsequent calls — steady state.
    warm_summaries = []
    serial_start = time.perf_counter()
    for p in paths:
        warm_summaries.append(_run_one(p, "WARM"))
    serial_wall = (time.perf_counter() - serial_start) * 1000
    print(f"\n=== SERIAL TOTAL (3 warm images, one at a time): {serial_wall:.0f} ms ===")

    parallel_wall, _ = _run_parallel(paths)

    print("\n=== SUMMARY ===")
    print(f"Cold first-call: {cold['total_ms']:.0f} ms (warmup baked in)")
    avg_warm = sum(s["total_ms"] for s in warm_summaries) / len(warm_summaries)
    print(f"Warm avg per image: {avg_warm:.0f} ms")
    print(f"Serial 3 warm images:   {serial_wall:.0f} ms")
    print(f"Parallel 3 warm images: {parallel_wall:.0f} ms  (speedup {serial_wall / max(parallel_wall, 1):.2f}x)")


if __name__ == "__main__":
    main()
