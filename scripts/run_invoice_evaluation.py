from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT = SCRIPT_DIR.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app import IMAGE_EXTS, process_image_ocr
from docflow_core import DocFlowProcessor
from docflow_support import build_error_info
from invoice_extractor import InvoiceExtractor

DEFAULT_MANIFEST = ROOT / "发票样本" / "invoice_test_manifest.csv"
DEFAULT_REPORT_ROOT = ROOT / "reports"
DEFAULT_DATASET_ROOT = ROOT / "发票样本" / "InvoiceDatasets-master" / "dataset" / "images"
FIELD_NAMES = [
    "invoice_code",
    "invoice_number",
    "invoice_date",
    "amount",
    "tax",
    "total",
    "buyer_name",
    "seller_name",
    "buyer_tax_id",
    "seller_tax_id",
    "invoice_type",
]

MONEY_FIELDS = {"amount", "tax", "total"}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="run_invoice_evaluation",
        description="运行发票测试集评测并生成结构化报告。",
    )
    parser.add_argument(
        "manifest",
        nargs="?",
        default=str(DEFAULT_MANIFEST),
        help="发票测试清单 CSV 路径。",
    )
    parser.add_argument(
        "--dataset-root",
        default=str(DEFAULT_DATASET_ROOT),
        help="图片数据集根目录，默认指向 InvoiceDatasets-master/dataset/images。",
    )
    parser.add_argument(
        "--report-root",
        default=str(DEFAULT_REPORT_ROOT),
        help="评测报告输出目录，默认 reports/。",
    )
    parser.add_argument(
        "--sample-type",
        action="append",
        default=[],
        help="仅评测指定 sample_type，可重复传入多个。",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="仅运行前 N 个样本，用于快速验证。",
    )
    return parser.parse_args()


def load_manifest(manifest_path: Path, sample_types: list[str], limit: int) -> list[dict[str, str]]:
    with manifest_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        rows = [dict(row) for row in reader]

    if sample_types:
        allowed = {item.strip() for item in sample_types if item.strip()}
        rows = [row for row in rows if row.get("sample_type", "").strip() in allowed]

    if limit and limit > 0:
        rows = rows[:limit]
    return rows


def _parse_bool(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    text = str(value).strip().lower()
    if not text:
        return default
    return text in {"1", "true", "yes", "y", "on"}


def _collapse_spaces(value: Any) -> str:
    return " ".join(str(value or "").strip().split())


def _normalize_text(value: Any) -> str:
    return _collapse_spaces(value).replace("：", ":")


def _normalize_identity(value: Any) -> str:
    text = _normalize_text(value).upper()
    return re.sub(r"[^A-Z0-9]", "", text)


def _normalize_amount(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, Decimal):
        amount = value
    elif isinstance(value, (int, float)):
        amount = Decimal(str(value))
    else:
        text = _normalize_text(value)
        if not text:
            return ""
        text = (
            text.replace(",", "")
            .replace("￥", "")
            .replace("¥", "")
            .replace("元", "")
            .replace("圆", "")
        )
        match = re.search(r"-?\d+(?:\.\d+)?", text)
        if not match:
            return ""
        try:
            amount = Decimal(match.group(0))
        except InvalidOperation:
            return ""

    quantized = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return format(quantized, "f")


def _normalize_date(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()

    text = _normalize_text(value)
    if not text:
        return ""

    direct_match = re.search(r"((?:19|20)\d{2})[-/.年]\s*(\d{1,2})[-/.月]\s*(\d{1,2})", text)
    if direct_match:
        year, month, day = direct_match.groups()
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"

    digits_match = re.search(r"\b((?:19|20)\d{2})(\d{2})(\d{2})\b", text)
    if digits_match:
        year, month, day = digits_match.groups()
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"

    return text


def normalize_field(field_name: str, value: Any) -> str:
    if field_name in MONEY_FIELDS:
        return _normalize_amount(value)
    if field_name == "invoice_date":
        return _normalize_date(value)
    if field_name in {"invoice_code", "invoice_number", "buyer_tax_id", "seller_tax_id"}:
        return _normalize_identity(value)
    return _normalize_text(value)


def _pct(numerator: int, denominator: int) -> float | None:
    if denominator <= 0:
        return None
    return round((numerator / denominator) * 100, 2)


def _fmt_pct(value: float | None) -> str:
    if value is None:
        return "N/A"
    return f"{value:.2f}%"


def _fmt_bool(value: Any) -> str:
    if value is True:
        return "true"
    if value is False:
        return "false"
    return ""


def resolve_file_path(row: dict[str, str], manifest_path: Path, dataset_root: Path) -> Path:
    file_name = str(row.get("file_name", "")).strip()
    source_dir = str(row.get("source_dir", "")).strip()
    if not file_name:
        return Path()

    path_candidates: list[Path] = []
    if source_dir:
        rel_dir = Path(source_dir)
        path_candidates.extend(
            [
                dataset_root / rel_dir / file_name,
                manifest_path.parent / rel_dir / file_name,
                ROOT / rel_dir / file_name,
            ]
        )
    else:
        path_candidates.extend(
            [
                dataset_root / file_name,
                manifest_path.parent / file_name,
                ROOT / file_name,
            ]
        )

    for candidate in path_candidates:
        if candidate.exists():
            return candidate
    return path_candidates[0]


def process_sample(
    processor: DocFlowProcessor,
    extractor: InvoiceExtractor,
    file_path: Path,
) -> dict[str, Any]:
    if not file_path.exists():
        return {
            "success": False,
            "file": file_path.name,
            "format": file_path.suffix.lstrip("."),
            "text": "",
            "tables": [],
            "metadata": {},
            "statistics": {},
            "processing_ms": 0.0,
            "error": f"文件不存在: {file_path}",
        }

    if file_path.suffix.lower() in IMAGE_EXTS:
        result = process_image_ocr(str(file_path), file_path.name)
        if result.get("success") and result.get("text"):
            result.setdefault("statistics", {})["invoice_fields"] = extractor.extract(result.get("text", ""))
        return result

    return processor.process(
        str(file_path),
        extract_keywords=False,
        extract_invoice=True,
        output_format="txt",
    )


def _build_failure_reason(
    file_exists: bool,
    raw_success: bool,
    invoice_match: bool,
    mismatch_fields: list[str],
    result_error: str,
    error_info: dict[str, Any],
) -> str:
    if not file_exists:
        return "样本文件不存在"
    if result_error:
        return result_error
    if not raw_success:
        return (error_info or {}).get("message") or "处理失败"
    if not invoice_match:
        return "发票判定不匹配"
    if mismatch_fields:
        return f"字段未命中: {', '.join(mismatch_fields[:5])}"
    return ""


def evaluate_row(
    processor: DocFlowProcessor,
    extractor: InvoiceExtractor,
    row: dict[str, str],
    manifest_path: Path,
    dataset_root: Path,
) -> dict[str, Any]:
    file_path = resolve_file_path(row, manifest_path, dataset_root)
    result = process_sample(processor, extractor, file_path)
    error_info = build_error_info(
        result.get("error", ""),
        file_name=file_path.name,
        file_ext=file_path.suffix.lower(),
        metadata_dict=result.get("metadata") or {},
        source="invoice_evaluation",
    )

    invoice_block = (result.get("statistics") or {}).get("invoice_fields") or {}
    fields = invoice_block.get("fields") or {}
    expected_invoice = _parse_bool(row.get("expected_invoice"), default=True)
    actual_invoice = bool(invoice_block.get("is_invoice"))
    invoice_match = expected_invoice == actual_invoice

    field_results: dict[str, dict[str, Any]] = {}
    checked_fields = 0
    covered_fields = 0
    matched_fields = 0
    mismatch_fields: list[str] = []

    for field_name in FIELD_NAMES:
        expected = normalize_field(field_name, row.get(field_name, ""))
        actual = normalize_field(field_name, (fields.get(field_name) or {}).get("value", ""))
        checked = bool(expected)
        predicted = bool(actual)
        matched = (expected == actual) if checked else None

        if checked:
            checked_fields += 1
            if predicted:
                covered_fields += 1
            if matched:
                matched_fields += 1
            else:
                mismatch_fields.append(field_name)

        field_results[field_name] = {
            "expected": expected,
            "actual": actual,
            "checked": checked,
            "predicted": predicted,
            "matched": matched,
        }

    has_field_labels = checked_fields > 0
    sample_passed = invoice_match and matched_fields == checked_fields
    failure_reason = _build_failure_reason(
        file_path.exists(),
        bool(result.get("success")),
        invoice_match,
        mismatch_fields,
        str(result.get("error") or ""),
        error_info,
    )

    return {
        "sample_id": row.get("sample_id", ""),
        "file_name": row.get("file_name", ""),
        "source_dir": row.get("source_dir", ""),
        "sample_type": row.get("sample_type", ""),
        "quality_level": row.get("quality_level", ""),
        "notes": row.get("notes", ""),
        "path": str(file_path),
        "file_exists": file_path.exists(),
        "expected_invoice": expected_invoice,
        "actual_invoice": actual_invoice,
        "invoice_match": invoice_match,
        "processing_ms": round(float(result.get("processing_ms", 0.0)), 2),
        "raw_success": bool(result.get("success")),
        "field_count": int(invoice_block.get("field_count", 0) or 0),
        "confidence": str(invoice_block.get("confidence", "none")),
        "has_field_labels": has_field_labels,
        "checked_fields": checked_fields,
        "covered_fields": covered_fields,
        "matched_fields": matched_fields,
        "field_accuracy_pct": _pct(matched_fields, checked_fields),
        "field_coverage_pct": _pct(covered_fields, checked_fields),
        "field_results": field_results,
        "mismatch_fields": mismatch_fields,
        "error": result.get("error", ""),
        "error_info": error_info,
        "sample_passed": sample_passed,
        "failure_reason": failure_reason,
    }


def _summarize_group(records: list[dict[str, Any]]) -> dict[str, Any]:
    total = len(records)
    labeled_samples = sum(1 for item in records if item["has_field_labels"])
    labeled_sample_passed = sum(1 for item in records if item["has_field_labels"] and item["sample_passed"])
    checked_fields = sum(int(item.get("checked_fields", 0)) for item in records)
    covered_fields = sum(int(item.get("covered_fields", 0)) for item in records)
    matched_fields = sum(int(item.get("matched_fields", 0)) for item in records)
    sample_passed = sum(1 for item in records if item["sample_passed"])
    invoice_match = sum(1 for item in records if item["invoice_match"])

    return {
        "total": total,
        "labeled_samples": labeled_samples,
        "unlabeled_samples": total - labeled_samples,
        "sample_passed": sample_passed,
        "sample_pass_rate_pct": _pct(sample_passed, total),
        "invoice_match": invoice_match,
        "invoice_match_rate_pct": _pct(invoice_match, total),
        "labeled_sample_passed": labeled_sample_passed,
        "labeled_sample_pass_rate_pct": _pct(labeled_sample_passed, labeled_samples),
        "labeled_field_total": checked_fields,
        "covered_field_total": covered_fields,
        "matched_field_total": matched_fields,
        "field_coverage_pct": _pct(covered_fields, checked_fields),
        "field_accuracy_pct": _pct(matched_fields, checked_fields),
        "avg_processing_ms": round(
            sum(float(item.get("processing_ms", 0.0)) for item in records) / total,
            2,
        ) if total else 0.0,
    }


def build_summary(items: list[dict[str, Any]], manifest_path: Path) -> dict[str, Any]:
    total = len(items)
    invoice_match_count = sum(1 for item in items if item["invoice_match"])
    sample_passed = sum(1 for item in items if item["sample_passed"])
    labeled_samples = sum(1 for item in items if item["has_field_labels"])
    labeled_sample_passed = sum(1 for item in items if item["has_field_labels"] and item["sample_passed"])
    labeled_field_total = sum(int(item.get("checked_fields", 0)) for item in items)
    covered_field_total = sum(int(item.get("covered_fields", 0)) for item in items)
    matched_field_total = sum(int(item.get("matched_fields", 0)) for item in items)
    avg_processing_ms = round(
        sum(float(item.get("processing_ms", 0.0)) for item in items) / total,
        2,
    ) if total else 0.0

    field_checked_counter = Counter()
    field_covered_counter = Counter()
    field_match_counter = Counter()
    for item in items:
        for field_name, info in item["field_results"].items():
            if info["checked"]:
                field_checked_counter[field_name] += 1
                if info["predicted"]:
                    field_covered_counter[field_name] += 1
                if info["matched"]:
                    field_match_counter[field_name] += 1

    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in items:
        grouped[item.get("sample_type") or "unknown"].append(item)

    by_type = {
        sample_type: _summarize_group(records)
        for sample_type, records in sorted(grouped.items())
    }

    return {
        "name": "发票测试集评测报告",
        "manifest": str(manifest_path),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "total": total,
        "sample_passed": sample_passed,
        "sample_failed": total - sample_passed,
        "sample_pass_rate_pct": _pct(sample_passed, total),
        "invoice_match_count": invoice_match_count,
        "invoice_match_rate_pct": _pct(invoice_match_count, total),
        "labeled_samples": labeled_samples,
        "unlabeled_samples": total - labeled_samples,
        "labeled_sample_passed": labeled_sample_passed,
        "labeled_sample_pass_rate_pct": _pct(labeled_sample_passed, labeled_samples),
        "labeled_field_total": labeled_field_total,
        "covered_field_total": covered_field_total,
        "matched_field_total": matched_field_total,
        "field_coverage_pct": _pct(covered_field_total, labeled_field_total),
        "field_accuracy_pct": _pct(matched_field_total, labeled_field_total),
        "avg_processing_ms": avg_processing_ms,
        "field_checked_counts": dict(field_checked_counter),
        "field_covered_counts": dict(field_covered_counter),
        "field_match_counts": dict(field_match_counter),
        "by_sample_type": by_type,
    }


def write_json(report_dir: Path, summary: dict[str, Any], items: list[dict[str, Any]]) -> Path:
    output_path = report_dir / "invoice_evaluation.json"
    output_path.write_text(
        json.dumps({"summary": summary, "items": items}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return output_path


def write_csv(report_dir: Path, items: list[dict[str, Any]]) -> Path:
    output_path = report_dir / "invoice_evaluation.csv"
    headers = [
        "sample_id",
        "file_name",
        "source_dir",
        "sample_type",
        "quality_level",
        "expected_invoice",
        "actual_invoice",
        "invoice_match",
        "has_field_labels",
        "checked_fields",
        "covered_fields",
        "matched_fields",
        "field_coverage_pct",
        "field_accuracy_pct",
        "field_count",
        "confidence",
        "processing_ms",
        "sample_passed",
        "failure_reason",
        "notes",
    ]
    for field_name in FIELD_NAMES:
        headers.extend(
            [
                f"{field_name}_expected",
                f"{field_name}_actual",
                f"{field_name}_predicted",
                f"{field_name}_matched",
            ]
        )

    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for item in items:
            row = [
                item.get("sample_id", ""),
                item.get("file_name", ""),
                item.get("source_dir", ""),
                item.get("sample_type", ""),
                item.get("quality_level", ""),
                _fmt_bool(item.get("expected_invoice")),
                _fmt_bool(item.get("actual_invoice")),
                _fmt_bool(item.get("invoice_match")),
                _fmt_bool(item.get("has_field_labels")),
                item.get("checked_fields", 0),
                item.get("covered_fields", 0),
                item.get("matched_fields", 0),
                _fmt_pct(item.get("field_coverage_pct")),
                _fmt_pct(item.get("field_accuracy_pct")),
                item.get("field_count", 0),
                item.get("confidence", ""),
                item.get("processing_ms", 0),
                _fmt_bool(item.get("sample_passed")),
                item.get("failure_reason", ""),
                item.get("notes", ""),
            ]
            for field_name in FIELD_NAMES:
                field_info = item["field_results"][field_name]
                row.extend(
                    [
                        field_info.get("expected", ""),
                        field_info.get("actual", ""),
                        _fmt_bool(field_info.get("predicted")),
                        _fmt_bool(field_info.get("matched")),
                    ]
                )
            writer.writerow(row)
    return output_path


def write_markdown(report_dir: Path, summary: dict[str, Any], items: list[dict[str, Any]]) -> Path:
    output_path = report_dir / "invoice_evaluation.md"
    lines = [
        f"# {summary['name']}",
        "",
        f"- 生成时间：{summary['generated_at']}",
        f"- 清单路径：`{summary['manifest']}`",
        "",
        "## 评测口径",
        "",
        f"- 样本总数：{summary['total']}",
        f"- 已标注字段样本：{summary['labeled_samples']}",
        f"- 未标注字段样本：{summary['unlabeled_samples']}",
    ]

    if summary["labeled_field_total"] == 0:
        lines.append("- 当前清单没有字段真值，只能统计发票判定匹配率，不能据此声称字段级命中率。")
    else:
        lines.extend(
            [
                f"- 已标注字段总数：{summary['labeled_field_total']}",
                "- 字段覆盖率：已标注字段中，系统给出非空结果的比例。",
                "- 字段准确率：已标注字段中，系统结果与真值完全一致的比例。",
            ]
        )

    lines.extend(
        [
            "",
            "## 总览",
            "",
            f"- 发票判定匹配率：{_fmt_pct(summary['invoice_match_rate_pct'])} ({summary['invoice_match_count']}/{summary['total']})",
            f"- 全样本通过率：{_fmt_pct(summary['sample_pass_rate_pct'])} ({summary['sample_passed']}/{summary['total']})",
            f"- 已标注样本通过率：{_fmt_pct(summary['labeled_sample_pass_rate_pct'])} ({summary['labeled_sample_passed']}/{summary['labeled_samples']})",
            f"- 字段覆盖率：{_fmt_pct(summary['field_coverage_pct'])} ({summary['covered_field_total']}/{summary['labeled_field_total']})",
            f"- 字段准确率：{_fmt_pct(summary['field_accuracy_pct'])} ({summary['matched_field_total']}/{summary['labeled_field_total']})",
            f"- 平均处理耗时：{summary['avg_processing_ms']} ms",
            "",
            "## 分类统计",
            "",
            "| 样本类别 | 总数 | 已标注样本 | 发票判定匹配率 | 字段准确率 | 已标注样本通过率 | 平均耗时(ms) |",
            "| --- | ---: | ---: | ---: | ---: | ---: | ---: |",
        ]
    )

    for sample_type, info in summary["by_sample_type"].items():
        lines.append(
            f"| {sample_type} | {info['total']} | {info['labeled_samples']} | "
            f"{_fmt_pct(info['invoice_match_rate_pct'])} | {_fmt_pct(info['field_accuracy_pct'])} | "
            f"{_fmt_pct(info['labeled_sample_pass_rate_pct'])} | {info['avg_processing_ms']} |"
        )

    lines.extend(
        [
            "",
            "## 字段统计",
            "",
            "| 字段 | 已标注数 | 已提取数 | 命中数 | 覆盖率 | 准确率 |",
            "| --- | ---: | ---: | ---: | ---: | ---: |",
        ]
    )

    for field_name in FIELD_NAMES:
        checked = int(summary["field_checked_counts"].get(field_name, 0))
        covered = int(summary["field_covered_counts"].get(field_name, 0))
        matched = int(summary["field_match_counts"].get(field_name, 0))
        lines.append(
            f"| {field_name} | {checked} | {covered} | {matched} | "
            f"{_fmt_pct(_pct(covered, checked))} | {_fmt_pct(_pct(matched, checked))} |"
        )

    lines.extend(
        [
            "",
            "## 样本明细",
            "",
            "| ID | 文件 | 分类 | 发票判定 | 字段命中 | 结果 | 说明 |",
            "| --- | --- | --- | --- | --- | --- | --- |",
        ]
    )

    for item in items:
        lines.append(
            f"| {item['sample_id']} | {item['file_name']} | {item['sample_type']} | "
            f"{'通过' if item['invoice_match'] else '失败'} | "
            f"{item['matched_fields']}/{item['checked_fields']} | "
            f"{'通过' if item['sample_passed'] else '未通过'} | "
            f"{item.get('failure_reason') or item.get('notes') or ''} |"
        )

    failed_items = [item for item in items if not item["sample_passed"]]
    if failed_items:
        lines.extend(["", "## 未通过样本", ""])
        for item in failed_items[:20]:
            lines.append(f"- `{item['sample_id']} {item['file_name']}`：{item.get('failure_reason') or '未命中'}")

    output_path.write_text("\n".join(lines), encoding="utf-8")
    return output_path


def main() -> int:
    args = parse_args()
    manifest_path = Path(args.manifest)
    if not manifest_path.is_absolute():
        manifest_path = ROOT / manifest_path
    if not manifest_path.exists():
        print(f"未找到发票清单：{manifest_path}")
        return 1

    dataset_root = Path(args.dataset_root)
    if not dataset_root.is_absolute():
        dataset_root = ROOT / dataset_root

    items = load_manifest(manifest_path, args.sample_type, args.limit)
    if not items:
        print("发票清单为空，或筛选后无样本。")
        return 1

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_root = Path(args.report_root)
    if not report_root.is_absolute():
        report_root = ROOT / report_root
    report_dir = report_root / f"invoice_evaluation_{timestamp}"
    report_dir.mkdir(parents=True, exist_ok=True)

    processor = DocFlowProcessor()
    extractor = InvoiceExtractor()
    records = [
        evaluate_row(processor, extractor, row, manifest_path, dataset_root)
        for row in items
    ]
    summary = build_summary(records, manifest_path)

    json_path = write_json(report_dir, summary, records)
    csv_path = write_csv(report_dir, records)
    md_path = write_markdown(report_dir, summary, records)

    print("发票评测完成。")
    print(f"- JSON: {json_path}")
    print(f"- CSV: {csv_path}")
    print(f"- Markdown: {md_path}")
    print(
        f"- 发票判定匹配率: {_fmt_pct(summary['invoice_match_rate_pct'])} "
        f"({summary['invoice_match_count']}/{summary['total']})"
    )
    print(
        f"- 已标注样本通过率: {_fmt_pct(summary['labeled_sample_pass_rate_pct'])} "
        f"({summary['labeled_sample_passed']}/{summary['labeled_samples']})"
    )
    print(
        f"- 字段覆盖率: {_fmt_pct(summary['field_coverage_pct'])} "
        f"({summary['covered_field_total']}/{summary['labeled_field_total']})"
    )
    print(
        f"- 字段准确率: {_fmt_pct(summary['field_accuracy_pct'])} "
        f"({summary['matched_field_total']}/{summary['labeled_field_total']})"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
