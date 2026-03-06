from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path

from app import IMAGE_EXTS, process_image_ocr
from docflow_core import DocFlowProcessor
from docflow_support import build_error_info


ROOT = Path(__file__).resolve().parent
DEFAULT_MANIFEST = ROOT / "evaluation_set" / "sample_manifest.json"
DEFAULT_REPORT_ROOT = ROOT / "reports"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="run_evaluation_set",
        description="运行 DocFlow 评价集并生成量化评估报告。",
    )
    parser.add_argument(
        "manifest",
        nargs="?",
        default=str(DEFAULT_MANIFEST),
        help="评价集清单 JSON 文件路径",
    )
    parser.add_argument(
        "--report-root",
        default=str(DEFAULT_REPORT_ROOT),
        help="评估报告输出目录，默认 reports/",
    )
    parser.add_argument(
        "--keywords",
        action="store_true",
        help="启用关键词与摘要提取",
    )
    return parser.parse_args()


def load_manifest(manifest_path: Path) -> dict:
    return json.loads(manifest_path.read_text(encoding="utf-8"))


def evaluate_item(processor: DocFlowProcessor, item: dict, extract_keywords: bool) -> dict:
    relative_path = item["path"]
    file_path = ROOT / relative_path
    expected_success = item.get("expected_success")
    min_char_count = item.get("min_char_count")
    must_contain = item.get("must_contain") or []
    must_not_contain = item.get("must_not_contain") or []

    if not file_path.exists():
        result = {
            "success": False,
            "file": file_path.name,
            "format": file_path.suffix.lstrip("."),
            "text": "",
            "tables": [],
            "metadata": {},
            "statistics": {},
            "processing_ms": 0.0,
            "error": f"文件不存在: {relative_path}",
        }
    else:
        if file_path.suffix.lower() in IMAGE_EXTS:
            result = process_image_ocr(str(file_path), file_path.name)
        else:
            result = processor.process(
                str(file_path),
                extract_keywords=extract_keywords,
                output_format="txt",
            )

    text = (result.get("text") or "").strip()
    char_count = result.get("statistics", {}).get("char_count", len(text))
    text_lower = text.lower()
    must_contain_hits = [token for token in must_contain if token.lower() in text_lower]
    must_not_contain_hits = [token for token in must_not_contain if token.lower() in text_lower]
    success_match = expected_success is None or bool(result.get("success")) == bool(expected_success)
    min_char_pass = min_char_count is None or char_count >= int(min_char_count)
    contains_pass = not must_contain or len(must_contain_hits) == len(must_contain)
    excludes_pass = not must_not_contain or not must_not_contain_hits

    active_rules = [success_match]
    if min_char_count is not None:
        active_rules.append(min_char_pass)
    if must_contain:
        active_rules.append(contains_pass)
    if must_not_contain:
        active_rules.append(excludes_pass)

    error_info = build_error_info(
        result.get("error", ""),
        file_name=file_path.name,
        file_ext=file_path.suffix.lower(),
        metadata_dict=result.get("metadata") or {},
        source="evaluation",
    )

    return {
        "path": relative_path,
        "filename": file_path.name,
        "expected_success": expected_success,
        "actual_success": bool(result.get("success")),
        "success_match": success_match,
        "min_char_count": min_char_count,
        "actual_char_count": char_count,
        "min_char_pass": min_char_pass,
        "must_contain": must_contain,
        "must_contain_hits": must_contain_hits,
        "contains_pass": contains_pass,
        "must_not_contain": must_not_contain,
        "must_not_contain_hits": must_not_contain_hits,
        "excludes_pass": excludes_pass,
        "processing_ms": round(float(result.get("processing_ms", 0.0)), 2),
        "error": result.get("error", ""),
        "error_info": error_info,
        "rule_passed": all(active_rules),
        "notes": item.get("notes", ""),
    }


def build_summary(items: list[dict], manifest: dict) -> dict:
    total = len(items)
    passed = sum(1 for item in items if item["rule_passed"])
    success_match_count = sum(1 for item in items if item["success_match"])
    contains_checked = sum(1 for item in items if item["must_contain"])
    contains_passed = sum(1 for item in items if item["must_contain"] and item["contains_pass"])
    min_char_checked = sum(1 for item in items if item["min_char_count"] is not None)
    min_char_passed = sum(1 for item in items if item["min_char_count"] is not None and item["min_char_pass"])
    avg_processing_ms = round(sum(item["processing_ms"] for item in items) / total, 2) if total else 0.0

    return {
        "name": manifest.get("name", "DocFlow 评价集"),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "total": total,
        "passed": passed,
        "failed": total - passed,
        "pass_rate_pct": round((passed / total) * 100, 2) if total else 0.0,
        "success_match_count": success_match_count,
        "contains_checked": contains_checked,
        "contains_passed": contains_passed,
        "min_char_checked": min_char_checked,
        "min_char_passed": min_char_passed,
        "avg_processing_ms": avg_processing_ms,
    }


def write_json(report_dir: Path, summary: dict, items: list[dict], manifest: dict) -> Path:
    output_path = report_dir / "evaluation.json"
    payload = {"manifest": manifest, "summary": summary, "items": items}
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path


def write_markdown(report_dir: Path, summary: dict, items: list[dict], manifest: dict) -> Path:
    output_path = report_dir / "evaluation.md"
    lines = [
        f"# {summary['name']}",
        "",
        f"- 生成时间：{summary['generated_at']}",
        f"- 描述：{manifest.get('description', '--')}",
        "",
        "## 总览",
        "",
        f"- 总样本：{summary['total']}",
        f"- 通过：{summary['passed']}",
        f"- 未通过：{summary['failed']}",
        f"- 通过率：{summary['pass_rate_pct']}%",
        f"- 成功状态匹配：{summary['success_match_count']}/{summary['total']}",
        f"- 关键词包含规则：{summary['contains_passed']}/{summary['contains_checked'] or 0}",
        f"- 最小字符规则：{summary['min_char_passed']}/{summary['min_char_checked'] or 0}",
        f"- 平均耗时：{summary['avg_processing_ms']} ms",
        "",
        "## 明细",
        "",
        "| 文件 | 结果 | 成功匹配 | 字符规则 | 包含规则 | 耗时(ms) | 备注 |",
        "| --- | --- | --- | --- | --- | --- | --- |",
    ]

    for item in items:
        lines.append(
            f"| {item['filename']} | {'通过' if item['rule_passed'] else '未通过'} | "
            f"{'是' if item['success_match'] else '否'} | "
            f"{'是' if item['min_char_pass'] else ('--' if item['min_char_count'] is None else '否')} | "
            f"{'是' if item['contains_pass'] else ('--' if not item['must_contain'] else '否')} | "
            f"{item['processing_ms']} | {item.get('notes', '') or (item.get('error_info', {}) or {}).get('category', '')} |"
        )

    failed_items = [item for item in items if not item["rule_passed"]]
    if failed_items:
        lines.extend(["", "## 未通过样本", ""])
        for item in failed_items:
            lines.append(f"- `{item['filename']}`：{item.get('error') or (item.get('error_info', {}) or {}).get('message', '规则未通过')}")

    output_path.write_text("\n".join(lines), encoding="utf-8")
    return output_path


def main() -> int:
    args = parse_args()
    manifest_path = Path(args.manifest)
    if not manifest_path.is_absolute():
        manifest_path = ROOT / manifest_path
    if not manifest_path.exists():
        print(f"未找到评价集清单：{manifest_path}")
        return 1

    manifest = load_manifest(manifest_path)
    items = manifest.get("items") or []
    if not items:
        print("评价集清单为空。")
        return 1

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_dir = Path(args.report_root) / f"evaluation_{timestamp}"
    report_dir.mkdir(parents=True, exist_ok=True)

    processor = DocFlowProcessor()
    records = [evaluate_item(processor, item, extract_keywords=args.keywords) for item in items]
    summary = build_summary(records, manifest)
    json_path = write_json(report_dir, summary, records, manifest)
    md_path = write_markdown(report_dir, summary, records, manifest)

    print("评价集运行完成。")
    print(f"- JSON: {json_path}")
    print(f"- Markdown: {md_path}")
    print(f"- 通过率: {summary['pass_rate_pct']}% ({summary['passed']}/{summary['total']})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
