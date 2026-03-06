"""
一键批量运行 DocFlow 测试并生成报告。

默认测试目录：
- test_documents
- test_documents_edge_cases

输出：
- reports/<timestamp>/report.md
- reports/<timestamp>/results.json
- reports/<timestamp>/results.csv
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import sys
from collections import Counter, defaultdict
from datetime import datetime
from html import escape
from pathlib import Path
from typing import Iterable

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT = SCRIPT_DIR.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app import IMAGE_EXTS, process_image_ocr
from docflow_core import DocFlowProcessor
from docflow_support import build_error_info, summarize_error_records

SAMPLE_DATA_DIR = ROOT / "sample_data"
SUITE_ALIASES = {
    "test_documents": SAMPLE_DATA_DIR / "test_documents",
    "test_documents_edge_cases": SAMPLE_DATA_DIR / "test_documents_edge_cases",
}
DEFAULT_SUITES = ["test_documents", "test_documents_edge_cases"]
DEFAULT_REPORT_ROOT = ROOT / "reports"


KNOWN_EXPECTATIONS = {
    "test_documents": defaultdict(lambda: True),
    "test_documents_edge_cases": {
        "00_empty.txt": True,
        "01_utf8_bom.txt": True,
        "02_gbk.txt": True,
        "03_utf16.txt": True,
        "04_long_line.txt": True,
        "05_special_chars.md": True,
        "06_带空格 和 中文 @2026!.txt": True,
        "10_nested.json": True,
        "11_invalid.json": False,
        "12_empty.csv": True,
        "13_irregular.csv": True,
        "14_large_table.csv": True,
        "20_empty.docx": True,
        "21_minimal.docx": True,
        "22_compat.doc": True,
        "23_fake.docx": False,
        "24_fake.doc": False,
        "30_empty.xlsx": True,
        "31_minimal.xlsx": True,
        "32_compat.xls": True,
        "33_fake.xlsx": False,
        "34_fake.xls": False,
        "40_empty.pptx": True,
        "41_minimal.pptx": True,
        "42_fake.pptx": False,
        "50_blank.png": True,
        "51_low_contrast.jpg": True,
        "52_tiny_text.png": True,
        "53_dense_text.webp": True,
        "54_tiff_scan.tiff": True,
        "60_blank_scan.pdf": False,
        "61_corrupt.pdf": False,
        "62_short_text.pdf": True,
        "70_unsupported.bin": False,
        "71_empty.json": False,
    },
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="run_batch_tests",
        description="批量运行 DocFlow 样本测试并生成报告。",
    )
    parser.add_argument(
        "suites",
        nargs="*",
        default=DEFAULT_SUITES,
        help="要测试的目录列表，默认同时运行 test_documents 和 test_documents_edge_cases",
    )
    parser.add_argument(
        "--report-root",
        default=str(DEFAULT_REPORT_ROOT),
        help="报告输出根目录，默认 reports/",
    )
    parser.add_argument(
        "--keywords",
        action="store_true",
        help="为非图片文件开启关键词/摘要提取（默认关闭以提升速度）",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="若结果不符合预期则返回非 0 状态码",
    )
    parser.add_argument(
        "--pdf-mode",
        choices=["accurate", "balanced", "fast"],
        default=os.getenv("DOCFLOW_DEFAULT_PDF_MODE", "balanced"),
        help="PDF 解析模式：高精度 / 平衡 / 快速",
    )
    return parser.parse_args()


def resolve_suite_path(suite: str) -> Path:
    suite_name = str(suite or "").strip()
    if suite_name in SUITE_ALIASES:
        return SUITE_ALIASES[suite_name]

    suite_path = Path(suite_name)
    if suite_path.is_absolute():
        return suite_path
    if (ROOT / suite_path).exists():
        return ROOT / suite_path
    return suite_path


def iter_files(suite_paths: Iterable[Path]) -> list[tuple[str, Path]]:
    items: list[tuple[str, Path]] = []
    for suite_path in suite_paths:
        if not suite_path.exists():
            continue
        for path in sorted(p for p in suite_path.iterdir() if p.is_file()):
            if path.name.lower() == "readme.md":
                continue
            items.append((suite_path.name, path))
    return items


def get_expected_result(suite_name: str, filename: str):
    suite_expectation = KNOWN_EXPECTATIONS.get(suite_name)
    if suite_expectation is None:
        return None
    if isinstance(suite_expectation, dict):
        return suite_expectation.get(filename)
    return suite_expectation[filename]


def run_single_case(
    processor: DocFlowProcessor,
    suite_name: str,
    file_path: Path,
    extract_keywords: bool,
    pdf_mode: str,
) -> dict:
    ext = file_path.suffix.lower()
    expected_success = get_expected_result(suite_name, file_path.name)

    try:
        if ext in IMAGE_EXTS:
            result = process_image_ocr(str(file_path), file_path.name)
        else:
            result = processor.process(
                str(file_path),
                extract_keywords=extract_keywords,
                output_format="txt",
                pdf_mode=pdf_mode,
            )
    except Exception as exc:
        result = {
            "success": False,
            "file": file_path.name,
            "format": ext.lstrip("."),
            "text": "",
            "tables": [],
            "metadata": {},
            "statistics": {},
            "processing_ms": 0.0,
            "formatted_output": "",
            "error": str(exc),
        }

    statistics = result.get("statistics") or {}
    metadata = result.get("metadata") or {}
    char_count = statistics.get("char_count", len(result.get("text", "")))
    table_count = statistics.get("table_count", len(result.get("tables", [])))
    actual_success = bool(result.get("success"))
    expectation_match = None if expected_success is None else (actual_success == expected_success)
    error_info = build_error_info(
        result.get("error", ""),
        file_name=file_path.name,
        file_ext=ext,
        metadata_dict=metadata,
        source="batch",
    )

    return {
        "suite": suite_name,
        "filename": file_path.name,
        "path": str(file_path),
        "extension": ext,
        "size_bytes": file_path.stat().st_size,
        "expected_success": expected_success,
        "success": actual_success,
        "matches_expectation": expectation_match,
        "format": result.get("format", ext.lstrip(".")),
        "char_count": char_count,
        "table_count": table_count,
        "processing_ms": round(float(result.get("processing_ms", 0.0)), 2),
        "error": result.get("error", ""),
        "error_info": error_info,
        "error_category": (error_info or {}).get("category", ""),
        "parse_method": metadata.get("解析方式", ""),
        "ocr_engine": metadata.get("ocr_engine") or metadata.get("engine", ""),
        "metadata": metadata,
    }


def build_summary(records: list[dict], pdf_mode: str = "balanced") -> dict:
    success_count = sum(1 for item in records if item["success"])
    fail_count = len(records) - success_count
    matched = [item for item in records if item["matches_expectation"] is not None]
    matched_count = sum(1 for item in matched if item["matches_expectation"])
    unexpected = [item for item in matched if not item["matches_expectation"]]
    avg_processing_ms = round(
        sum(item["processing_ms"] for item in records) / len(records),
        2,
    ) if records else 0.0

    suite_summary = {}
    for suite_name in sorted({item["suite"] for item in records}):
        suite_records = [item for item in records if item["suite"] == suite_name]
        suite_summary[suite_name] = {
            "total": len(suite_records),
            "success": sum(1 for item in suite_records if item["success"]),
            "failed": sum(1 for item in suite_records if not item["success"]),
            "char_count_total": sum(item["char_count"] for item in suite_records),
            "table_count_total": sum(item["table_count"] for item in suite_records),
            "success_rate_pct": round(
                sum(1 for item in suite_records if item["success"]) / len(suite_records) * 100,
                2,
            ) if suite_records else 0.0,
            "avg_processing_ms": round(
                sum(item["processing_ms"] for item in suite_records) / len(suite_records),
                2,
            ) if suite_records else 0.0,
        }

    format_summary = {}
    for extension in sorted({item["extension"] or "<none>" for item in records}):
        ext_records = [item for item in records if (item["extension"] or "<none>") == extension]
        format_summary[extension] = {
            "total": len(ext_records),
            "success": sum(1 for item in ext_records if item["success"]),
            "failed": sum(1 for item in ext_records if not item["success"]),
            "char_count_total": sum(item["char_count"] for item in ext_records),
            "table_count_total": sum(item["table_count"] for item in ext_records),
            "success_rate_pct": round(
                sum(1 for item in ext_records if item["success"]) / len(ext_records) * 100,
                2,
            ) if ext_records else 0.0,
            "avg_processing_ms": round(
                sum(item["processing_ms"] for item in ext_records) / len(ext_records),
                2,
            ) if ext_records else 0.0,
        }

    format_counter = Counter(item["extension"] or "<none>" for item in records)
    parse_method_counts = Counter(item["parse_method"] or "未标记" for item in records if item.get("parse_method"))
    ocr_engine_counts = Counter(item["ocr_engine"] or "未使用" for item in records if item.get("ocr_engine"))
    error_summary = summarize_error_records(records)
    return {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "pdf_mode": pdf_mode,
        "total": len(records),
        "success": success_count,
        "failed": fail_count,
        "cancelled": 0,
        "avg_processing_ms": avg_processing_ms,
        "success_rate_pct": round((success_count / len(records)) * 100, 2) if records else 0.0,
        "expected_checked": len(matched),
        "expected_matched": matched_count,
        "unexpected_count": len(unexpected),
        "suite_summary": suite_summary,
        "format_summary": format_summary,
        "format_counter": dict(sorted(format_counter.items())),
        "parse_method_counts": dict(parse_method_counts),
        "ocr_engine_counts": dict(ocr_engine_counts),
        "char_count_total": sum(item["char_count"] for item in records),
        "table_count_total": sum(item["table_count"] for item in records),
        "avg_char_count": round(sum(item["char_count"] for item in records) / len(records), 2) if records else 0.0,
        "unexpected_files": [item["filename"] for item in unexpected],
        **error_summary,
    }


def write_json(report_dir: Path, summary: dict, records: list[dict]) -> Path:
    output_path = report_dir / "results.json"
    payload = {"summary": summary, "records": records}
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path


def write_csv(report_dir: Path, records: list[dict]) -> Path:
    output_path = report_dir / "results.csv"
    fieldnames = [
        "suite",
        "filename",
        "extension",
        "size_bytes",
        "expected_success",
        "success",
        "matches_expectation",
        "char_count",
        "table_count",
        "processing_ms",
        "parse_method",
        "ocr_engine",
        "error_category",
        "error",
        "path",
    ]
    with output_path.open("w", encoding="utf-8-sig", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for record in records:
            writer.writerow({key: record.get(key, "") for key in fieldnames})
    return output_path


def markdown_table(rows: list[list[str]]) -> str:
    if not rows:
        return ""
    header = "| " + " | ".join(rows[0]) + " |"
    sep = "| " + " | ".join(["---"] * len(rows[0])) + " |"
    body = ["| " + " | ".join(row) + " |" for row in rows[1:]]
    return "\n".join([header, sep, *body])


def write_markdown(report_dir: Path, summary: dict, records: list[dict], suite_names: list[str]) -> Path:
    output_path = report_dir / "report.md"

    overview_rows = [["指标", "值"]]
    overview_rows.extend(
        [
            ["生成时间", summary["generated_at"]],
            ["PDF模式", summary.get("pdf_mode", "balanced")],
            ["测试总数", str(summary["total"])],
            ["成功", str(summary["success"])],
            ["失败", str(summary["failed"])],
            ["已校验预期", str(summary["expected_checked"])],
            ["符合预期", str(summary["expected_matched"])],
            ["异常项", str(summary["unexpected_count"])],
        ]
    )

    suite_rows = [["测试集", "总数", "成功", "失败", "平均耗时(ms)"]]
    for suite_name in suite_names:
        item = summary["suite_summary"].get(suite_name, {})
        suite_rows.append(
            [
                suite_name,
                str(item.get("total", 0)),
                str(item.get("success", 0)),
                str(item.get("failed", 0)),
                str(item.get("avg_processing_ms", 0.0)),
            ]
        )

    failure_rows = [["测试集", "文件", "预期", "实际", "错误信息"]]
    failures = [item for item in records if not item["success"]]
    for item in failures[:50]:
        failure_rows.append(
            [
                item["suite"],
                item["filename"],
                str(item["expected_success"]),
                str(item["success"]),
                (item["error"] or "").replace("\n", " ")[:120],
            ]
        )

    unexpected_rows = [["测试集", "文件", "预期", "实际"]]
    unexpected = [item for item in records if item["matches_expectation"] is False]
    for item in unexpected:
        unexpected_rows.append(
            [item["suite"], item["filename"], str(item["expected_success"]), str(item["success"])]
        )

    top_slowest = sorted(records, key=lambda item: item["processing_ms"], reverse=True)[:10]
    slow_rows = [["测试集", "文件", "耗时(ms)", "字符数", "解析方式"]]
    for item in top_slowest:
        slow_rows.append(
            [
                item["suite"],
                item["filename"],
                str(item["processing_ms"]),
                str(item["char_count"]),
                item["parse_method"] or item["ocr_engine"] or "-",
            ]
        )

    format_rows = [["格式", "总数", "成功", "失败", "成功率", "平均耗时(ms)"]]
    for extension, item in summary.get("format_summary", {}).items():
        format_rows.append(
            [
                extension,
                str(item.get("total", 0)),
                str(item.get("success", 0)),
                str(item.get("failed", 0)),
                f"{item.get('success_rate_pct', 0)}%",
                str(item.get("avg_processing_ms", 0.0)),
            ]
        )

    error_rows = [["错误类型", "次数"]]
    for category, count in (summary.get("error_category_counts") or {}).items():
        error_rows.append([category, str(count)])

    lines = [
        "# DocFlow 批量测试报告",
        "",
        f"- 报告目录：`{report_dir}`",
        f"- 测试集：{', '.join(suite_names)}",
        "",
        "## 总览",
        "",
        markdown_table(overview_rows),
        "",
        "## 分测试集统计",
        "",
        markdown_table(suite_rows),
        "",
        "## 分格式统计",
        "",
        markdown_table(format_rows),
        "",
        "## 失败样本",
        "",
        markdown_table(failure_rows) if len(failure_rows) > 1 else "无",
        "",
        "## 不符合预期",
        "",
        markdown_table(unexpected_rows) if len(unexpected_rows) > 1 else "无",
        "",
        "## 最慢的 10 个样本",
        "",
        markdown_table(slow_rows),
        "",
        "## 错误分类",
        "",
        markdown_table(error_rows) if len(error_rows) > 1 else "无",
        "",
        "## 格式分布",
        "",
    ]

    for ext, count in summary["format_counter"].items():
        lines.append(f"- `{ext}`: {count}")

    output_path.write_text("\n".join(lines), encoding="utf-8")
    return output_path


def write_html_dashboard(report_dir: Path, summary: dict, records: list[dict], suite_names: list[str]) -> Path:
    output_path = report_dir / "summary.html"
    top_failures = [item for item in records if not item["success"]][:20]
    unexpected_rows_data = [item for item in records if item["matches_expectation"] is False][:20]
    top_slowest = sorted(records, key=lambda item: item["processing_ms"], reverse=True)[:12]
    suite_summary = summary.get("suite_summary") or {}
    format_summary = summary.get("format_summary") or {}
    error_counts = summary.get("error_category_counts") or {}
    parse_counts = summary.get("parse_method_counts") or {}
    ocr_counts = summary.get("ocr_engine_counts") or {}
    error_samples = summary.get("error_samples") or []
    max_suite_total = max((item.get("total", 0) for item in suite_summary.values()), default=1)
    max_format_total = max((item.get("total", 0) for item in format_summary.values()), default=1)
    max_error_count = max(error_counts.values(), default=1)
    max_parse_count = max(parse_counts.values(), default=1)
    max_ocr_count = max(ocr_counts.values(), default=1)

    def esc(value) -> str:
        if value is None or value == "":
            return "-"
        return escape(str(value))

    def fmt_number(value) -> str:
        if isinstance(value, float):
            if value.is_integer():
                return f"{int(value):,}"
            return f"{value:,.2f}"
        if isinstance(value, int):
            return f"{value:,}"
        return esc(value)

    def render_table_rows(rows: list[str], colspan: int, empty_text: str = "暂无数据") -> str:
        return "".join(rows) if rows else f'<tr><td colspan="{colspan}" class="empty">{esc(empty_text)}</td></tr>'

    def render_bar_rows(items: list[tuple[str, int | float]], max_value: float, empty_text: str) -> str:
        if not items:
            return f'<div class="empty-card">{esc(empty_text)}</div>'
        tones = ["green", "blue", "violet", "amber", "pink"]
        rows = []
        safe_max = max(max_value, 1)
        for index, (label, value) in enumerate(items):
            width = 0 if value <= 0 else max(8.0, round((float(value) / safe_max) * 100, 2))
            tone = tones[index % len(tones)]
            rows.append(
                f"""
                <div class="bar-item">
                  <div class="bar-head"><span>{esc(label)}</span><strong>{fmt_number(value)}</strong></div>
                  <div class="bar-track"><div class="bar-fill {tone}" style="width:{min(width, 100)}%"></div></div>
                </div>
                """
            )
        return "".join(rows)

    total = summary.get("total", 0)
    failed = summary.get("failed", 0)
    success = summary.get("success", 0)
    success_rate = float(summary.get("success_rate_pct", 0.0) or 0.0)
    if failed == 0 and summary.get("unexpected_count", 0) == 0:
        health_label = "状态稳定"
        health_class = "ok"
    elif failed <= max(1, total // 10) and summary.get("unexpected_count", 0) <= 1:
        health_label = "基本可用"
        health_class = "warn"
    else:
        health_label = "需要关注"
        health_class = "danger"

    suite_cards = []
    for suite_name in suite_names:
        item = suite_summary.get(suite_name, {})
        rate = float(item.get("success_rate_pct", 0.0) or 0.0)
        suite_cards.append(
            f"""
            <div class="suite-card">
              <div class="suite-top">
                <div>
                  <div class="suite-name">{esc(suite_name)}</div>
                  <div class="suite-sub">样本 {fmt_number(item.get('total', 0))} · 平均耗时 {fmt_number(item.get('avg_processing_ms', 0.0))} ms</div>
                </div>
                <span class="pill {'ok' if rate >= 90 else ('warn' if rate >= 70 else 'danger')}">{fmt_number(rate)}%</span>
              </div>
              <div class="suite-progress"><div class="bar-fill {'green' if rate >= 90 else ('amber' if rate >= 70 else 'pink')}" style="width:{min(rate, 100)}%"></div></div>
              <div class="suite-metrics">
                <span>成功 {fmt_number(item.get('success', 0))}</span>
                <span>失败 {fmt_number(item.get('failed', 0))}</span>
                <span>字符 {fmt_number(item.get('char_count_total', 0))}</span>
                <span>表格 {fmt_number(item.get('table_count_total', 0))}</span>
              </div>
            </div>
            """
        )

    fmt_rows = []
    for extension, item in format_summary.items():
        rate = float(item.get("success_rate_pct", 0.0) or 0.0)
        share = 0 if max_format_total <= 0 else round((item.get("total", 0) / max_format_total) * 100, 2)
        fmt_rows.append(
            f"""
            <tr>
              <td><code>{esc(extension)}</code></td>
              <td>{fmt_number(item.get('total', 0))}</td>
              <td>{fmt_number(item.get('success', 0))}</td>
              <td>{fmt_number(item.get('failed', 0))}</td>
              <td>
                <div class="inline-metric">
                  <span>{fmt_number(rate)}%</span>
                  <div class="mini-track"><div class="mini-fill {'green' if rate >= 90 else ('amber' if rate >= 70 else 'pink')}" style="width:{min(rate, 100)}%"></div></div>
                </div>
              </td>
              <td>{fmt_number(item.get('avg_processing_ms', 0.0))}</td>
              <td>{fmt_number(item.get('char_count_total', 0))}</td>
              <td>{fmt_number(item.get('table_count_total', 0))}</td>
              <td>{fmt_number(share)}%</td>
            </tr>
            """
        )

    fail_rows = [
        f"""
        <tr>
          <td>{esc(item['suite'])}</td>
          <td title="{esc(item['path'])}">{esc(item['filename'])}</td>
          <td><span class="pill danger">{esc(item.get('error_category') or '-')}</span></td>
          <td>{esc(item.get('parse_method') or item.get('ocr_engine') or '-')}</td>
          <td>{esc((item.get('error') or '-').replace(chr(10), ' '))[:180]}</td>
        </tr>
        """
        for item in top_failures
    ]

    unexpected_rows = [
        f"""
        <tr>
          <td>{esc(item['suite'])}</td>
          <td>{esc(item['filename'])}</td>
          <td>{esc(item.get('expected_success'))}</td>
          <td>{esc(item.get('success'))}</td>
          <td>{esc(item.get('error_category') or ('成功' if item.get('success') else '失败'))}</td>
        </tr>
        """
        for item in unexpected_rows_data
    ]

    slow_rows = [
        f"""
        <tr>
          <td>{esc(item['suite'])}</td>
          <td>{esc(item['filename'])}</td>
          <td>{fmt_number(item['processing_ms'])}</td>
          <td>{fmt_number(item['char_count'])}</td>
          <td>{esc(item.get('parse_method') or '-')}</td>
          <td>{esc(item.get('ocr_engine') or '-')}</td>
        </tr>
        """
        for item in top_slowest
    ]

    insight_tags = [
        f'<span class="tag">总样本 {fmt_number(total)}</span>',
        f'<span class="tag">成功率 {fmt_number(success_rate)}%</span>',
        f'<span class="tag">异常项 {fmt_number(summary.get("unexpected_count", 0))}</span>',
        f'<span class="tag">总字符 {fmt_number(summary.get("char_count_total", 0))}</span>',
        f'<span class="tag">总表格 {fmt_number(summary.get("table_count_total", 0))}</span>',
        f'<span class="tag">PDF 模式 {esc(summary.get("pdf_mode", "balanced"))}</span>',
    ]

    error_sample_cards = "".join(
        f"""
        <div class="sample-card">
          <div class="sample-file">{esc(item.get('filename'))}</div>
          <div class="sample-cat">{esc(item.get('category'))}</div>
          <div class="sample-msg">{esc(item.get('message'))}</div>
          <div class="sample-hint">{esc(item.get('hint'))}</div>
        </div>
        """
        for item in error_samples[:6]
    ) or '<div class="empty-card">当前没有失败样本。</div>'

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>DocFlow 批测汇总</title>
  <style>
    :root{{color-scheme:dark;--bg:#0b1020;--panel:#131a2f;--panel2:#10172a;--line:#283350;--text:#edf2ff;--muted:#93a4c8;--green:#3de8a0;--blue:#7cb4ff;--violet:#9b8cff;--amber:#ffb340;--pink:#ff6b8a;}}
    *{{box-sizing:border-box}}
    body{{font-family:Segoe UI,Arial,sans-serif;background:radial-gradient(circle at top,#192444 0,#0b1020 48%);color:var(--text);margin:0;padding:24px}}
    .wrap{{max-width:1320px;margin:0 auto}}
    h1,h2,h3{{margin:0}}
    h2{{font-size:22px;margin-bottom:16px}}
    section{{margin-top:28px}}
    .hero{{display:grid;grid-template-columns:1.4fr .9fr;gap:18px;align-items:stretch}}
    .hero-card,.panel,table{{background:rgba(19,26,47,.92);border:1px solid var(--line);border-radius:18px;box-shadow:0 16px 40px rgba(0,0,0,.18)}}
    .hero-card{{padding:22px}}
    .hero-title{{display:flex;align-items:center;justify-content:space-between;gap:16px;margin-bottom:10px}}
    .meta{{color:var(--muted);font-size:14px;line-height:1.7}}
    .status-badge{{display:inline-flex;align-items:center;padding:8px 12px;border-radius:999px;font-size:13px;font-weight:700}}
    .status-badge.ok{{background:rgba(61,232,160,.14);color:var(--green);border:1px solid rgba(61,232,160,.28)}}
    .status-badge.warn{{background:rgba(255,179,64,.12);color:var(--amber);border:1px solid rgba(255,179,64,.28)}}
    .status-badge.danger{{background:rgba(255,107,138,.12);color:var(--pink);border:1px solid rgba(255,107,138,.28)}}
    .hero-progress{{padding:22px}}
    .big-metric{{font-size:44px;font-weight:800;line-height:1}}
    .sub-metric{{color:var(--muted);margin-top:8px}}
    .hero-bar{{height:14px;background:#0e1528;border-radius:999px;border:1px solid var(--line);overflow:hidden;margin:18px 0 14px}}
    .hero-bar-fill{{height:100%;background:linear-gradient(90deg,var(--green),#7cf0b9)}}
    .hero-split{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
    .split-box{{padding:12px 14px;border-radius:14px;background:var(--panel2);border:1px solid var(--line)}}
    .split-box strong{{display:block;font-size:20px;margin-top:6px}}
    .tag-row{{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px}}
    .tag{{display:inline-flex;align-items:center;padding:7px 11px;border-radius:999px;background:#1a2542;color:#dfe8ff;font-size:13px;border:1px solid #27365b}}
    .cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-top:18px}}
    .card{{padding:16px 18px}}
    .lbl{{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}}
    .val{{font-size:28px;font-weight:700;margin-top:8px}}
    .subval{{margin-top:6px;color:var(--muted);font-size:13px}}
    .panel{{padding:18px}}
    .two-col{{display:grid;grid-template-columns:1.1fr .9fr;gap:18px}}
    .three-col{{display:grid;grid-template-columns:repeat(3,1fr);gap:18px}}
    .suite-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px}}
    .suite-card{{padding:16px;border-radius:16px;background:var(--panel2);border:1px solid var(--line)}}
    .suite-top{{display:flex;justify-content:space-between;gap:12px;align-items:flex-start}}
    .suite-name{{font-size:16px;font-weight:700}}
    .suite-sub{{margin-top:6px;color:var(--muted);font-size:13px}}
    .suite-progress{{height:10px;background:#0e1528;border-radius:999px;overflow:hidden;border:1px solid #202b46;margin:14px 0}}
    .suite-metrics{{display:grid;grid-template-columns:repeat(2,1fr);gap:8px;color:var(--muted);font-size:13px}}
    .pill{{display:inline-flex;align-items:center;padding:6px 10px;border-radius:999px;font-size:12px;font-weight:700;border:1px solid transparent;white-space:nowrap}}
    .pill.ok{{background:rgba(61,232,160,.14);color:var(--green);border-color:rgba(61,232,160,.28)}}
    .pill.warn{{background:rgba(255,179,64,.12);color:var(--amber);border-color:rgba(255,179,64,.28)}}
    .pill.danger{{background:rgba(255,107,138,.12);color:var(--pink);border-color:rgba(255,107,138,.28)}}
    .bar-stack{{display:flex;flex-direction:column;gap:12px}}
    .bar-item{{padding:12px 14px;border-radius:14px;background:var(--panel2);border:1px solid var(--line)}}
    .bar-head{{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:10px}}
    .bar-head span{{color:var(--muted)}}
    .bar-track,.mini-track{{height:10px;background:#0d1426;border-radius:999px;overflow:hidden;border:1px solid #1f2a45}}
    .bar-fill,.mini-fill{{height:100%;border-radius:999px}}
    .bar-fill.green,.mini-fill.green{{background:linear-gradient(90deg,var(--green),#7bf1bb)}}
    .bar-fill.blue,.mini-fill.blue{{background:linear-gradient(90deg,var(--blue),#a8ccff)}}
    .bar-fill.violet,.mini-fill.violet{{background:linear-gradient(90deg,var(--violet),#beb0ff)}}
    .bar-fill.amber,.mini-fill.amber{{background:linear-gradient(90deg,var(--amber),#ffd180)}}
    .bar-fill.pink,.mini-fill.pink{{background:linear-gradient(90deg,var(--pink),#ff9cb2)}}
    .sample-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}}
    .sample-card{{padding:14px;border-radius:14px;background:var(--panel2);border:1px solid var(--line)}}
    .sample-file{{font-weight:700}}
    .sample-cat{{margin-top:8px;color:var(--pink);font-size:13px}}
    .sample-msg{{margin-top:8px;font-size:13px;line-height:1.6}}
    .sample-hint{{margin-top:8px;color:var(--muted);font-size:12px;line-height:1.6}}
    table{{width:100%;border-collapse:separate;border-spacing:0;overflow:hidden}}
    th,td{{padding:12px 14px;border-bottom:1px solid var(--line);text-align:left;font-size:14px;vertical-align:top}}
    th{{background:#18233f;color:#b8c6e6;font-weight:600;position:sticky;top:0}}
    tr:last-child td{{border-bottom:none}}
    tbody tr:hover td{{background:rgba(124,180,255,.05)}}
    code{{padding:3px 7px;border-radius:999px;background:#1a2542;color:#cfe0ff}}
    .inline-metric{{display:flex;align-items:center;gap:10px;min-width:160px}}
    .inline-metric span{{white-space:nowrap}}
    .mini-track{{flex:1}}
    .empty,.empty-card{{padding:18px;color:var(--muted);text-align:center}}
    .nav{{display:flex;flex-wrap:wrap;gap:8px;margin-top:16px}}
    .nav a{{color:var(--blue);text-decoration:none;padding:8px 12px;border-radius:999px;background:#121a31;border:1px solid var(--line)}}
    .nav a:hover{{border-color:#3a4c78;background:#172140}}
    @media (max-width: 980px){{.hero,.two-col,.three-col{{grid-template-columns:1fr}} body{{padding:16px}}}}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="hero">
      <div class="hero-card">
        <div class="hero-title">
          <div>
            <h1>DocFlow 批测可视化汇总</h1>
            <div class="meta">生成时间：{esc(summary.get('generated_at', '--'))}<br>测试集：{esc(', '.join(suite_names))} ｜ PDF 模式：{esc(summary.get('pdf_mode', 'balanced'))}</div>
          </div>
          <span class="status-badge {health_class}">{health_label}</span>
        </div>
        <div class="tag-row">{''.join(insight_tags)}</div>
        <div class="nav">
          <a href="#suite">测试集表现</a>
          <a href="#format">格式表现</a>
          <a href="#engine">解析链路</a>
          <a href="#errors">错误与建议</a>
          <a href="#risk">风险样本</a>
        </div>
      </div>
      <div class="hero-progress">
        <div class="lbl">整体成功率</div>
        <div class="big-metric">{fmt_number(success_rate)}%</div>
        <div class="sub-metric">成功 {fmt_number(success)} / 失败 {fmt_number(failed)} / 预期匹配 {fmt_number(summary.get('expected_matched', 0))}/{fmt_number(summary.get('expected_checked', 0))}</div>
        <div class="hero-bar"><div class="hero-bar-fill" style="width:{min(success_rate, 100)}%"></div></div>
        <div class="hero-split">
          <div class="split-box"><span class="lbl">平均耗时</span><strong>{fmt_number(summary.get('avg_processing_ms', 0))}</strong></div>
          <div class="split-box"><span class="lbl">总字符</span><strong>{fmt_number(summary.get('char_count_total', 0))}</strong></div>
          <div class="split-box"><span class="lbl">总表格</span><strong>{fmt_number(summary.get('table_count_total', 0))}</strong></div>
        </div>
      </div>
    </div>

    <div class="cards">
      <div class="card hero-card"><div class="lbl">总样本</div><div class="val">{fmt_number(total)}</div><div class="subval">覆盖 {fmt_number(len(suite_names))} 个测试集</div></div>
      <div class="card hero-card"><div class="lbl">成功</div><div class="val">{fmt_number(success)}</div><div class="subval">通过率 {fmt_number(success_rate)}%</div></div>
      <div class="card hero-card"><div class="lbl">失败</div><div class="val">{fmt_number(failed)}</div><div class="subval">错误分类 {fmt_number(len(error_counts))} 种</div></div>
      <div class="card hero-card"><div class="lbl">异常项</div><div class="val">{fmt_number(summary.get('unexpected_count', 0))}</div><div class="subval">与预期不符的样本</div></div>
      <div class="card hero-card"><div class="lbl">平均字符数</div><div class="val">{fmt_number(summary.get('avg_char_count', 0))}</div><div class="subval">单样本平均产出</div></div>
      <div class="card hero-card"><div class="lbl">格式数</div><div class="val">{fmt_number(len(format_summary))}</div><div class="subval">当前报告覆盖的后缀类型</div></div>
    </div>

    <section id="suite">
      <h2>测试集表现</h2>
      <div class="suite-grid">
        {''.join(suite_cards) or '<div class="empty-card">暂无测试集统计。</div>'}
      </div>
    </section>

    <section id="format" class="two-col">
      <div class="panel">
        <h2>格式表现</h2>
        <table>
          <thead><tr><th>格式</th><th>总数</th><th>成功</th><th>失败</th><th>成功率</th><th>平均耗时(ms)</th><th>字符</th><th>表格</th><th>占比</th></tr></thead>
          <tbody>{render_table_rows(fmt_rows, 9)}</tbody>
        </table>
      </div>
      <div class="panel">
        <h2>格式分布</h2>
        <div class="bar-stack">
          {render_bar_rows(list(summary.get('format_counter', {}).items()), max_format_total, '暂无格式分布数据')}
        </div>
      </div>
    </section>

    <section id="engine" class="three-col">
      <div class="panel">
        <h2>错误分类</h2>
        <div class="bar-stack">
          {render_bar_rows(list(error_counts.items()), max_error_count, '当前没有错误分类数据')}
        </div>
      </div>
      <div class="panel">
        <h2>解析方式分布</h2>
        <div class="bar-stack">
          {render_bar_rows(list(parse_counts.items()), max_parse_count, '当前没有解析方式记录')}
        </div>
      </div>
      <div class="panel">
        <h2>OCR 引擎分布</h2>
        <div class="bar-stack">
          {render_bar_rows(list(ocr_counts.items()), max_ocr_count, '当前没有 OCR 引擎记录')}
        </div>
      </div>
    </section>

    <section id="errors" class="two-col">
      <div class="panel">
        <h2>错误样本与修复建议</h2>
        <div class="sample-grid">
          {error_sample_cards}
        </div>
      </div>
      <div class="panel">
        <h2>不符合预期的样本</h2>
        <table>
          <thead><tr><th>测试集</th><th>文件</th><th>预期</th><th>实际</th><th>结果</th></tr></thead>
          <tbody>{render_table_rows(unexpected_rows, 5, '当前没有与预期不符的样本')}</tbody>
        </table>
      </div>
    </section>

    <section id="risk" class="two-col">
      <div class="panel">
        <h2>失败样本</h2>
        <table>
          <thead><tr><th>测试集</th><th>文件</th><th>错误类型</th><th>解析链路</th><th>错误信息</th></tr></thead>
          <tbody>{render_table_rows(fail_rows, 5)}</tbody>
        </table>
      </div>
      <div class="panel">
        <h2>最慢样本</h2>
        <table>
          <thead><tr><th>测试集</th><th>文件</th><th>耗时(ms)</th><th>字符数</th><th>解析方式</th><th>OCR</th></tr></thead>
          <tbody>{render_table_rows(slow_rows, 6)}</tbody>
        </table>
      </div>
    </section>
  </div>
</body>
</html>"""
    output_path.write_text(html, encoding="utf-8")
    return output_path


def main() -> int:
    args = parse_args()

    suite_paths = [resolve_suite_path(suite) for suite in args.suites]
    suite_paths = [path for path in suite_paths if path.exists() and path.is_dir()]
    if not suite_paths:
        print("未找到可测试目录。")
        return 1

    cases = iter_files(suite_paths)
    if not cases:
        print("测试目录中没有找到可测试文件。")
        return 1

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_dir = Path(args.report_root) / f"batch_test_{timestamp}"
    report_dir.mkdir(parents=True, exist_ok=True)

    processor = DocFlowProcessor()
    records = []

    print(f"开始批量测试，共 {len(cases)} 个文件...")
    for index, (suite_name, file_path) in enumerate(cases, start=1):
        print(f"[{index}/{len(cases)}] {suite_name} -> {file_path.name}")
        records.append(
            run_single_case(
                processor=processor,
                suite_name=suite_name,
                file_path=file_path,
                extract_keywords=args.keywords,
                pdf_mode=args.pdf_mode,
            )
        )

    summary = build_summary(records, pdf_mode=args.pdf_mode)
    json_path = write_json(report_dir, summary, records)
    csv_path = write_csv(report_dir, records)
    md_path = write_markdown(report_dir, summary, records, [path.name for path in suite_paths])
    html_path = write_html_dashboard(report_dir, summary, records, [path.name for path in suite_paths])

    print("批量测试完成。")
    print(f"- HTML 汇总页: {html_path}")
    print(f"- Markdown 报告: {md_path}")
    print(f"- JSON 结果: {json_path}")
    print(f"- CSV 汇总: {csv_path}")
    print(
        f"- 成功/失败: {summary['success']}/{summary['failed']}，"
        f"预期匹配: {summary['expected_matched']}/{summary['expected_checked']}"
    )

    if args.strict and summary["unexpected_count"] > 0:
        print("存在不符合预期的样本，按 strict 模式返回失败。")
        return 2

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
