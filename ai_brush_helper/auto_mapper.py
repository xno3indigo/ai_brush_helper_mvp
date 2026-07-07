from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Any

from ai_brush_helper.common import clean_text, normalize_token_text, read_csv, tokenize, write_csv


BRAND_ALIASES = {
    "百泽安": ["百泽安", "替雷利珠", "百济神州"],
    "达伯舒": ["达伯舒", "信迪利", "信达生物"],
    "艾瑞卡": ["艾瑞卡", "卡瑞利珠", "恒瑞"],
    "拓益": ["拓益", "特瑞普利", "君实"],
    "可瑞达": ["可瑞达", "帕博利珠", "默沙东", "k药"],
    "欧狄沃": ["欧狄沃", "纳武利尤", "百时美施贵宝", "o药"],
    "汉斯状": ["汉斯状", "斯鲁利", "复宏汉霖"],
    "择捷美": ["择捷美", "舒格利"],
}


def load_manifest(import_dir: Path) -> list[dict[str, Any]]:
    return json.loads((import_dir / "_manifest.json").read_text(encoding="utf-8"))


def load_chart_inventory(inspect_dir: Path) -> list[dict[str, Any]]:
    return json.loads((inspect_dir / "ppt_chart_inventory.json").read_text(encoding="utf-8"))


def table_row_labels(import_dir: Path, table: dict[str, Any]) -> list[str]:
    path = import_dir / table["file"]
    if not path.exists():
        return []
    rows = read_csv(path)
    return [clean_text(row.get("label", ""), 300) for row in rows if clean_text(row.get("label", ""), 300)]


def brand_hits(text: str) -> set[str]:
    normalized = normalize_token_text(text)
    hits: set[str] = set()
    for brand, aliases in BRAND_ALIASES.items():
        if any(normalize_token_text(alias) in normalized for alias in aliases):
            hits.add(brand)
    return hits


def table_profile(import_dir: Path, table: dict[str, Any]) -> dict[str, Any]:
    labels = table_row_labels(import_dir, table)
    text = " ".join(
        [
            table.get("table_name", ""),
            table.get("question_code", ""),
            table.get("question_title", ""),
            " ".join(table.get("columns", [])),
            " ".join(labels),
        ]
    )
    return {
        "table_name": table["table_name"],
        "question_code": table.get("question_code", ""),
        "question_title": table.get("question_title", ""),
        "file": table.get("file", ""),
        "tokens": tokenize(text),
        "brands": brand_hits(text),
        "labels": labels,
        "columns": table.get("columns", []),
    }


def parse_spec_page_rules(spec_text: str) -> dict[int, set[str]]:
    page_rules: dict[int, set[str]] = {}
    current_page: int | None = None
    for raw_line in spec_text.splitlines():
        page_match = re_match_page(raw_line)
        if page_match is not None:
            current_page = page_match
            page_rules.setdefault(current_page, set())
            continue
        if current_page is None:
            continue
        if "expected_questions" in raw_line:
            values = raw_line.split("[", 1)[-1].split("]", 1)[0]
            for item in values.split(","):
                code = normalize_token_text(item)
                if code:
                    page_rules[current_page].add(code)
    return page_rules


def re_match_page(raw_line: str) -> int | None:
    stripped = raw_line.strip()
    if not stripped.endswith(":"):
        return None
    value = stripped[:-1].strip().strip("'\"")
    if value.isdigit():
        return int(value)
    return None


def score_table_for_chart(
    chart: dict[str, Any],
    table: dict[str, Any],
    spec_text: str = "",
    page_rules: dict[int, set[str]] | None = None,
) -> tuple[int, list[str]]:
    chart_text = " ".join(
        [
            chart.get("slide_title", ""),
            chart.get("slide_text", ""),
            " ".join(chart.get("series_names", [])),
            " ".join(chart.get("category_values", [])),
            " ".join(chart.get("workbook_first_row", [])),
            " ".join(chart.get("workbook_first_col", [])),
            " ".join(chart.get("tokens", [])),
        ]
    )
    chart_tokens = tokenize(chart_text)
    chart_brands = brand_hits(chart_text)
    table_tokens = table["tokens"]
    overlap = sorted(chart_tokens & table_tokens)
    brand_overlap = sorted(chart_brands & table["brands"])
    score = 0
    reasons: list[str] = []
    expected_codes = (page_rules or {}).get(int(chart.get("slide", 0)), set())
    table_code = normalize_token_text(table.get("question_code", ""))
    table_name = normalize_token_text(table.get("table_name", ""))
    if expected_codes:
        if table_code in expected_codes or any(code and code in table_name for code in expected_codes):
            score += 35
            reasons.append("spec_page_expected_question")
        else:
            score -= 8
            reasons.append("spec_page_other_question_penalty")

    if overlap:
        score += min(6, len(overlap))
        reasons.append("token_overlap:" + ",".join(overlap[:10]))
    if brand_overlap:
        score += min(5, len(brand_overlap))
        reasons.append("brand_overlap:" + ",".join(brand_overlap))

    question_code = clean_text(table.get("question_code", ""), 50)
    if question_code and question_code.lower() in chart_text.lower():
        score += 20
        reasons.append("question_code_in_chart_context")

    title_tokens = tokenize(table.get("question_title", ""))
    title_overlap = sorted(title_tokens & chart_tokens)
    if title_overlap:
        score += min(12, len(title_overlap) * 2)
        reasons.append("question_title_overlap:" + ",".join(title_overlap[:8]))

    metric_score, metric_reasons = score_metric_intent(chart_text, table)
    if metric_score:
        score += metric_score
        reasons.extend(metric_reasons)

    first_row_col_text = " ".join(chart.get("workbook_first_row", []) + chart.get("workbook_first_col", []))
    label_hits = [
        label
        for label in table.get("labels", [])
        if label and normalize_token_text(label) and normalize_token_text(label) in normalize_token_text(first_row_col_text)
    ]
    if label_hits:
        score += min(16, len(label_hits) * 2)
        reasons.append("embedded_label_overlap:" + ",".join(label_hits[:8]))

    if chart.get("shape_guess") != "unknown":
        score += 2
        reasons.append("target_shape_detected")

    if spec_text and table.get("question_code", "") and table.get("question_code", "") in spec_text:
        score += 4
        reasons.append("question_code_in_spec")

    return score, reasons


def score_metric_intent(chart_text: str, table: dict[str, Any]) -> tuple[int, list[str]]:
    text = normalize_token_text(chart_text)
    title = normalize_token_text(table.get("question_title", ""))
    code = normalize_token_text(table.get("question_code", ""))
    rules = [
        (["b1a", "第一提及"], ["第一提及", "tom"], "metric_first_mention"),
        (["b1b", "其他还有吗"], ["其他自发", "自发提及"], "metric_other_spontaneous"),
        (["b1c", "提示后知晓"], ["总知晓", "知晓"], "metric_awareness"),
        (["c1a", "处方使用"], ["p3m处方过", "处方过"], "metric_prescribed"),
        (["c1b", "经常处方"], ["p3m经常处方", "经常处方"], "metric_regular"),
        (["c1c", "最常处方"], ["p3m最常处方", "最常处方"], "metric_most"),
        (["c2", "处方占比"], ["份额", "share", "soc"], "metric_share"),
        (["nps"], ["nps", "promotor", "detractor"], "metric_nps"),
        (["e1a", "拜访覆盖"], ["覆盖", "sov"], "metric_call_coverage"),
        (["e2", "拜访次数"], ["频率", "次数"], "metric_call_frequency"),
        (["f1", "活动接触次数"], ["活动", "会议", "接触次数"], "metric_activity"),
    ]
    for source_terms, chart_terms, reason in rules:
        source_hit = any(normalize_token_text(term) in title or normalize_token_text(term) == code for term in source_terms)
        chart_hit = any(normalize_token_text(term) in text for term in chart_terms)
        if source_hit and chart_hit:
            return 18, [reason]
    return 0, []


def confidence_from_score(score: int) -> str:
    if score >= 30:
        return "high"
    if score >= 18:
        return "medium"
    if score >= 8:
        return "low"
    return "no_match"


def generate_mapping(import_dir: Path, inspect_dir: Path, out_dir: Path, wave: str = "", spec: Path | None = None) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest = load_manifest(import_dir)
    charts = load_chart_inventory(inspect_dir)
    spec_text = spec.read_text(encoding="utf-8") if spec and spec.exists() else ""
    page_rules = parse_spec_page_rules(spec_text)
    profiles = [table_profile(import_dir, table) for table in manifest]

    bindings: list[dict[str, Any]] = []
    review_rows: list[dict[str, Any]] = []
    for chart in charts:
        scored = []
        for table in profiles:
            score, reasons = score_table_for_chart(chart, table, spec_text, page_rules)
            scored.append((score, reasons, table))
        scored.sort(key=lambda item: (-item[0], item[2]["table_name"]))
        best_score, best_reasons, best_table = scored[0]
        confidence = confidence_from_score(best_score)
        binding = {
            "page": chart["slide"],
            "chart_sort": chart["chart_sort"],
            "chart_path": chart["chart_path"],
            "embedded_workbook": chart.get("embedded_workbook", ""),
            "embedded_status": chart.get("embedded_status", ""),
            "chart_type": chart.get("chart_type", ""),
            "target_shape": chart.get("shape_guess", "unknown"),
            "series_count": chart.get("series_count", 0),
            "source_table": best_table["table_name"] if confidence != "no_match" else "",
            "source_file": best_table["file"] if confidence != "no_match" else "",
            "source_question_code": best_table["question_code"] if confidence != "no_match" else "",
            "source_question_title": best_table["question_title"] if confidence != "no_match" else "",
            "segment_filter": "Total",
            "target_wave": wave,
            "confidence": confidence,
            "score": best_score,
            "reason": best_reasons,
        }
        bindings.append(binding)
        review_rows.append(
            {
                "page": binding["page"],
                "chart_sort": binding["chart_sort"],
                "chart_path": binding["chart_path"],
                "embedded_status": binding["embedded_status"],
                "target_shape": binding["target_shape"],
                "source_table": binding["source_table"],
                "source_question_code": binding["source_question_code"],
                "score": binding["score"],
                "confidence": binding["confidence"],
                "reason": ";".join(binding["reason"]),
                "slide_title": chart.get("slide_title", ""),
                "series_names": " | ".join(chart.get("series_names", [])),
                "category_values": " | ".join(chart.get("category_values", [])[:30]),
                "workbook_first_row": " | ".join(chart.get("workbook_first_row", [])[:30]),
                "workbook_first_col": " | ".join(chart.get("workbook_first_col", [])[:30]),
            }
        )

    spec_doc = {"wave": wave, "bindings": bindings}
    (out_dir / "mapping_spec.enhanced.json").write_text(json.dumps(spec_doc, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "mapping_spec.enhanced.yaml").write_text(to_review_yaml(spec_doc), encoding="utf-8")
    write_csv(
        out_dir / "mapping_review.enhanced.csv",
        review_rows,
        [
            "page",
            "chart_sort",
            "chart_path",
            "embedded_status",
            "target_shape",
            "source_table",
            "source_question_code",
            "score",
            "confidence",
            "reason",
            "slide_title",
            "series_names",
            "category_values",
            "workbook_first_row",
            "workbook_first_col",
        ],
    )
    counts: dict[str, int] = {}
    for binding in bindings:
        counts[binding["confidence"]] = counts.get(binding["confidence"], 0) + 1
    return {"bindings": len(bindings), "confidence_counts": counts, "out_dir": str(out_dir)}


def to_review_yaml(spec_doc: dict[str, Any]) -> str:
    lines = ["# 增强版自动映射 spec，供审查和 enhanced renderer 使用。", f"wave: {spec_doc.get('wave', '')}", "bindings:"]
    for binding in spec_doc["bindings"]:
        lines.append(f"  - page: {binding['page']}")
        lines.append(f"    chart_sort: {binding['chart_sort']}")
        lines.append(f"    chart_path: {json.dumps(binding['chart_path'], ensure_ascii=False)}")
        lines.append(f"    embedded_workbook: {json.dumps(binding['embedded_workbook'], ensure_ascii=False)}")
        lines.append(f"    target_shape: {json.dumps(binding['target_shape'], ensure_ascii=False)}")
        lines.append(f"    source_table: {json.dumps(binding['source_table'], ensure_ascii=False)}")
        lines.append(f"    source_question_code: {json.dumps(binding['source_question_code'], ensure_ascii=False)}")
        lines.append(f"    segment_filter: {json.dumps(binding['segment_filter'], ensure_ascii=False)}")
        lines.append(f"    confidence: {binding['confidence']}")
        lines.append(f"    score: {binding['score']}")
    return "\n".join(lines) + "\n"
