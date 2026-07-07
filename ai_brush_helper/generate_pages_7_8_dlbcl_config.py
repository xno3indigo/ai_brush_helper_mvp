#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pymysql

from ai_brush_helper.ppt_chart_inspector import inspect_ppt_charts, write_chart_inventory


DISPLAY_NAME_RULES = [
    ("赞必佳", "赞必佳"),
    ("免疫", "免疫+化疗"),
    ("福可维", "福可维"),
    ("安泰适", "安泰适"),
]

DB_CONFIG_TABLE_COLUMNS = [
    "id",
    "last_updated_time",
    "name",
    "base_f",
    "fix_order",
    "fix_order_column",
    "page",
    "refresh_axis",
    "sort",
    "sort_column",
    "sort_order",
    "group_sign",
    "pivot",
    "group_id",
    "view_type",
    "original_name",
    "data_source_id",
    "n_value",
    "row_filters",
    "created_by",
    "last_updated_by",
    "created_time",
]

DB_CONFIG_FIELD_COLUMNS = [
    "table_id",
    "origin_name",
    "name",
    "column_index",
    "column_reference",
    "database_column_name",
    "type",
    "size",
    "n_value",
    "template_field",
    "created_time",
    "created_time_ms",
]


def connect(args: argparse.Namespace) -> pymysql.connections.Connection:
    return pymysql.connect(
        host=args.host,
        port=args.port,
        user=args.user,
        password=args.password,
        database=args.database,
        charset="utf8mb4",
    )


def table_columns(conn: pymysql.connections.Connection, database: str, table: str) -> list[str]:
    with conn.cursor(pymysql.cursors.DictCursor) as cur:
        cur.execute(
            """
            select column_name
            from information_schema.columns
            where table_schema=%s and table_name=%s
            order by ordinal_position
            """,
            (database, table),
        )
        return [row["column_name"] for row in cur.fetchall()]


def table_sample(conn: pymysql.connections.Connection, table: str, limit: int = 5) -> list[dict[str, Any]]:
    with conn.cursor(pymysql.cursors.DictCursor) as cur:
        cur.execute(f"select * from `{table}` limit {int(limit)}")
        return list(cur.fetchall())


def result_table_names(conn: pymysql.connections.Connection) -> list[str]:
    with conn.cursor() as cur:
        cur.execute("show tables")
        return [
            str(row[0])
            for row in cur.fetchall()
            if not str(row[0]).startswith("bh_") and not str(row[0]).startswith("sys_")
        ]


def build_table_profiles(conn: pymysql.connections.Connection, database: str) -> list[dict[str, Any]]:
    profiles = []
    for table in result_table_names(conn):
        columns = table_columns(conn, database, table)
        sample = table_sample(conn, table, 3)
        sample_text = json.dumps(sample, ensure_ascii=False, default=str)
        profiles.append(
            {
                "name": table,
                "columns": columns,
                "sample": sample,
                "text": " ".join([table, " ".join(columns), sample_text]).lower(),
            }
        )
    return profiles


def score_profile(profile: dict[str, Any], required_columns: set[str], name_terms: list[str]) -> tuple[int, list[str]]:
    score = 0
    reasons: list[str] = []
    name = profile["name"].lower()
    columns = set(profile["columns"])
    missing = sorted(required_columns - columns)
    if not missing:
        score += 40
        reasons.append("required_columns_exist")
    else:
        score -= 30 * len(missing)
        reasons.append("missing_columns:" + ",".join(missing))
    for term in name_terms:
        if term.lower() in name:
            score += 15
            reasons.append(f"name_contains:{term}")
    if "_total_" in name or name.endswith("_total_26w2"):
        score += 8
        reasons.append("prefer_total_table")
    return score, reasons


def best_profile(
    profiles: list[dict[str, Any]],
    required_columns: set[str],
    name_terms: list[str],
) -> tuple[dict[str, Any] | None, list[str], int]:
    scored = []
    for profile in profiles:
        score, reasons = score_profile(profile, required_columns, name_terms)
        scored.append((score, profile["name"], profile, reasons))
    scored.sort(key=lambda item: (-item[0], item[1]))
    if not scored or scored[0][0] < 20:
        return None, ["no_database_table_passed_threshold"], 0
    score, _, profile, reasons = scored[0]
    return profile, reasons, score


def display_name(column: str) -> str:
    fixed = {
        "brand": "品牌",
        "attr": "属性",
        "ranking": "排名",
        "n": "N",
        "tom": "TOM",
        "NPS": "NPS",
        "sov": "SOV",
        "soc": "SOC",
    }
    if column in fixed:
        return fixed[column]
    for token, name in DISPLAY_NAME_RULES:
        if token in column:
            return name
    return re.sub(r"_\d+$", "", column)


def group_sign(table_name: str) -> str:
    match = re.match(r"(.+?)_(?:total|[^_]+)_26w\d+$", table_name)
    if match:
        return match.group(1) + "_total"
    return re.sub(r"_26w\d+$", "", table_name)


def chart_context(chart: dict[str, Any]) -> str:
    return " ".join(
        [
            str(chart.get("slide_title", "")),
            str(chart.get("slide_text", "")),
            " ".join(chart.get("series_names", [])),
            " ".join(chart.get("category_values", [])),
            " ".join(chart.get("workbook_first_row", [])),
            " ".join(chart.get("workbook_first_col", [])),
        ]
    ).lower()


def chart_number(chart: dict[str, Any], key: str) -> int:
    try:
        return int(chart.get(key, "") or 0)
    except ValueError:
        return 0


def infer_slide7_visual_role(chart: dict[str, Any]) -> dict[str, Any] | None:
    x = chart_number(chart, "chart_x")
    y = chart_number(chart, "chart_y")
    if x < 2_000_000 and y > 3_000_000:
        return {
            "metric": "adoption_rate",
            "chart_role": "Adoption Rate",
            "required_columns": {"brand", "penetration", "n"},
            "name_terms": ["adoption", "penetration", "total", "26w2"],
            "target_shape": "adoption_rate_table_or_chart",
        }
    if 4_000_000 <= x < 8_000_000 and y < 3_000_000:
        return {
            "metric": "NPS",
            "chart_role": "NPS",
            "required_columns": {"brand", "NPS", "n"},
            "name_terms": ["nps_score", "total", "26w2"],
            "target_shape": "brand_rows_single_metric",
        }
    if x >= 8_000_000 and y < 3_000_000:
        return {
            "metric": "sov",
            "chart_role": "SOV",
            "required_columns": {"brand", "sov", "n"},
            "name_terms": ["sov", "total", "26w2"],
            "target_shape": "brand_rows_single_metric",
        }
    if 4_000_000 <= x < 8_000_000 and y > 3_000_000:
        return {
            "metric": "tom",
            "chart_role": "TOM",
            "required_columns": {"brand", "tom", "n"},
            "name_terms": ["brand_awareness", "total", "26w2"],
            "target_shape": "brand_rows_wave_columns",
        }
    if x >= 8_000_000 and y > 3_000_000:
        return {
            "metric": "soc",
            "chart_role": "SOC",
            "required_columns": {"brand", "soc", "n"},
            "name_terms": ["soc", "share_of_choice", "total", "26w2"],
            "target_shape": "brand_rows_single_metric",
        }
    return None


def fields_for_metric(profile: dict[str, Any], metric: str) -> list[tuple[str, str]]:
    if metric == "raw_score" or metric == "score_ratio":
        value_columns = [column for column in profile["columns"] if column not in {"id", "attr", "ranking"}]
        return [("attr", "属性"), ("ranking", "排名")] + [(column, display_name(column)) for column in value_columns]
    dimension = "brand" if "brand" in profile["columns"] else profile["columns"][0]
    fields = [(dimension, display_name(dimension))]
    if metric in profile["columns"]:
        fields.append((metric, display_name(metric)))
    if "n" in profile["columns"]:
        fields.append(("n", "N"))
    return fields


def infer_bindings(page_charts: list[dict[str, Any]], profiles: list[dict[str, Any]]) -> list[dict[str, Any]]:
    bindings: list[dict[str, Any]] = []
    for chart in page_charts:
        page = int(chart.get("slide", 0))
        sort = int(chart.get("chart_sort", 0))
        profile: dict[str, Any] | None = None
        metric = ""
        target_shape = ""
        chart_role = ""
        fields: list[tuple[str, str]] = []
        reasons: list[str] = []
        score = 0

        if page == 7:
            role = infer_slide7_visual_role(chart)
            if not role:
                continue
            metric = role["metric"]
            target_shape = role["target_shape"]
            chart_role = role["chart_role"]
            profile, reasons, score = best_profile(profiles, role["required_columns"], role["name_terms"])
            reasons.append(
                f"slide7_visual_role_by_position:x={chart.get('chart_x')},y={chart.get('chart_y')}->{chart_role}"
            )
            if profile:
                fields = fields_for_metric(profile, metric)
        elif page == 8 and sort in {1, 2}:
            profile, reasons, score = best_profile(profiles, {"attr", "ranking"}, ["brand_equity", "total", "26w2"])
            metric = "raw_score" if sort == 1 else "score_ratio"
            target_shape = "attribute_rows_product_score_columns" if sort == 1 else "attribute_rows_product_score_ratio_columns"
            chart_role = "品牌属性打分 原始分数" if sort == 1 else "品牌属性打分 比例"
            if profile:
                fields = fields_for_metric(profile, metric)
            if "scatter" in str(chart.get("chart_type", "")).lower():
                reasons.append("ppt_chart_type_scatter")
            reasons.append(f"brand_equity_metric_inferred_from_chart_order:{sort}->{metric}")

        if not profile:
            bindings.append(
                {
                    "page": page,
                    "sort": sort,
                    "name": "",
                    "metric": metric,
                    "group_sign": "",
                    "chart_role": chart_role,
                    "target_shape": target_shape,
                    "fields": [],
                    "reason": "; ".join(reasons + [f"no_matching_database_table; auto_score:{score}"]),
                }
            )
            continue
        bindings.append(
            {
                "page": page,
                "sort": sort,
                "name": profile["name"],
                "metric": metric,
                "group_sign": group_sign(profile["name"]),
                "chart_role": chart_role,
                "target_shape": target_shape,
                "fields": fields,
                "reason": "; ".join(reasons + [f"auto_score:{score}"]),
            }
        )
    return bindings


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fieldnames})


def none_if_blank(value: Any) -> Any:
    return None if value == "" else value


def config_table_counts(conn: pymysql.connections.Connection) -> dict[str, int]:
    counts: dict[str, int] = {}
    with conn.cursor(pymysql.cursors.DictCursor) as cur:
        for table in ["bh_database_table", "bh_database_table_field", "bh_charts_replaces"]:
            cur.execute(f"select count(*) as row_count from `{table}`")
            counts[table] = int(cur.fetchone()["row_count"])
    return counts


def next_database_table_id(conn: pymysql.connections.Connection) -> int:
    with conn.cursor(pymysql.cursors.DictCursor) as cur:
        cur.execute("select coalesce(max(id), 0) + 1 as next_id from `bh_database_table`")
        return int(cur.fetchone()["next_id"])


def remap_table_ids(
    db_tables: list[dict[str, Any]],
    db_fields: list[dict[str, Any]],
    start_id: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    id_map: dict[int, int] = {}
    mapped_tables: list[dict[str, Any]] = []
    for offset, row in enumerate(db_tables):
        old_id = int(row["id"])
        new_id = start_id + offset
        id_map[old_id] = new_id
        mapped = dict(row)
        mapped["id"] = new_id
        mapped_tables.append(mapped)

    mapped_fields: list[dict[str, Any]] = []
    for row in db_fields:
        mapped = dict(row)
        mapped["table_id"] = id_map[int(row["table_id"])]
        mapped_fields.append(mapped)
    return mapped_tables, mapped_fields


def insert_rows(
    conn: pymysql.connections.Connection,
    table: str,
    rows: list[dict[str, Any]],
    columns: list[str],
) -> None:
    if not rows:
        return
    placeholders = ", ".join(["%s"] * len(columns))
    column_sql = ", ".join(f"`{column}`" for column in columns)
    sql = f"insert into `{table}` ({column_sql}) values ({placeholders})"
    values = [[none_if_blank(row.get(column)) for column in columns] for row in rows]
    with conn.cursor() as cur:
        cur.executemany(sql, values)


def write_config_to_database(
    conn: pymysql.connections.Connection,
    db_tables: list[dict[str, Any]],
    db_fields: list[dict[str, Any]],
    replace_existing: bool = False,
) -> dict[str, Any]:
    before_counts = config_table_counts(conn)
    non_empty = {table: count for table, count in before_counts.items() if count}
    if non_empty and not replace_existing:
        return {
            "status": "skipped_non_empty_config_tables",
            "before_counts": before_counts,
            "message": "Use --replace-existing-config to delete existing config rows before inserting generated config.",
        }

    now = datetime.now()
    now_ms = int(now.timestamp() * 1000)
    with conn.cursor() as cur:
        if replace_existing:
            cur.execute("delete from `bh_database_table_field`")
            cur.execute("delete from `bh_charts_replaces`")
            cur.execute("delete from `bh_database_table`")

    start_id = next_database_table_id(conn)
    mapped_tables, mapped_fields = remap_table_ids(db_tables, db_fields, start_id)

    strict_tables: list[dict[str, Any]] = []
    for row in mapped_tables:
        strict = {column: row.get(column, "") for column in DB_CONFIG_TABLE_COLUMNS}
        strict["last_updated_time"] = now
        strict["created_time"] = now
        strict_tables.append(strict)

    strict_fields: list[dict[str, Any]] = []
    for row in mapped_fields:
        strict = {column: row.get(column, "") for column in DB_CONFIG_FIELD_COLUMNS}
        strict["created_time"] = now
        strict["created_time_ms"] = now_ms
        strict_fields.append(strict)

    insert_rows(conn, "bh_database_table", strict_tables, DB_CONFIG_TABLE_COLUMNS)
    insert_rows(conn, "bh_database_table_field", strict_fields, DB_CONFIG_FIELD_COLUMNS)
    conn.commit()

    return {
        "status": "inserted",
        "before_counts": before_counts,
        "after_counts": config_table_counts(conn),
        "inserted_bh_database_table": len(strict_tables),
        "inserted_bh_database_table_field": len(strict_fields),
        "table_id_start": start_id,
    }


def find_chart(charts: list[dict[str, Any]], page: int, sort: int) -> dict[str, Any]:
    for chart in charts:
        if int(chart.get("slide", 0)) == page and int(chart.get("chart_sort", 0)) == sort:
            return chart
    return {}


def generate_config(args: argparse.Namespace) -> dict[str, Any]:
    args.out.mkdir(parents=True, exist_ok=True)
    conn = connect(args)
    charts = inspect_ppt_charts(args.pptx)
    page_charts = [chart for chart in charts if int(chart.get("slide", 0)) in {7, 8}]
    write_chart_inventory(page_charts, args.out / "ppt_inspect")
    table_profiles = build_table_profiles(conn, args.database)
    bindings = infer_bindings(page_charts, table_profiles)

    db_tables: list[dict[str, Any]] = []
    db_fields: list[dict[str, Any]] = []
    review_rows: list[dict[str, Any]] = []
    spec_bindings: list[dict[str, Any]] = []
    sample_bundle: dict[str, Any] = {}

    table_id = 1
    for binding in bindings:
        chart = find_chart(page_charts, binding["page"], binding["sort"])
        if not binding["name"]:
            review_rows.append(
                {
                    "page": binding["page"],
                    "chart_sort": binding["sort"],
                    "chart_x": chart.get("chart_x", ""),
                    "chart_y": chart.get("chart_y", ""),
                    "chart_path": chart.get("chart_path", ""),
                    "embedded_workbook": chart.get("embedded_workbook", ""),
                    "embedded_status": chart.get("embedded_status", ""),
                    "chart_type": chart.get("chart_type", ""),
                    "series_count": chart.get("series_count", ""),
                    "workbook_first_row": " | ".join(chart.get("workbook_first_row", [])),
                    "workbook_first_col": " | ".join(chart.get("workbook_first_col", [])),
                    "db_table": "",
                    "db_metric": binding["metric"],
                    "db_columns": "",
                    "missing_columns": "source_table_not_found",
                    "confidence": "needs_review",
                    "reason": binding["reason"],
                }
            )
            continue
        columns = table_columns(conn, args.database, binding["name"])
        missing_columns = [field[0] for field in binding["fields"] if field[0] not in columns]
        db_tables.append(
            {
                "id": table_id,
                "name": binding["name"],
                "base_f": "",
                "fix_order": "",
                "fix_order_column": "",
                "page": binding["page"],
                "refresh_axis": "chart",
                "sort": binding["sort"],
                "sort_column": "",
                "sort_order": "",
                "group_sign": binding["group_sign"],
                "pivot": 0,
                "group_id": "",
                "view_type": "chart",
                "original_name": binding["chart_role"],
                "data_source_id": args.database,
                "n_value": "n",
                "row_filters": json.dumps({"metric": binding["metric"], "target_shape": binding["target_shape"]}, ensure_ascii=False),
                "confidence": "high" if not missing_columns and chart else "needs_review",
                "reason": binding["reason"],
            }
        )
        for column_index, (column, display_name) in enumerate(binding["fields"], start=1):
            db_fields.append(
                {
                    "table_id": table_id,
                    "origin_name": column,
                    "name": display_name,
                    "column_index": column_index,
                    "column_reference": "",
                    "database_column_name": column,
                    "type": "value" if column not in {"brand", "attr", "ranking", "n"} else "dimension",
                    "size": "",
                    "n_value": "n" if column == "n" else "",
                    "template_field": 0,
                    "confidence": "high" if column in columns else "missing",
                    "reason": "field exists in database table" if column in columns else "field missing in database table",
                }
            )

        review_rows.append(
            {
                "page": binding["page"],
                "chart_sort": binding["sort"],
                "chart_x": chart.get("chart_x", ""),
                "chart_y": chart.get("chart_y", ""),
                "chart_path": chart.get("chart_path", ""),
                "embedded_workbook": chart.get("embedded_workbook", ""),
                "embedded_status": chart.get("embedded_status", ""),
                "chart_type": chart.get("chart_type", ""),
                "series_count": chart.get("series_count", ""),
                "workbook_first_row": " | ".join(chart.get("workbook_first_row", [])),
                "workbook_first_col": " | ".join(chart.get("workbook_first_col", [])),
                "db_table": binding["name"],
                "db_metric": binding["metric"],
                "db_columns": " | ".join(columns),
                "missing_columns": " | ".join(missing_columns),
                "confidence": "high" if not missing_columns and chart else "needs_review",
                "reason": binding["reason"],
            }
        )

        spec_bindings.append(
            {
                "page": binding["page"],
                "chart_sort": binding["sort"],
                "chart_path": chart.get("chart_path", ""),
                "embedded_workbook": chart.get("embedded_workbook", ""),
                "source_database": args.database,
                "source_table": binding["name"],
                "source_metric": binding["metric"],
                "target_shape": binding["target_shape"],
                "fields": [{"db_column": column, "ppt_name": display_name} for column, display_name in binding["fields"]],
                "render_transform": "divide_score_by_100" if binding["metric"] == "score_ratio" else "",
                "confidence": "high" if not missing_columns and chart else "needs_review",
                "reason": binding["reason"],
            }
        )
        sample_bundle.setdefault(binding["name"], {"columns": columns, "sample": table_sample(conn, binding["name"], 5)})
        table_id += 1

    write_csv(
        args.out / "bh_database_table.generated.csv",
        db_tables,
        [
            "id",
            "name",
            "base_f",
            "fix_order",
            "fix_order_column",
            "page",
            "refresh_axis",
            "sort",
            "sort_column",
            "sort_order",
            "group_sign",
            "pivot",
            "group_id",
            "view_type",
            "original_name",
            "data_source_id",
            "n_value",
            "row_filters",
            "confidence",
            "reason",
        ],
    )
    write_csv(
        args.out / "bh_database_table_field.generated.csv",
        db_fields,
        [
            "table_id",
            "origin_name",
            "name",
            "column_index",
            "column_reference",
            "database_column_name",
            "type",
            "size",
            "n_value",
            "template_field",
            "confidence",
            "reason",
        ],
    )
    write_csv(
        args.out / "mapping_review.pages_7_8.csv",
        review_rows,
        [
            "page",
            "chart_sort",
            "chart_x",
            "chart_y",
            "chart_path",
            "embedded_workbook",
            "embedded_status",
            "chart_type",
            "series_count",
            "workbook_first_row",
            "workbook_first_col",
            "db_table",
            "db_metric",
            "db_columns",
            "missing_columns",
            "confidence",
            "reason",
        ],
    )
    (args.out / "mapping_spec.pages_7_8.json").write_text(
        json.dumps({"database": args.database, "bindings": spec_bindings}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (args.out / "mapping_spec.pages_7_8.yaml").write_text(to_yaml(args.database, spec_bindings), encoding="utf-8")
    (args.out / "db_samples.pages_7_8.json").write_text(json.dumps(sample_bundle, ensure_ascii=False, indent=2, default=str), encoding="utf-8")

    db_write_report = {"status": "not_requested"}
    if args.write_db:
        try:
            db_write_report = write_config_to_database(
                conn,
                db_tables,
                db_fields,
                replace_existing=args.replace_existing_config,
            )
        except Exception:
            conn.rollback()
            raise
    (args.out / "db_write_report.pages_7_8.json").write_text(
        json.dumps(db_write_report, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    conn.close()

    summary = {
        "database": args.database,
        "pptx": str(args.pptx),
        "bindings": len(spec_bindings),
        "table_rows": len(db_tables),
        "field_rows": len(db_fields),
        "candidate_db_tables": len(table_profiles),
        "high_confidence_bindings": sum(1 for row in review_rows if row["confidence"] == "high"),
        "needs_review_bindings": sum(1 for row in review_rows if row["confidence"] != "high"),
        "db_write": db_write_report,
        "out": str(args.out),
    }
    (args.out / "summary.pages_7_8.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def to_yaml(database: str, bindings: list[dict[str, Any]]) -> str:
    lines = [
        "# PPT 第 7/8 页与 tracking_dlbcl 数据库的局部配置 spec",
        f"database: {database}",
        "bindings:",
    ]
    for item in bindings:
        lines.append(f"  - page: {item['page']}")
        lines.append(f"    chart_sort: {item['chart_sort']}")
        lines.append(f"    chart_path: {json.dumps(item['chart_path'], ensure_ascii=False)}")
        lines.append(f"    embedded_workbook: {json.dumps(item['embedded_workbook'], ensure_ascii=False)}")
        lines.append(f"    source_table: {json.dumps(item['source_table'], ensure_ascii=False)}")
        lines.append(f"    source_metric: {json.dumps(item['source_metric'], ensure_ascii=False)}")
        lines.append(f"    target_shape: {json.dumps(item['target_shape'], ensure_ascii=False)}")
        if item.get("render_transform"):
            lines.append(f"    render_transform: {json.dumps(item['render_transform'], ensure_ascii=False)}")
        lines.append(f"    confidence: {item['confidence']}")
        lines.append("    fields:")
        for field in item["fields"]:
            lines.append(f"      - db_column: {json.dumps(field['db_column'], ensure_ascii=False)}")
            lines.append(f"        ppt_name: {json.dumps(field['ppt_name'], ensure_ascii=False)}")
    return "\n".join(lines) + "\n"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate page 7/8 config tables from PPT template and tracking_dlbcl database.")
    parser.add_argument("--pptx", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--host", default="192.168.20.7")
    parser.add_argument("--port", type=int, default=3306)
    parser.add_argument("--user", default="root")
    parser.add_argument("--password", default="root")
    parser.add_argument("--database", default="tracking_dlbcl")
    parser.add_argument("--write-db", action="store_true", help="Insert generated high-confidence config rows into bh_* config tables.")
    parser.add_argument(
        "--replace-existing-config",
        action="store_true",
        help="Delete existing bh_database_table, bh_database_table_field and bh_charts_replaces rows before writing.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    print(json.dumps(generate_config(args), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
