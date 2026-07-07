#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from ai_brush_helper.auto_mapper import generate_mapping
from ai_brush_helper.dp_workbook_importer import import_dp_workbook
from ai_brush_helper.ppt_chart_inspector import inspect_ppt_charts, write_chart_inventory
from ai_brush_helper.ppt_renderer import render_enhanced


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the enhanced DP-to-PPT brush-data workflow.")
    parser.add_argument("--mode", choices=["all", "import-dp", "inspect-ppt", "auto-map", "render"], default="all")
    parser.add_argument("--excel", type=Path, help="DP workbook path, required for all/import-dp.")
    parser.add_argument("--sheet", default="DP_问卷", help="DP result sheet name.")
    parser.add_argument("--pptx", type=Path, help="PPT template path, required for all/inspect-ppt/render.")
    parser.add_argument("--spec", type=Path, help="Optional metric/spec rules file.")
    parser.add_argument("--wave", default="", help="Source data wave label, for example 26h1.")
    parser.add_argument("--target-wave", default="", help="PPT target wave label, for example 26W1.")
    parser.add_argument("--out", type=Path, required=True, help="Output directory.")
    parser.add_argument("--mapping-spec", type=Path, help="Enhanced mapping JSON path for render mode.")
    parser.add_argument("--import-dir", type=Path, help="Existing import output dir for auto-map/render.")
    parser.add_argument("--inspect-dir", type=Path, help="Existing PPT inspect output dir for auto-map.")
    parser.add_argument("--output-pptx", type=Path, help="Rendered PPTX path.")
    parser.add_argument("--min-confidence", choices=["low", "medium", "high"], default="medium")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    args.out.mkdir(parents=True, exist_ok=True)

    import_dir = args.import_dir or (args.out / "01_dp_import")
    inspect_dir = args.inspect_dir or (args.out / "02_ppt_inspect")
    mapping_dir = args.out / "03_mapping"
    render_dir = args.out / "04_render"
    target_wave = args.target_wave or args.wave
    summary: dict[str, object] = {"mode": args.mode, "out": str(args.out)}

    if args.mode in {"all", "import-dp"}:
        if not args.excel:
            raise SystemExit("--excel is required.")
        summary["dp_import"] = import_dp_workbook(args.excel, import_dir, sheet_name=args.sheet, wave=args.wave)

    if args.mode in {"all", "inspect-ppt"}:
        if not args.pptx:
            raise SystemExit("--pptx is required.")
        chart_rows = inspect_ppt_charts(args.pptx)
        summary["ppt_inspect"] = write_chart_inventory(chart_rows, inspect_dir)

    if args.mode in {"all", "auto-map"}:
        summary["auto_map"] = generate_mapping(import_dir, inspect_dir, mapping_dir, wave=target_wave, spec=args.spec)

    if args.mode in {"all", "render"}:
        if not args.pptx:
            raise SystemExit("--pptx is required.")
        mapping_spec = args.mapping_spec or (mapping_dir / "mapping_spec.enhanced.json")
        output_pptx = args.output_pptx or (render_dir / f"{args.pptx.stem}.enhanced.rendered.pptx")
        render_dir.mkdir(parents=True, exist_ok=True)
        summary["render"] = render_enhanced(
            args.pptx,
            import_dir,
            mapping_spec,
            output_pptx,
            render_dir,
            min_confidence=args.min_confidence,
        )

    (args.out / "summary.enhanced.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

