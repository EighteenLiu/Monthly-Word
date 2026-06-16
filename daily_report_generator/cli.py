from __future__ import annotations

import argparse
import shutil
import uuid
from datetime import date
from pathlib import Path

from .config import DEFAULT_CLEAN_TEMPLATE, DEFAULT_TRANSFER_TEMPLATE, OUTPUT_ROOT
from .services.aggregator import aggregate_records, preview_payload
from .services.ledger_reader import read_ledger
from .services.normalizer import CLEAN_TYPE, TRANSFER_TYPE, normalize_station_type
from .services.renderer import make_zip, render_reports


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="从检查台账生成中转站/密闭式清洁站日报")
    parser.add_argument("--ledger", type=Path, required=True, help="检查台账 .xls/.xlsx")
    parser.add_argument("--date", required=True, help="检查日期，例如 2026-06-05")
    parser.add_argument("--transfer-template", type=Path, default=DEFAULT_TRANSFER_TEMPLATE, help="中转站日报 Jinja docx 模板")
    parser.add_argument("--clean-template", type=Path, default=DEFAULT_CLEAN_TEMPLATE, help="密闭式清洁站日报 Jinja docx 模板")
    parser.add_argument("--types", nargs="+", default=[TRANSFER_TYPE, CLEAN_TYPE], help="日报类型：中转站 密闭式清洁站")
    parser.add_argument("--zip", action="store_true", help="将多个日报打包为 zip")
    parser.add_argument("--output", type=Path, default=OUTPUT_ROOT, help="输出目录")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    report_date = date.fromisoformat(args.date)
    work_dir = args.output / "_cli_work" / uuid.uuid4().hex
    parsed = read_ledger(args.ledger, work_dir)
    result = aggregate_records(parsed.records, report_date)
    print(preview_payload(result))
    types = [normalize_station_type(value) for value in args.types]
    files = render_reports(result, args.transfer_template, args.clean_template, args.output, types)
    if args.zip or len(files) > 1:
        files = [make_zip(files, args.output, report_date)]
    for file in files:
        print(f"[完成] {file}")
    shutil.rmtree(work_dir, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
