"""
One-command entry point for the monthly bulletin workflow.

Conversion is handled inside monthly_generator.py. It reuses existing .docx
files when they are already up to date and converts only missing or stale .doc
daily reports.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from monthly_generator import (
    DEFAULT_CONVERTED_DIR,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_RAW_INPUT_DIR,
    DEFAULT_STATION_TYPE,
    ConversionError,
    generate_monthly_report,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="执行日报转换检查并生成月度检查通报")
    parser.add_argument("--raw-input", type=Path, default=DEFAULT_RAW_INPUT_DIR, help="原始 .doc 日报根目录")
    parser.add_argument("--converted-output", type=Path, default=DEFAULT_CONVERTED_DIR, help="转换后 .docx 日报根目录")
    parser.add_argument("--monthly-output", type=Path, default=DEFAULT_OUTPUT_DIR, help="月报输出目录")
    parser.add_argument("--month", required=True, help="月份，例如 4、04 或 2026-04")
    parser.add_argument("--year", type=int, help="报告年份，默认从 --month 解析，解析不到则为 2026")
    parser.add_argument("--station-type", default=DEFAULT_STATION_TYPE, help="日报子目录名称，默认：清洁站")
    parser.add_argument(
        "--engine",
        choices=("auto", "word", "libreoffice"),
        default="auto",
        help="doc 转 docx 引擎，默认 auto",
    )
    parser.add_argument("--skip-convert", action="store_true", help="跳过 doc 转 docx 检查，直接汇总现有 docx")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        output_path = generate_monthly_report(
            raw_input_dir=args.raw_input,
            converted_dir=args.converted_output,
            output_dir=args.monthly_output,
            month=args.month,
            year=args.year,
            station_type=args.station_type,
            engine=args.engine,
            skip_convert=args.skip_convert,
        )
    except (ConversionError, FileNotFoundError, ValueError) as exc:
        print(f"[失败] {exc}", file=sys.stderr)
        return 1

    print(f"[完成] 月报已生成：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
