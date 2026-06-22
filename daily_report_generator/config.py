from __future__ import annotations

import sys
from pathlib import Path


def get_project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


PROJECT_ROOT = get_project_root()
APP_ROOT = PROJECT_ROOT if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
UPLOAD_ROOT = APP_ROOT / "uploads"
OUTPUT_ROOT = PROJECT_ROOT / "output" / "日报"
DEFAULT_INPUT_DIR = PROJECT_ROOT / "input"
DEFAULT_TRANSFER_TEMPLATE = DEFAULT_INPUT_DIR / "中转站日报_jinja模板.docx"
DEFAULT_CLEAN_TEMPLATE = DEFAULT_INPUT_DIR / "密闭式清洁站日报_jinja模板.docx"
