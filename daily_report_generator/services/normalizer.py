from __future__ import annotations

import re
from datetime import date, datetime


TRANSFER_TYPE = "中转站"
CLEAN_TYPE = "密闭式清洁站"
SUPPORTED_TYPES = (TRANSFER_TYPE, CLEAN_TYPE)

IGNORE_ISSUE_TEXTS = {
    "",
    "无",
    "无问题",
    "没问题",
    "良好",
    "良好未发现问题",
    "未发现问题",
    "无异常",
    "正常",
    "合格",
    "1",
    "未开门",
    "升级改造",
}

PHOTO_GROUP_NAMES = {"问题照片", "问题图片", "问题照片及附件"}


def clean_cell(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def compact_text(value: str) -> str:
    return re.sub(r"\s+", "", value or "").strip("，,。；;：:")


def is_no_problem_text(value: str) -> bool:
    text = compact_text(value)
    return text in IGNORE_ISSUE_TEXTS


def normalize_problem_sentence(value: str) -> str:
    text = re.sub(r"\s+", "", value or "").strip()
    text = text.strip("，,。；;：:")
    return text


def build_summary(issue_texts: list[str]) -> str:
    problems: list[str] = []
    seen: set[str] = set()
    for text in issue_texts:
        if is_no_problem_text(text):
            continue
        normalized = normalize_problem_sentence(text)
        if not normalized or normalized in seen:
            continue
        problems.append(normalized)
        seen.add(normalized)

    if not problems:
        return "无问题。"
    joined = "；".join(f"（{index}）{problem}" for index, problem in enumerate(problems, start=1))
    return ensure_period(f"存在的问题是：{joined}")


def ensure_period(text: str) -> str:
    text = text.strip()
    if not text:
        return "无问题。"
    return text if text.endswith(("。", "！", "？")) else f"{text}。"


def parse_date_like(value: object) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value).strip()
    if not text:
        return None
    text = text.replace("/", "-").replace("年", "-").replace("月", "-").replace("日", "")
    text = re.sub(r"\s+", " ", text)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(text[: len(datetime.now().strftime(fmt))] if "%H" in fmt else text, fmt).date()
        except ValueError:
            continue
    match = re.search(r"(?P<year>\d{4})[-.]?(?P<month>\d{1,2})[-.]?(?P<day>\d{1,2})", text)
    if match:
        try:
            return date(int(match.group("year")), int(match.group("month")), int(match.group("day")))
        except ValueError:
            return None
    return None


def parse_datetime_like(value: object) -> datetime | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)

    text = str(value).strip()
    if not text:
        return None
    text = text.replace("/", "-").replace("年", "-").replace("月", "-").replace("日", "")
    text = re.sub(r"\s+", " ", text)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    parsed_date = parse_date_like(text)
    if parsed_date:
        return datetime(parsed_date.year, parsed_date.month, parsed_date.day)
    return None


def report_date_from_created_time(value: datetime | None) -> date | None:
    if value is None:
        return None
    if value.hour >= 12:
        from datetime import timedelta

        return (value + timedelta(days=1)).date()
    return value.date()


def normalize_station_type(value: str) -> str:
    text = compact_text(value)
    lower = text.lower()
    if "中转" in text or lower in {"transfer", "recycle", "2"}:
        return TRANSFER_TYPE
    if "清洁站" in text or "密闭" in text or lower in {"clean", "cleaning", "sealed", "1"}:
        return CLEAN_TYPE
    return value.strip()


def item_display_name(indicator_name: str, issue_text: str) -> str:
    indicator = indicator_name.strip()
    if indicator and not is_no_problem_text(indicator):
        return indicator
    problem = normalize_problem_sentence(issue_text)
    return problem or "检查情况"


def level4_point_display_name(value: str) -> str:
    text = (value or "").strip()
    for suffix in ("密闭式清洁站", "清洁站"):
        if text.endswith(suffix):
            text = text[: -len(suffix)].strip()
    return text


def chinese_section_number(index: int) -> str:
    numerals = "零一二三四五六七八九"
    if index <= 0:
        return str(index)
    if index < 10:
        return numerals[index]
    if index == 10:
        return "十"
    if index < 20:
        return f"十{numerals[index - 10]}"
    if index < 100:
        tens, ones = divmod(index, 10)
        return f"{numerals[tens]}十{numerals[ones] if ones else ''}"
    return str(index)
