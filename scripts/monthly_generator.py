"""
Generate the monthly sealed cleaning station inspection bulletin.

The script first checks whether raw .doc daily reports have already been
converted to .docx. Existing converted files are reused when they are newer
than the corresponding raw file; only missing or stale files are converted.
Then it summarizes all converted daily reports for the requested month and
creates the monthly .docx bulletin.
"""

from __future__ import annotations

import argparse
import io
import re
import shutil
import subprocess
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_RAW_INPUT_DIR = PROJECT_ROOT / "01_原始日报"
DEFAULT_CONVERTED_DIR = PROJECT_ROOT / "02_转换后日报"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "04_输出月报"
DEFAULT_STATION_TYPE = "清洁站"
DEFAULT_YEAR = 2026

DAILY_NAME_RE = re.compile(r"(?P<month>\d{1,2})月(?P<day>\d{1,2})日")

CLEANING_ISSUE_CATEGORIES: list[tuple[str, tuple[str, ...]]] = [
    ("无称重系统或称重系统损坏", ("无称重系统", "称重系统损坏", "无称重", "称重设备损坏", "称重屏幕破损")),
    ("小型收集车混装混运", ("小型收集车混装混运", "混装混运")),
    ("箱体内垃圾混投", ("箱体内垃圾混投", "垃圾混投", "混投")),
    ("未开门运行", ("未开门运行", "未开门")),
    ("拒收单不准确", ("拒收单不准确", "拒收单")),
]

TRANSFER_ISSUE_CATEGORIES: list[tuple[str, tuple[str, ...]]] = [
    ("清运不及时、可回收物大量积存", ("清运不及时", "大量积存", "可回收物大量积存")),
    ("周边环境脏乱", ("周边环境脏乱", "环境脏乱")),
    ("无可回收价格表或价格表损坏", ("无可回收价格表", "无价格表", "价格表损坏", "价格表破损", "无可回收物价格表", "价目表", "价目价目表")),
    ("无备案公示", ("无备案公示", "无备案信息", "备案信息表", "备案公示")),
    ("消防水源不合格", ("消防水源不合格",)),
    ("无消防安全水源", ("无消防安全水源", "无消防水源")),
    ("无营业执照", ("无营业执照",)),
    ("无安全风险公告", ("无安全风险公告", "无公示牌", "安全风险公告")),
    ("称重系统损坏", ("称重系统损坏", "称重屏幕破损", "称重设备损坏", "计量称不能使用")),
    ("灭火器过期", ("灭火器过期",)),
    ("灭火器不合格", ("灭火器不合格",)),
    ("未按规定区域存放物品", ("未按规定区域存放", "未按规定存放", "暂存区未按规定存放", "堆放混乱")),
    ("无七禁收八不准承诺书", ("无七禁收八不准承诺书", "七禁收八不准")),
    ("配电箱处堆放杂物", ("配电箱处堆放杂物", "配电箱堆放杂物")),
    ("安全员未按时上岗", ("安全员未按时上岗",)),
    ("安全员无明显身份标识", ("安全员无明显身份标识", "全员无明显身份标识", "无明显身份标识")),
    ("未按时开门运行", ("未按时开门运行", "未开门运行", "未开门")),
    ("无企安安", ("无企安安", "企安安无法正常登录")),
    ("无灭蝇措施", ("无灭蝇措施",)),
]

STATION_PROFILES = {
    "清洁站": {
        "title": "西城区密闭式清洁站存在问题检查通报",
        "output_name": "西城区{year}年{month}月密闭式清洁站检查通报.docx",
        "subject": "密闭式清洁站",
        "attachment_subject": "密闭式清洁站",
        "detail_heading": "二、各街道情况",
        "categories": CLEANING_ISSUE_CATEGORIES,
        "unopened_category": "未开门运行",
        "non_problem_status_keywords": (),
    },
    "中转站": {
        "title": "西城区可回收物中转站检查情况通报",
        "output_name": "西城区{year}年{month}月中转站检查通报.docx",
        "subject": "可回收物中转站",
        "attachment_subject": "可回收物中转站",
        "detail_heading": "二、各街道案例",
        "categories": TRANSFER_ISSUE_CATEGORIES,
        "unopened_category": "未按时开门运行",
        "non_problem_status_keywords": ("升级改造", "施工停运", "停运"),
    },
}


class ConversionError(RuntimeError):
    """Raised when a daily report cannot be converted."""


@dataclass(frozen=True)
class ImageSpec:
    blob: bytes
    width_pt: float | None
    height_pt: float | None


@dataclass(frozen=True)
class ImageRow:
    images: tuple[ImageSpec, ...]
    centered: bool = True


@dataclass(frozen=True)
class DailyRecord:
    source: Path
    month: int
    day: int
    street: str
    station: str
    issue_text: str
    issue_counts: dict[str, int]
    image_rows: tuple[ImageRow, ...] = ()

    @property
    def date_label(self) -> str:
        return f"{self.month}月{self.day}日"

    @property
    def has_problem(self) -> bool:
        return normalize_issue_text(self.issue_text) != "无问题。"


@dataclass
class MergedRecord:
    station: str
    issue_texts: list[str]
    image_rows: list[ImageRow]


@dataclass
class ConversionSummary:
    converted: int = 0
    skipped: int = 0
    failed: list[tuple[Path, str]] = field(default_factory=list)


def iter_doc_files(input_dir: Path) -> list[Path]:
    if not input_dir.exists():
        return []
    return sorted(
        path
        for path in input_dir.rglob("*.doc")
        if path.is_file() and path.suffix.lower() == ".doc"
    )


def find_soffice() -> str | None:
    for executable in ("soffice", "libreoffice"):
        found = shutil.which(executable)
        if found:
            return found
    return None


def convert_with_word(source: Path, target: Path) -> None:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ConversionError("pywin32 is not installed") from exc

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(source.resolve()))
        target.parent.mkdir(parents=True, exist_ok=True)
        document.SaveAs2(str(target.resolve()), FileFormat=16)
    except Exception as exc:
        raise ConversionError(str(exc)) from exc
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def convert_with_libreoffice(source: Path, target: Path) -> None:
    soffice = find_soffice()
    if not soffice:
        raise ConversionError("LibreOffice/soffice was not found in PATH")

    target.parent.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to",
            "docx",
            "--outdir",
            str(target.parent.resolve()),
            str(source.resolve()),
        ],
        capture_output=True,
        text=True,
        check=False,
    )
    generated = target.parent / f"{source.stem}.docx"
    if result.returncode != 0 or not generated.exists():
        message = (result.stderr or result.stdout or "LibreOffice conversion failed").strip()
        raise ConversionError(message)
    if generated.resolve() != target.resolve():
        generated.replace(target)


def convert_one(source: Path, target: Path, engine: str) -> bool:
    if target.exists() and target.stat().st_mtime >= source.stat().st_mtime:
        return False

    engines = ["word", "libreoffice"] if engine == "auto" else [engine]
    errors: list[str] = []
    for selected in engines:
        try:
            if selected == "word":
                convert_with_word(source, target)
            elif selected == "libreoffice":
                convert_with_libreoffice(source, target)
            else:
                raise ConversionError(f"Unknown conversion engine: {selected}")
            return True
        except ConversionError as exc:
            errors.append(f"{selected}: {exc}")

    raise ConversionError("; ".join(errors))


def ensure_converted(
    raw_input_dir: Path,
    converted_dir: Path,
    station_type: str = DEFAULT_STATION_TYPE,
    engine: str = "auto",
) -> ConversionSummary:
    station_type = normalize_station_type(station_type)
    source_root = raw_input_dir / station_type
    if not source_root.exists() and raw_input_dir.name == station_type:
        source_root = raw_input_dir
    target_root = converted_dir / station_type
    summary = ConversionSummary()

    for source in iter_doc_files(source_root):
        relative = source.relative_to(source_root)
        target = (target_root / relative).with_suffix(".docx")
        try:
            was_converted = convert_one(source, target, engine)
        except ConversionError as exc:
            summary.failed.append((source, str(exc)))
            continue

        if was_converted:
            summary.converted += 1
            print(f"[转换] {source.name} -> {target.name}")
        else:
            summary.skipped += 1

    if summary.skipped:
        print(f"[跳过] {summary.skipped} 个日报已完成 doc -> docx 转换")
    if summary.converted:
        print(f"[完成] 本次新转换 {summary.converted} 个日报")
    if not summary.converted and not summary.skipped:
        print(f"[提示] 未发现原始 .doc 日报：{source_root}")
    return summary


def parse_month(value: str | int | None) -> int:
    if value is None:
        raise ValueError("请通过 --month 指定月份，例如 4 或 2026-04")
    text = str(value).strip()
    match = re.search(r"(?:(?:\d{4})[-年])?(?P<month>\d{1,2})(?:月)?$", text)
    if not match:
        raise ValueError(f"无法识别月份：{value}")
    month = int(match.group("month"))
    if not 1 <= month <= 12:
        raise ValueError(f"月份超出范围：{value}")
    return month


def parse_year(value: int | None, month_value: str | int | None) -> int:
    if value:
        return value
    if month_value:
        match = re.search(r"(?P<year>\d{4})[-年]", str(month_value))
        if match:
            return int(match.group("year"))
    return DEFAULT_YEAR


def previous_month(year: int, month: int) -> tuple[int, int]:
    if month == 1:
        return year - 1, 12
    return year, month - 1


def report_period(year: int, month: int) -> tuple[date, date]:
    start_year, start_month = previous_month(year, month)
    return date(start_year, start_month, 20), date(year, month, 19)


def record_date_for_period(record_month: int, record_day: int, start: date, end: date) -> date:
    if record_month == start.month:
        return date(start.year, record_month, record_day)
    if record_month == end.month:
        return date(end.year, record_month, record_day)
    return date(end.year, record_month, record_day)


def format_date_cn(value: date) -> str:
    return f"{value.year}年{value.month}月{value.day}日"


def daily_date_from_name(path: Path) -> tuple[int, int] | None:
    match = DAILY_NAME_RE.search(path.stem)
    if not match:
        return None
    return int(match.group("month")), int(match.group("day"))


def normalize_issue_text(text: str) -> str:
    text = re.sub(r"\s+", "", text or "")
    if not text or text in {"无问题", "未发现问题"}:
        return "无问题。"
    return text if text.endswith(("。", "！", "？")) else f"{text}。"


def split_street_station(title: str) -> tuple[str, str]:
    marker = "街道"
    idx = title.find(marker)
    if idx == -1:
        return "未识别街道", title.strip()
    street = title[: idx + len(marker)].strip()
    station = title[idx + len(marker) :].strip()
    return street, station or title.strip()


def get_station_profile(station_type: str) -> dict:
    aliases = {
        "1": "清洁站",
        "cleaning": "清洁站",
        "clean": "清洁站",
        "sealed": "清洁站",
        "2": "中转站",
        "transfer": "中转站",
        "recycle": "中转站",
    }
    station_type = aliases.get(station_type, station_type)
    if station_type not in STATION_PROFILES:
        supported = "、".join(STATION_PROFILES)
        raise ValueError(f"不支持的日报类型：{station_type}。目前支持：{supported}")
    return STATION_PROFILES[station_type]


def normalize_station_type(station_type: str) -> str:
    aliases = {
        "1": "清洁站",
        "cleaning": "清洁站",
        "clean": "清洁站",
        "sealed": "清洁站",
        "2": "中转站",
        "transfer": "中转站",
        "recycle": "中转站",
    }
    return aliases.get(station_type, station_type)


def count_issue_categories(issue_text: str, categories: list[tuple[str, tuple[str, ...]]]) -> dict[str, int]:
    normalized = normalize_issue_text(issue_text)
    counts = {category: 0 for category, _ in categories}
    if normalized == "无问题。":
        return counts
    for category, keywords in categories:
        if any(keyword in normalized for keyword in keywords):
            counts[category] = 1
    return counts


def previous_nonempty_paragraph_text(paragraphs, index: int) -> str:
    for cursor in range(index - 1, -1, -1):
        text = paragraphs[cursor].text.strip()
        if text:
            return text
    return ""


def is_overall_heading(text: str) -> bool:
    return text in {"整体情况", "1.整体情况", "一、整体情况"} or text.endswith("整体情况")


def is_detail_heading(text: str) -> bool:
    return text in {"具体情况", "2.具体情况", "二、具体情况"} or text.endswith("具体情况")


def next_overall_index(paragraphs, index: int) -> int:
    for cursor in range(index + 1, len(paragraphs)):
        if is_overall_heading(paragraphs[cursor].text.strip()):
            return cursor
    return len(paragraphs)


def parse_points_from_style(style: str) -> tuple[float | None, float | None]:
    width_match = re.search(r"width:([0-9.]+)pt", style or "")
    height_match = re.search(r"height:([0-9.]+)pt", style or "")
    width = float(width_match.group(1)) if width_match else None
    height = float(height_match.group(1)) if height_match else None
    return width, height


def image_row_from_paragraph(document: Document, paragraph) -> ImageRow | None:
    images: list[ImageSpec] = []
    shapes = paragraph._p.xpath('.//*[local-name()="shape"]')
    for shape in shapes:
        image_data = shape.xpath('.//*[local-name()="imagedata"]')
        if not image_data:
            continue
        rid = image_data[0].get(qn("r:id"))
        style = shape.get("style", "")
        related_part = document.part.related_parts.get(rid)
        if related_part is None:
            continue
        width, height = parse_points_from_style(style)
        images.append(ImageSpec(blob=related_part.blob, width_pt=width, height_pt=height))
    if not images:
        return None
    return ImageRow(images=tuple(images), centered=paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER)


def detail_heading_matches_issue(heading: str, issue_text: str) -> bool:
    issue = normalize_issue_text(issue_text)
    rules = [
        (("混投", "垃圾", "箱体"), ("内部", "箱体", "垃圾")),
        (("运输", "喷涂", "车辆"), ("运输", "车辆")),
        (("拒收单",), ("拒收", "证件", "文件", "内部")),
        (("称重",), ("称重", "电箱", "公示", "备案", "证件", "文件", "内部")),
        (("价格表", "价目表", "价目"), ("价格", "价目", "公示", "公告", "营业执照")),
        (("备案", "公示牌", "安全风险公告", "七禁收八不准", "企安安", "营业执照"), ("备案", "公告", "公示", "营业执照", "承诺书", "证件", "负责人", "安全制度")),
        (("安全员", "身份标识"), ("负责人", "证件", "安全员", "门牌", "正门")),
        (("周边环境", "环境脏乱"), ("周边环境", "内部环境", "内部及环境")),
        (("消防水源",), ("消防", "水源", "内部")),
        (("配电箱", "电源箱"), ("配电箱", "电源箱", "内部")),
        (("堆放混乱", "未按规定", "暂存区"), ("内部环境", "内部", "暂存")),
        (("灭火器",), ("灭火器",)),
        (("门牌", "正门"), ("正门", "门牌")),
    ]
    for issue_keywords, heading_keywords in rules:
        if any(keyword in issue for keyword in issue_keywords):
            return any(keyword in heading for keyword in heading_keywords)
    return False


def extract_problem_image_rows(document: Document, detail_start: int, detail_end: int, issue_text: str) -> tuple[ImageRow, ...]:
    sections: list[tuple[str, list[ImageRow]]] = []
    current_heading = ""
    current_rows: list[ImageRow] = []

    for paragraph in document.paragraphs[detail_start:detail_end]:
        text = paragraph.text.strip()
        row = image_row_from_paragraph(document, paragraph)
        if text:
            if current_heading or current_rows:
                sections.append((current_heading, current_rows))
            current_heading = text
            current_rows = []
        elif row is not None:
            current_rows.append(row)

    if current_heading or current_rows:
        sections.append((current_heading, current_rows))

    all_rows = [row for _, rows in sections for row in rows]
    issue_items = split_issue_items(issue_text)
    if not issue_items:
        return tuple(all_rows)

    matched_rows: list[ImageRow] = []
    seen_signatures: set[tuple[tuple[int, bytes], ...]] = set()
    for item in issue_items:
        item_rows = [
            row
            for heading, rows in sections
            if detail_heading_matches_issue(heading, item)
            for row in rows
        ]
        if not item_rows:
            return tuple(all_rows)
        for row in item_rows:
            signature = image_row_signature(row)
            if signature in seen_signatures:
                continue
            matched_rows.append(row)
            seen_signatures.add(signature)

    return tuple(matched_rows)


def parse_daily_report(path: Path, categories: list[tuple[str, tuple[str, ...]]]) -> list[DailyRecord]:
    date = daily_date_from_name(path)
    if date is None:
        return []
    month, day = date
    document = Document(str(path))
    paragraphs = [p.text.strip() for p in document.paragraphs]

    records: list[DailyRecord] = []
    for index, text in enumerate(paragraphs):
        if not is_overall_heading(text) or index == 0:
            continue

        title = previous_nonempty_paragraph_text(document.paragraphs, index)
        issue_parts: list[str] = []
        cursor = index + 1
        while cursor < len(paragraphs) and not is_detail_heading(paragraphs[cursor]):
            if paragraphs[cursor]:
                issue_parts.append(paragraphs[cursor])
            cursor += 1

        street, station = split_street_station(title)
        issue_text = normalize_issue_text("".join(issue_parts))
        detail_start = cursor + 1 if cursor < len(paragraphs) else cursor
        detail_end = max(detail_start, next_overall_index(document.paragraphs, index) - 1)
        records.append(
            DailyRecord(
                source=path,
                month=month,
                day=day,
                street=street,
                station=station,
                issue_text=issue_text,
                issue_counts=count_issue_categories(issue_text, categories),
                image_rows=extract_problem_image_rows(document, detail_start, detail_end, issue_text),
            )
        )
    return records


def collect_records(converted_dir: Path, station_type: str, year: int, month: int) -> list[DailyRecord]:
    station_type = normalize_station_type(station_type)
    source_dir = converted_dir / station_type
    categories = get_station_profile(station_type)["categories"]
    start, end = report_period(year, month)
    records: list[DailyRecord] = []
    for path in sorted(source_dir.glob("*.docx")):
        parsed_date = daily_date_from_name(path)
        if parsed_date is None:
            continue
        record_date = record_date_for_period(parsed_date[0], parsed_date[1], start, end)
        if not start <= record_date <= end:
            continue
        records.extend(parse_daily_report(path, categories))
    return sorted(
        records,
        key=lambda item: (record_date_for_period(item.month, item.day, start, end), item.street, item.station),
    )


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = tc_pr.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        tc_pr.append(shading)
    shading.set(qn("w:fill"), fill)


def set_cell_text(cell, text: str, bold: bool = False) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = "仿宋"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")
    run.font.size = Pt(10.5)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def remove_empty_paragraph(paragraph) -> None:
    paragraph._element.getparent().remove(paragraph._element)
    paragraph._p = paragraph._element = None


def cleanup_empty_paragraphs(document: Document) -> None:
    for paragraph in list(document.paragraphs):
        xml = paragraph._p.xml
        if paragraph.text.strip():
            continue
        if "a:blip" in xml or "v:imagedata" in xml or "w:sectPr" in xml:
            continue
        remove_empty_paragraph(paragraph)


def configure_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "仿宋"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")
    normal.font.size = Pt(16)

    title = document.styles["Title"]
    title.font.name = "黑体"
    title._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
    title.font.size = Pt(18)
    title.font.bold = True

    heading1 = document.styles["Heading 1"]
    heading1.font.name = "黑体"
    heading1._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
    heading1.font.size = Pt(16)
    heading1.font.bold = True
    heading1.font.color.rgb = RGBColor(0, 0, 0)

    heading2 = document.styles["Heading 2"]
    heading2.font.name = "黑体"
    heading2._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
    heading2.font.size = Pt(16)
    heading2.font.bold = True
    heading2.font.color.rgb = RGBColor(0, 0, 0)


def add_paragraph(document: Document, text: str, *, bold: bool = False, align=None) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.first_line_indent = None if bold else Cm(0.74)
    paragraph.paragraph_format.line_spacing = 1.5
    paragraph.paragraph_format.space_after = Pt(0)
    if align is not None:
        paragraph.alignment = align
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = "黑体" if bold else "仿宋"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体" if bold else "仿宋")
    run.font.size = Pt(16)


def add_image_rows(document: Document, image_rows: tuple[ImageRow, ...]) -> None:
    images = [image for image_row in image_rows for image in image_row.images]
    for index in range(0, len(images), 2):
        row_images = images[index : index + 2]
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.line_spacing = 1.0
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        for image in row_images:
            run = paragraph.add_run()
            kwargs = {}
            if image.width_pt is not None:
                kwargs["width"] = Pt(image.width_pt)
            if image.height_pt is not None:
                kwargs["height"] = Pt(image.height_pt)
            image_stream = io.BytesIO()
            with Image.open(io.BytesIO(image.blob)) as pil_image:
                pil_image.save(image_stream, format="PNG")
            image_stream.seek(0)
            run.add_picture(image_stream, **kwargs)


def add_heading(document: Document, text: str, level: int) -> None:
    paragraph = document.add_paragraph(style=f"Heading {level}")
    paragraph.paragraph_format.first_line_indent = None
    paragraph.paragraph_format.line_spacing = 1.5
    paragraph.paragraph_format.space_before = Pt(6 if level == 1 else 3)
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(text)
    run.bold = True
    run.font.name = "黑体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0, 0, 0)


def chinese_section_number(index: int) -> str:
    numerals = "零一二三四五六七八九十"
    if index <= 10:
        return numerals[index]
    if index < 20:
        return f"十{numerals[index - 10]}"
    tens, ones = divmod(index, 10)
    suffix = numerals[ones] if ones else ""
    return f"{numerals[tens]}十{suffix}"


def is_unopened_record(record: DailyRecord, profile: dict) -> bool:
    category = profile["unopened_category"]
    return record.issue_counts.get(category, 0) > 0 and sum(record.issue_counts.values()) == 1


def is_body_record(record: DailyRecord, profile: dict) -> bool:
    if not record.has_problem or is_unopened_record(record, profile):
        return False
    if sum(record.issue_counts.values()) == 0 and any(
        keyword in record.issue_text for keyword in profile.get("non_problem_status_keywords", ())
    ):
        return False
    return True


def normalize_display_issue_text(text: str) -> str:
    text = text.strip()
    if re.match(r"^（\d+）", text):
        return text
    return f"（1）{text}"


def split_issue_items(text: str) -> list[str]:
    normalized = normalize_issue_text(text)
    matches = list(re.finditer(r"（\d+）", normalized))
    if not matches:
        return [strip_trailing_punctuation(normalized)]
    items: list[str] = []
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(normalized)
        item = strip_trailing_punctuation(normalized[start:end].strip())
        if item:
            items.append(item)
    return items


def strip_trailing_punctuation(text: str) -> str:
    return text.strip().rstrip("；;。")


def format_issue_items(items: list[str]) -> str:
    unique_items: list[str] = []
    seen: set[str] = set()
    for item in items:
        normalized = strip_trailing_punctuation(item)
        if normalized and normalized not in seen:
            unique_items.append(normalized)
            seen.add(normalized)
    return "；".join(f"（{index}）{item}" for index, item in enumerate(unique_items, start=1)) + "。"


def issue_item_key(item: str, categories: list[tuple[str, tuple[str, ...]]]) -> str:
    for category, keywords in categories:
        if any(keyword in item for keyword in keywords):
            return f"category:{category}"
    return f"text:{item}"


def image_row_signature(row: ImageRow) -> tuple[tuple[int, bytes], ...]:
    return tuple((len(image.blob), image.blob[:32]) for image in row.images)


def merge_street_records(records: list[DailyRecord], categories: list[tuple[str, tuple[str, ...]]]) -> list[MergedRecord]:
    merged_items: dict[str, list[str]] = {}
    merged_texts: dict[str, dict[str, str]] = {}
    merged_images: dict[str, dict[str, tuple[ImageRow, ...]]] = {}

    for record in records:
        issue_keys = merged_items.setdefault(record.station, [])
        issue_texts = merged_texts.setdefault(record.station, {})
        issue_images = merged_images.setdefault(record.station, {})

        for item in split_issue_items(record.issue_text):
            key = issue_item_key(item, categories)
            if key not in issue_texts:
                issue_keys.append(key)
                issue_texts[key] = item
            # Records are sorted by date, so assignment keeps only the latest
            # image rows for duplicated problem categories.
            issue_images[key] = record.image_rows

    result: list[MergedRecord] = []
    for station, issue_keys in merged_items.items():
        rows: list[ImageRow] = []
        seen_rows: set[tuple[tuple[int, bytes], ...]] = set()
        for key in issue_keys:
            for row in merged_images.get(station, {}).get(key, ()):
                signature = image_row_signature(row)
                if signature in seen_rows:
                    continue
                rows.append(row)
                seen_rows.add(signature)
        result.append(
            MergedRecord(
                station=station,
                issue_texts=[merged_texts[station][key] for key in issue_keys],
                image_rows=rows,
            )
        )
    return result


def summarize_records(
    records: list[DailyRecord],
    categories: list[tuple[str, tuple[str, ...]]],
) -> tuple[dict[str, list[DailyRecord]], dict[str, dict[str, int]]]:
    by_street: dict[str, list[DailyRecord]] = defaultdict(list)
    counts: dict[str, dict[str, int]] = defaultdict(lambda: {category: 0 for category, _ in categories})

    for record in records:
        by_street[record.street].append(record)
        for category, count in record.issue_counts.items():
            counts[record.street][category] += count

    return dict(sorted(by_street.items())), dict(sorted(counts.items()))


def build_summary_sentence(
    records: list[DailyRecord],
    year: int,
    month: int,
    profile: dict,
    street_count: int,
) -> str:
    start, end = report_period(year, month)
    problem_records = [record for record in records if is_body_record(record, profile)]
    start_label = format_date_cn(start)
    end_label = format_date_cn(end)

    if start_label == end_label:
        prefix = f"{start_label}对本区{street_count}个街道{profile['subject']}进行了{len(records)}个次检查。"
    else:
        prefix = f"{start_label}至{end_label}对本区{street_count}个街道{profile['subject']}进行了{len(records)}个次检查。"

    if not problem_records:
        return f"{prefix}除未开门运行情况外，检查未发现其他问题。"
    return f"{prefix}检查发现需通报问题站次{len(problem_records)}个。"


def add_issue_table(document: Document, counts: dict[str, dict[str, int]], categories: list[tuple[str, tuple[str, ...]]]) -> None:
    headers = ["街道名称", *[category for category, _ in categories], "问题总数"]
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = True

    for index, header in enumerate(headers):
        cell = table.rows[0].cells[index]
        set_cell_text(cell, header, bold=True)
        set_cell_shading(cell, "D9EAF7")

    total_by_category = {category: 0 for category, _ in categories}
    for street_counts in counts.values():
        for category, value in street_counts.items():
            total_by_category[category] += value

    rows = [("总计", total_by_category), *counts.items()]
    for street, street_counts in rows:
        row = table.add_row()
        values = [street]
        problem_total = 0
        for category, _ in categories:
            value = street_counts.get(category, 0)
            problem_total += value
            values.append("" if value == 0 else str(value))
        values.append("" if problem_total == 0 else str(problem_total))
        for index, value in enumerate(values):
            set_cell_text(row.cells[index], value)


def create_monthly_docx(records: list[DailyRecord], output_path: Path, year: int, month: int, station_type: str) -> Path:
    profile = get_station_profile(station_type)
    categories = profile["categories"]
    if not records:
        raise ValueError(f"未找到 {month} 月{profile['subject']}日报记录")

    document = Document()
    configure_styles(document)

    section = document.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.18)
    section.right_margin = Cm(3.18)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(18)
    run = title.add_run(profile["title"])
    run.bold = True
    run.font.name = "黑体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
    run.font.size = Pt(22)

    by_street, counts = summarize_records(
        [record for record in records if not is_unopened_record(record, profile)],
        categories,
    )
    streets_with_body_records = [
        (street, [record for record in street_records if is_body_record(record, profile)])
        for street, street_records in by_street.items()
    ]
    streets_with_body_records = [
        (street, street_records)
        for street, street_records in streets_with_body_records
        if street_records
    ]

    add_heading(document, "一、总体情况", level=1)
    add_paragraph(document, build_summary_sentence(records, year, month, profile, len(streets_with_body_records)))

    add_heading(document, profile["detail_heading"], level=1)
    if not streets_with_body_records:
        add_paragraph(document, "本月除未开门运行情况外，未发现需列入街道情况的其他问题。")
    for street_index, (street, street_records) in enumerate(streets_with_body_records, start=1):
        add_heading(document, f"（{chinese_section_number(street_index)}）{street}", level=2)
        merged_records = merge_street_records(street_records, categories)
        use_station_numbers = len(merged_records) > 1
        for station_index, merged_record in enumerate(merged_records, start=1):
            issue_text = format_issue_items(merged_record.issue_texts)
            station_label = f"{station_index}.{merged_record.station}" if use_station_numbers else merged_record.station
            add_paragraph(document, f"{station_label}：{issue_text}")
            add_image_rows(document, tuple(merged_record.image_rows))

    landscape = document.add_section(WD_ORIENT.LANDSCAPE)
    landscape.orientation = WD_ORIENT.LANDSCAPE
    landscape.page_width, landscape.page_height = landscape.page_height, landscape.page_width
    landscape.top_margin = Cm(2.54)
    landscape.bottom_margin = Cm(2.54)
    landscape.left_margin = Cm(2.54)
    landscape.right_margin = Cm(2.54)

    start, end = report_period(year, month)
    attachment = document.add_paragraph()
    attachment.alignment = WD_ALIGN_PARAGRAPH.CENTER
    attachment.paragraph_format.space_before = Pt(8)
    attachment.paragraph_format.space_after = Pt(8)
    run = attachment.add_run(
        f"附件：{format_date_cn(start)}至{format_date_cn(end)}"
        f"{profile['attachment_subject']}各项指标问题数"
    )
    run.font.name = "仿宋"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "仿宋")
    run.font.size = Pt(14)

    add_issue_table(document, counts, categories)

    cleanup_empty_paragraphs(document)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    return output_path


def generate_monthly_report(
    raw_input_dir: Path = DEFAULT_RAW_INPUT_DIR,
    converted_dir: Path = DEFAULT_CONVERTED_DIR,
    output_dir: Path = DEFAULT_OUTPUT_DIR,
    month: str | int | None = None,
    year: int | None = None,
    station_type: str = DEFAULT_STATION_TYPE,
    engine: str = "auto",
    skip_convert: bool = False,
) -> Path:
    station_type = normalize_station_type(station_type)
    report_month = parse_month(month)
    report_year = parse_year(year, month)

    if not skip_convert:
        summary = ensure_converted(raw_input_dir, converted_dir, station_type, engine)
        if summary.failed:
            messages = "\n".join(f"{path}: {error}" for path, error in summary.failed)
            raise ConversionError(f"部分日报转换失败：\n{messages}")

    records = collect_records(converted_dir, station_type, report_year, report_month)
    profile = get_station_profile(station_type)
    output_path = output_dir / profile["output_name"].format(year=report_year, month=report_month)
    return create_monthly_docx(records, output_path, report_year, report_month, station_type)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="汇总日报并生成密闭式清洁站月度检查通报")
    parser.add_argument("--raw-input", type=Path, default=DEFAULT_RAW_INPUT_DIR, help="原始 .doc 日报根目录")
    parser.add_argument("--converted-output", type=Path, default=DEFAULT_CONVERTED_DIR, help="转换后 .docx 日报根目录")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_DIR, help="月报输出目录")
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
            output_dir=args.output,
            month=args.month,
            year=args.year,
            station_type=args.station_type,
            engine=args.engine,
            skip_convert=args.skip_convert,
        )
    except PermissionError as exc:
        print(f"[失败] 无法写入目标文件，可能正在被 Word 打开：{exc.filename}", file=sys.stderr)
        print("[提示] 请关闭该文档后重新运行脚本，或通过 --output 指定其他输出目录。", file=sys.stderr)
        return 1
    except (ConversionError, FileNotFoundError, ValueError) as exc:
        print(f"[失败] {exc}", file=sys.stderr)
        return 1

    print(f"[完成] 月报已生成：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
