from __future__ import annotations

import zipfile
from types import SimpleNamespace
from datetime import date
from pathlib import Path

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm

from .aggregator import result_for_type
from .models import AggregationResult, StationItem, StationSummary
from .normalizer import CLEAN_TYPE, TRANSFER_TYPE, chinese_section_number, compact_text, is_no_problem_text, level4_point_display_name
from .transfer_station_mapping import TRANSFER_STATION_DETAIL_MAPPING


PHOTO_WIDTH = Cm(7.22)
PHOTO_HEIGHT = Cm(4.23)
NO_PROBLEM_PHOTO_LIMIT = 4
CLEAN_STATION_GOOD_TITLE = "良好，未发现问题"
TRANSFER_FRONT_DOOR_GROUP = "正门及门牌"
CLEAN_STATION_TITLE_MAP = {
    "标志不完整不清晰、喷涂不规范": "正门及门牌",
    "收集车辆敞口运输": "运输车辆",
    "小型收集车混装混运": "运输车辆",
    "箱体内垃圾混投": "内部环境",
    "无称重系统或称重系统损坏": "内部环境",
    "拒收单不准确": "操作规范及流程",
    "未开门运行": "操作规范及流程",
    "正门及门牌": "正门及门牌",
    "公告及文件": "公告及文件",
    "运输车辆": "运输车辆",
    "内部环境": "内部环境",
    "灭火器": "灭火器",
    "操作规范及流程": "操作规范及流程",
}


def output_dir_for_type(output_root: Path, station_type: str) -> Path:
    folder_name = "清洁站" if station_type == CLEAN_TYPE else station_type
    return output_root / folder_name


def render_reports(
    result: AggregationResult,
    transfer_template: Path | None,
    clean_template: Path | None,
    output_root: Path,
    types: list[str],
) -> list[Path]:
    rendered: list[Path] = []
    for station_type in types:
        if not result_for_type(result, station_type):
            continue
        output_dir = output_dir_for_type(output_root, station_type)
        output_dir.mkdir(parents=True, exist_ok=True)
        if station_type == TRANSFER_TYPE:
            if not transfer_template:
                raise ValueError("生成中转站日报需要上传中转站模板")
            output = output_dir / f"{result.report_date_text}中转站检查情况.docx"
            render_one(result, station_type, transfer_template, output)
        elif station_type == CLEAN_TYPE:
            if not clean_template:
                raise ValueError("生成密闭式清洁站日报需要上传密闭式清洁站模板")
            output = output_dir / f"{result.report_date_text}密闭式清洁站检查情况.docx"
            render_one(result, station_type, clean_template, output)
        else:
            raise ValueError(f"不支持的日报类型：{station_type}")
        rendered.append(output)
    if not rendered:
        raise ValueError(f"{result.report_date_text}未找到可生成的日报数据")
    return rendered


def render_one(result: AggregationResult, station_type: str, template_path: Path, output_path: Path) -> Path:
    stations = result_for_type(result, station_type)
    if not stations:
        raise ValueError(f"{result.report_date_text}未找到{station_type}数据")

    doc = DocxTemplate(str(template_path))
    context = {
        "report_date": result.report_date.isoformat(),
        "report_date_text": result.report_date_text,
        "transfer_stations": [],
        "clean_stations": [],
    }
    station_payload = [
        station_to_template(station, doc, index, station_type)
        for index, station in enumerate(stations, start=1)
    ]
    if station_type == TRANSFER_TYPE:
        context["transfer_stations"] = station_payload
    else:
        context["clean_stations"] = station_payload
    doc.render(context)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def station_to_template(station: StationSummary, doc: DocxTemplate, index: int, station_type: str = "") -> SimpleNamespace:
    items = [item_to_template(item, doc) for item in station.items]
    return SimpleNamespace(
        index=index,
        index_cn=chinese_section_number(index),
        street=station.street,
        street_name=station.street,
        level3_point=station.street,
        level4_point=level4_point_display_name(station.name),
        name=station.name,
        display_name=station.display_name,
        summary=station.summary,
        items=items,
        detail_items=build_detail_items(items, station_type),
    )


def item_to_template(item: StationItem, doc: DocxTemplate) -> dict:
    photos = [InlineImage(doc, str(path), width=PHOTO_WIDTH, height=PHOTO_HEIGHT) for path in item.photo_paths if path.exists()]
    return SimpleNamespace(
        name=item.name,
        indicator_level_1=item.indicator_level_1,
        indicator_name=item.indicator_name,
        result=item.result,
        issue_text=item.issue_text,
        photo_paths=item.photo_paths,
        photos=photos,
        photo_rows=[
            SimpleNamespace(
                left=photos[index],
                right=photos[index + 1] if index + 1 < len(photos) else None,
            )
            for index in range(0, len(photos), 2)
        ],
    )


def norm_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().replace(" ", "").replace("　", "")


def item_context_text(item: SimpleNamespace) -> str:
    parts = [
        getattr(item, "indicator_level_1", ""),
        getattr(item, "indicator_name", ""),
        getattr(item, "result", ""),
        getattr(item, "issue_text", ""),
        getattr(item, "name", ""),
        " ".join(str(path.name) for path in getattr(item, "photo_paths", []) if hasattr(path, "name")),
    ]
    return "|".join(norm_text(part) for part in parts if part)


def is_no_problem_item(item: SimpleNamespace) -> bool:
    return any(
        is_no_problem_text(str(value))
        for value in (
            getattr(item, "indicator_name", ""),
            getattr(item, "result", ""),
            getattr(item, "issue_text", ""),
            getattr(item, "name", ""),
        )
    )


def is_good_no_problem_indicator(item: SimpleNamespace) -> bool:
    return compact_text(getattr(item, "indicator_name", "")) == compact_text(CLEAN_STATION_GOOD_TITLE)


def is_item_match(item: SimpleNamespace, config: dict) -> bool:
    exact_fields = {
        norm_text(getattr(item, "indicator_name", "")),
        norm_text(getattr(item, "result", "")),
        norm_text(getattr(item, "issue_text", "")),
        norm_text(getattr(item, "name", "")),
    }
    context = item_context_text(item)
    indicators = [norm_text(value) for value in config.get("indicator_names", [])]
    aliases = [norm_text(value) for value in config.get("aliases", [])]
    if any(value and value in exact_fields for value in indicators):
        return True
    if any(value and value in context for value in indicators):
        return True
    return any(value and value in context for value in aliases)


def make_photo_rows(photos: list[InlineImage]) -> list[SimpleNamespace]:
    return [
        SimpleNamespace(
            left=photos[index],
            right=photos[index + 1] if index + 1 < len(photos) else None,
        )
        for index in range(0, len(photos), 2)
    ]


def build_detail_items(items: list[SimpleNamespace], station_type: str) -> list[SimpleNamespace]:
    if station_type == TRANSFER_TYPE:
        return build_transfer_station_detail_items(items)
    if station_type == CLEAN_TYPE:
        return build_clean_station_detail_items(items)
    return items


def clean_station_title_for_item(item: SimpleNamespace) -> str:
    indicator_name = compact_text(getattr(item, "indicator_name", ""))
    if indicator_name == compact_text(CLEAN_STATION_GOOD_TITLE):
        result = getattr(item, "result", "")
        return CLEAN_STATION_TITLE_MAP.get(result) or result or CLEAN_STATION_GOOD_TITLE
    return CLEAN_STATION_TITLE_MAP.get(indicator_name) or getattr(item, "name", "") or getattr(item, "indicator_name", "")


def build_clean_station_detail_items(items: list[SimpleNamespace]) -> list[SimpleNamespace]:
    grouped: dict[str, SimpleNamespace] = {}
    order: list[str] = []
    for item in items:
        title = clean_station_title_for_item(item)
        if not title:
            continue
        detail = grouped.get(title)
        if detail is None:
            detail = SimpleNamespace(
                name=title,
                photos=[],
                photo_paths=[],
            )
            grouped[title] = detail
            order.append(title)
        for photo_index, photo in enumerate(getattr(item, "photos", [])):
            photo_paths = getattr(item, "photo_paths", [])
            photo_id = str(photo_paths[photo_index]) if photo_index < len(photo_paths) else f"{id(item)}:{photo_index}"
            if photo_id in detail.photo_paths:
                continue
            detail.photos.append(photo)
            detail.photo_paths.append(photo_id)

    return [
        SimpleNamespace(
            name=grouped[title].name,
            photos=grouped[title].photos,
            photo_rows=make_photo_rows(grouped[title].photos),
        )
        for title in order
    ]


def build_transfer_station_detail_items(items: list[SimpleNamespace], show_empty: bool = False) -> list[SimpleNamespace]:
    used_photo_ids: set[str] = set()
    detail_items: list[SimpleNamespace] = []

    def add_detail_item(name: str, photos: list[InlineImage]) -> None:
        if not photos and not show_empty:
            return
        detail_items.append(
            SimpleNamespace(
                name=name,
                photos=photos,
                photo_rows=make_photo_rows(photos),
            )
        )

    def front_door_detail() -> tuple[SimpleNamespace, bool]:
        for detail in detail_items:
            if detail.name == TRANSFER_FRONT_DOOR_GROUP:
                return detail, False
        detail = SimpleNamespace(name=TRANSFER_FRONT_DOOR_GROUP, photos=[], photo_rows=[])
        detail_items.insert(0, detail)
        return detail, True

    def append_unused_photos(target: SimpleNamespace, item: SimpleNamespace, limit: int | None = None) -> int:
        added = 0
        for photo_index, photo in enumerate(getattr(item, "photos", [])):
            photo_paths = getattr(item, "photo_paths", [])
            photo_id = str(photo_paths[photo_index]) if photo_index < len(photo_paths) else f"{id(item)}:{photo_index}"
            if photo_id in used_photo_ids:
                continue
            target.photos.append(photo)
            used_photo_ids.add(photo_id)
            added += 1
            if limit is not None and added >= limit:
                break
        target.photo_rows = make_photo_rows(target.photos)
        return added

    for config in TRANSFER_STATION_DETAIL_MAPPING:
        matched_photos: list[InlineImage] = []
        for item in items:
            if not is_item_match(item, config):
                continue
            for photo_index, photo in enumerate(getattr(item, "photos", [])):
                photo_paths = getattr(item, "photo_paths", [])
                photo_id = str(photo_paths[photo_index]) if photo_index < len(photo_paths) else f"{id(item)}:{photo_index}"
                if photo_id in used_photo_ids:
                    continue
                matched_photos.append(photo)
                used_photo_ids.add(photo_id)
        add_detail_item(config["name"], matched_photos)
    for item in items:
        if not is_no_problem_item(item):
            continue
        if is_good_no_problem_indicator(item):
            detail, created = front_door_detail()
            added = append_unused_photos(detail, item, NO_PROBLEM_PHOTO_LIMIT)
            if created and added == 0:
                detail_items.remove(detail)
            continue
        fallback = SimpleNamespace(name=getattr(item, "name", "") or "无问题", photos=[], photo_rows=[])
        append_unused_photos(fallback, item, NO_PROBLEM_PHOTO_LIMIT)
        if fallback.photos:
            detail_items.append(fallback)
    return detail_items


def make_zip(files: list[Path], output_root: Path, report_date: date) -> Path:
    zip_path = output_root / f"{report_date.strftime('%Y%m%d')}_日报.zip"
    return zip_path
