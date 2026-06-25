from __future__ import annotations

import zipfile
from types import SimpleNamespace
from datetime import date
from pathlib import Path

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm

from .aggregator import result_for_type
from .models import AggregationResult, StationItem, StationSummary
from .normalizer import CLEAN_TYPE, TRANSFER_TYPE, chinese_section_number, is_no_problem_text, level4_point_display_name
from .transfer_station_mapping import TRANSFER_STATION_DETAIL_MAPPING


PHOTO_WIDTH = Cm(7.22)
PHOTO_HEIGHT = Cm(5.72)
NO_PROBLEM_PHOTO_LIMIT = 4


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
        detail_items=build_transfer_station_detail_items(items) if station_type == TRANSFER_TYPE else items,
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


def build_transfer_station_detail_items(items: list[SimpleNamespace], show_empty: bool = False) -> list[SimpleNamespace]:
    used_photo_ids: set[str] = set()
    detail_items: list[SimpleNamespace] = []
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
        if matched_photos or show_empty:
            detail_items.append(
                SimpleNamespace(
                    name=config["name"],
                    photos=matched_photos,
                    photo_rows=make_photo_rows(matched_photos),
                )
            )
    for item in items:
        if not is_no_problem_item(item):
            continue
        fallback_photos: list[InlineImage] = []
        for photo_index, photo in enumerate(getattr(item, "photos", [])):
            photo_paths = getattr(item, "photo_paths", [])
            photo_id = str(photo_paths[photo_index]) if photo_index < len(photo_paths) else f"{id(item)}:{photo_index}"
            if photo_id in used_photo_ids:
                continue
            fallback_photos.append(photo)
            used_photo_ids.add(photo_id)
            if len(fallback_photos) >= NO_PROBLEM_PHOTO_LIMIT:
                break
        if fallback_photos:
            detail_items.append(
                SimpleNamespace(
                    name=getattr(item, "name", "") or "无问题",
                    photos=fallback_photos,
                    photo_rows=make_photo_rows(fallback_photos),
                )
            )
    return detail_items


def make_zip(files: list[Path], output_root: Path, report_date: date) -> Path:
    zip_path = output_root / f"{report_date.strftime('%Y%m%d')}_日报.zip"
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for file in files:
            try:
                arcname = file.relative_to(output_root)
            except ValueError:
                arcname = file.name
            archive.write(file, arcname=arcname.as_posix() if isinstance(arcname, Path) else arcname)
    return zip_path
