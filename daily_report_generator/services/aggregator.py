from __future__ import annotations

from collections import OrderedDict
from datetime import date

from .models import AggregationResult, LedgerRecord, StationItem, StationSummary
from .normalizer import (
    CLEAN_TYPE,
    TRANSFER_TYPE,
    build_summary,
    chinese_section_number,
    item_display_name,
    level4_point_display_name,
)


def aggregate_records(records: list[LedgerRecord], report_date: date) -> AggregationResult:
    daily_records = [record for record in records if record.check_date == report_date]
    transfer = aggregate_type(daily_records, TRANSFER_TYPE)
    clean = aggregate_type(daily_records, CLEAN_TYPE)
    return AggregationResult(report_date=report_date, transfer_stations=transfer, clean_stations=clean)


def aggregate_type(records: list[LedgerRecord], station_type: str) -> list[StationSummary]:
    grouped: OrderedDict[tuple[str, str], list[LedgerRecord]] = OrderedDict()
    for record in sorted(records, key=lambda item: (item.street, item.station_name, item.row_index)):
        if record.point_type != station_type:
            continue
        grouped.setdefault((record.street, record.station_name), []).append(record)

    stations: list[StationSummary] = []
    for (street, station_name), station_records in grouped.items():
        items = build_items(station_records)
        stations.append(
            StationSummary(
                street=street,
                name=station_name,
                display_name=f"{street}{station_name}",
                summary=build_summary([record.issue_text for record in station_records]),
                items=items,
            )
        )
    return stations


def build_items(records: list[LedgerRecord]) -> list[StationItem]:
    grouped: OrderedDict[str, StationItem] = OrderedDict()
    for record in records:
        name = item_display_name(record.indicator_name, record.issue_text)
        item = grouped.setdefault(
            name,
            StationItem(
                name=name,
                indicator_level_1=record.indicator_level_1,
                indicator_name=record.indicator_name,
                result=record.indicator_result,
                issue_text=record.issue_text,
            ),
        )
        if not item.indicator_level_1 and record.indicator_level_1:
            item.indicator_level_1 = record.indicator_level_1
        if not item.indicator_name and record.indicator_name:
            item.indicator_name = record.indicator_name
        if not item.result and record.indicator_result:
            item.result = record.indicator_result
        if not item.issue_text and record.issue_text:
            item.issue_text = record.issue_text
        for photo in record.problem_photos:
            if photo not in item.photo_paths:
                item.photo_paths.append(photo)
    return list(grouped.values())


def result_for_type(result: AggregationResult, station_type: str) -> list[StationSummary]:
    if station_type == TRANSFER_TYPE:
        return result.transfer_stations
    if station_type == CLEAN_TYPE:
        return result.clean_stations
    raise ValueError(f"不支持的日报类型：{station_type}")


def preview_payload(result: AggregationResult, station_type: str | None = None) -> dict:
    selected_types = [station_type] if station_type else [TRANSFER_TYPE, CLEAN_TYPE]
    type_payloads: dict[str, dict] = {}
    total_stations = 0
    total_items = 0
    total_photos = 0
    for selected in selected_types:
        stations = result_for_type(result, selected)
        total_stations += len(stations)
        item_count = sum(len(station.items) for station in stations)
        photo_count = sum(len(item.photo_paths) for station in stations for item in station.items)
        total_items += item_count
        total_photos += photo_count
        type_payloads[selected] = {
            "station_count": len(stations),
            "item_count": item_count,
            "photo_count": photo_count,
            "stations": [
                {
                    "index": index,
                    "index_cn": chinese_section_number(index),
                    "street": station.street,
                    "street_name": station.street,
                    "level3_point": station.street,
                    "level4_point": level4_point_display_name(station.name),
                    "name": station.name,
                    "display_name": station.display_name,
                    "summary": station.summary,
                    "item_count": len(station.items),
                    "photo_count": sum(len(item.photo_paths) for item in station.items),
                    "items": [
                        {
                            "name": item.name,
                            "result": item.result,
                            "issue_text": item.issue_text,
                            "photo_count": len(item.photo_paths),
                        }
                        for item in station.items
                    ],
                }
                for index, station in enumerate(stations, start=1)
            ],
        }
    return {
        "report_date": result.report_date.isoformat(),
        "report_date_text": result.report_date_text,
        "total_station_count": total_stations,
        "total_item_count": total_items,
        "total_photo_count": total_photos,
        "types": type_payloads,
    }
