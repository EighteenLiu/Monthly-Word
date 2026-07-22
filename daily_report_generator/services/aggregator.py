from __future__ import annotations

from collections import OrderedDict
from datetime import date

from .models import AggregationResult, LedgerRecord, StationItem, StationSummary
from .normalizer import (
    CLEAN_TYPE,
    TRANSFER_TYPE,
    compact_text,
    chinese_section_number,
    ensure_period,
    item_display_name,
    level4_point_display_name,
    normalize_problem_sentence,
)


ISSUE_TEXT_NO_PROBLEM = {
    "无",
    "无问题",
    "良好",
    "良好未发现问题",
    "未发现问题",
    "无异常",
    "正常",
    "合格",
}

PROBLEM_INDICATOR_RESULTS = {
    CLEAN_TYPE: {
        "标志不完整不清晰、喷涂不规范": {"有问题"},
        "收集车辆敞口运输": {"有问题"},
        "未开门运行": {"有问题", "不运行", "未运行"},
        "无称重系统或称重系统损坏": {"有问题"},
        "小型收集车混装混运": {"有问题"},
        "箱体内垃圾混投": {"有问题"},
        "拒收单不准确": {"有问题"},
    },
    TRANSFER_TYPE: {
        "周边环境脏乱情况": {"周边环境脏乱"},
        "可回收价格表": {"无可回收价格表"},
        "备案公示": {"无备案公示"},
        "消防水源是否合格": {"消防水源不合格"},
        "消防安全水源": {"无消防安全水源"},
        "营业执照": {"无营业执照"},
        "安全风险公告": {"无安全风险公告"},
        "称重系统损坏": {"称重系统损坏"},
        "灭火器": {"灭火器不合格", "灭火器过期"},
        "按规定区域存放灭火器等物品": {"灭火器等未按规定放置"},
        "七禁收八不准承诺书情况": {"无七禁收八不准承诺书"},
        "运输车辆防遗撒检查台账": {"无车辆信息、出入时间、装载品类及装载后的影像资料"},
        "安全员情况": {"安全员上岗但无明显身份标识", "安全员未按时上岗"},
        "灭火器和消防栓箱内有每月检查记录表": {"无"},
        "开门运行情况": {"未按时开门运行", "不运行", "未运行"},
        "企安安情况": {"无企安安"},
        "灭蝇措施情况": {"无灭蝇措施"},
    },
}

NON_PROBLEM_INDICATOR_RESULTS = {
    CLEAN_TYPE: {
        "标志不完整不清晰、喷涂不规范": {"无问题"},
        "收集车辆敞口运输": {"无问题"},
        "良好，未发现问题": {"操作规范及流程", "公告及文件", "灭火器", "内部环境", "运输车辆", "正门及门牌"},
        "未开门运行": {"无问题"},
        "无称重系统或称重系统损坏": {"无问题"},
        "小型收集车混装混运": {"无问题"},
        "箱体内垃圾混投": {"无问题"},
        "拒收单不准确": {"无问题"},
    },
    TRANSFER_TYPE: {
        "良好，未发现问题": {""},
        "周边环境脏乱情况": {"良好，未发现问题"},
        "可回收价格表": {"有可回收价格表"},
        "备案公示": {"有备案公示"},
        "消防水源是否合格": {"消防水源合格"},
        "消防安全水源": {"有消防安全水源"},
        "营业执照": {"有营业执照"},
        "安全风险公告": {"有安全风险公告"},
        "称重系统损坏": {"良好，未发现问题"},
        "灭火器": {"灭火器合格"},
        "按规定区域存放灭火器等物品": {"灭火器按规定放置"},
        "七禁收八不准承诺书情况": {"有七禁八不准承诺书"},
        "运输车辆防遗撒检查台账": {"有车辆信息、出入时间、装载品类及装载后的影像资料"},
        "安全员情况": {"安全员按时上岗"},
        "灭火器和消防栓箱内有每月检查记录表": {"有"},
        "开门运行情况": {"按时开门运行"},
        "企安安情况": {"有企安安"},
        "灭蝇措施情况": {"有灭蝇措施"},
    },
}

COMPACT_PROBLEM_INDICATOR_RESULTS = {
    station_type: {
        compact_text(indicator): {compact_text(result) for result in results}
        for indicator, results in station_rules.items()
    }
    for station_type, station_rules in PROBLEM_INDICATOR_RESULTS.items()
}

COMPACT_NON_PROBLEM_INDICATOR_RESULTS = {
    station_type: {
        compact_text(indicator): {compact_text(result) for result in results}
        for indicator, results in station_rules.items()
    }
    for station_type, station_rules in NON_PROBLEM_INDICATOR_RESULTS.items()
}

COMPACT_ISSUE_TEXT_NO_PROBLEM = {compact_text(text) for text in ISSUE_TEXT_NO_PROBLEM}
COMPACT_NO_PROBLEM_INDICATORS = {
    compact_text("良好，未发现问题"),
    compact_text("良好未发现问题"),
}


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
                summary=build_station_summary(station_records, station_type),
                items=items,
            )
        )
    return stations


def build_station_summary(records: list[LedgerRecord], station_type: str) -> str:
    problems: list[str] = []
    seen: set[str] = set()
    for record in records:
        if not is_problem_record(record, station_type):
            continue
        problem = problem_text(record)
        if not problem or problem in seen:
            continue
        problems.append(problem)
        seen.add(problem)

    if not problems:
        return "无问题。"
    joined = "；".join(f"（{index}）{problem}" for index, problem in enumerate(problems, start=1))
    return ensure_period(f"存在的问题是：{joined}")


def issue_text_is_no_problem(text: str) -> bool:
    return compact_text(text) in COMPACT_ISSUE_TEXT_NO_PROBLEM


def is_problem_record(record: LedgerRecord, station_type: str) -> bool:
    if issue_text_is_no_problem(record.issue_text):
        return False

    indicator = compact_text(record.indicator_name)
    indicator_result = compact_text(record.indicator_result)
    if indicator in COMPACT_NO_PROBLEM_INDICATORS:
        return bool(indicator_result) and not issue_text_is_no_problem(indicator_result)
    if not indicator:
        return False

    if indicator_result in COMPACT_NON_PROBLEM_INDICATOR_RESULTS.get(station_type, {}).get(indicator, set()):
        return False
    return indicator_result in COMPACT_PROBLEM_INDICATOR_RESULTS.get(station_type, {}).get(indicator, set())


def problem_text(record: LedgerRecord) -> str:
    indicator = compact_text(record.indicator_name)
    indicator_result = compact_text(record.indicator_result)
    if indicator == compact_text("未开门运行") and indicator_result in {compact_text("不运行"), compact_text("未运行")}:
        return "未开门运行"
    if indicator == compact_text("开门运行情况") and indicator_result in {compact_text("不运行"), compact_text("未运行")}:
        return "未按时开门运行"

    issue_text = normalize_problem_sentence(record.issue_text)
    if issue_text and not issue_text_is_no_problem(issue_text):
        return issue_text
    if indicator in COMPACT_NO_PROBLEM_INDICATORS:
        indicator_result = normalize_problem_sentence(record.indicator_result)
        return indicator_result if indicator_result and not issue_text_is_no_problem(indicator_result) else ""
    indicator_result = normalize_problem_sentence(record.indicator_result)
    if indicator_result:
        return indicator_result
    return normalize_problem_sentence(record.indicator_name)


def should_skip_detail_record(record: LedgerRecord) -> bool:
    indicator = compact_text(record.indicator_name)
    indicator_result = compact_text(record.indicator_result)
    return (
        indicator in COMPACT_NO_PROBLEM_INDICATORS
        and not indicator_result
        and issue_text_is_no_problem(record.issue_text)
    )


def build_items(records: list[LedgerRecord]) -> list[StationItem]:
    grouped: OrderedDict[str, StationItem] = OrderedDict()
    for record in records:
        if should_skip_detail_record(record):
            continue
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
