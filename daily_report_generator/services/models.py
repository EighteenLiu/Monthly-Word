from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path


@dataclass(frozen=True)
class LedgerImage:
    row_index: int
    column_index: int
    path: Path


@dataclass(frozen=True)
class LedgerRecord:
    row_index: int
    serial_no: str
    point_level_1: str
    point_type: str
    street: str
    station_name: str
    issue_text: str
    indicator_level_1: str
    indicator_name: str
    indicator_result: str
    report_time: datetime | None
    created_time: datetime | None
    check_date: date | None
    problem_photos: tuple[Path, ...] = ()


@dataclass
class StationItem:
    name: str
    indicator_level_1: str = ""
    indicator_name: str = ""
    result: str = ""
    issue_text: str = ""
    photo_paths: list[Path] = field(default_factory=list)


@dataclass
class StationSummary:
    street: str
    name: str
    display_name: str
    summary: str
    items: list[StationItem] = field(default_factory=list)


@dataclass
class AggregationResult:
    report_date: date
    transfer_stations: list[StationSummary]
    clean_stations: list[StationSummary]

    @property
    def report_date_text(self) -> str:
        return f"{self.report_date.month}月{self.report_date.day}日"

    @property
    def report_date_code(self) -> str:
        return self.report_date.strftime("%Y%m%d")
