from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Annotated

from fastapi import FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from .config import (
    APP_ROOT,
    DEFAULT_CLEAN_TEMPLATE,
    DEFAULT_TRANSFER_TEMPLATE,
    OUTPUT_ROOT,
    UPLOAD_ROOT,
)
from .services.aggregator import aggregate_records, preview_payload
from .services.archive import SessionStore
from .services.normalizer import CLEAN_TYPE, SUPPORTED_TYPES, TRANSFER_TYPE, normalize_station_type
from .services.renderer import make_zip, render_reports


app = FastAPI(title="中转站、密闭式清洁站日报生成系统")
store = SessionStore(UPLOAD_ROOT, OUTPUT_ROOT)


class GenerateRequest(BaseModel):
    upload_id: str
    date: date
    types: list[str]
    zip: bool = False


@app.post("/api/upload")
def upload_files(
    ledger: Annotated[UploadFile, File(description="检查台账 .xls/.xlsx")],
    transfer_template: Annotated[UploadFile | None, File(description="中转站日报模板 .docx")] = None,
    clean_template: Annotated[UploadFile | None, File(description="密闭式清洁站日报模板 .docx")] = None,
) -> dict:
    try:
        validate_extension(ledger.filename, {".xls", ".xlsx"}, "检查台账")
        if transfer_template is not None:
            validate_extension(transfer_template.filename, {".docx"}, "中转站模板")
        if clean_template is not None:
            validate_extension(clean_template.filename, {".docx"}, "密闭式清洁站模板")
        session = store.save_upload(
            ledger_file=ledger,
            transfer_template_file=transfer_template,
            clean_template_file=clean_template,
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    parsed = session.parsed_ledger
    assert parsed is not None
    return {
        "upload_id": session.upload_id,
        "ledger": session.ledger_path.name if session.ledger_path else "",
        "record_count": len(parsed.records),
        "image_count": parsed.image_count,
        "warnings": parsed.warnings,
        "has_transfer_template": bool(session.transfer_template_path or DEFAULT_TRANSFER_TEMPLATE.exists()),
        "has_clean_template": bool(session.clean_template_path or DEFAULT_CLEAN_TEMPLATE.exists()),
        "dates": [{"date": key.isoformat(), "count": value} for key, value in parsed.available_dates.items()],
    }


@app.get("/api/dates")
def dates(upload_id: str) -> dict:
    session = require_session(upload_id)
    parsed = require_parsed(session)
    return {
        "upload_id": upload_id,
        "dates": [{"date": key.isoformat(), "count": value} for key, value in parsed.available_dates.items()],
        "warnings": parsed.warnings,
    }


@app.get("/api/preview")
def preview(
    upload_id: str,
    report_date: Annotated[date, Query(alias="date")],
    station_type: Annotated[str | None, Query(alias="type")] = None,
) -> dict:
    session = require_session(upload_id)
    parsed = require_parsed(session)
    result = aggregate_records(parsed.records, report_date)
    normalized_type = normalize_optional_type(station_type)
    return preview_payload(result, normalized_type)


@app.post("/api/generate")
def generate(request: GenerateRequest) -> dict:
    session = require_session(request.upload_id)
    parsed = require_parsed(session)
    selected_types = normalize_types(request.types)
    transfer_template = session.transfer_template_path or (DEFAULT_TRANSFER_TEMPLATE if DEFAULT_TRANSFER_TEMPLATE.exists() else None)
    clean_template = session.clean_template_path or (DEFAULT_CLEAN_TEMPLATE if DEFAULT_CLEAN_TEMPLATE.exists() else None)
    result = aggregate_records(parsed.records, request.date)
    try:
        files = render_reports(
            result=result,
            transfer_template=transfer_template,
            clean_template=clean_template,
            output_root=OUTPUT_ROOT,
            types=selected_types,
        )
        downloadable = make_zip(files, OUTPUT_ROOT, request.date) if request.zip or len(files) > 1 else files[0]
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    file_id = store.register_file(downloadable)
    return {
        "file_id": file_id,
        "filename": downloadable.name,
        "download_url": f"/api/download/{file_id}",
        "files": [file.name for file in files],
        "output_path": str(downloadable),
    }


@app.get("/api/download/{file_id}")
def download(file_id: str) -> FileResponse:
    try:
        path = store.get_file(file_id)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    media_type = "application/zip" if path.suffix.lower() == ".zip" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return FileResponse(path, media_type=media_type, filename=path.name)


def validate_extension(filename: str | None, allowed: set[str], label: str) -> None:
    suffix = Path(filename or "").suffix.lower()
    if suffix not in allowed:
        raise ValueError(f"{label}文件格式不正确，仅支持：{', '.join(sorted(allowed))}")


def require_session(upload_id: str):
    try:
        return store.get(upload_id)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc


def require_parsed(session):
    if session.parsed_ledger is None:
        raise HTTPException(status_code=400, detail="该上传会话尚未解析台账")
    return session.parsed_ledger


def normalize_optional_type(value: str | None) -> str | None:
    if not value:
        return None
    normalized = normalize_station_type(value)
    if normalized not in SUPPORTED_TYPES:
        raise HTTPException(status_code=400, detail=f"不支持的日报类型：{value}")
    return normalized


def normalize_types(values: list[str]) -> list[str]:
    if not values:
        raise HTTPException(status_code=400, detail="请至少选择一种日报类型")
    normalized: list[str] = []
    for value in values:
        station_type = normalize_station_type(value)
        if station_type not in SUPPORTED_TYPES:
            raise HTTPException(status_code=400, detail=f"不支持的日报类型：{value}")
        if station_type not in normalized:
            normalized.append(station_type)
    return normalized


web_dir = APP_ROOT / "web"
app.mount("/", StaticFiles(directory=web_dir, html=True), name="web")
