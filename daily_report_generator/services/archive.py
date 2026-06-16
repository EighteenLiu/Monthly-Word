from __future__ import annotations

import shutil
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from .ledger_reader import ParsedLedger, read_ledger


@dataclass
class UploadSession:
    upload_id: str
    root: Path
    ledger_path: Path | None = None
    transfer_template_path: Path | None = None
    clean_template_path: Path | None = None
    parsed_ledger: ParsedLedger | None = None
    created_at: datetime = field(default_factory=datetime.now)


class SessionStore:
    def __init__(self, upload_root: Path, output_root: Path) -> None:
        self.upload_root = upload_root
        self.output_root = output_root
        self.upload_root.mkdir(parents=True, exist_ok=True)
        self.output_root.mkdir(parents=True, exist_ok=True)
        self._sessions: dict[str, UploadSession] = {}
        self._files: dict[str, Path] = {}

    def create_session(self) -> UploadSession:
        upload_id = uuid.uuid4().hex
        root = self.upload_root / upload_id
        root.mkdir(parents=True, exist_ok=True)
        session = UploadSession(upload_id=upload_id, root=root)
        self._sessions[upload_id] = session
        return session

    def get(self, upload_id: str) -> UploadSession:
        session = self._sessions.get(upload_id)
        if not session:
            raise KeyError(f"upload_id 不存在或已过期：{upload_id}")
        return session

    def save_upload(
        self,
        *,
        ledger_file,
        transfer_template_file=None,
        clean_template_file=None,
    ) -> UploadSession:
        session = self.create_session()
        session.ledger_path = save_file(ledger_file, session.root / "ledger")
        if transfer_template_file is not None:
            session.transfer_template_path = save_file(transfer_template_file, session.root / "templates")
        if clean_template_file is not None:
            session.clean_template_path = save_file(clean_template_file, session.root / "templates")
        if session.ledger_path is None:
            raise ValueError("必须上传检查台账")
        session.parsed_ledger = read_ledger(session.ledger_path, session.root / "work")
        return session

    def register_file(self, path: Path) -> str:
        file_id = uuid.uuid4().hex
        self._files[file_id] = path
        return file_id

    def get_file(self, file_id: str) -> Path:
        path = self._files.get(file_id)
        if not path or not path.exists():
            raise KeyError(f"下载文件不存在或已过期：{file_id}")
        return path


def save_file(upload_file, target_dir: Path) -> Path | None:
    if upload_file is None:
        return None
    target_dir.mkdir(parents=True, exist_ok=True)
    filename = Path(upload_file.filename or "upload.bin").name
    target = target_dir / filename
    with target.open("wb") as handle:
        shutil.copyfileobj(upload_file.file, handle)
    return target

