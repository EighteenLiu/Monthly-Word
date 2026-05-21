"""
Batch-convert Word .doc daily reports to .docx.

Default input:
    01_原始日报/<station>/*.doc

Default output:
    02_转换后日报/<station>/*.docx

The converter tries Microsoft Word COM automation first on Windows, then falls
back to LibreOffice/soffice if available in PATH.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT_DIR = PROJECT_ROOT / "01_原始日报"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "02_转换后日报"


class ConversionError(RuntimeError):
    """Raised when a document cannot be converted."""


def iter_doc_files(input_dir: Path) -> list[Path]:
    return sorted(
        path
        for path in input_dir.rglob("*.doc")
        if path.is_file() and path.suffix.lower() == ".doc"
    )


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
    except Exception as exc:  # COM errors vary by Office version/localization.
        raise ConversionError(str(exc)) from exc
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def find_soffice() -> str | None:
    for executable in ("soffice", "libreoffice"):
        found = shutil.which(executable)
        if found:
            return found
    return None


def convert_with_libreoffice(source: Path, target: Path) -> None:
    soffice = find_soffice()
    if not soffice:
        raise ConversionError("LibreOffice/soffice was not found in PATH")

    target.parent.mkdir(parents=True, exist_ok=True)
    command = [
        soffice,
        "--headless",
        "--convert-to",
        "docx",
        "--outdir",
        str(target.parent.resolve()),
        str(source.resolve()),
    ]
    result = subprocess.run(command, capture_output=True, text=True, check=False)
    generated = target.parent / f"{source.stem}.docx"
    if result.returncode != 0 or not generated.exists():
        message = (result.stderr or result.stdout or "LibreOffice conversion failed").strip()
        raise ConversionError(message)
    if generated.resolve() != target.resolve():
        generated.replace(target)


def convert_one(source: Path, target: Path, engine: str) -> None:
    if target.exists() and target.stat().st_mtime >= source.stat().st_mtime:
        print(f"[跳过] {source} -> {target} 已是最新")
        return

    engines = ["word", "libreoffice"] if engine == "auto" else [engine]
    errors: list[str] = []
    for selected in engines:
        try:
            if selected == "word":
                convert_with_word(source, target)
            elif selected == "libreoffice":
                convert_with_libreoffice(source, target)
            else:
                raise ConversionError(f"未知转换引擎: {selected}")
            print(f"[完成] {source} -> {target}")
            return
        except ConversionError as exc:
            errors.append(f"{selected}: {exc}")

    raise ConversionError("; ".join(errors))


def convert_all(input_dir: Path, output_dir: Path, engine: str = "auto") -> int:
    files = iter_doc_files(input_dir)
    if not files:
        print(f"[提示] 未发现 .doc 文件: {input_dir}")
        return 0

    failed = 0
    for source in files:
        relative = source.relative_to(input_dir)
        target = (output_dir / relative).with_suffix(".docx")
        try:
            convert_one(source, target, engine)
        except ConversionError as exc:
            failed += 1
            print(f"[失败] {source}: {exc}", file=sys.stderr)

    if failed:
        print(f"[结果] {len(files) - failed} 个成功，{failed} 个失败")
        return 1

    print(f"[结果] 全部转换完成，共 {len(files)} 个文件")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="批量将原始日报 .doc 转换为 .docx")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT_DIR, help="原始日报目录")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_DIR, help="转换后日报目录")
    parser.add_argument(
        "--engine",
        choices=("auto", "word", "libreoffice"),
        default="auto",
        help="转换引擎，默认 auto",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    return convert_all(args.input, args.output, args.engine)


if __name__ == "__main__":
    raise SystemExit(main())
