from __future__ import annotations

import queue
import sys
import threading
import traceback
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, X, Button, Entry, Frame, Label, StringVar, Tk, Text, filedialog, messagebox, ttk


def get_project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


PROJECT_ROOT = get_project_root()

from scripts.monthly_generator import (  # noqa: E402
    DEFAULT_CONVERTED_DIR,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_RAW_INPUT_DIR,
    generate_monthly_report,
    normalize_station_type,
)


class TextWriter:
    def __init__(self, messages: queue.Queue[str]) -> None:
        self.messages = messages

    def write(self, text: str) -> None:
        if text:
            self.messages.put(text)

    def flush(self) -> None:
        pass


class MonthlyReportApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("西城区月报生成工具")
        self.root.geometry("820x560")
        self.root.minsize(760, 500)
        self.ensure_default_dirs()

        self.year_var = StringVar(value="2026")
        self.month_var = StringVar(value="5")
        self.type_var = StringVar(value="1")
        self.raw_dir_var = StringVar(value=str(DEFAULT_RAW_INPUT_DIR))
        self.output_dir_var = StringVar(value=str(DEFAULT_OUTPUT_DIR))
        self.status_var = StringVar(value="待运行")
        self.messages: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None

        self.build_ui()
        self.root.after(120, self.drain_messages)

    def build_ui(self) -> None:
        outer = Frame(self.root, padx=18, pady=16)
        outer.pack(fill=BOTH, expand=True)

        title = Label(outer, text="西城区月报生成工具", font=("Microsoft YaHei", 16, "bold"))
        title.pack(anchor="w")

        form = Frame(outer, pady=14)
        form.pack(fill=X)

        row1 = Frame(form)
        row1.pack(fill=X, pady=6)
        self.add_labeled_entry(row1, "年份", self.year_var, width=16)
        self.add_labeled_entry(row1, "报告月份", self.month_var, width=16)

        type_box = Frame(row1)
        type_box.pack(side=LEFT, padx=(18, 0))
        Label(type_box, text="报告类型").pack(anchor="w")
        type_select = ttk.Combobox(
            type_box,
            textvariable=self.type_var,
            values=("1 - 密闭式清洁站", "2 - 中转站"),
            width=20,
            state="readonly",
        )
        type_select.current(0)
        type_select.pack()
        type_select.bind("<<ComboboxSelected>>", self.on_type_selected)

        self.add_folder_row(form, "日报所在文件夹", self.raw_dir_var, self.choose_raw_dir)
        self.add_folder_row(form, "月报输出文件夹", self.output_dir_var, self.choose_output_dir)

        actions = Frame(outer)
        actions.pack(fill=X, pady=(4, 12))
        self.generate_button = Button(actions, text="生成月报", width=14, command=self.start_generation)
        self.generate_button.pack(side=LEFT)
        Button(actions, text="清空日志", width=12, command=self.clear_log).pack(side=LEFT, padx=(10, 0))
        Label(actions, textvariable=self.status_var, fg="#1d4ed8").pack(side=RIGHT)

        Label(outer, text="运行日志").pack(anchor="w")
        log_frame = Frame(outer)
        log_frame.pack(fill=BOTH, expand=True)
        self.log = Text(log_frame, height=16, wrap="word", font=("Consolas", 10))
        self.log.pack(fill=BOTH, expand=True)
        self.write_log("请选择日报所在文件夹后点击“生成月报”。\n")
        self.write_log("文件夹可以选择 01_原始日报 根目录，也可以直接选择 清洁站 或 中转站 文件夹。\n")

    def ensure_default_dirs(self) -> None:
        for directory in (
            DEFAULT_RAW_INPUT_DIR / "清洁站",
            DEFAULT_RAW_INPUT_DIR / "中转站",
            DEFAULT_CONVERTED_DIR / "清洁站",
            DEFAULT_CONVERTED_DIR / "中转站",
            DEFAULT_OUTPUT_DIR,
        ):
            directory.mkdir(parents=True, exist_ok=True)

    def add_labeled_entry(self, parent: Frame, label: str, variable: StringVar, width: int) -> None:
        box = Frame(parent)
        box.pack(side=LEFT)
        Label(box, text=label).pack(anchor="w")
        Entry(box, textvariable=variable, width=width).pack()

    def add_folder_row(self, parent: Frame, label: str, variable: StringVar, command) -> None:
        row = Frame(parent)
        row.pack(fill=X, pady=6)
        Label(row, text=label, width=14, anchor="w").pack(side=LEFT)
        Entry(row, textvariable=variable).pack(side=LEFT, fill=X, expand=True)
        Button(row, text="选择", width=8, command=command).pack(side=LEFT, padx=(8, 0))

    def on_type_selected(self, _event=None) -> None:
        selected = self.type_var.get()
        self.type_var.set("2" if selected.startswith("2") else "1")

    def choose_raw_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择日报所在文件夹", initialdir=self.raw_dir_var.get() or str(PROJECT_ROOT))
        if selected:
            self.raw_dir_var.set(selected)

    def choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="选择月报输出文件夹", initialdir=self.output_dir_var.get() or str(PROJECT_ROOT))
        if selected:
            self.output_dir_var.set(selected)

    def clear_log(self) -> None:
        self.log.delete("1.0", END)

    def write_log(self, text: str) -> None:
        self.log.insert(END, text)
        self.log.see(END)

    def drain_messages(self) -> None:
        while True:
            try:
                message = self.messages.get_nowait()
            except queue.Empty:
                break
            self.write_log(message)
        self.root.after(120, self.drain_messages)

    def start_generation(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("正在运行", "月报正在生成，请稍候。")
            return

        year = self.year_var.get().strip()
        month = self.month_var.get().strip()
        station_type = self.type_var.get().strip()
        if station_type.startswith("2"):
            station_type = "2"
        elif station_type.startswith("1"):
            station_type = "1"

        if not year.isdigit() or not month.isdigit():
            messagebox.showerror("参数错误", "年份和月份必须是数字。")
            return

        raw_dir = Path(self.raw_dir_var.get().strip())
        output_dir = Path(self.output_dir_var.get().strip())
        if not raw_dir.exists():
            messagebox.showerror("路径错误", "日报所在文件夹不存在。")
            return

        self.status_var.set("运行中")
        self.generate_button.config(state="disabled")
        self.write_log("\n========== 开始生成 ==========\n")

        args = (int(year), month, station_type, raw_dir, output_dir)
        self.worker = threading.Thread(target=self.run_generation, args=args, daemon=True)
        self.worker.start()

    def run_generation(self, year: int, month: str, station_type: str, raw_dir: Path, output_dir: Path) -> None:
        original_stdout = sys.stdout
        sys.stdout = TextWriter(self.messages)
        try:
            normalized_type = normalize_station_type(station_type)
            output_path = generate_monthly_report(
                raw_input_dir=raw_dir,
                converted_dir=DEFAULT_CONVERTED_DIR,
                output_dir=output_dir,
                month=month,
                year=year,
                station_type=normalized_type,
                engine="auto",
                skip_convert=False,
            )
            self.messages.put(f"\n[完成] 月报已生成：{output_path}\n")
            self.root.after(0, lambda: self.status_var.set("完成"))
            self.root.after(0, lambda: messagebox.showinfo("生成完成", f"月报已生成：\n{output_path}"))
        except Exception:
            self.messages.put("\n[失败]\n")
            self.messages.put(traceback.format_exc())
            self.root.after(0, lambda: self.status_var.set("失败"))
            self.root.after(0, lambda: messagebox.showerror("生成失败", "生成失败，请查看运行日志。"))
        finally:
            sys.stdout = original_stdout
            self.root.after(0, lambda: self.generate_button.config(state="normal"))


def main() -> None:
    root = Tk()
    MonthlyReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
