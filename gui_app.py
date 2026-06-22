from __future__ import annotations

import queue
import shutil
import sys
import threading
import traceback
import uuid
from datetime import date
from pathlib import Path
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    X,
    BooleanVar,
    Button,
    Canvas,
    Entry,
    Frame,
    Label,
    StringVar,
    Tk,
    Text,
    filedialog,
    messagebox,
    ttk,
)


def get_project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


PROJECT_ROOT = get_project_root()

from daily_report_generator.config import (  # noqa: E402
    DEFAULT_CLEAN_TEMPLATE,
    DEFAULT_TRANSFER_TEMPLATE,
    OUTPUT_ROOT as DAILY_OUTPUT_ROOT,
)
from daily_report_generator.services.aggregator import aggregate_records  # noqa: E402
from daily_report_generator.services.ledger_reader import read_ledger, read_ledger_dates  # noqa: E402
from daily_report_generator.services.normalizer import CLEAN_TYPE, TRANSFER_TYPE  # noqa: E402
from daily_report_generator.services.renderer import render_reports  # noqa: E402
from scripts.monthly_generator import (  # noqa: E402
    DEFAULT_CONVERTED_DIR,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_RAW_INPUT_DIR,
    DEFAULT_TEMPLATE_DIR,
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


class ReportToolApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("西城区日报、月报生成工具")
        self.root.geometry("980x680")
        self.root.minsize(900, 620)

        self.messages: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.status_var = StringVar(value="待运行")
        self.daily_preview_text = StringVar(value="未解析日期")
        self.daily_available_dates: list[tuple[str, int]] = []
        self.daily_date_vars: list[tuple[str, BooleanVar]] = []
        self.daily_select_all_var = BooleanVar(value=True)
        self.syncing_daily_dates = False

        self.init_variables()
        self.ensure_default_dirs()
        self.build_ui()
        self.root.after(120, self.drain_messages)

    def init_variables(self) -> None:
        self.daily_ledger_var = StringVar(value=str(PROJECT_ROOT / "input" / "202606006中转站.xls"))
        self.daily_transfer_template_var = StringVar(value=str(DEFAULT_TRANSFER_TEMPLATE))
        self.daily_clean_template_var = StringVar(value=str(DEFAULT_CLEAN_TEMPLATE))
        self.daily_output_dir_var = StringVar(value=str(DAILY_OUTPUT_ROOT))
        self.daily_transfer_var = BooleanVar(value=True)
        self.daily_clean_var = BooleanVar(value=True)

        self.month_year_var = StringVar(value="2026")
        self.month_month_var = StringVar(value="5")
        self.month_type_var = StringVar(value="1 - 密闭式清洁站")
        self.month_raw_dir_var = StringVar(value=str(DEFAULT_RAW_INPUT_DIR))
        self.month_template_path_var = StringVar(value="")
        self.month_output_dir_var = StringVar(value=str(DEFAULT_OUTPUT_DIR))

    def ensure_default_dirs(self) -> None:
        for directory in (
            DEFAULT_RAW_INPUT_DIR / "清洁站",
            DEFAULT_RAW_INPUT_DIR / "中转站",
            DEFAULT_CONVERTED_DIR / "清洁站",
            DEFAULT_CONVERTED_DIR / "中转站",
            DEFAULT_TEMPLATE_DIR,
            DEFAULT_OUTPUT_DIR,
            DAILY_OUTPUT_ROOT,
        ):
            directory.mkdir(parents=True, exist_ok=True)

    def build_ui(self) -> None:
        outer = Frame(self.root, padx=18, pady=16)
        outer.pack(fill=BOTH, expand=True)

        header = Frame(outer)
        header.pack(fill=X)
        Label(header, text="西城区日报、月报生成工具", font=("Microsoft YaHei", 16, "bold")).pack(side=LEFT)
        Label(header, textvariable=self.status_var, fg="#1d4ed8").pack(side=RIGHT)

        notebook = ttk.Notebook(outer)
        notebook.pack(fill=X, expand=False, pady=(10, 6))

        daily_tab = Frame(notebook, padx=14, pady=12)
        monthly_tab = Frame(notebook, padx=14, pady=12)
        notebook.add(daily_tab, text="日报生成")
        notebook.add(monthly_tab, text="月报生成")

        self.build_daily_tab(daily_tab)
        self.build_monthly_tab(monthly_tab)

        footer = Frame(outer)
        footer.pack(fill=BOTH, expand=True)
        Label(footer, text="运行日志").pack(anchor="w")
        self.log = Text(footer, height=15, wrap="word", font=("Consolas", 10))
        self.log.pack(fill=BOTH, expand=True, pady=(4, 0))
        Button(footer, text="清空日志", width=12, command=self.clear_log).pack(anchor="e", pady=(6, 0))
        self.write_log("请选择“日报生成”或“月报生成”页签后开始。\n")

    def build_daily_tab(self, parent: Frame) -> None:
        Label(parent, text="从检查台账生成中转站、密闭式清洁站日报", font=("Microsoft YaHei", 12, "bold")).pack(anchor="w")
        form = Frame(parent, pady=6)
        form.pack(fill=X)

        self.add_file_row(form, "检查台账", self.daily_ledger_var, self.choose_daily_ledger, width=16)
        self.add_file_row(form, "中转站模板", self.daily_transfer_template_var, self.choose_daily_transfer_template, width=16)
        self.add_file_row(form, "清洁站模板", self.daily_clean_template_var, self.choose_daily_clean_template, width=16)
        self.add_folder_row(form, "日报输出文件夹", self.daily_output_dir_var, self.choose_daily_output_dir, width=16)

        options = Frame(parent)
        options.pack(fill=X, pady=(4, 8))
        Label(options, text="日报类型").pack(side=LEFT)
        ttk.Checkbutton(options, text="中转站", variable=self.daily_transfer_var).pack(side=LEFT, padx=(0, 10))
        ttk.Checkbutton(options, text="密闭式清洁站", variable=self.daily_clean_var).pack(side=LEFT, padx=(0, 10))

        actions = Frame(parent)
        actions.pack(fill=X, pady=(2, 6))
        self.parse_daily_button = Button(actions, text="解析日期", width=12, command=self.start_daily_parse)
        self.parse_daily_button.pack(side=LEFT)
        self.generate_daily_button = Button(actions, text="生成日报", width=12, command=self.start_daily_generation)
        self.generate_daily_button.pack(side=LEFT, padx=(10, 0))

        Label(parent, textvariable=self.daily_preview_text, fg="#374151").pack(anchor="w")
        date_header = Frame(parent)
        date_header.pack(fill=X, pady=(4, 0))
        Label(date_header, text="检查日期").pack(side=LEFT)
        ttk.Checkbutton(
            date_header,
            text="全选",
            variable=self.daily_select_all_var,
            command=self.toggle_all_daily_dates,
        ).pack(side=LEFT, padx=(12, 0))
        daily_date_box = Frame(parent)
        daily_date_box.pack(fill=X, pady=(3, 4))
        self.daily_date_canvas = Canvas(daily_date_box, height=112, highlightthickness=1, highlightbackground="#d1d5db")
        daily_date_scrollbar = ttk.Scrollbar(daily_date_box, orient="vertical", command=self.daily_date_canvas.yview)
        self.daily_date_canvas.configure(yscrollcommand=daily_date_scrollbar.set)
        self.daily_date_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        daily_date_scrollbar.pack(side=RIGHT, fill="y")
        self.daily_date_frame = Frame(self.daily_date_canvas)
        self.daily_date_window = self.daily_date_canvas.create_window((0, 0), window=self.daily_date_frame, anchor="nw")
        self.daily_date_frame.bind(
            "<Configure>",
            lambda _event: self.daily_date_canvas.configure(scrollregion=self.daily_date_canvas.bbox("all")),
        )
        self.daily_date_canvas.bind(
            "<Configure>",
            lambda event: self.daily_date_canvas.itemconfigure(self.daily_date_window, width=event.width),
        )
        self.daily_date_canvas.bind("<Enter>", self.bind_daily_date_wheel)
        self.daily_date_canvas.bind("<Leave>", self.unbind_daily_date_wheel)

    def build_monthly_tab(self, parent: Frame) -> None:
        Label(parent, text="从已转换日报汇总生成月报", font=("Microsoft YaHei", 12, "bold")).pack(anchor="w")
        form = Frame(parent, pady=6)
        form.pack(fill=X)

        row1 = Frame(form)
        row1.pack(fill=X, pady=6)
        self.add_labeled_entry(row1, "年份", self.month_year_var, width=14)
        self.add_labeled_entry(row1, "报告月份", self.month_month_var, width=14)
        type_box = Frame(row1)
        type_box.pack(side=LEFT, padx=(18, 0))
        Label(type_box, text="报告类型").pack(anchor="w")
        type_select = ttk.Combobox(
            type_box,
            textvariable=self.month_type_var,
            values=("1 - 密闭式清洁站", "2 - 中转站"),
            width=20,
            state="readonly",
        )
        type_select.pack()

        self.add_folder_row(form, "日报所在文件夹", self.month_raw_dir_var, self.choose_month_raw_dir, width=16)
        self.add_file_row(form, "月报docx模板", self.month_template_path_var, self.choose_month_template_file, width=16)
        self.add_folder_row(form, "月报输出文件夹", self.month_output_dir_var, self.choose_month_output_dir, width=16)

        actions = Frame(parent)
        actions.pack(fill=X, pady=(4, 10))
        self.generate_monthly_button = Button(actions, text="生成月报", width=12, command=self.start_monthly_generation)
        self.generate_monthly_button.pack(side=LEFT)
        Label(parent, text="文件夹可以选择 01_原始日报 根目录，也可以直接选择 清洁站 或 中转站 文件夹。", fg="#374151").pack(anchor="w")

    def add_labeled_entry(self, parent: Frame, label: str, variable: StringVar, width: int) -> None:
        box = Frame(parent)
        box.pack(side=LEFT)
        Label(box, text=label).pack(anchor="w")
        Entry(box, textvariable=variable, width=width).pack()

    def add_folder_row(self, parent: Frame, label: str, variable: StringVar, command, width: int = 14) -> None:
        row = Frame(parent)
        row.pack(fill=X, pady=6)
        Label(row, text=label, width=width, anchor="w").pack(side=LEFT)
        Entry(row, textvariable=variable).pack(side=LEFT, fill=X, expand=True)
        Button(row, text="选择", width=8, command=command).pack(side=LEFT, padx=(8, 0))

    def add_file_row(self, parent: Frame, label: str, variable: StringVar, command, width: int = 14) -> None:
        row = Frame(parent)
        row.pack(fill=X, pady=6)
        Label(row, text=label, width=width, anchor="w").pack(side=LEFT)
        Entry(row, textvariable=variable).pack(side=LEFT, fill=X, expand=True)
        Button(row, text="选择", width=8, command=command).pack(side=LEFT, padx=(8, 0))

    def choose_daily_ledger(self) -> None:
        selected = filedialog.askopenfilename(
            title="选择检查台账",
            initialdir=str(PROJECT_ROOT / "input"),
            filetypes=(("Excel 台账", "*.xls *.xlsx"), ("所有文件", "*.*")),
        )
        if selected:
            self.daily_ledger_var.set(selected)

    def choose_daily_transfer_template(self) -> None:
        self.choose_docx_file("选择中转站日报模板", self.daily_transfer_template_var, PROJECT_ROOT / "input")

    def choose_daily_clean_template(self) -> None:
        self.choose_docx_file("选择密闭式清洁站日报模板", self.daily_clean_template_var, PROJECT_ROOT / "input")

    def choose_daily_output_dir(self) -> None:
        self.choose_folder("选择日报输出文件夹", self.daily_output_dir_var)

    def choose_month_raw_dir(self) -> None:
        self.choose_folder("选择日报所在文件夹", self.month_raw_dir_var)

    def choose_month_output_dir(self) -> None:
        self.choose_folder("选择月报输出文件夹", self.month_output_dir_var)

    def choose_month_template_file(self) -> None:
        self.choose_docx_file("选择月报docx模板", self.month_template_path_var, DEFAULT_TEMPLATE_DIR)

    def choose_docx_file(self, title: str, variable: StringVar, initial_dir: Path) -> None:
        selected = filedialog.askopenfilename(
            title=title,
            initialdir=str(initial_dir if initial_dir.exists() else PROJECT_ROOT),
            filetypes=(("Word 文档", "*.docx"), ("所有文件", "*.*")),
        )
        if selected:
            variable.set(selected)

    def choose_folder(self, title: str, variable: StringVar) -> None:
        selected = filedialog.askdirectory(title=title, initialdir=variable.get() or str(PROJECT_ROOT))
        if selected:
            variable.set(selected)

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

    def ensure_not_running(self, label: str) -> bool:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("正在运行", f"{label}正在执行，请稍候。")
            return False
        return True

    def start_daily_parse(self) -> None:
        if not self.ensure_not_running("日报解析"):
            return
        ledger_path = Path(self.daily_ledger_var.get().strip())
        if not ledger_path.exists():
            messagebox.showerror("路径错误", "检查台账不存在。")
            return
        self.status_var.set("解析日报台账中")
        self.set_daily_buttons("disabled")
        self.write_log("\n========== 开始解析日报台账 ==========\n")
        self.worker = threading.Thread(target=self.run_daily_parse, args=(ledger_path,), daemon=True)
        self.worker.start()

    def run_daily_parse(self, ledger_path: Path) -> None:
        try:
            available_dates = read_ledger_dates(ledger_path)
            dates = [(key.isoformat(), value) for key, value in available_dates.items()]
            self.root.after(0, lambda: self.set_daily_dates(dates))
            total = sum(value for _, value in dates)
            self.messages.put(f"[完成] 日期解析完成：{len(dates)} 个日期，{total} 条记录。未转换台账，未抽取图片。\n")
            self.root.after(0, lambda: self.status_var.set("日期已解析"))
        except Exception:
            self.messages.put("\n[失败]\n")
            self.messages.put(traceback.format_exc())
            self.root.after(0, lambda: self.status_var.set("解析失败"))
            self.root.after(0, lambda: messagebox.showerror("解析失败", "解析台账失败，请查看运行日志。"))
        finally:
            self.root.after(0, lambda: self.set_daily_buttons("normal"))

    def set_daily_dates(self, dates: list[tuple[str, int]]) -> None:
        self.daily_available_dates = dates
        self.daily_date_vars = []
        for child in self.daily_date_frame.winfo_children():
            child.destroy()

        self.daily_select_all_var.set(bool(dates))
        for iso_date, count in dates:
            selected = BooleanVar(value=True)
            self.daily_date_vars.append((iso_date, selected))
            ttk.Checkbutton(
                self.daily_date_frame,
                text=f"{iso_date}（{count} 条）",
                variable=selected,
                command=self.update_daily_select_all,
            ).pack(anchor="w")

        if dates:
            self.daily_preview_text.set(f"已识别 {len(dates)} 个检查日期，默认全部勾选")
        else:
            self.daily_preview_text.set("未识别到可用日期")

    def selected_daily_dates(self) -> list[date]:
        return [date.fromisoformat(value) for value, selected in self.daily_date_vars if selected.get()]

    def toggle_all_daily_dates(self) -> None:
        if self.syncing_daily_dates:
            return
        self.syncing_daily_dates = True
        selected = self.daily_select_all_var.get()
        for _, variable in self.daily_date_vars:
            variable.set(selected)
        self.syncing_daily_dates = False

    def update_daily_select_all(self) -> None:
        if self.syncing_daily_dates:
            return
        self.syncing_daily_dates = True
        self.daily_select_all_var.set(bool(self.daily_date_vars) and all(variable.get() for _, variable in self.daily_date_vars))
        self.syncing_daily_dates = False

    def selected_daily_types(self) -> list[str]:
        types: list[str] = []
        if self.daily_transfer_var.get():
            types.append(TRANSFER_TYPE)
        if self.daily_clean_var.get():
            types.append(CLEAN_TYPE)
        return types

    def start_daily_generation(self) -> None:
        if not self.ensure_not_running("日报生成"):
            return
        selected_dates = self.selected_daily_dates()
        if not selected_dates:
            messagebox.showerror("参数错误", "请先解析并选择至少一个检查日期。")
            return
        types = self.selected_daily_types()
        if not types:
            messagebox.showerror("参数错误", "请至少选择一种日报类型。")
            return
        ledger_path = Path(self.daily_ledger_var.get().strip())
        transfer_template = Path(self.daily_transfer_template_var.get().strip()) if self.daily_transfer_template_var.get().strip() else None
        clean_template = Path(self.daily_clean_template_var.get().strip()) if self.daily_clean_template_var.get().strip() else None
        output_dir = Path(self.daily_output_dir_var.get().strip())
        if not ledger_path.exists():
            messagebox.showerror("路径错误", "检查台账不存在。")
            return
        if TRANSFER_TYPE in types and (not transfer_template or not transfer_template.exists()):
            messagebox.showerror("路径错误", "中转站模板不存在。")
            return
        if CLEAN_TYPE in types and (not clean_template or not clean_template.exists()):
            messagebox.showerror("路径错误", "密闭式清洁站模板不存在。")
            return
        self.status_var.set("正在生成")
        self.set_daily_buttons("disabled")
        self.generate_daily_button.config(text="正在生成")
        self.write_log("\n========== 开始生成日报 ==========\n")
        args = (ledger_path, selected_dates, types, transfer_template, clean_template, output_dir)
        self.worker = threading.Thread(target=self.run_daily_generation, args=args, daemon=True)
        self.worker.start()

    def run_daily_generation(
        self,
        ledger_path: Path,
        selected_dates: list[date],
        types: list[str],
        transfer_template: Path | None,
        clean_template: Path | None,
        output_dir: Path,
    ) -> None:
        work_dir: Path | None = None
        try:
            work_dir = output_dir / "_gui_work" / uuid.uuid4().hex
            parsed = read_ledger(ledger_path, work_dir)
            generated_files: list[Path] = []
            skipped: list[str] = []
            for selected_date in selected_dates:
                result = aggregate_records(parsed.records, selected_date)
                try:
                    files = render_reports(result, transfer_template, clean_template, output_dir, types)
                except ValueError as exc:
                    skipped.append(str(exc))
                    continue
                generated_files.extend(files)
            if not generated_files:
                message = "所选日期没有可生成的日报数据。"
                if skipped:
                    message += "\n" + "\n".join(skipped)
                raise ValueError(message)
            self.messages.put("[完成] 日报已生成：\n" + "\n".join(str(path) for path in generated_files) + "\n")
            if skipped:
                self.messages.put("[跳过]\n" + "\n".join(skipped) + "\n")
            self.root.after(0, lambda: self.status_var.set("日报生成完成"))
            self.root.after(0, lambda: messagebox.showinfo("生成完成", "日报已生成：\n" + "\n".join(str(path) for path in generated_files)))
        except Exception:
            self.messages.put("\n[失败]\n")
            self.messages.put(traceback.format_exc())
            self.root.after(0, lambda: self.status_var.set("日报生成失败"))
            self.root.after(0, lambda: messagebox.showerror("生成失败", "生成日报失败，请查看运行日志。"))
        finally:
            if work_dir is not None:
                shutil.rmtree(work_dir, ignore_errors=True)
            self.root.after(0, lambda: self.set_daily_buttons("normal"))

    def set_daily_buttons(self, state: str) -> None:
        for button in (self.parse_daily_button, self.generate_daily_button):
            button.config(state=state)
        if state == "normal":
            self.generate_daily_button.config(text="生成日报")

    def bind_daily_date_wheel(self, _event: object) -> None:
        self.root.bind_all("<MouseWheel>", self.scroll_daily_dates)

    def unbind_daily_date_wheel(self, _event: object) -> None:
        self.root.unbind_all("<MouseWheel>")

    def scroll_daily_dates(self, event: object) -> None:
        delta = getattr(event, "delta", 0)
        if delta:
            self.daily_date_canvas.yview_scroll(int(-1 * (delta / 120)), "units")

    def start_monthly_generation(self) -> None:
        if not self.ensure_not_running("月报生成"):
            return
        year = self.month_year_var.get().strip()
        month = self.month_month_var.get().strip()
        station_type = "2" if self.month_type_var.get().startswith("2") else "1"
        if not year.isdigit() or not month.isdigit():
            messagebox.showerror("参数错误", "年份和月份必须是数字。")
            return
        raw_dir = Path(self.month_raw_dir_var.get().strip())
        template_text = self.month_template_path_var.get().strip()
        template_path = Path(template_text) if template_text else None
        output_dir = Path(self.month_output_dir_var.get().strip())
        if not raw_dir.exists():
            messagebox.showerror("路径错误", "日报所在文件夹不存在。")
            return
        if template_path and not template_path.exists():
            messagebox.showerror("路径错误", "月报docx模板不存在。")
            return

        self.status_var.set("生成月报中")
        self.generate_monthly_button.config(state="disabled")
        self.write_log("\n========== 开始生成月报 ==========\n")
        args = (int(year), month, station_type, raw_dir, output_dir, template_path)
        self.worker = threading.Thread(target=self.run_monthly_generation, args=args, daemon=True)
        self.worker.start()

    def run_monthly_generation(
        self,
        year: int,
        month: str,
        station_type: str,
        raw_dir: Path,
        output_dir: Path,
        template_path: Path | None,
    ) -> None:
        original_stdout = sys.stdout
        sys.stdout = TextWriter(self.messages)
        try:
            output_path = generate_monthly_report(
                raw_input_dir=raw_dir,
                converted_dir=DEFAULT_CONVERTED_DIR,
                output_dir=output_dir,
                month=month,
                year=year,
                station_type=normalize_station_type(station_type),
                engine="auto",
                skip_convert=False,
                template_path=template_path,
            )
            self.messages.put(f"\n[完成] 月报已生成：{output_path}\n")
            self.root.after(0, lambda: self.status_var.set("月报生成完成"))
            self.root.after(0, lambda: messagebox.showinfo("生成完成", f"月报已生成：\n{output_path}"))
        except Exception:
            self.messages.put("\n[失败]\n")
            self.messages.put(traceback.format_exc())
            self.root.after(0, lambda: self.status_var.set("月报生成失败"))
            self.root.after(0, lambda: messagebox.showerror("生成失败", "生成月报失败，请查看运行日志。"))
        finally:
            sys.stdout = original_stdout
            self.root.after(0, lambda: self.generate_monthly_button.config(state="normal"))


def main() -> None:
    root = Tk()
    ReportToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
