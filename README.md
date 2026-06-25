# CleanAndCenter Word 报告生成工具

西城区中转站、密闭式清洁站日报和月报生成工具。项目现在提供一个统一桌面窗口，可在同一程序中完成：

- 从检查台账生成每日“中转站检查情况”和“密闭式清洁站检查情况”Word 日报。
- 从原始 Word 日报汇总生成月度检查通报。

## 功能简介

### 日报生成

- 支持读取 `.xls` / `.xlsx` 检查台账。
- 支持双层表头台账，按“创建时间”优先、“案件上报时间”兜底解析日期。
- `解析日期` 只读取文字和日期，不转换文件、不抽取图片，适合快速选择日期。
- 支持一次选择多个日期批量生成日报。
- 支持同时勾选中转站和密闭式清洁站；若某天某类没有数据，会自动跳过，不报错。
- 生成日报时才转换 `.xls`、抽取嵌入图片、按行列锚点映射照片。
- 日报按类型输出到 `output/日报/清洁站/` 和 `output/日报/中转站/`。

### 月报生成

- 批量将原始 `.doc` 日报转换为 `.docx`。
- 自动复用已转换且未过期的 `.docx` 文件。
- 按默认周期“上月 20 日至本月 19 日”汇总日报。
- 支持两类月报：
  - `清洁站`
  - `中转站`
- 输出月度检查通报 Word 文档。

## 推荐运行方式

安装依赖：

```powershell
pip install -r requirements.txt
```

启动统一桌面窗口：

```powershell
run_gui.bat
```

窗口内包含两个页签：

- `日报生成`
- `月报生成`

不需要打开浏览器。

## 日报生成流程

1. 进入 `日报生成` 页签。
2. 选择检查台账。
3. 选择中转站日报模板和密闭式清洁站日报模板。
4. 选择日报输出文件夹。
5. 点击 `解析日期`。
6. 在日期列表中按 `Ctrl` 或 `Shift` 多选日期。
7. 勾选要生成的日报类型。
8. 点击 `生成日报`。

输出示例：

```text
output/日报/中转站/6月5日中转站检查情况.docx
output/日报/清洁站/6月6日密闭式清洁站检查情况.docx
```

默认日报模板位于：

```text
input/中转站日报_jinja模板.docx
input/密闭式清洁站日报_jinja模板.docx
```

## 月报生成流程

1. 将原始 `.doc` 日报放入对应目录：

```text
01_原始日报/清洁站/
01_原始日报/中转站/
```

文件名需要包含日期，例如：

```text
4月20日密闭式清洁站检查情况.doc
5月9日中转站检查情况.doc
```

2. 进入 `月报生成` 页签。
3. 填写年份、月份和报告类型。
4. 选择日报所在文件夹、月报模板和输出目录。
5. 点击 `生成月报`。

默认输出目录：

```text
04_输出月报/
```

输出示例：

```text
西城区2026年5月密闭式清洁站检查通报.docx
西城区2026年5月中转站检查通报.docx
```

## 命令行方式

### 生成日报

```powershell
python -m daily_report_generator.cli --ledger input\202606006中转站.xls --date 2026-06-05 --types 中转站
```

### 生成月报

```powershell
python scripts\run_all.py --year 2026 --month 5 --station-type 清洁站
python scripts\run_all.py --year 2026 --month 5 --station-type 中转站
```

指定转换引擎：

```powershell
python scripts\run_all.py --year 2026 --month 5 --station-type 清洁站 --engine word
python scripts\run_all.py --year 2026 --month 5 --station-type 清洁站 --engine libreoffice
```

如果已经有转换好的 `.docx`，可以跳过转换：

```powershell
python scripts\run_all.py --year 2026 --month 5 --station-type 清洁站 --skip-convert
```

单独转换原始日报：

```powershell
python scripts\convert_script.py
python scripts\convert_script.py --input 01_原始日报 --output 02_转换后日报
```

## 项目结构

```text
CleanAndCenter_Word/
├─ 01_原始日报/                         # 月报输入：原始 .doc 日报
│  ├─ 清洁站/
│  └─ 中转站/
├─ 02_转换后日报/                       # 月报中间文件：转换后的 .docx 日报
│  ├─ 清洁站/
│  └─ 中转站/
├─ 03_月报模板/                         # 月报模板目录
├─ 04_输出月报/                         # 月报输出目录
├─ input/                               # 日报模板和样例台账
│  ├─ 中转站日报_jinja模板.docx
│  └─ 密闭式清洁站日报_jinja模板.docx
├─ daily_report_generator/
│  ├─ services/                         # 日报解析、聚合、渲染服务
│  ├─ output/                           # 日报输出目录
│  ├─ app.py                            # 可选 Web API 入口
│  └─ cli.py                            # 日报命令行入口
├─ scripts/
│  ├─ convert_script.py                 # 单独执行日报 doc -> docx 转换
│  ├─ monthly_generator.py              # 月报核心逻辑
│  └─ run_all.py                        # 月报一键命令行入口
├─ gui_app.py                           # 统一桌面窗口入口
├─ run_gui.bat                          # Windows 启动脚本
├─ requirements.txt
└─ README.md
```

## 环境要求

- Python 3.10 或更高版本。
- Windows 推荐安装 Microsoft Word。
- 如不使用 Microsoft Word，可安装 LibreOffice，并确保 `soffice` 已加入系统 `PATH`。

主要依赖：

- `python-docx`
- `docxtpl`
- `openpyxl`
- `xlrd`
- `Pillow`
- `lxml`
- `pywin32`，仅 Windows 下用于调用 Microsoft Word/Excel 转换。

## 常见问题

### 日期解析很快，但生成日报慢

这是正常的。日期解析只读取文字和日期；生成日报时才会转换 `.xls`、抽取图片并写入 Word。图片越多、台账越大，生成耗时越长。

### 某天没有清洁站或中转站数据

程序会自动跳过该类型，不会报错。若某个日期两类都没有数据，才会提示该日期没有可生成的日报。

### 提示转换失败

请确认：

- Windows 已安装 Microsoft Word/Excel，或已安装 LibreOffice。
- 如果使用 LibreOffice，`soffice` 可以在命令行中直接运行。
- 原始 `.doc`、`.xls` 文件没有被 Word、Excel 或其他程序占用。

### 提示无法写入目标文件

通常是因为生成的 Word 文件正在被 Word 打开。关闭对应文档后重新运行即可。

## 开发说明

日报核心流程：

1. `daily_report_generator/services/ledger_reader.py` 读取台账、快速解析日期、转换 `.xls`、抽取图片。
2. `daily_report_generator/services/aggregator.py` 按日期、类型、街道和站点聚合。
3. `daily_report_generator/services/renderer.py` 使用 docxtpl 渲染 Word 日报。

月报核心流程：

1. `scripts/monthly_generator.py` 转换、解析和汇总日报。
2. `scripts/run_all.py` 提供命令行一键入口。

统一桌面窗口入口：

```text
gui_app.py
```
