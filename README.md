# Monthly-Word

西城区检查日报月报生成工具。项目用于将每日 Word 检查日报汇总为月度检查通报，支持原始 `.doc` 日报自动转换为 `.docx`，并从转换后的日报中提取检查问题、图片和统计表，生成最终月报 Word 文档。

## 功能简介

- 批量将原始 `.doc` 日报转换为 `.docx`
- 按月汇总日报内容并生成月度检查通报
- 支持两类日报：
  - `清洁站`
  - `中转站`
- 自动复用已转换且未过期的 `.docx` 文件
- 支持图形界面和命令行两种运行方式
- 默认统计周期为上月 20 日至本月 19 日

## 项目结构

```text
Monthly-Word/
├─ 01_原始日报/          # 放置原始 .doc 日报
│  ├─ 清洁站/
│  └─ 中转站/
├─ 02_转换后日报/        # 自动生成或复用的 .docx 日报
│  ├─ 清洁站/
│  └─ 中转站/
├─ 03_月报模板/          # 预留模板目录
├─ 04_输出月报/          # 生成的月报文件
├─ scripts/
│  ├─ convert_script.py  # 单独执行日报转换
│  ├─ monthly_generator.py
│  └─ run_all.py         # 转换并生成月报的一键命令行入口
├─ gui_app.py            # 图形界面入口
├─ run_gui.bat           # Windows 双击启动脚本
├─ requirements.txt
└─ README.md
```

## 环境要求

- Python 3.10 或更高版本
- Windows 推荐安装 Microsoft Word
- 如不使用 Microsoft Word，可安装 LibreOffice，并确保 `soffice` 已加入系统 `PATH`

安装依赖：

```powershell
pip install -r requirements.txt
```

依赖包括：

- `python-docx`
- `Pillow`
- `lxml`
- `pywin32`，仅 Windows 下用于调用 Microsoft Word 转换 `.doc`

## 快速开始

### 1. 准备日报文件

将原始 `.doc` 日报放入对应目录：

```text
01_原始日报/清洁站/
01_原始日报/中转站/
```

文件名需要包含日期，例如：

```text
4月20日密闭式清洁站检查情况.doc
5月9日中转站检查情况.doc
```

### 2. 使用图形界面生成

Windows 下可直接双击：

```text
run_gui.bat
```

启动后填写年份、月份、日报类型，选择原始日报目录和输出目录，然后点击“生成月报”。

### 3. 使用命令行生成

生成 2026 年 5 月清洁站月报：

```powershell
python scripts\run_all.py --year 2026 --month 5 --station-type 清洁站
```

生成 2026 年 5 月中转站月报：

```powershell
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

## 单独转换日报

只执行 `.doc` 到 `.docx` 的批量转换：

```powershell
python scripts\convert_script.py
```

指定输入和输出目录：

```powershell
python scripts\convert_script.py --input 01_原始日报 --output 02_转换后日报
```

## 输出结果

默认输出到：

```text
04_输出月报/
```

输出文件名示例：

```text
西城区2026年5月密闭式清洁站检查通报.docx
西城区2026年5月中转站检查通报.docx
```

## 常见问题

### 提示转换失败

请确认以下事项：

- Windows 已安装 Microsoft Word，或已安装 LibreOffice
- 如果使用 LibreOffice，`soffice` 可以在命令行中直接运行
- 原始 `.doc` 文件没有被 Word 或其他程序占用

### 提示无法写入目标文件

通常是因为生成的月报文件正在被 Word 打开。关闭对应文档后重新运行即可。

### 没有找到日报记录

请检查：

- 日报是否放在正确的类型目录下，例如 `清洁站` 或 `中转站`
- 文件名是否包含类似 `4月20日`、`5月9日` 的日期
- 生成月份是否正确。工具默认汇总周期为上月 20 日至本月 19 日

## 开发说明

核心流程在 `scripts/monthly_generator.py`：

1. 检查原始 `.doc` 日报是否需要转换
2. 将缺失或过期的 `.doc` 转为 `.docx`
3. 解析指定周期内的日报
4. 汇总问题分类、街道情况和图片
5. 生成月度检查通报 `.docx`

图形界面入口为 `gui_app.py`，命令行一键入口为 `scripts/run_all.py`。
