# 变更记录

本项目遵循 [语义化版本](https://semver.org/lang/zh-CN/)，格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。

## [0.2.0] - 2026-09-25

从"单文件小脚本"升级为可维护的内部工具。本版修复项来自代码评审记录
[docs/reviews/review-01-代码评审.md](docs/reviews/review-01-代码评审.md)。

### 新增

- **工程化**：单文件脚本拆成 `src/excel2sql/` 包（cli / reader / headers / sqlgen / dialects / clipboard），
  新增 `pyproject.toml`、单元测试、GitHub Actions CI、`.editorconfig`
- **命令行入口 `excel2sql`**：0.1.0 的 `xlsx2sql` 作为兼容别名同时保留
- **`--format insert`**：输出 `INSERT INTO ... VALUES`，配合 `--batch-size` 分批，大表更实用
- **`--empty-as-null` / `--all-string`**：显式控制空字符串与列类型策略
- **`--delimiter` / `--input-encoding`**：CSV 分隔符与编码可控（默认自动探测）
- **`--encoding utf-8-sig`**：解决老版本 SSMS 打开 UTF-8 无 BOM 文件中文乱码
- **`--force`**：非交互模式下允许自动修复表头后继续
- **跨平台剪贴板**：Windows `Set-Clipboard` / macOS `pbcopy` / Linux `wl-copy`·`xclip`·`xsel`
- **MIT License**（`LICENSE`），`pyproject.toml` 同步声明 license 与 classifier
- **`excel2sql.bat` 自动探测 python**：依次尝试 `python-path.txt` → `EXCEL2SQL_PYTHON` → 仓库 `.venv`
  → `PATH`（跳过 Microsoft Store 占位程序）→ `py -3` → 常见安装目录（含 `C:\Python3*`、`C:\Python3*`），
  并检查版本 ≥ 3.9；全部失败时打印三种配置办法

### 安全

- **列名/表名标识符转义**：SQL Server `]`→`]]`、MySQL `` ` ``→` `` ` ``、标准方言 `"`→`""`。
  修复前，形如 `a] FROM x; DROP TABLE y; --` 的表头会破坏 SQL 结构
- **MySQL 反斜杠转义**：字符串中的 `\` 转义为 `\\`，避免默认 `sql_mode` 下内容被改写
- 剪贴板临时文件改为 `finally` 清理，异常路径不再残留

### 修复

- **Oracle 日期格式**：`date` 与 `datetime` 分别使用 `YYYY-MM-DD` 与 `YYYY-MM-DD HH24:MI:SS`，
  修复纯日期值报 `ORA-01840: input value not long enough for date format`
- **`--header-row 0` 未校验**：修复前会取 `rows[-1]`（最后一行）当表头，现在直接报错
- **非交互模式不再调用 `input()`**：多 sheet 未指定 `-s`、表头不规范、未指定 `-o` 时都改为明确报错/走默认值，
  便于 CI 与批处理（退出码 `2` 表示表头不规范）
- **空值语义统一**：`.xls` 经 pandas 读到的 `NaN`、`.xlsx` 的空单元格统一为 `None` → `NULL`
- **类型推断收紧**：任何字符串（含空串，除非已按 NULL 处理）都参与列级类型统一，避免同列 `1` 与 `''` 混用
- **重复列名修复冲突**：`a, a, a_2` 这类输入不再修出新的重复列名
- **`bool` 误判**：表头体检不再把 `bool` 当数字（`bool` 是 `int` 子类）
- **大表策略**：SQL 改为流式写入文件；`union` 超过 2000 行给出改用 `insert` 的提示；超过 200000 行直接拒绝
- **依赖缺失提示**：缺少 `openpyxl` / `pandas` 时给出可直接执行的安装命令

### 变更

- 输出 SQL 头部注释改为 `-- 由 excel2sql 生成`，并提示内联数据不适用于生产批量导入
- 输出文件默认名保持 `<文件名>_<sheet名>_hardcode.sql`

### 移除

- 0.1.0 的 `建议2.txt`（0 字节空文件）

## [0.1.0] - 2026-09-08

- 初版单文件脚本 `xlsx2sql.py`：扫描目录、编号选文件与 sheet、表头校验、四方言、
  `UNION ALL` / CTE 两种输出、剪贴板复制。归档于 `docs/legacy/xlsx2sql-0.1.0.py`
