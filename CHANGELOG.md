# 变更记录

本项目遵循 [语义化版本](https://semver.org/lang/zh-CN/)，格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。

## [0.3.1] - 2026-09-25

首次公开推送后 CI 抓到的 Windows 专属崩溃，修掉。

### 修复

- **Windows 下中文提示崩溃**（`UnicodeEncodeError: 'charmap' codec ...`）：
  stdout 不是终端时（CI、`> out.log`、`| more`）Python 用的是 ANSI 码页，
  en-US 环境即 `cp1252`，打印「`[OK] 30 行 x 24 列`」这类中文直接抛异常并返回 1，
  而 SQL 文件其实已经正确写出——**只在最后一行汇报上翻车**。
  现新增 `cli.setup_console()`：管道/重定向场景切到 UTF-8，交互终端保留原编码
  只放宽错误处理，任何情况下都不再崩。
  （Linux/macOS 默认 UTF-8，所以此前只有 Windows 两个 job 红。）
- 补两条回归测试：在 `cp1252` 的 stdout/stderr 下跑正常流程与 `--help`。

## [0.3.0] - 2026-09-25

按实际交互体验反馈重做交互层，并把常用选择固化成配置文件。

### 新增

- **配置文件 `excel2sql.ini`**：放在数据目录（项目级）或 `~/.excel2sql/config.ini`（用户级）即可固化
  方言、输出格式、表名、输出目录、编码、空串策略、剪贴板等设置，交互时不再逐项询问；
  命令行参数优先级最高。`--config PATH` 指定路径，`--init-config` 生成带注释的模板，
  仓库内附 `excel2sql.ini.example`
- **表头行自动识别**（`header_row = auto`）：按"文本多、互不重复、下方紧跟数据"打分选出表头行，
  交互时打印带行号的数据预览并允许输入行号修正
- **输出目录配置化**：`output_dir` 支持 `source`（同目录）/ `sub`（默认，源目录下 `excel2sql-out/`）/
  任意路径；`overwrite = false` 时同名文件自动追加 `-2`、`-3` 后缀防冲突
- **交互简化**：目录里只有一个文件时回车即用（不再要求输入编号）；工作簿只有一个 sheet 时自动使用；
  方言改为 `1~4` 编号菜单；输出形式三个选项分行显示
- 布尔开关新增反向参数 `--no-empty-as-null`、`--no-all-string`、`--no-copy-clipboard`，
  便于覆盖配置文件里的设置

### 修复

- **`UNION ALL` 未独立成行**：此前第一条 `SELECT` 与 `UNION ALL` 粘连，
  Oracle 下会生成 `... FROM dualUNION ALL` 直接语法错误；现固定为 `...\nUNION ALL\n...`
- **`excel2sql.bat` 报 `'rc' is not recognized`**：批处理里混入了弯引号等非 ASCII 字符，
  在 `chcp 65001` 下会被 cmd 解析错位、执行到下一行的片段；现改为**纯 ASCII**，
  中文提示全部由 Python 打印
- `excel2sql.bat` 的 python 探测链：`python-path.txt` → `EXCEL2SQL_PYTHON` → 仓库 `.venv` →
  常见安装目录 → PATH（跳过 Microsoft Store 占位程序）→ `py` 启动器，并校验版本 ≥ 3.9

### 变更

- 非交互模式未指定 `-o` 时同样遵守配置的 `output_dir` 与防冲突策略；显式 `-o` 则按原样写入（可覆盖），
  交互模式与此保持一致
- 交互模式补上与非交互一致的**行数保护**：超过 200000 行先确认再生成，超过 2000 行的 `union`
  会提示改用 `insert`
- 交互结束不再询问"是否复制到剪贴板"，改由配置 `copy_clipboard` 决定

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
  → `PATH`（跳过 Microsoft Store 占位程序）→ `py -3` → 常见安装目录（`%LOCALAPPDATA%\Programs\Python\Python3*`、`C:\Python3*`），
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
