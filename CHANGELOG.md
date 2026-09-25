# 变更记录

本项目遵循 [语义化版本](https://semver.org/lang/zh-CN/)，格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)。

## [未发布]

### 变更（内部重构，无行为差异）

- **渲染层抽出「输出写法」注册表**：`sqlgen.py` 的 `_render()` 不再于中段用
  `if opts.fmt == 'insert'` 分支，改为 `RENDERERS = {'union': …, 'insert': …}` 分派；
  两种写法共用的头注释收进 `_header_comments()`。结构与 `dialects.DIALECTS` 同构 ——
  加一种写法 = 加一个函数 + 注册一行，不必再进 `_render` 中段改分支。
  模块导入时校验 `RENDERERS` 与 `FORMATS` 不漂移。
- **已验证零行为差异**：186 条单测通过；60 条端到端用例的日志与产物**逐字节一致**；
  另有 192 条（4 输入 × 4 方言 × 3 输出形式 × 4 选项集）新旧实现对撞，
  `rc` / `stdout` / `stderr` / 产物字节全部一致。
- **明确边界**：`--format` 只表示**同一种 SQL 文本的不同写法**。换**输出载体**
  （json/yaml）或换**产品形态**（ddl/orm）属于新能力，其渲染器入参协议不同
  （`ddl`/`orm` 还需要列类型与约束的来源），不并入 `RENDERERS`、不污染 `--format` 语义。
  等出现第三种写法时再拆出 `renderers/`，且届时按**载体**建目录而非按写法平铺。

## [0.3.2] - 2026-09-25

一次跨 SQL Server / MySQL / PostgreSQL 的端到端实测（51,290 行大表 + 多组边界夹具）暴露的问题。
**SQL 渲染与转义全部通过**（标识符转义、单引号/反斜杠/emoji、GBK、多行文本、NULL 与空串语义、
`insert` 分批、大表流式写入内存有界），问题集中在**表头/列取舍**与**类型推断**两处。
修复后用同一套用例在三库**复验通过**（load / empty / insert / read 全部 rc=0）。

### 修复

- **表头行被误判 → 静默丢一行数据**（退出码仍为 0）。表头里只要出现一个**空列名**，
  该行的 `filled` 项就掉 `0.5×0.5=0.25` 分，而"位置靠前"的偏好只有 `0.05/行`，
  数据行因此反超成为"表头"。新增两道保护：
  1. 候选行与第 1 行**分差 < 0.3** 时保守选第 1 行；
  2. 被跳过的行只要**不止一个非空格**就不跳 —— 标题/说明行总是稀疏的，被误判的真表头是填满的。

  实测：表头 `装运方式 | 2024` 得分 1.75 vs 数据行 3.95（分差 2.2），单靠阈值拦不住，靠第 2 条拦下。
- **表头为空白的尾列被整列裁掉**（连该列数据一起，且不告警）。原实现只看表头是否为空，
  于是"表头留空但下面有数据"的列被无声删除 —— 两种不同输入的产物曾**逐字节相同**。
  现在只有「表头为空 **且** 该列数据全空」才裁，否则保留并由 `col_N` 补名。
- **MySQL 下多行文本列被推成 `varbinary`**：MySQL 里裸 `CHAR(10)` 返回的是**二进制串**，
  `CONCAT()` 一旦遇到二进制参数结果也是二进制（列字符集为 NULL、排序按字节走）。
  改为 `CHAR(10 USING utf8mb4)`。SQL Server / Oracle / PostgreSQL 的写法实测都正常，无需改动。

### 新增

- **`--strict-header`**：要求表头必须在第 1 行。自动识别需要跳过行时直接以退出码 `2` 报错，
  便于脚本 / CI fail fast，而不是默默接受一个可能丢数据的判断。
- **`--infer-types` / `--no-infer-types`**（配置项 `infer_types`，默认关）：
  CSV 没有类型信息，此前**所有列**都会被字符串化，`union` 产物随之彻底丢类型
  （`SUM()` 在 SS 报 `Msg 8117`、PG 报 `function sum(text) does not exist`）。
  打开后，**整列都是数字文本**的列按数字输出；带前导零（`'007'`）、`NaN`/`inf`
  或混有非数字的列一律保持文本 —— 不做逐格猜测，默认行为与之前完全一致。

### 行为变化（升级前请注意）

- **表头里带空列名的输入，现在会"响亮报错"，不再"能跑出结果"**。
  修复前，这类文件会被选错表头行然后**成功**输出（少一行数据，退出码 `0`）；
  修复后正确选中第 1 行，随即被表头规范校验拦下，返回退出码 `2`
  （提示补列名，或加 `--force` 继续）。加 `--force` 后按第 1 行正常切分、空列名补 `col_N`。
  也就是说：**原来"能出结果"的输入现在可能改为报错** —— 这是有意的，因为原来的结果是错的。

### 文档

- README 补三条此前未记载的限制：CSV 无类型信息、`union` 产物反推表结构会丢类型、
  MySQL 执行端默认 `latin1`（中文/emoji 按字节落库）；并补自动识别表头的两道保护说明。
- 产物头部的「按字符串输出」注释改为**分档说明原因**（用户指定 `--all-string` /
  同列混类型 / 源数据是文本）。此前纯数字 CSV 会被误标成"混类型列统一为字符串"。
- 新增 [docs/testing/e2e-testing.md](docs/testing/e2e-testing.md)：实测方法、
  7 个问题的**可粘贴最小复现**、复现要避的坑（`sqlcmd` 的 `GO`、`psql` 的 NULL 显示、
  MySQL 连接字符集等）。

### 测试

- 新增 `tests/test_regressions.py`：24 条回归用例，按实测问题编号组织。
  断言全部落在**行数 / 列数 / 列名**上 —— 这些都是"退出码为 0 的静默丢数据"，只看退出码抓不住。

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
