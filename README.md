# excel2sql

[![CI](https://github.com/delux2557/excel2sql/actions/workflows/ci.yml/badge.svg)](https://github.com/delux2557/excel2sql/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

把 Excel / CSV 的每一行转成**可直接执行的硬编码 SQL**：既可以是 `UNION ALL` 内联表，也可以是 `INSERT INTO ... VALUES`。

适合的场景：临时把几十到几千行数据贴进数据库做**测试、修数、造 mock**。不适合生产批量导入——那种场景请用数据库原生的 `BULK INSERT` / `LOAD DATA` / `COPY`（见[大数据量建议](#大数据量建议)）。

## 特性

- **交互式向导**：自动扫描当前目录 → 选文件 → 选 sheet → 表头确认 → 生成。只有一个文件/一个 sheet 时自动跳过，
  表头行自动识别并打印预览，全程基本只需按回车
- **配置文件**（ssh config 风格）：方言、输出格式、输出目录、编码、空串策略、剪贴板等写进 `excel2sql.ini`，
  交互时不再逐项询问；命令行参数可临时覆盖
- **完全非交互模式**：给定文件参数后不再有任何 `input()`，可安全用于批处理与 CI
- **多方言**：SQL Server / MySQL / Oracle / PostgreSQL / SQLite，标识符引用、字符串转义、日期包装、换行拼接各不相同
- **表头体检**：识别空列名、重复列名、"表头其实是数据行"，可自动修复（`col_N` / `_2` 后缀）
- **列级类型统一**：同列混了数字和字符串时整列字符串化，避免 `UNION ALL` 触发隐式转换报错或算术溢出
- **转义安全**：列名按方言转义内部引号，字符串转义单引号（MySQL 额外处理反斜杠），避免列名/单元格内容破坏 SQL 结构
- **大表友好**：流式写文件，不把整份 SQL 拼在内存里；超阈值提醒改用 `INSERT` 分批
- **跨平台剪贴板**：生成后可直接复制到剪贴板（Windows / macOS / Linux）

## 安装

```bash
# 源码目录内开发式安装（推荐）
python -m pip install -e .

# 需要读老式 .xls 时
python -m pip install -e ".[xls]"
```

不安装也可以直接跑：

```bash
# 标准 Python：把 src 加进模块搜索路径
PYTHONPATH=src python -m excel2sql --help             # bash / macOS / Linux
set "PYTHONPATH=src" && python -m excel2sql --help    # Windows cmd

# embed 版 Python（带 ._pth，会忽略 PYTHONPATH，需用 runpy 引导）
python -c "import sys,runpy; sys.path.insert(0, r'<仓库路径>\src'); runpy.run_module('excel2sql', run_name='__main__')" --help
```

### Windows 双击运行

双击 `excel2sql.bat` 即进入交互向导（**免安装**，直接用仓库里的 `src/`）。它会按下面的顺序自动找 python：

1. 仓库根目录的 `python-path.txt`（一行，写 python.exe 完整路径）
2. 环境变量 `EXCEL2SQL_PYTHON`
3. 仓库内的 `.venv\Scripts\python.exe`
4. `PATH` 里的 `python`（自动跳过 Microsoft Store 的占位程序）
5. `py -3` 启动器
6. 常见安装位置：`%LOCALAPPDATA%\Programs\Python\Python310`~`313`、`C:\Python311`~`C:\Python313`

如果都找不到（或版本低于 3.9），窗口里会给出具体解决办法。**最省事的办法**是在仓库根建一个 `python-path.txt`：

```text
C:\Python312\python.exe
```

> 注意：用记事本保存时编码选 **ANSI**（UTF-8 带 BOM 会让首字符变成乱码，导致路径失效）。`python-path.txt` 属于本机配置，已在 `.gitignore` 里排除。
>
> 另外，`excel2sql.bat` 扫描的是**当前目录**——双击时即 bat 所在目录；若那里没有数据文件，向导会提示你输入目标文件夹，也可以直接把 Excel 文件拖到 bat 图标上。

## 快速开始

### 交互式

```bash
excel2sql
```

```
==========================================================
  excel2sql 0.4.0  |  Excel / CSV  ->  硬编码 SQL
==========================================================
配置：使用内置默认值（sqlserver / union / cte / 表名 HARDCODE）
提示：在数据目录放一个 excel2sql.ini 就能固化方言/格式/输出目录等设置。

发现 1 个文件（目录：D:\data）
  superstore-sample.csv   (8 KB)  （回车直接使用）
回车继续，或粘贴其它路径，q 退出:

读取：D:\data\superstore-sample.csv
  唯一 sheet：superstore-sample  (31 行 x 24 列)，自动使用
  读取数据中…（几万行的 Excel 可能要十几秒）
  表头识别：第 1 行最像表头（文本列名 + 下方是数据）

表头预览（← 标记的是当前选中的表头行）：
     1 | 行 ID | 订单 ID | 订购日期 | 装运日期 | 装运方式 | 客户 ID | 客户名称 | …(+17 列)   ← 表头
     2 | 40098 | CA-2014-AB1001… | 2024-11-11 00:… | 2024-11-13 00:… | 一级 | AB-100151402 | Aaron Bergman | …(+17 列)
     3 | 26341 | IN-2014-JR1621… | 2024-02-05 00:… | 2024-02-07 00:… | 二级 | JR-162107 | Justin Ritter | …(+17 列)
     … | …（共 31 行）
回车确认，或输入其它行号修正 [1]:
  表头校验通过：24 列 / 30 行数据

[OK] 30 行 x 24 列 -> D:\data\excel2sql-out\superstore-sample_superstore_sample_hardcode.sql   (24 KB)
预览：
  SELECT N'40098' AS [行 ID], N'CA-2014-AB10015140-41954' AS [订单 ID], N'2024-11-11 00:00:00' AS [订购日期], ...
```

**能自动推断的都不问**：目录里只有一个文件时回车即用、工作簿只有一个 sheet 时自动选择、表头行自动识别并给预览。
方言、输出格式、表名、编码、输出目录这些**由配置文件决定**（见[配置文件](#配置文件)），
需要临时改就用命令行参数（例如 `excel2sql -d mysql`）。

仓库自带一份可以直接拿来跑的示例数据（就是上面这份，来自公开的零售订单样例集）：
[`examples/superstore-sample.csv`](examples/superstore-sample.csv)，30 行 × 24 列。

### 非交互（脚本 / CI）

```bash
# SQL Server，CTE 内联表（examples/ 里就有这份数据）
excel2sql examples/superstore-sample.csv -o out.sql

# MySQL，改成 INSERT 分批，空串按 NULL
excel2sql examples/superstore-sample.csv -d mysql --format insert --empty-as-null -o out.sql

# Oracle，只要纯 UNION ALL 块（方便嵌进已有 SQL）
excel2sql examples/superstore-sample.csv -d oracle --wrap plain -o out.sql

# 多 sheet 的 Excel 要指定工作表
excel2sql 订单表.xlsx -s 明细 -o out.sql

# CSV，表头在第 2 行，分号分隔
excel2sql data.csv --header-row 2 --delimiter ";" -o out.sql
```

## 命令行参数

不带参数时默认值来自[配置文件](#配置文件)；**命令行参数优先级最高**，可临时覆盖。

| 参数 | 说明 |
|---|---|
| `FILE` | 源文件；**省略则进入交互向导** |
| `-s, --sheet NAME` | 工作表名（多 sheet 时必填，否则报错并列出可选值） |
| `-o, --out PATH` | 输出路径；**显式指定则按原样写入**（可覆盖），不指定则按配置的 `output_dir` 生成并防冲突 |
| `-d, --dialect NAME` | `sqlserver` / `mysql` / `oracle` / `postgresql` / `sqlite`，也可用 `1`~`5`（也接受 `mssql`/`mariadb`/`pg`/`ora`/`sqlite3` 等别名）。**无法识别即报错退出，不会静默回退成别的方言** |
| `--header-row N` | 表头行号；`auto`（默认）表示自动识别 |
| `--format union\|insert` | 输出格式 |
| `--wrap cte\|plain` | `union` 模式下是否用 `WITH ... AS (...)` 包裹 |
| `--table NAME` | 内联表名 |
| `--empty-as-null` / `--no-empty-as-null` | 空字符串是否按 `NULL` 输出 |
| `--all-string` / `--no-all-string` | 是否所有列强制按字符串输出 |
| `--infer-types` / `--no-infer-types` | 强制 / 关闭数字类型推断（默认 `auto`：**只对 CSV 生效**，见[类型与空值规则](#类型与空值规则)） |
| `--copy-clipboard` / `--no-copy-clipboard` | 生成后是否复制到剪贴板 |
| `--batch-size N` | `insert` 模式每批行数 |
| `--encoding ENC` | 输出文件编码，如 `utf-8-sig`（老版 SSMS 中文乱码时用） |
| `--input-encoding ENC` | CSV 输入编码，默认自动尝试 `utf-8-sig → gb18030 → utf-8 → gbk` |
| `--delimiter CHAR` | CSV 分隔符，默认在 `, ; \t \|` 中自动选切分最细的 |
| `--force` | 非交互模式下，表头不规范也自动修复并继续 |
| `--strict-header` | 要求表头必须在第 1 行；自动识别若需跳过行则以退出码 `2` 报错（供脚本 / CI fail fast） |
| `--config PATH` | 指定配置文件 |
| `--init-config` | 在当前目录生成 `excel2sql.ini` 模板后退出 |
| `-V, --version` | 版本号 |

**退出码**：`0` 成功；`1` 一般错误（读不到文件、sheet 不存在、参数非法）；`2` 表头不规范且未加 `--force`。

## 配置文件

把"不用每次选"的东西固化下来，写法和 ssh config 一样是纯文本键值：

```bash
excel2sql --init-config          # 生成带注释的模板 excel2sql.ini
```

查找顺序（找到即用）：`--config` 指定 → `./excel2sql.ini`（数据目录，项目级）→ `~/.excel2sql/config.ini`（用户级）→ 内置默认值。

```ini
[output]
dialect = mysql                 # sqlserver | mysql | oracle | postgresql | sqlite
format = union                  # union | insert
wrap = cte                      # cte | plain
table = HARDCODE
batch_size = 500
output_dir = sub                # source=同目录 | sub=源目录下建子目录 | 任意路径
output_subdir = excel2sql-out
overwrite = false               # false = 同名自动加 -2/-3 后缀
encoding = utf-8                # utf-8 | utf-8-sig
filename = {name}_{sheet}_hardcode.sql   # 可用 {name} {sheet} {date}

[data]
empty_as_null = false
all_string = false
header_row = auto               # auto = 自动识别，或写 1 / 2 / ...

[ui]
ask_advanced = false            # true = 交互时逐项询问方言/格式/编码/剪贴板
copy_clipboard = false
```

- 配置写错不会中断：无法解析的项会打印一条 `[配置] ...` 警告，并回退到默认值
- 键名多余或未知会被忽略，方便把同一份 ini 在不同版本间复用
- `ask_advanced = true` 时才逐项询问，适合想每次手动确认的场景
- 完整示例见 [`excel2sql.ini.example`](excel2sql.ini.example)（**注意它是模板**：程序只认 `excel2sql.ini`，
  需要复制或改名才生效；`excel2sql.ini` 属本机配置，已在 `.gitignore` 里排除）

## 输出格式

`union` + `cte`（默认）：

```sql
WITH [HARDCODE] AS (
    SELECT N'一级' AS [装运方式], N'消费者' AS [细分市场]
    UNION ALL
    SELECT N'二级' AS [装运方式], N'公司' AS [细分市场]
)
SELECT * FROM [HARDCODE];
```

CTE 体内的 `SELECT` / `UNION ALL` 缩进一级、闭合括号回到行首；`--wrap plain` 则**不缩进**，
因为那是给「嵌进已有 SQL」用的，缩进交给调用方按所在层级对齐。

`insert`：

```sql
INSERT INTO [HARDCODE] ([装运方式], [细分市场]) VALUES
('一级', '消费者'),
('二级', '公司');
```

多行文本单元格会被拼成单行表达式，避免一条 `SELECT` 被换行截断：

```sql
SELECT N'上海市' + CHAR(10) + N'浦东新区' AS [收货地址]   -- SQL Server
SELECT CONCAT('上海市', CHAR(10), '浦东新区') AS `收货地址` -- MySQL
SELECT '上海市' || CHR(10) || '浦东新区' AS "收货地址"     -- Oracle / PostgreSQL
SELECT '上海市' || CHAR(10) || '浦东新区' AS "收货地址"    -- SQLite
```

## 类型与空值规则

| 情况 | 输出 |
|---|---|
| 空单元格（Excel 的 `None`、pandas 的 `NaN`） | `NULL` |
| 空字符串 | `''`，加 `--empty-as-null` 后变 `NULL` |
| 数字 | 原样，如 `2`、`221.98`；`NaN`/`inf` → `NULL` |
| 日期 | `'2024-11-11'`（Oracle 为 `TO_DATE(...,'YYYY-MM-DD')`；SQLite 同为纯文本，它没有日期类型） |
| 日期时间 | `'2024-11-11 00:00:00'`（Oracle 为 `TO_DATE(...,'YYYY-MM-DD HH24:MI:SS')`） |
| 布尔 | `'1'` / `'0'` |
| **同列混有数字和字符串** | 整列按字符串输出（避免 `UNION ALL` 类型冲突） |
| 加了 `--all-string` | 所有非空值都按字符串输出（此时不做数字推断） |
| **CSV 输入（默认 `auto`）** | CSV 没有类型信息，读出来全是文本。默认把**整列都能无损解析成数字**的列还原成数字，其余列保持文本 |
| `infer_types = off` / `--no-infer-types` | CSV 的列一律按字符串输出（0.3.x 的旧行为，仍可完整复现） |
| `infer_types = on` / `--infer-types` | 强制推断，**Excel 源也照做**（默认 `auto` 会放过 Excel 源） |
| 推断时**刻意不转**的列 | 带前导零的整数（`'007'` 是编号不是数量）、有效数字超过 **15 位**的、含 `NaN`/`inf` 的、混有非数字的 |

> **为什么 `auto` 只对 CSV 生效**：xlsx/xls 的单元格自带类型 —— 写成文本就是用户有意的文本，
> 工具不该去改写它已经明确表达过的意图。CSV 没有这个信息，才需要靠内容推断。
>
> **为什么卡在 15 位有效数字**：Excel 本身只保证 15 位有效数字。越过这条线的"数字"几乎一定是
> 编号/账号而不是数量，而且 19 位以上会超出 SQL Server / MySQL / PostgreSQL 的 `bigint` 范围 ——
> 硬转成数字字面量会让**灌库直接报错**。宁可留在文本。
>
> ★ **已知边界**：11 位手机号这类能无损装进 `bigint` 的编号**仍会被转成数字**。
> 要保住它们用 `--no-infer-types`，或让编号带前导零（那样会自动保留文本）。

> 为什么必须做列级统一：SQL Server 会按数据类型优先级把 `nvarchar` 隐式转成 `int`，于是 `'t'` 报 `Conversion failed`、`7900454710` 直接算术溢出。
>
> 反过来说，**被字符串化的列在三库里都算不了数**：实测 `SUM()` 在 SQL Server 报 `Msg 8117 Operand data type nvarchar is invalid for sum operator`、PostgreSQL 报 `function sum(text) does not exist`（MySQL 会隐式转换且结果正确）。CSV 输入在 0.4.0 起默认自动还原数字列，正是为了让这条不成为默认陷阱。

## 各方言差异

| | SQL Server | MySQL | Oracle | PostgreSQL | SQLite |
|---|---|---|---|---|---|
| 标识符 | `[col]`，`]`→`]]` | `` `col` ``，`` ` ``→` `` ` `` | `"col"`，`"`→`""` | `"col"`，`"`→`""` | `"col"`，`"`→`""` |
| 字符串前缀 | `N'...'` | `'...'` | `'...'` | `'...'` | `'...'` |
| 单引号转义 | `''` | `''` | `''` | `''` | `''` |
| 反斜杠 | 无特殊含义 | **额外转义为 `\\`** | 无特殊含义 | 无特殊含义 | 无特殊含义 |
| 拼接 | `+` | `CONCAT(...)` | `\|\|` | `\|\|` | `\|\|` |
| 换行 | `CHAR(10)` | `CHAR(10 USING utf8mb4)` | `CHR(10)` | `CHR(10)` | `CHAR(10)` |
| 日期 | 隐式转换 | 隐式转换 | `TO_DATE(...)` | 隐式转换 | 纯 ISO 文本 |
| 无表查询 | 不需要 `FROM` | 不需要 | `FROM dual` | 不需要 | 不需要 |

日期包装只在**方言自己需要**时才加：目前只有 Oracle（裸字符串参与日期比较会失败，必须 `TO_DATE`）。
SQL Server / MySQL / PostgreSQL 由上下文自动转换；SQLite 没有日期类型，ISO-8601 文本就是它的规范表示，
`date()` / `strftime()` 能直接识别（`tests/test_sqlite.py` 里有实测守着这条前提）。

## 落地到数据库时的注意

产物是**纯 `SELECT` / `INSERT` 文本**，本身不带类型声明：

- **`union` 产物是裸 `SELECT`**。用 `SELECT * INTO`（SQL Server）/ `CREATE TABLE AS`（PG、MySQL、SQLite）
  让数据库**从字面量推类型**时，日期列只会落成字符串类型（SS `nvarchar`、PG `text`、MySQL `varchar`），
  数字列也取决于字面量。**正式建表请自己写 `CREATE TABLE`**，再用 `insert` 产物灌数据 ——
  实测这样都能无损接受。
- **MySQL 执行端要显式指定字符集**。`docker exec` 或任何非交互会话里，MySQL 客户端默认用 `latin1`，
  中文和 emoji 会**按字节**落进 `latin1` 列（`CHAR_LENGTH('📦🚚✅')` 返回 11 而不是 3）。
  执行产物前先 `SET NAMES utf8mb4;`，或给客户端加 `--default-character-set=utf8mb4`。
- **SQLite：列的「亲和性」会改写你存进去的值**。SQLite 是动态类型，但若目标列声明了数值亲和性，
  `'0001'` 会被**静默存成整数 `1`**（实测：`CREATE TABLE t (a NUMERIC)` 之后 `INSERT ... ('0001')`
  读回是 `1`；声明 `TEXT` 则原样保留）。所以**编号类列务必声明 `TEXT`**：

  ```sql
  CREATE TABLE orders ("订单编号" TEXT, "数量" INTEGER, "金额" REAL, "订购日期" TEXT);
  ```

  日期没有原生类型，ISO-8601 文本就是规范表示，`date()` / `strftime()` 直接可用：
  `SELECT strftime('%Y', "订购日期") FROM orders;`

## 表头体检

生成前会检查表头，命中任意一条即提示（交互模式可确认后继续，非交互模式需加 `--force`）：

- 表头有空列 → 自动命名为 `col_2`、`col_5`……
- 列名重复（大小写不敏感）→ 自动加 `_2`、`_3` 后缀，且保证结果全局唯一
- 表头行**全是数字或日期** → 高度怀疑表头不在这一行，请改 `--header-row N`
- 表头下方**没有任何数据行** → 直接判定无法转换

**自动识别表头时的两道保护**（防止把数据行当表头、静默丢一行）：

- 候选行与第 1 行**分差过小**（< 0.3）时保守选第 1 行。表头里只要有一个空列名，该行得分就会掉
  `0.5×0.5=0.25`，而"位置靠前"的偏好只有 `0.05/行` —— 证据不足时不该翻盘。
- **被跳过的行只要不止一个非空格就不跳**。标题/说明行总是稀疏的（往往只有第一格有内容），
  而"表头被误判成标题"的那种行是填满的。实测表头 `装运方式 | 2024` 得分 1.75 vs 数据行 3.95，
  分差 2.2，单靠阈值拦不住，靠的就是这条。

两条都会在提示里说明原因，并给出 `--header-row N` 的修正建议。脚本 / CI 里想更严格就用 `--strict-header`。

## 安全说明

生成的 SQL 里所有标识符和字符串都按方言做了转义，因此**不可信的列名或单元格内容不会破坏 SQL 结构**。例如 SQL Server 列名 `a] FROM x; DROP TABLE y; --` 会被写成 `[a]] FROM x; DROP TABLE y; --]`，仍是一个合法的列名。

不过仍请注意：输出的 SQL 是**纯文本可信脚本**，其中的数据内容来自你的表格，请勿对来源不明的文件直接在生产库执行。

## 大数据量建议

| 行数 | 建议 |
|---|---|
| < 2000 | `UNION ALL` 完全够用，方言兼容性最好 |
| 2000 ~ 20000 | 用 `--format insert --batch-size 500`，解析开销小很多 |
| > 20000 | 别用本工具导入。改用数据库原生通道：SQL Server `BULK INSERT`/`bcp`、MySQL `LOAD DATA INFILE`、Oracle `SQL*Loader`/外部表、PostgreSQL `COPY`、SQLite `.import` |
| > 200000 | 工具会直接拒绝，避免生成一个谁也打不开的 SQL 文件 |

## 目录结构

```
.
├── src/excel2sql/          # 包源码
│   ├── cli.py              # 命令行入口：交互向导 + 批处理
│   ├── config.py           # excel2sql.ini 配置读写与输出路径防冲突
│   ├── reader.py           # Excel/CSV 读取与编码、分隔符探测
│   ├── headers.py          # 表头识别、校验与修复
│   ├── sqlgen.py           # 字面量渲染与 SQL 生成
│   ├── dialects.py         # 各方言规则
│   └── clipboard.py        # 跨平台剪贴板
├── tests/                  # pytest / unittest 均可跑
│   ├── test_regressions.py # 实测发现的缺陷回归集
│   └── test_sqlite.py      # SQLite：把产物真灌进 sqlite3 再逐格核对（标准库，能进 CI）
├── examples/               # 示例数据（可直接喂给 CLI）
├── docs/
│   ├── reviews/            # 代码评审记录
│   ├── testing/            # 端到端实测：方法、问题清单、最小复现、避坑清单
│   └── legacy/             # 0.1.0 单文件脚本归档
├── excel2sql.bat           # Windows 双击启动（自动探测 python）
├── excel2sql.ini.example   # 配置文件模板（复制成 excel2sql.ini 即生效）
├── LICENSE                 # MIT
├── python-path.txt         # 可选：本机 python 路径（已 gitignore）
└── pyproject.toml
```

## 开发

```bash
python -m pip install -e ".[dev]"
python -m pytest                                    # 或：
python -m unittest discover -s tests -t .            # 必须带 -s / -t：
                                                    # tests/__init__.py 靠它把 src/ 注入模块路径
```

测试不依赖网络，xlsx 用例会在运行时用 openpyxl 现场生成临时文件。
SQLite 用例用的是标准库 `sqlite3`，所以「生成 → 灌库 → 读回」这条端到端链路也会在 CI 里真跑，
不像 SQL Server / Oracle 那样需要外部容器。

## 已知限制

- 不含 xlsb / ods / parquet 支持
- `.xls`、`.xlsm` 依赖可选包（`pandas`+`xlrd` / `openpyxl`）
- 一个 sheet 只处理一张连续表，不支持多块表头或合并单元格的"花式"版式
- 超宽表（几百列）生成的 SQL 单行会很长，部分客户端显示吃力
- 数字类型推断的边界：**能无损装进 64 位整数的编号仍会被转成数字**（如 11 位手机号）。
  要保住这类列用 `--no-infer-types`，或让编号带前导零

## 变更记录

见 [CHANGELOG.md](CHANGELOG.md)。0.2.0 相对 0.1.0 的完整改动与评审来源见 [docs/reviews/review-01-代码评审.md](docs/reviews/review-01-代码评审.md)。

0.3.2 那批问题的来源、**可粘贴的最小复现**与真库验证方法见 [docs/testing/e2e-testing.md](docs/testing/e2e-testing.md)。

## 许可

[MIT License](LICENSE) © 2026 delux2557 —— 可自由使用、修改、分发，保留版权声明即可。
