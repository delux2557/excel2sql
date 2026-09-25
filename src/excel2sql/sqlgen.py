"""SQL 生成：把二维数据渲染成硬编码的内联表。

安全要点：
- 标识符（列名/表名）一律走 Dialect.quote_ident 转义内部引号；
- 字符串按方言转义（MySQL 额外转义反斜杠）；
- 类型按列统一，保证 UNION ALL / 多行 VALUES 的列类型一致。
"""
from __future__ import annotations

import datetime
import math
import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple

from .dialects import Dialect
from .headers import is_blank

FORMATS = ('union', 'insert')
WRAPS = ('cte', 'plain')

# 超过该行数仍用 UNION ALL 会明显拖慢解析，默认建议 insert
UNION_ROW_WARN = 2000

# CTE 体内每行缩进的宽度
CTE_INDENT = '    '

# 纯数字文本（--infer-types 用）
_INT_RE = re.compile(r'[+-]?\d+')
_FLOAT_RE = re.compile(r'[+-]?(\d+\.\d*|\.\d+|\d+)([eE][+-]?\d+)?')

# 有效数字上限。取 15 是因为 Excel 本身只保证 15 位有效数字：
# 越过这条线的"数字"几乎一定是编号/账号而不是数量，而且 19 位以上还会超出
# SQL Server / MySQL / PostgreSQL 的 bigint 范围 —— 硬转成数字字面量会直接报错。
MAX_SIG_DIGITS = 15

# 「按字符串输出」的原因分档 —— 写进产物头注释，让"用户要求的"与"工具被迫的"可区分
REASON_ALL_STRING = '用户指定 --all-string'
REASON_MIXED = '同列混有数字/日期与文本，整列统一为字符串'
REASON_TEXT = '源数据是文本（CSV 无类型信息）'


class SqlGenError(Exception):
    """无法生成 SQL（如没有数据行）。"""


@dataclass
class SqlOptions:
    dialect: Dialect
    table: str = 'HARDCODE'
    fmt: str = 'union'              # union | insert
    wrap: str = 'cte'               # cte | plain（仅 union 生效）
    empty_as_null: bool = False     # 空字符串视为 NULL
    all_string: bool = False        # 所有列强制按字符串输出
    infer_types: bool = False       # 源数据无类型信息（CSV）：整列是数字时按数字输出
                                    # 注：auto/on/off 的判定与「源文件是否自带类型」在 cli 层完成，
                                    # 到这一层已经是确定的是/否
    batch_size: int = 500           # insert 模式每批行数

    def __post_init__(self) -> None:
        if self.fmt not in FORMATS:
            raise SqlGenError('未知输出格式：{}'.format(self.fmt))
        if self.wrap not in WRAPS:
            raise SqlGenError('未知包裹方式：{}'.format(self.wrap))
        if self.batch_size < 1:
            raise SqlGenError('batch_size 必须 >= 1')


# ---------------------------------------------------------------- 字面量
def infer_column_types(header: Sequence[object], rows: Sequence[Sequence[object]],
                       all_string: bool = False, empty_as_null: bool = False) -> List[bool]:
    """逐列判断是否需要强制按字符串输出。

    只要该列出现过「非空字符串」就整列字符串化，避免同列混 int/varchar 导致
    UNION ALL 隐式转换报错或算术溢出。
    """
    ncols = len(header)
    if all_string:
        return [True] * ncols
    force = [False] * ncols
    for row in rows:
        for i in range(min(ncols, len(row))):
            v = row[i]
            if isinstance(v, str):
                if empty_as_null and v.strip() == '':
                    continue                       # 会被当 NULL，不影响列类型
                force[i] = True
    return force


def _sig_digits(text: str) -> int:
    """数字文本的有效位数。

    用 Decimal 解析，省得自己处理前导零与指数记号；解析不了就返回一个
    大于任何阈值的大数，让调用方按「不安全」处理。
    """
    try:
        return len(Decimal(text).as_tuple().digits)
    except (InvalidOperation, ValueError):
        return MAX_SIG_DIGITS + 1


def _as_number(text: str):
    """把纯数字文本转成 int/float；不是数字则返回 None。

    两道刻意的保守规则：
    - **带前导零的整数（'007'）不转** —— 转成 7 会丢信息，这类列本来更可能是编号；
    - **有效位数超过 15 位的不转** —— 见 MAX_SIG_DIGITS 的说明。

    小数不受前导零规则约束：'007.5' 与 '7.5' 是同一个数，前导零只是写法。
    """
    t = text.strip()
    if not t:
        return None
    if _INT_RE.fullmatch(t):
        digits = t.lstrip('+-')
        if len(digits) > 1 and digits[0] == '0':
            return None                            # '007' / '000123' 保持文本
        if _sig_digits(digits) > MAX_SIG_DIGITS:
            return None                            # 多半是编号，且再长会有溢出风险
        return int(t)
    if _FLOAT_RE.fullmatch(t):
        if _sig_digits(t) > MAX_SIG_DIGITS:
            return None                            # 超出双精度可无损表达的范围
        try:
            return float(t)
        except ValueError:
            return None
    return None


def coerce_numeric_columns(rows: Sequence[Sequence[object]]) -> List[List[object]]:
    """把「整列都是数字文本」的列换成真正的数字（--infer-types 的实现）。

    为什么需要：CSV 没有类型信息，读出来一律是 `str`，于是 infer_column_types()
    会把**所有列**都字符串化，`union` 产物随之彻底丢类型（2026-09-25 实测：
    SQL Server 落成 `nvarchar`、PG 落成 `text`、`SUM()` 三库全都报错）。

    保守起见只做**整列可解析**才转换：只要有一格不是数字（含 NaN/inf/前导零）
    就整列保持文本，绝不逐格猜。该列的空白单元格会一并转成 `NULL`
    —— 数字列本来就容不下空字符串。
    """
    out = [list(r) for r in rows]
    ncols = max((len(r) for r in out), default=0)
    for i in range(ncols):
        col = [(r[i] if i < len(r) else None) for r in out]
        idx = [j for j, v in enumerate(col) if not is_blank(v)]
        if not idx or not all(isinstance(col[j], str) for j in idx):
            continue                               # 含非文本值 -> 不动
        nums = [_as_number(col[j]) for j in idx]
        if any(n is None for n in nums):
            continue                               # 有一格不像数字 -> 整列不动
        for j, n in zip(idx, nums):
            out[j][i] = n
        for j, v in enumerate(col):                # 空白 -> NULL
            if is_blank(v):
                out[j][i] = None
    return out


def forced_groups(header: Sequence[object], rows: Sequence[Sequence[object]],
                  force: Sequence[bool], all_string: bool) -> List[Tuple[str, List[str]]]:
    """给「按字符串输出」的列分档，供产物头注释区分原因。

    分档依据是**该列实际观测到的值**，不依赖调用方声明来源：
    - 用户显式 `--all-string`；
    - 列里既有文本又有非文本 -> 同列混类型（工具被迫统一）；
    - 列里的非空值全都是文本 -> 源数据本来就是文本（CSV 的典型特征）。
    """
    names = [str(h) for h, f in zip(header, force) if f]
    if not names:
        return []
    if all_string:
        return [(REASON_ALL_STRING, names)]

    text_like: List[str] = []
    mixed: List[str] = []
    for i, (name, f) in enumerate(zip(header, force)):
        if not f:
            continue
        vals = [(r[i] if i < len(r) else None) for r in rows]
        vals = [v for v in vals
                if not (v is None or (isinstance(v, str) and v.strip() == ''))]
        if vals and all(isinstance(v, str) for v in vals):
            text_like.append(str(name))
        else:
            mixed.append(str(name))

    out: List[Tuple[str, List[str]]] = []
    if mixed:
        out.append((REASON_MIXED, mixed))
    if text_like:
        out.append((REASON_TEXT, text_like))
    return out


def literal(value, force_str: bool, opts: SqlOptions) -> str:
    dia = opts.dialect

    if value is None:
        return 'NULL'
    if isinstance(value, float) and math.isnan(value):
        return 'NULL'
    if isinstance(value, bool):
        return "'1'" if value else "'0'"
    if isinstance(value, str) and opts.empty_as_null and value.strip() == '':
        return 'NULL'

    if isinstance(value, datetime.datetime):
        s = value.strftime('%Y-%m-%d %H:%M:%S')
        return dia.datetime_prefix + dia.quote_string(s) + dia.datetime_suffix
    if isinstance(value, datetime.date):
        s = value.strftime('%Y-%m-%d')
        return "{}{}{}".format(dia.date_prefix, dia.quote_string(s), dia.date_suffix)

    if isinstance(value, (int, float)) and not force_str:
        if isinstance(value, float):
            if math.isinf(value):
                return 'NULL'
            return repr(value)
        return str(value)

    text = str(value).replace('\r\n', '\n').replace('\r', '\n')
    segments = [dia.quote_string(part) for part in text.split('\n')]
    if len(segments) == 1:
        return segments[0]
    return dia.concat_strings(segments, dia.newline_expr)


def render_row(header: Sequence[object], row: Sequence[object], force: Sequence[bool],
               opts: SqlOptions, with_alias: bool = True) -> str:
    """一行数据的字面量列表。with_alias=False 用于 INSERT（列名已在列清单里）。"""
    cells = []
    for i, name in enumerate(header):
        v = row[i] if i < len(row) else None
        cell = literal(v, force[i], opts)
        if with_alias:
            cell += ' AS ' + opts.dialect.quote_ident(name)
        cells.append(cell)
    return ', '.join(cells)


# ---------------------------------------------------------------- 渲染
def prepare(header: Sequence[object], rows: Sequence[Sequence[object]], opts: SqlOptions):
    """把原始行列整理成可直接渲染的 (表头, 有效数据行, 类型标记)。"""
    if not header:
        raise SqlGenError('表头为空，无法生成 SQL')
    ncols = len(header)
    data = [list(r[:ncols]) for r in rows if not all(is_blank(v) for v in list(r)[:ncols])]
    if not data:
        raise SqlGenError('没有有效数据行')
    # --all-string 是「所有列都按字符串输出」，此时再转数字既无意义，
    # 又会把数字列里的空格子变成 NULL —— 与"全部按字符串"自相矛盾。
    if opts.infer_types and not opts.all_string:
        data = coerce_numeric_columns(data)
    force = infer_column_types(header, data, opts.all_string, opts.empty_as_null)
    return list(header), data, force


def iter_sql(header: Sequence[object], rows: Sequence[Sequence[object]],
             opts: SqlOptions) -> Iterator[str]:
    """流式产出 SQL 片段，避免大表一次性拼成巨型字符串。"""
    header, data, force = prepare(header, rows, opts)
    return _render(header, data, force, opts)


def _header_comments(header: Sequence[object], data: Sequence[Sequence[object]],
                     force: Sequence[bool], opts: SqlOptions) -> Iterator[str]:
    """产物头注释。所有渲染器共用。"""
    yield '-- 由 excel2sql 生成：{} 行 x {} 列\n'.format(len(data), len(header))
    yield '-- 方言：{}   输出格式：{}\n'.format(opts.dialect.name, opts.fmt)
    for reason, names in forced_groups(header, data, force, opts.all_string):
        yield '-- 按字符串输出的列（{}）：{}\n'.format(reason, ', '.join(names))
    yield '-- 提示：内联数据仅供测试/修数，请勿直接用于生产批量导入\n'


def _render_union(header: Sequence[object], data: Sequence[Sequence[object]],
                  force: Sequence[bool], opts: SqlOptions) -> Iterator[str]:
    """每行一条 SELECT 再用 UNION ALL 串起来，可选 CTE 包裹。"""
    dia = opts.dialect
    lines = ['SELECT ' + render_row(header, r, force, opts) + dia.dual for r in data]

    if opts.wrap == 'plain':
        # plain 是给「嵌进已有 SQL」用的，不缩进 —— 由调用方按所在层级自行对齐
        yield '\nUNION ALL\n'.join(lines) + '\n'
        return

    # CTE 体内缩进一级，闭合括号回到行首（sqlfluff 等格式化工具的默认风格）
    sep = '\n' + CTE_INDENT + 'UNION ALL\n' + CTE_INDENT
    yield 'WITH {} AS (\n'.format(dia.quote_ident(opts.table))
    yield CTE_INDENT + sep.join(lines) + '\n'
    yield ')\nSELECT * FROM {};\n'.format(dia.quote_ident(opts.table))


def _render_insert(header: Sequence[object], data: Sequence[Sequence[object]],
                   force: Sequence[bool], opts: SqlOptions) -> Iterator[str]:
    """INSERT INTO ... VALUES (...), (...)；超过 batch_size 就另起一条语句。"""
    dia = opts.dialect
    cols = ', '.join(dia.quote_ident(h) for h in header)
    yield 'INSERT INTO {} ({}) VALUES\n'.format(dia.quote_ident(opts.table), cols)
    total = len(data)
    for start in range(0, total, opts.batch_size):
        chunk = data[start:start + opts.batch_size]
        body = ',\n'.join('(' + render_row(header, r, force, opts, with_alias=False) + ')'
                          for r in chunk)
        last = start + len(chunk) >= total
        yield body + (';\n' if last else ';\n\nINSERT INTO {} ({}) VALUES\n'.format(
            dia.quote_ident(opts.table), cols))


# 渲染器注册表 —— 与 dialects.DIALECTS 同构：加一种**输出写法** = 加一个函数 + 注册一行，
# 不必再进 _render 的中段改分支。
#
# 注意 format 只管「同一种 SQL 文本的不同写法」。要换**输出载体**（json/yaml）或
# **产品形态**（ddl/orm）属于新能力，那些渲染器的入参协议与本表不同（ddl/orm 还需要
# 列类型与约束的来源），不该塞进这里污染 --format 的语义。
#
# 演进路径：出现第三种写法、或 _render 里的分支超过 3 处时，再拆成 renderers/ 独立模块；
# 届时按「载体」建目录（sql/、json/），而不是按「写法」平铺。
SqlRenderer = Callable[
    [Sequence[object], Sequence[Sequence[object]], Sequence[bool], SqlOptions],
    Iterator[str],
]

RENDERERS: Dict[str, SqlRenderer] = {
    'union': _render_union,
    'insert': _render_insert,
}

if tuple(RENDERERS) != FORMATS:      # 注册表与声明的格式清单不得漂移
    raise RuntimeError('RENDERERS 与 FORMATS 不一致：{}'.format(tuple(RENDERERS)))


def _render(header: Sequence[object], data: Sequence[Sequence[object]],
            force: Sequence[bool], opts: SqlOptions) -> Iterator[str]:
    """公共头注释 + 分派到 opts.fmt 对应的渲染器。"""
    yield from _header_comments(header, data, force, opts)
    yield from RENDERERS[opts.fmt](header, data, force, opts)


def build_sql(header: Sequence[object], rows: Sequence[Sequence[object]], opts: SqlOptions) -> str:
    """小数据量/测试用的便捷封装。"""
    return ''.join(iter_sql(header, rows, opts))


def write_sql(path, header: Sequence[object], rows: Sequence[Sequence[object]],
              opts: SqlOptions, encoding: str = 'utf-8') -> int:
    """流式写文件，返回写入的数据行数。"""
    header, data, force = prepare(header, rows, opts)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', encoding=encoding, newline='') as f:
        for chunk in _render(header, data, force, opts):
            f.write(chunk)
    return len(data)
