"""SQL 生成：把二维数据渲染成硬编码的内联表。

安全要点：
- 标识符（列名/表名）一律走 Dialect.quote_ident 转义内部引号；
- 字符串按方言转义（MySQL 额外转义反斜杠）；
- 类型按列统一，保证 UNION ALL / 多行 VALUES 的列类型一致。
"""
from __future__ import annotations

import datetime
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Iterator, List, Optional, Sequence

from .dialects import Dialect
from .headers import is_blank

FORMATS = ('union', 'insert')
WRAPS = ('cte', 'plain')

# 超过该行数仍用 UNION ALL 会明显拖慢解析，默认建议 insert
UNION_ROW_WARN = 2000


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
    force = infer_column_types(header, data, opts.all_string, opts.empty_as_null)
    return list(header), data, force


def iter_sql(header: Sequence[object], rows: Sequence[Sequence[object]],
             opts: SqlOptions) -> Iterator[str]:
    """流式产出 SQL 片段，避免大表一次性拼成巨型字符串。"""
    header, data, force = prepare(header, rows, opts)
    return _render(header, data, force, opts)


def _render(header: Sequence[object], data: Sequence[Sequence[object]],
            force: Sequence[bool], opts: SqlOptions) -> Iterator[str]:
    dia = opts.dialect
    cols = ', '.join(dia.quote_ident(h) for h in header)
    forced = [h for h, f in zip(header, force) if f]

    yield '-- 由 excel2sql 生成：{} 行 x {} 列\n'.format(len(data), len(header))
    yield '-- 方言：{}   输出格式：{}\n'.format(dia.name, opts.fmt)
    if forced:
        yield '-- 混类型列统一为字符串：{}\n'.format(', '.join(str(x) for x in forced))
    yield '-- 提示：内联数据仅供测试/修数，请勿直接用于生产批量导入\n'

    if opts.fmt == 'insert':
        yield 'INSERT INTO {} ({}) VALUES\n'.format(dia.quote_ident(opts.table), cols)
        total = len(data)
        for start in range(0, total, opts.batch_size):
            chunk = data[start:start + opts.batch_size]
            body = ',\n'.join('(' + render_row(header, r, force, opts, with_alias=False) + ')'
                              for r in chunk)
            last = start + len(chunk) >= total
            yield body + (';\n' if last else ';\n\nINSERT INTO {} ({}) VALUES\n'.format(
                dia.quote_ident(opts.table), cols))
        return

    lines = ['SELECT ' + render_row(header, r, force, opts) + dia.dual for r in data]
    if opts.wrap == 'plain':
        yield '\nUNION ALL\n'.join(lines) + '\n'
    else:
        yield 'WITH {} AS (\n'.format(dia.quote_ident(opts.table))
        yield '\nUNION ALL\n'.join(lines)
        yield '\n)\nSELECT * FROM {};\n'.format(dia.quote_ident(opts.table))


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
