"""SQL 方言定义：标识符引用、字符串转义、日期包装、换行拼接。"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict


@dataclass(frozen=True)
class Dialect:
    """一种 SQL 方言的字面量/标识符规则。"""

    key: str
    name: str
    ident_open: str
    ident_close: str
    string_prefix: str = ''
    concat: str = 'plus'          # plus: a + b | pipe: a || b | func: CONCAT(a, b)
    newline_expr: str = 'CHAR(10)'
    date_prefix: str = ''
    date_suffix: str = ''
    datetime_prefix: str = ''
    datetime_suffix: str = ''
    dual: str = ''
    backslash_escape: bool = False

    # -- 标识符 --
    def quote_ident(self, name: object) -> str:
        """引用列名/表名，并转义内部引号（防语法破坏与注入）。"""
        if self.ident_open == '[':
            inner = str(name).replace(']', ']]')
        elif self.ident_open == '`':
            inner = str(name).replace('`', '``')
        else:                                     # 标准双引号方言
            inner = str(name).replace('"', '""')
        return '{}{}{}'.format(self.ident_open, inner, self.ident_close)

    # -- 字符串 --
    def quote_string(self, body: str) -> str:
        """单行字符串字面量（不含换行）。"""
        s = body
        if self.backslash_escape:                 # MySQL 默认把 \\ 当转义符
            s = s.replace('\\', '\\\\')
        s = s.replace("'", "''")
        return "{}'{}'".format(self.string_prefix, s)

    def concat_strings(self, segments, newline_expr: str) -> str:
        """把多行文本拼成单个表达式。"""
        parts = []
        for i, seg in enumerate(segments):
            if i:
                parts.append(newline_expr)
            parts.append(seg)
        if self.concat == 'func':
            return 'CONCAT({})'.format(', '.join(parts))
        op = ' || ' if self.concat == 'pipe' else ' + '
        return op.join(parts)

    @property
    def date_format(self) -> str:
        return 'YYYY-MM-DD'

    @property
    def datetime_format(self) -> str:
        return 'YYYY-MM-DD HH24:MI:SS'


SQLSERVER = Dialect(
    key='sqlserver', name='SQL Server', ident_open='[', ident_close=']',
    string_prefix='N', concat='plus', newline_expr='CHAR(10)',
)

MYSQL = Dialect(
    key='mysql', name='MySQL', ident_open='`', ident_close='`',
    string_prefix='', concat='func',
    # ★ 必须带 USING：MySQL 里裸 CHAR(10) 返回的是**二进制字符串**，
    #   CONCAT() 一旦遇到二进制参数，结果也是二进制 ——
    #   `CREATE TABLE AS SELECT CONCAT('a', CHAR(10), 'b')` 会推出 varbinary 列
    #   （排序/比较按字节走，字符集为 NULL，不能直接用于生产）。
    #   实测 2026-09-25：SQL Server / PostgreSQL / Oracle 的三种写法都正常，只有 MySQL 需要改。
    newline_expr='CHAR(10 USING utf8mb4)',
    backslash_escape=True,
)

ORACLE = Dialect(
    key='oracle', name='Oracle', ident_open='"', ident_close='"',
    string_prefix='', concat='pipe', newline_expr='CHR(10)',
    # 前后缀不含引号：字面量由 quote_string 负责加引号
    date_prefix='TO_DATE(', date_suffix=",'YYYY-MM-DD')",
    datetime_prefix='TO_DATE(', datetime_suffix=",'YYYY-MM-DD HH24:MI:SS')",
    dual=' FROM dual',
)

POSTGRESQL = Dialect(
    key='postgresql', name='PostgreSQL', ident_open='"', ident_close='"',
    string_prefix='', concat='pipe', newline_expr='CHR(10)',
)

SQLITE = Dialect(
    key='sqlite', name='SQLite', ident_open='"', ident_close='"',
    string_prefix='', concat='pipe', newline_expr='CHAR(10)',
    # 日期不加包装：SQLite 没有日期类型，ISO-8601 文本就是它的规范表示，
    # 而 '2024-11-11' / '2024-11-11 13:05:00' 这样的字面量能被 date()/strftime()
    # 直接识别（实测 strftime('%Y', "订购日期") 正常返回 '2024'）。
    #
    # 规律是「只有真正需要包装的方言才加前缀」—— 目前只有 Oracle 需要
    # （裸字符串参与日期比较会失败，必须 TO_DATE）。SQL Server / MySQL / PostgreSQL
    # 都由上下文自动转换，SQLite 则根本不需要转。
)

DIALECTS: Dict[str, Dialect] = {d.key: d for d in (SQLSERVER, MYSQL, ORACLE, POSTGRESQL, SQLITE)}

# 命令行/交互输入的别名 -> 方言 key
ALIASES = {
    '1': 'sqlserver', 'sqlserver': 'sqlserver', 'mssql': 'sqlserver',
    '2': 'mysql', 'mysql': 'mysql', 'mariadb': 'mysql',
    '3': 'oracle', 'oracle': 'oracle', 'ora': 'oracle',
    '4': 'postgresql', 'postgresql': 'postgresql', 'postgres': 'postgresql', 'pg': 'postgresql',
    '5': 'sqlite', 'sqlite': 'sqlite', 'sqlite3': 'sqlite',
}

ORDER = ('sqlserver', 'mysql', 'oracle', 'postgresql', 'sqlite')

MENU = '   '.join('{}={}'.format(k, DIALECTS[k].name) for k in ORDER)
NUMBERED = {str(i): k for i, k in enumerate(ORDER, start=1)}


def resolve(value: str) -> Dialect:
    """把 '-d 1' / 'mysql' / 'PostgreSQL' 统一解析成 Dialect；无法识别时返回 SQL Server。"""
    return DIALECTS[ALIASES.get(str(value).strip().lower(), 'sqlserver')]
