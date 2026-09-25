# -*- coding: utf-8 -*-
"""SQLite 方言：不只比对字面量文本，而是**把产物真灌进 sqlite3 再逐格核对**。

SQLite 的价值在于标准库自带 `sqlite3` —— 这条「生成 -> 落库 -> 读回」的端到端链路
能直接进 CI，不像 SQL Server / Oracle 那样必须依赖外部容器（见 docs/testing/e2e-testing.md）。

夹具是故意挑的：编号带前导零（auto 推断必须把它留在文本）、金额由 int 与 float 混排、
备注含真实换行、数量列有一个空格子。
"""
import datetime
import io
import os
import sqlite3
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path

from excel2sql.cli import EXIT_OK, main
from excel2sql.dialects import SQLITE
from excel2sql.sqlgen import SqlOptions, build_sql

openpyxl = None
try:
    import openpyxl  # noqa: F401
except ImportError:                                   # pragma: no cover
    pass

HEADER = ['订单编号', '数量', '金额', '订购日期', '备注']

CSV = (
    '订单编号,数量,金额,订购日期,备注\n'
    '0001,2,221.98,2024-11-11,"第一行\n第二行"\n'
    '0002,9,3709.39,2024-02-05,\n'
    '0003,,0,2024-01-01,甲\n'
)

EXPECTED = [
    ('0001', 2, 221.98, '2024-11-11', '第一行\n第二行'),
    ('0002', 9, 3709.39, '2024-02-05', ''),
    ('0003', None, 0, '2024-01-01', '甲'),
]


class _Dir:
    """临时目录 + 空白 excel2sql.ini。

    那个 ini 是必需的：配置文件查找顺序是 `./excel2sql.ini` -> `~/.excel2sql/config.ini`，
    没有它就可能会被**本机用户级配置**影响，测试结果随机器而变。
    """

    def __enter__(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self._cwd = Path.cwd()
        os.chdir(str(self.root))
        (self.root / 'excel2sql.ini').write_text('# 测试用：阻断用户级配置\n', encoding='utf-8')
        return self.root

    def __exit__(self, *exc):
        os.chdir(str(self._cwd))
        self.tmp.cleanup()
        return False


def run_cli(root, extra):
    src = root / 'orders.csv'
    src.write_text(CSV, encoding='utf-8')
    out = root / 'out.sql'
    argv = [str(src), '-d', 'sqlite', '--header-row', '1',
            '-o', str(out), '--no-copy-clipboard'] + extra
    buf_out, buf_err = io.StringIO(), io.StringIO()
    with redirect_stdout(buf_out), redirect_stderr(buf_err):
        code = main(argv)
    text = out.read_text(encoding='utf-8') if out.is_file() else ''
    return code, text, buf_err.getvalue()


def strip_comments(sql):
    """去掉 `--` 开头的注释行 —— 它要嵌进 CREATE TABLE AS 里，注释行夹在中间没必要。"""
    return '\n'.join(l for l in sql.split('\n') if not l.lstrip().startswith('--'))


def load_union(sql):
    """union 产物本身就是一条 SELECT（带 CTE），直接 CTAS 进表。

    SQLite 的 CTAS 不给列声明类型（BLOB 亲和），所以不会顺手改写存入的值 ——
    这正是能拿它来验证「产物本身无损」的前提。
    """
    conn = sqlite3.connect(':memory:')
    conn.executescript('CREATE TABLE loaded AS ' + strip_comments(sql).strip())
    return conn


def load_insert(sql):
    """insert 产物要求目标表已存在；按表头建一张**不带类型声明**的表。"""
    cols = ', '.join('"{}"'.format(h.replace('"', '""')) for h in HEADER)
    conn = sqlite3.connect(':memory:')
    conn.executescript('CREATE TABLE loaded ({}) ;\n{}'.format(
        cols, strip_comments(sql).strip()))
    return conn


class _RoundTrip:
    """三种输出形式共用同一套断言：灌库之后逐格核对，一格都不许变。"""

    ARGS = ()
    LOAD = staticmethod(load_union)

    def _conn(self):
        with _Dir() as root:
            code, sql, err = run_cli(root, list(self.ARGS))
            self.assertEqual(code, EXIT_OK, err)
            self.assertTrue(sql, '没有产出 SQL')
            return self.LOAD(sql)                  # 连库，后面直接查

    def _rows(self):
        return [tuple(r) for r in self._conn().execute('SELECT * FROM loaded').fetchall()]

    def test_row_count(self):
        self.assertEqual(len(self._rows()), len(EXPECTED))

    def test_column_names(self):
        got = [d[0] for d in self._conn().execute('SELECT * FROM loaded').description]
        self.assertEqual(got, HEADER)

    def test_values_are_identical(self):
        self.assertEqual(self._rows(), [tuple(r) for r in EXPECTED])

    def test_text_stays_text(self):
        """编号/日期/备注必须是 text —— 前导零被转成数字是最典型的数据失真。"""
        got = self._conn().execute(
            'SELECT typeof("订单编号"), typeof("订购日期"), typeof("备注") FROM loaded'
        ).fetchall()
        self.assertEqual(got, [('text', 'text', 'text')] * len(EXPECTED))

    def test_numbers_are_numbers(self):
        got = self._conn().execute(
            'SELECT typeof("数量"), typeof("金额") FROM loaded').fetchall()
        self.assertEqual(got, [('integer', 'real'),
                               ('integer', 'real'),
                               ('null', 'integer')])

    def test_leading_zero_identifier_is_intact(self):
        got = [r[0] for r in self._conn().execute(
            'SELECT "订单编号" FROM loaded ORDER BY "订单编号"').fetchall()]
        self.assertEqual(got, ['0001', '0002', '0003'])

    def test_multiline_newline_survives(self):
        """多行文本靠 `|| CHAR(10) ||` 拼出来，落库后必须是**真的换行**。"""
        val = self._conn().execute(
            'SELECT "备注" FROM loaded WHERE "订单编号" = ?', ('0001',)).fetchone()[0]
        self.assertEqual(val, '第一行\n第二行')
        self.assertIn('\n', val)

    def test_blank_numeric_cell_becomes_null(self):
        val = self._conn().execute(
            'SELECT "数量" FROM loaded WHERE "订单编号" = ?', ('0003',)).fetchone()[0]
        self.assertIsNone(val)

    def test_empty_text_stays_empty_string(self):
        """空串不是 NULL —— 想变成 NULL 要显式加 --empty-as-null。"""
        val = self._conn().execute(
            'SELECT "备注" FROM loaded WHERE "订单编号" = ?', ('0002',)).fetchone()[0]
        self.assertEqual(val, '')

    def test_date_is_still_queryable_as_a_date(self):
        """DATE() 包一层不影响比较与排序：两边都是 ISO 文本。"""
        got = self._conn().execute(
            "SELECT COUNT(*) FROM loaded WHERE \"订购日期\" >= '2024-02-01'").fetchone()[0]
        self.assertEqual(got, 2)


class TestUnionCteRoundTrip(_RoundTrip, unittest.TestCase):
    ARGS = ['--format', 'union', '--wrap', 'cte', '--table', 'src']


class TestUnionPlainRoundTrip(_RoundTrip, unittest.TestCase):
    ARGS = ['--format', 'union', '--wrap', 'plain', '--table', 'src']


class TestInsertRoundTrip(_RoundTrip, unittest.TestCase):
    ARGS = ['--format', 'insert', '--table', 'loaded']
    LOAD = staticmethod(load_insert)


class TestSqliteLiterals(unittest.TestCase):
    """字面量层面的方言规则（比落库更快，先在小处钉死）。"""

    def _sql(self, header, rows, **kw):
        kw.setdefault('dialect', SQLITE)
        return build_sql(header, rows, SqlOptions(**kw))

    def test_dates_are_plain_iso_text(self):
        """SQLite 没有日期类型：直接写 ISO-8601 文本，不套 DATE()/DATETIME()。

        这个决定的前提是「SQLite 的日期函数能直接识别这种文本字面量」，
        前提由 TestSqliteDateFunctions 守着 —— 前提没了就得把包装加回来。
        """
        sql = self._sql(['d', 'ts'], [[datetime.date(2024, 11, 11),
                                       datetime.datetime(2024, 11, 11, 13, 5, 0)]])
        self.assertIn('\'2024-11-11\' AS "d"', sql)
        self.assertIn('\'2024-11-11 13:05:00\' AS "ts"', sql)
        self.assertNotIn('DATE(', sql)
        self.assertNotIn('DATETIME(', sql)

    def test_multiline_uses_pipe_concat_and_plain_quotes(self):
        sql = self._sql(['t'], [['a\nb']])
        self.assertIn('\'a\' || CHAR(10) || \'b\' AS "t"', sql)
        self.assertNotIn("N'", sql)                       # SQLite 没有 N 前缀

    def test_no_dual_in_select(self):
        sql = self._sql(['a'], [[1]])
        self.assertNotIn('dual', sql)
        self.assertIn('SELECT 1 AS "a"\n', sql)

    def test_identifier_doublequote_escaping(self):
        sql = self._sql(['a"b'], [[1]])
        self.assertIn('AS "a""b"', sql)


class TestSqliteDateFunctions(unittest.TestCase):
    """日期以 ISO 文本落库后，SQLite 的日期函数必须能直接认出来。"""

    def _conn(self):
        with _Dir() as root:
            code, sql, err = run_cli(root, ['--table', 'src'])
            self.assertEqual(code, EXIT_OK, err)
            return load_union(sql)

    def test_strftime_reads_the_year(self):
        got = self._conn().execute(
            "SELECT strftime('%Y', \"订购日期\") FROM loaded ORDER BY \"订单编号\""
        ).fetchall()
        self.assertEqual([r[0] for r in got], ['2024', '2024', '2024'])

    def test_date_comparison_uses_text_order(self):
        got = self._conn().execute(
            "SELECT COUNT(*) FROM loaded WHERE \"订购日期\" >= '2024-02-01'").fetchone()[0]
        self.assertEqual(got, 2)

    def test_date_column_is_stored_as_text(self):
        got = self._conn().execute(
            'SELECT DISTINCT typeof("订购日期") FROM loaded').fetchall()
        self.assertEqual(got, [('text',)])


@unittest.skipIf(openpyxl is None, '需要 openpyxl')
class TestSqliteExcelDateCell(unittest.TestCase):
    """xlsx 里**真正的日期单元格**走到 SQLite 的完整链路。

    注意：openpyxl 把日期单元格读成 `datetime.datetime`，所以产物里带一个
    `00:00:00`（这在所有方言下都一样，是既有行为）。SQLite 的 date() 能从中取出日期部分。
    """

    def test_excel_date_cell_lands_as_iso_text(self):
        with _Dir() as root:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(['日期', '备注'])
            ws.append([datetime.date(2024, 11, 11), '甲'])
            xlsx = root / 'd.xlsx'
            wb.save(xlsx)

            out = root / 'd.sql'
            buf = io.StringIO()
            with redirect_stdout(buf), redirect_stderr(buf):
                code = main([str(xlsx), '-d', 'sqlite', '--header-row', '1',
                             '-o', str(out), '--no-copy-clipboard'])
            self.assertEqual(code, EXIT_OK, buf.getvalue())

            sql = out.read_text(encoding='utf-8')
            self.assertNotIn('DATE(', sql)                # 不加包装
            conn = sqlite3.connect(':memory:')
            conn.executescript('CREATE TABLE loaded AS ' + strip_comments(sql).strip())
            got = conn.execute('SELECT date("日期"), typeof("日期") FROM loaded').fetchone()
            self.assertEqual(got, ('2024-11-11', 'text'))


if __name__ == '__main__':
    unittest.main()
