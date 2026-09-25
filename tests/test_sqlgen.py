# -*- coding: utf-8 -*-
import datetime
import tempfile
import unittest
from pathlib import Path

from excel2sql import dialects, sqlgen
from excel2sql.sqlgen import SqlGenError, SqlOptions, build_sql, write_sql


def opts(**kw):
    kw.setdefault('dialect', dialects.SQLSERVER)
    return SqlOptions(**kw)


def selects(sql):
    """产物里「一行一条数据」的 SELECT 语句（CTE 体内有缩进，先 strip 再判）。"""
    lines = (l.strip() for l in sql.split('\n'))
    return [l for l in lines if l.startswith('SELECT ') and not l.startswith('SELECT *')]


class TestLiterals(unittest.TestCase):
    def test_string_gets_quoted(self):
        self.assertEqual(build_sql(['a'], [['x']], opts()).count("N'x'"), 1)

    def test_number_stays_numeric(self):
        self.assertIn('SELECT 25 AS [数量]', build_sql(['数量'], [[25]], opts()))
        self.assertIn('SELECT 62.9 AS [r]', build_sql(['r'], [[62.9]], opts()))

    def test_none_becomes_null(self):
        self.assertIn('SELECT NULL AS [a], 1 AS [b]', build_sql(['a', 'b'], [[None, 1]], opts()))

    def test_empty_string_default_keeps_empty_literal(self):
        self.assertIn("SELECT N'' AS [a], 1 AS [b]", build_sql(['a', 'b'], [['', 1]], opts()))

    def test_empty_as_null_option(self):
        self.assertIn('SELECT NULL AS [a], 1 AS [b]',
                      build_sql(['a', 'b'], [['', 1]], opts(empty_as_null=True)))

    def test_bool(self):
        self.assertIn('SELECT \'1\' AS [a]', build_sql(['a'], [[True]], opts()))

    def test_nan_and_inf(self):
        sql = build_sql(['a', 'b'], [[float('nan'), float('inf')]], opts())
        self.assertIn('SELECT NULL AS [a], NULL AS [b]', sql)

    def test_quote_inside_value(self):
        self.assertIn("SELECT N'it''s' AS [a]", build_sql(['a'], [["it's"]], opts()))


class TestTypeUnification(unittest.TestCase):
    """评审第 4 条：UNION ALL 同列类型必须一致。"""

    def test_mixed_string_and_int_becomes_all_string(self):
        sql = build_sql(['序号'], [['t'], [2]], opts())
        self.assertIn("SELECT N't' AS [序号]", sql)
        self.assertIn("SELECT N'2' AS [序号]", sql)

    def test_big_number_beside_text_is_quoted(self):
        # 7900454710 超过 int 上限，若不加引号会算术溢出
        sql = build_sql(['PO'], [['7900414002@2'], [7900454710]], opts())
        self.assertIn("SELECT N'7900454710' AS [PO]", sql)

    def test_empty_string_participates_in_inference(self):
        sql = build_sql(['a', 'b'], [['', 1], [2, 3]], opts())
        self.assertIn("SELECT N'' AS [a], 1 AS [b]", sql)
        self.assertIn("SELECT N'2' AS [a], 3 AS [b]", sql)

    def test_empty_string_skipped_when_treated_as_null(self):
        sql = build_sql(['a', 'b'], [['', 1], [2, 3]], opts(empty_as_null=True))
        self.assertIn('SELECT NULL AS [a], 1 AS [b]', sql)
        self.assertIn('SELECT 2 AS [a], 3 AS [b]', sql)

    def test_all_string_option(self):
        sql = build_sql(['a'], [[25]], opts(all_string=True))
        self.assertIn("SELECT N'25' AS [a]", sql)

    def test_pure_numeric_column_stays_numeric(self):
        sql = build_sql(['a'], [[1], [2]], opts())
        self.assertIn('SELECT 1 AS [a]', sql)
        self.assertIn('SELECT 2 AS [a]', sql)


class TestDates(unittest.TestCase):
    def test_sqlserver_datetime_literal(self):
        sql = build_sql(['t'], [[datetime.datetime(2026, 9, 8, 8, 25, 46)]], opts())
        self.assertIn("SELECT N'2026-09-08 08:25:46' AS [t]", sql)

    def test_oracle_date_uses_date_format(self):
        """评审第 3 条：纯 date 不能用 HH24:MI:SS 格式，否则 ORA-01840。"""
        sql = build_sql(['d'], [[datetime.date(2026, 9, 8)]],
                        opts(dialect=dialects.ORACLE))
        self.assertIn("TO_DATE('2026-09-08','YYYY-MM-DD')", sql)
        self.assertNotIn('HH24', sql)

    def test_oracle_datetime_uses_datetime_format(self):
        sql = build_sql(['d'], [[datetime.datetime(2026, 9, 8, 8, 25, 46)]],
                        opts(dialect=dialects.ORACLE))
        self.assertIn("TO_DATE('2026-09-08 08:25:46','YYYY-MM-DD HH24:MI:SS')", sql)

    def test_oracle_select_has_dual(self):
        sql = build_sql(['a'], [[1]], opts(dialect=dialects.ORACLE))
        self.assertIn('FROM dual', sql)


class TestMultiline(unittest.TestCase):
    """单元格内换行要拼成单行表达式，避免一条 SELECT 被换行截断。"""

    def test_sqlserver_concat(self):
        sql = build_sql(['t'], [['上海市\n浦东新区']], opts())
        self.assertIn("SELECT N'上海市' + CHAR(10) + N'浦东新区' AS [t]", sql)

    def test_mysql_concat(self):
        # MySQL 必须带 USING utf8mb4：裸 CHAR(10) 是二进制串，
        # 会让 CONCAT 的结果整列变成 varbinary（CTAS 时列类型/字符集全错）
        sql = build_sql(['t'], [['上海市\n浦东新区']], opts(dialect=dialects.MYSQL))
        self.assertIn("SELECT CONCAT('上海市', CHAR(10 USING utf8mb4), '浦东新区') AS `t`", sql)

    def test_oracle_chr(self):
        sql = build_sql(['t'], [['上海市\n浦东新区']], opts(dialect=dialects.ORACLE))
        self.assertIn("'上海市' || CHR(10) || '浦东新区'", sql)

    def test_crlf_normalized(self):
        sql = build_sql(['t'], [['a\r\nb']], opts())
        self.assertIn("N'a' + CHAR(10) + N'b'", sql)
        self.assertNotIn('\\r', sql)


class TestIdentifiers(unittest.TestCase):
    def test_malicious_column_name_is_escaped(self):
        sql = build_sql(['a] FROM x; DROP TABLE y; --'], [[1]], opts())
        self.assertIn('[a]] FROM x; DROP TABLE y; --]', sql)
        self.assertNotIn('DROP TABLE y; --] AS', sql)

    def test_table_name_escaped(self):
        sql = build_sql(['a'], [[1]], opts(table='t]x'))
        self.assertIn('WITH [t]]x] AS', sql)
        self.assertIn('SELECT * FROM [t]]x];', sql)


class TestFormats(unittest.TestCase):
    def test_union_cte_wrapper(self):
        sql = build_sql(['a'], [[1], [2]], opts())
        self.assertTrue(sql.rstrip().endswith('SELECT * FROM [HARDCODE];'))
        self.assertEqual(len(selects(sql)), 2)

    def test_union_plain_has_no_wrapper(self):
        sql = build_sql(['a'], [[1], [2]], opts(wrap='plain'))
        self.assertNotIn('WITH ', sql)
        self.assertIn('UNION ALL', sql)

    def test_union_all_count(self):
        sql = build_sql(['a'], [[i] for i in range(5)], opts())
        self.assertEqual(sql.count('UNION ALL'), 4)

    def test_union_all_is_always_a_line_of_its_own(self):
        """UNION ALL 必须独占一行：与上一行粘连会造出 [a]UNION 这类坏 token。

        两种包裹方式都要守：CTE 体内允许缩进，但整行除了缩进不能有别的东西。
        """
        for wrap, expected in (('cte', '    UNION ALL'), ('plain', 'UNION ALL')):
            with self.subTest(wrap=wrap):
                sql = build_sql(['a'], [[1], [2]], opts(wrap=wrap))
                self.assertEqual([l for l in sql.split('\n') if 'UNION ALL' in l],
                                 [expected])

    def test_oracle_dual_not_glued_to_union(self):
        """Oracle 下 SELECT ... FROM dual 后面紧贴 UNION ALL 会解析失败。"""
        sql = build_sql(['a'], [[1], [2]], opts(dialect=dialects.ORACLE, wrap='plain'))
        self.assertIn('FROM dual\nUNION ALL\n', sql)

    def test_plain_wrap_union_on_own_line(self):
        sql = build_sql(['a'], [[1], [2]], opts(wrap='plain'))
        self.assertIn('SELECT 1 AS [a]\nUNION ALL\nSELECT 2 AS [a]', sql)

    def test_insert_single_batch(self):
        sql = build_sql(['a', 'b'], [[1, 'x'], [2, 'y']], opts(fmt='insert'))
        self.assertIn('INSERT INTO [HARDCODE] ([a], [b]) VALUES', sql)
        self.assertIn("(1, N'x'),", sql)
        self.assertTrue(sql.rstrip().endswith("(2, N'y');"))

    def test_insert_batching(self):
        sql = build_sql(['a'], [[i] for i in range(5)], opts(fmt='insert', batch_size=2))
        self.assertEqual(sql.count('INSERT INTO'), 3)   # 2 + 2 + 1
        self.assertEqual(sql.count(';'), 3)

    def test_insert_mysql_backticks(self):
        sql = build_sql(['a'], [[1]], opts(dialect=dialects.MYSQL, fmt='insert'))
        self.assertIn('INSERT INTO `HARDCODE` (`a`) VALUES', sql)


class TestCteLayout(unittest.TestCase):
    """CTE 体内缩进一级、闭合括号回到行首（对齐 sqlfluff 等工具的默认风格）。"""

    def test_inner_selects_and_union_are_indented(self):
        sql = build_sql(['a'], [[1], [2]], opts())
        self.assertIn('WITH [HARDCODE] AS (\n', sql)
        self.assertIn('\n    SELECT 1 AS [a]\n', sql)
        self.assertIn('\n    UNION ALL\n', sql)
        self.assertIn('\n    SELECT 2 AS [a]\n)\nSELECT * FROM [HARDCODE];\n', sql)

    def test_closing_paren_returns_to_column_zero(self):
        sql = build_sql(['a'], [[1], [2]], opts())
        lines = sql.split('\n')
        self.assertEqual(lines[lines.index('    SELECT 2 AS [a]') + 1], ')')

    def test_every_row_line_is_indented_exactly_once(self):
        sql = build_sql(['a'], [[i] for i in range(4)], opts())
        body = [l for l in sql.split('\n')
                if l.strip().startswith('SELECT ') and not l.strip().startswith('SELECT *')]
        self.assertEqual(len(body), 4)
        for line in body:
            self.assertTrue(line.startswith('    '), line)
            self.assertFalse(line.startswith('     '), line)   # 只缩进一级

    def test_plain_wrap_is_not_indented(self):
        """plain 是给「嵌进已有 SQL」用的，缩进交给调用方按所在层级对齐。"""
        sql = build_sql(['a'], [[1], [2]], opts(wrap='plain'))
        self.assertIn('SELECT 1 AS [a]\nUNION ALL\nSELECT 2 AS [a]', sql)
        self.assertNotIn('    SELECT', sql)

    def test_indent_does_not_leak_into_insert(self):
        sql = build_sql(['a', 'b'], [[1, 'x']], opts(fmt='insert'))
        self.assertNotIn('    ', sql)

    def test_header_comments_are_not_indented(self):
        sql = build_sql(['a'], [[1]], opts())
        head = [l for l in sql.split('\n') if l.startswith('--')]
        self.assertTrue(head)
        for line in head:
            self.assertFalse(line.startswith(' '), line)


class TestRowsAndErrors(unittest.TestCase):
    def test_blank_rows_are_dropped(self):
        sql = build_sql(['a'], [[1], [None], ['  '], [2]], opts())
        self.assertEqual(len(selects(sql)), 2)

    def test_short_row_padded_with_null(self):
        sql = build_sql(['a', 'b'], [[1]], opts())
        self.assertIn('SELECT 1 AS [a], NULL AS [b]', sql)

    def test_all_blank_row_is_dropped(self):
        sql = build_sql(['a', 'b'], [[1, 2], ['', '  '], [None, None]], opts())
        self.assertEqual(len(selects(sql)), 1)

    def test_extra_columns_ignored(self):
        sql = build_sql(['a'], [[1, 2, 3]], opts())
        self.assertIn('SELECT 1 AS [a]', sql)

    def test_no_data_raises(self):
        with self.assertRaises(SqlGenError):
            build_sql(['a'], [[None], [None]], opts())

    def test_empty_header_raises(self):
        with self.assertRaises(SqlGenError):
            build_sql([], [], opts())

    def test_bad_format_raises(self):
        with self.assertRaises(SqlGenError):
            opts(fmt='update')

    def test_header_comment_lists_forced_columns(self):
        # 整列都是文本 -> 归因"源数据是文本"（CSV 的典型情形）
        sql = build_sql(['文字', '数字'], [['a', 1]], opts())
        self.assertIn('按字符串输出的列（源数据是文本（CSV 无类型信息））：文字', sql)

    def test_header_comment_says_mixed_when_column_has_both(self):
        # 同列既有文本又有数字 -> 归因"同列混类型"（工具被迫统一的那档）
        sql = build_sql(['混'], [['a'], [2]], opts())
        self.assertIn('同列混有数字/日期与文本，整列统一为字符串）：混', sql)

    def test_header_comment_says_user_asked_for_all_string(self):
        # 用户显式要求 -> 单独一档，不再与"工具被迫"混为一谈
        sql = build_sql(['数字'], [[1]], opts(all_string=True))
        self.assertIn('按字符串输出的列（用户指定 --all-string）：数字', sql)

    def test_no_comment_when_nothing_forced(self):
        sql = build_sql(['数字'], [[1]], opts())
        self.assertNotIn('按字符串输出的列', sql)


class TestWrite(unittest.TestCase):
    def test_write_returns_row_count(self):
        with tempfile.TemporaryDirectory() as d:
            out = Path(d) / 'sub' / 'x.sql'
            n = write_sql(out, ['a'], [[1], [2], [3]], opts())
            self.assertEqual(n, 3)
            self.assertTrue(out.is_file())
            self.assertIn('SELECT 3 AS [a]', out.read_text(encoding='utf-8'))

    def test_utf8_sig_encoding_option(self):
        """评审小问题 5：老版 SSMS 需要 BOM 才不乱码。"""
        with tempfile.TemporaryDirectory() as d:
            out = Path(d) / 'x.sql'
            write_sql(out, ['列'], [[1]], opts(), encoding='utf-8-sig')
            self.assertTrue(out.read_bytes().startswith(b'\xef\xbb\xbf'))

    def test_plain_utf8_has_no_bom(self):
        with tempfile.TemporaryDirectory() as d:
            out = Path(d) / 'x.sql'
            write_sql(out, ['列'], [[1]], opts())
            self.assertFalse(out.read_bytes().startswith(b'\xef\xbb\xbf'))


class TestRendererRegistry(unittest.TestCase):
    """渲染器注册表：加一种输出写法 = 加一个函数 + 注册一行。"""

    def test_registry_keys_match_declared_formats(self):
        self.assertEqual(tuple(sqlgen.RENDERERS), sqlgen.FORMATS)

    def test_every_renderer_is_callable(self):
        for name, fn in sqlgen.RENDERERS.items():
            with self.subTest(fmt=name):
                self.assertTrue(callable(fn))

    def test_render_delegates_and_keeps_header_comments(self):
        """分派器必须自己产出头注释，否则改写法会静默丢掉告警。"""
        sql = build_sql(['文字'], [['x']], opts())
        self.assertIn('-- 由 excel2sql 生成：1 行 x 1 列', sql)
        self.assertIn('-- 方言：SQL Server', sql)
        self.assertIn('按字符串输出的列', sql)

    def test_header_comments_are_shared_by_both_renderers(self):
        """两种写法的头注释只有「输出格式」那一行不同 —— 公共部分不允许各写一份。"""
        def head(s):
            return [l for l in s.split('\n') if l.startswith('--')]

        union = head(build_sql(['a'], [[1]], opts(fmt='union')))
        insert = head(build_sql(['a'], [[1]], opts(fmt='insert')))

        self.assertEqual(len(union), len(insert))
        self.assertIn('输出格式：union', union[1])
        self.assertIn('输出格式：insert', insert[1])
        self.assertEqual(union[:1] + union[2:], insert[:1] + insert[2:])

    def test_non_sql_carrier_is_still_rejected(self):
        """json/csv 是另一个载体，不是 format 的取值 —— 这一层语义要守住。"""
        for carrier in ('json', 'csv', 'ddl', 'orm'):
            with self.subTest(fmt=carrier):
                with self.assertRaises(SqlGenError):
                    SqlOptions(dialect=dialects.SQLSERVER, fmt=carrier)


if __name__ == '__main__':
    unittest.main()
