# -*- coding: utf-8 -*-
import unittest

from excel2sql import dialects


class TestIdentifierQuoting(unittest.TestCase):
    """评审第 1 条：标识符必须转义内部引号，否则列名能改写 SQL 结构。"""

    def test_sqlserver_escapes_bracket(self):
        name = 'a] FROM x; DROP TABLE y; --'
        self.assertEqual(dialects.SQLSERVER.quote_ident(name),
                         '[a]] FROM x; DROP TABLE y; --]')

    def test_mysql_escapes_backtick(self):
        self.assertEqual(dialects.MYSQL.quote_ident('a`b'), '`a``b`')

    def test_ansi_escapes_doublequote(self):
        self.assertEqual(dialects.ORACLE.quote_ident('a"b'), '"a""b"')
        self.assertEqual(dialects.POSTGRESQL.quote_ident('a"b'), '"a""b"')

    def test_chinese_column_names(self):
        self.assertEqual(dialects.SQLSERVER.quote_ident('装运方式号'), '[装运方式号]')


class TestStringQuoting(unittest.TestCase):
    def test_single_quote_doubling(self):
        self.assertEqual(dialects.SQLSERVER.quote_string("it's"), "N'it''s'")
        self.assertEqual(dialects.POSTGRESQL.quote_string("it's"), "'it''s'")

    def test_mysql_escapes_backslash(self):
        """评审第 2 条：MySQL 默认 sql_mode 下反斜杠是转义符。"""
        self.assertEqual(dialects.MYSQL.quote_string('C:\\temp'), "'C:\\\\temp'")

    def test_non_mysql_keeps_backslash(self):
        self.assertEqual(dialects.SQLSERVER.quote_string('C:\\temp'), "N'C:\\temp'")
        self.assertEqual(dialects.ORACLE.quote_string('C:\\temp'), "'C:\\temp'")

    def test_mysql_backslash_before_quote(self):
        # 输入 \' -> 反斜杠加倍 + 单引号加倍 -> '\\'''（6 字符）
        self.assertEqual(dialects.MYSQL.quote_string("\\'"), "'\\\\'''")


class TestConcat(unittest.TestCase):
    SEGS = ["N'a'", "N'b'"]

    def test_sqlserver_plus(self):
        self.assertEqual(dialects.SQLSERVER.concat_strings(self.SEGS, 'CHAR(10)'),
                         "N'a' + CHAR(10) + N'b'")

    def test_mysql_concat_function(self):
        self.assertEqual(dialects.MYSQL.concat_strings(["'a'", "'b'"], 'CHAR(10)'),
                         "CONCAT('a', CHAR(10), 'b')")

    def test_oracle_pipe(self):
        self.assertEqual(dialects.ORACLE.concat_strings(["'a'", "'b'"], 'CHR(10)'),
                         "'a' || CHR(10) || 'b'")


class TestResolve(unittest.TestCase):
    def test_numeric_and_name_aliases(self):
        self.assertIs(dialects.resolve('1'), dialects.SQLSERVER)
        self.assertIs(dialects.resolve('MySQL'), dialects.MYSQL)
        self.assertIs(dialects.resolve('ORA'), dialects.ORACLE)
        self.assertIs(dialects.resolve('pg'), dialects.POSTGRESQL)

    def test_unknown_falls_back_to_sqlserver(self):
        self.assertIs(dialects.resolve('db2'), dialects.SQLSERVER)


class TestDateFormats(unittest.TestCase):
    def test_oracle_date_and_datetime_differ(self):
        """评审第 3 条：格式模型必须与值匹配，否则 ORA-01840。"""
        self.assertIn('YYYY-MM-DD', dialects.ORACLE.date_suffix)
        self.assertNotIn('HH24', dialects.ORACLE.date_suffix)
        self.assertIn('HH24:MI:SS', dialects.ORACLE.datetime_suffix)
        self.assertTrue(dialects.ORACLE.date_prefix.startswith('TO_DATE('))
        self.assertNotIn("'", dialects.ORACLE.date_prefix)   # 引号交给 quote_string

    def test_non_oracle_has_no_date_wrapper(self):
        self.assertEqual(dialects.SQLSERVER.date_prefix, '')
        self.assertEqual(dialects.SQLSERVER.datetime_suffix, '')


if __name__ == '__main__':
    unittest.main()
