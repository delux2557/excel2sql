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
        self.assertEqual(dialects.SQLSERVER.quote_ident('客户 ID'), '[客户 ID]')


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


class TestSqlite(unittest.TestCase):
    """SQLite：无 N 前缀、标识符用双引号、没有原生日期类型。"""

    def test_identifier_quoting(self):
        self.assertEqual(dialects.SQLITE.quote_ident('a"b'), '"a""b"')
        self.assertEqual(dialects.SQLITE.quote_ident('客户 ID'), '"客户 ID"')

    def test_no_n_prefix(self):
        self.assertEqual(dialects.SQLITE.string_prefix, '')
        self.assertEqual(dialects.SQLITE.quote_string("it's"), "'it''s'")

    def test_backslash_is_not_an_escape(self):
        """只有 MySQL 默认把反斜杠当转义符，SQLite 不是。"""
        self.assertEqual(dialects.SQLITE.quote_string('C:\\temp'), "'C:\\temp'")

    def test_pipe_concat(self):
        self.assertEqual(dialects.SQLITE.concat_strings(["'a'", "'b'"], 'CHAR(10)'),
                         "'a' || CHAR(10) || 'b'")

    def test_date_has_no_wrapper(self):
        """SQLite 没有日期类型，ISO 文本就是规范表示，不需要（也不该）包一层。

        只有 Oracle 需要包装（裸字符串参与日期比较会失败）。
        """
        self.assertEqual(dialects.SQLITE.date_prefix, '')
        self.assertEqual(dialects.SQLITE.date_suffix, '')
        self.assertEqual(dialects.SQLITE.datetime_prefix, '')
        self.assertEqual(dialects.SQLITE.datetime_suffix, '')
        self.assertNotIn('TO_DATE', dialects.SQLITE.date_suffix)

    def test_no_dual(self):
        self.assertEqual(dialects.SQLITE.dual, '')

    def test_resolve_aliases(self):
        for value in ('sqlite', 'SQLite', 'SQLITE', 'sqlite3', '5'):
            with self.subTest(value=value):
                self.assertIs(dialects.resolve(value), dialects.SQLITE)


class TestDialectRegistry(unittest.TestCase):
    """方言清单只有一处定义（dialects.ORDER），编号/菜单都从它派生，不许各写一份。"""

    def test_no_orphan_dialect(self):
        self.assertEqual(set(dialects.DIALECTS), set(dialects.ORDER))

    def test_numbered_matches_order(self):
        self.assertEqual(tuple(dialects.NUMBERED),
                         tuple(str(i) for i in range(1, len(dialects.ORDER) + 1)))
        for num, key in dialects.NUMBERED.items():
            self.assertIs(dialects.resolve(num), dialects.DIALECTS[key])

    def test_menu_lists_every_dialect(self):
        for key in dialects.ORDER:
            self.assertIn(dialects.DIALECTS[key].name, dialects.MENU)

    def test_every_alias_points_to_a_real_dialect(self):
        for alias, key in dialects.ALIASES.items():
            with self.subTest(alias=alias):
                self.assertIn(key, dialects.DIALECTS)


if __name__ == '__main__':
    unittest.main()
