# -*- coding: utf-8 -*-
import io
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock

from excel2sql.cli import EXIT_ERROR, EXIT_OK, EXIT_UNCLEAN, main

openpyxl = None
try:
    import openpyxl  # noqa: F401
except ImportError:                                   # pragma: no cover
    pass


class _Sandbox:
    """临时目录 + 无 input() 保护的非交互测试环境。"""

    def __init__(self, testcase):
        self.testcase = testcase

    def __enter__(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self._patch = mock.patch('builtins.input', side_effect=AssertionError(
            '非交互模式不应调用 input()'))
        self._patch.start()
        return self.root

    def __exit__(self, *exc):
        self._patch.stop()
        self.tmp.cleanup()
        return False


def run(argv):
    out, err = io.StringIO(), io.StringIO()
    with redirect_stdout(out), redirect_stderr(err):
        code = main(argv)
    return code, out.getvalue(), err.getvalue()


def make_csv(root: Path, name='data.csv', text='装运方式,数量\nA-1,20\nA-2,30\n') -> Path:
    p = root / name
    p.write_text(text, encoding='utf-8')
    return p


def make_xlsx(root: Path, name='data.xlsx', sheets=None) -> Path:
    sheets = sheets or [('Sheet1', [['a', 'b'], [1, 2]])]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sname, rows in sheets:
        ws = wb.create_sheet(sname)
        for row in rows:
            ws.append(row)
    p = root / name
    wb.save(p)
    return p


class TestArgumentValidation(unittest.TestCase):
    def test_version_exits_zero(self):
        with self.assertRaises(SystemExit) as ctx:
            main(['--version'])
        self.assertEqual(ctx.exception.code, 0)

    def test_header_row_zero_is_rejected(self):
        """评审第 8 条：修复前 header-row 0 会取最后一行当表头。"""
        with _Sandbox(self) as root:
            code, _, err = run([str(make_csv(root)), '--header-row', '0'])
            self.assertEqual(code, EXIT_ERROR)
            self.assertIn('--header-row', err)

    def test_header_row_negative_is_rejected(self):
        with _Sandbox(self) as root:
            code, _, _ = run([str(make_csv(root)), '--header-row', '-3'])
            self.assertEqual(code, EXIT_ERROR)

    def test_batch_size_zero_is_rejected(self):
        with _Sandbox(self) as root:
            code, _, _ = run([str(make_csv(root)), '--batch-size', '0'])
            self.assertEqual(code, EXIT_ERROR)

    def test_missing_file(self):
        code, _, err = run(['no-such-file.xlsx'])
        self.assertEqual(code, EXIT_ERROR)
        self.assertIn('文件不存在', err)


class TestNonInteractiveContract(unittest.TestCase):
    """评审第 7 条：带文件参数时不该再有任何交互。"""

    def test_end_to_end_csv(self):
        with _Sandbox(self) as root:
            src = make_csv(root)
            out = root / 'o.sql'
            code, stdout, err = run([str(src), '-o', str(out)])
            self.assertEqual(code, EXIT_OK, err)
            sql = out.read_text(encoding='utf-8')
            self.assertIn("SELECT N'A-1' AS [装运方式], N'20' AS [数量]", sql)
            self.assertIn('WITH [HARDCODE] AS', sql)
            self.assertIn('SELECT * FROM [HARDCODE];', sql)
            self.assertIn('[OK]', stdout)

    def test_default_output_name(self):
        with _Sandbox(self) as root:
            src = make_csv(root, name='orders.csv')
            code, _, _ = run([str(src)])
            self.assertEqual(code, EXIT_OK)
            self.assertTrue((root / 'orders_orders_hardcode.sql').is_file())

    def test_multi_sheet_without_sheet_flag_fails_fast(self):
        if openpyxl is None:
            self.skipTest('需要 openpyxl')
        with _Sandbox(self) as root:
            src = make_xlsx(root, sheets=[('甲', [['a'], [1]]), ('乙', [['a'], [2]])])
            code, _, err = run([str(src)])
            self.assertEqual(code, EXIT_ERROR)
            self.assertIn('-s', err)
            self.assertIn('甲', err)          # 应列出可选项
            self.assertIn('乙', err)

    def test_sheet_flag_selects_sheet(self):
        if openpyxl is None:
            self.skipTest('需要 openpyxl')
        with _Sandbox(self) as root:
            src = make_xlsx(root, sheets=[('甲', [['a'], [1]]), ('乙', [['a'], [2]])])
            out = root / 'o.sql'
            code, _, err = run([str(src), '-s', '乙', '-o', str(out)])
            self.assertEqual(code, EXIT_OK, err)
            self.assertIn('SELECT 2 AS [a]', out.read_text(encoding='utf-8'))

    def test_unknown_sheet(self):
        if openpyxl is None:
            self.skipTest('需要 openpyxl')
        with _Sandbox(self) as root:
            code, _, err = run([str(make_xlsx(root)), '-s', '不存在'])
            self.assertEqual(code, EXIT_ERROR)
            self.assertIn('找不到 sheet', err)


class TestUncleanHeader(unittest.TestCase):
    CSV = 'a,,a\n1,2,3\n4,5,6\n'

    def test_fails_without_force(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text=self.CSV)
            code, _, err = run([str(src), '-o', str(root / 'o.sql')])
            self.assertEqual(code, EXIT_UNCLEAN)
            self.assertIn('不规范', err)
            self.assertIn('--force', err)
            self.assertFalse((root / 'o.sql').exists())

    def test_force_repairs_and_continues(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text=self.CSV)
            out = root / 'o.sql'
            code, _, err = run([str(src), '--force', '-o', str(out)])
            self.assertEqual(code, EXIT_OK, err)
            sql = out.read_text(encoding='utf-8')
            self.assertIn('[a]', sql)
            self.assertIn('[col_2]', sql)
            self.assertIn('[a_2]', sql)

    def test_no_data_rows(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text='a,b\n')
            code, _, err = run([str(src)])
            self.assertEqual(code, EXIT_ERROR)
            self.assertIn('没有数据', err)


class TestOptions(unittest.TestCase):
    def test_insert_format(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text='a\n1\n2\n')
            out = root / 'o.sql'
            code, _, err = run([str(src), '--format', 'insert', '--batch-size', '1', '-o', str(out)])
            self.assertEqual(code, EXIT_OK, err)
            sql = out.read_text(encoding='utf-8')
            self.assertEqual(sql.count('INSERT INTO'), 2)

    def test_plain_wrap(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text='a\n1\n2\n')
            out = root / 'o.sql'
            run([str(src), '--wrap', 'plain', '-o', str(out)])
            sql = out.read_text(encoding='utf-8')
            self.assertNotIn('WITH ', sql)
            self.assertIn('UNION ALL', sql)

    def test_dialects_by_number_and_name(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text='a\n1\n')
            for flag, expect in (('1', '[a]'), ('mysql', '`a`'), ('3', '"a"'), ('postgresql', '"a"')):
                out = root / 'o.sql'
                code, _, err = run([str(src), '-d', flag, '-o', str(out)])
                self.assertEqual(code, EXIT_OK, err)
                self.assertIn(expect, out.read_text(encoding='utf-8'))

    def test_empty_as_null(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text='a,b\n,1\n')
            out = root / 'o.sql'
            run([str(src), '--empty-as-null', '-o', str(out)])
            self.assertIn('SELECT NULL AS [a], N\'1\' AS [b]', out.read_text(encoding='utf-8'))

    def test_utf8_sig_output(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text='列\n1\n')
            out = root / 'o.sql'
            run([str(src), '--encoding', 'utf-8-sig', '-o', str(out)])
            self.assertTrue(out.read_bytes().startswith(b'\xef\xbb\xbf'))

    def test_table_name(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text='a\n1\n')
            out = root / 'o.sql'
            run([str(src), '--table', 'TMP', '-o', str(out)])
            self.assertIn('WITH [TMP] AS', out.read_text(encoding='utf-8'))

    def test_encoding_of_output_directory_is_created(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text='a\n1\n')
            out = root / 'sub' / 'deep' / 'o.sql'
            code, _, err = run([str(src), '-o', str(out)])
            self.assertEqual(code, EXIT_OK, err)
            self.assertTrue(out.is_file())

    def test_help_exits_zero(self):
        with self.assertRaises(SystemExit) as ctx:
            with redirect_stdout(io.StringIO()):
                main(['--help'])
        self.assertEqual(ctx.exception.code, 0)


if __name__ == '__main__':
    unittest.main()
