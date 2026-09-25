# -*- coding: utf-8 -*-
import io
import os
import sys
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
        self._cwd = Path.cwd()
        os.chdir(str(self.root))            # 让配置文件查找也落在临时目录里
        self._patch = mock.patch('builtins.input', side_effect=AssertionError(
            '非交互模式不应调用 input()'))
        self._patch.start()
        return self.root

    def __exit__(self, *exc):
        self._patch.stop()
        os.chdir(str(self._cwd))
        self.tmp.cleanup()
        return False


def run(argv):
    out, err = io.StringIO(), io.StringIO()
    with redirect_stdout(out), redirect_stderr(err):
        code = main(argv)
    return code, out.getvalue(), err.getvalue()


def make_csv(root: Path, name='data.csv', text='装运方式,细分市场\n一级,消费者\n二级,公司\n') -> Path:
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
            self.assertIn("SELECT N'一级' AS [装运方式], N'消费者' AS [细分市场]", sql)
            self.assertIn('WITH [HARDCODE] AS', sql)
            self.assertIn('SELECT * FROM [HARDCODE];', sql)
            self.assertIn('[OK]', stdout)

    def test_default_output_goes_to_subdir(self):
        """默认输出到源文件目录下的 excel2sql-out/ 子目录，不污染数据目录。"""
        with _Sandbox(self) as root:
            src = make_csv(root, name='orders.csv')
            code, _, err = run([str(src)])
            self.assertEqual(code, EXIT_OK, err)
            self.assertTrue((root / 'excel2sql-out' / 'orders_orders_hardcode.sql').is_file())

    def test_default_output_avoids_collision(self):
        with _Sandbox(self) as root:
            src = make_csv(root, name='orders.csv')
            run([str(src)])
            run([str(src)])
            files = sorted(p.name for p in (root / 'excel2sql-out').iterdir())
            self.assertEqual(files, ['orders_orders_hardcode-2.sql', 'orders_orders_hardcode.sql'])

    def test_explicit_out_overwrites(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text='a\n1\n')
            out = root / 'o.sql'
            run([str(src), '-o', str(out)])
            run([str(src), '-o', str(out)])          # 显式指定就按原样写
            self.assertTrue(out.is_file())
            self.assertFalse((root / 'o-2.sql').exists())

    def test_header_row_auto_detected(self):
        """前两行是标题/空行时，自动识别应跳到真正的表头行。"""
        with _Sandbox(self) as root:
            src = make_csv(root, text='销售报表\n\n行ID,数量\n1,20\n')
            out = root / 'o.sql'
            code, _, err = run([str(src), '-o', str(out)])
            self.assertEqual(code, EXIT_OK, err)
            sql = out.read_text(encoding='utf-8')
            self.assertIn('[行ID]', sql)
            self.assertNotIn('销售报表', sql)
            self.assertIn('表头自动识别', err)

    def test_header_row_explicit_wins(self):
        with _Sandbox(self) as root:
            src = make_csv(root, text='销售报表\n\n行ID,数量\n1,20\n')
            out = root / 'o.sql'
            code, _, err = run([str(src), '-o', str(out), '--header-row', '1'])
            self.assertEqual(code, EXIT_OK, err)
            self.assertIn('销售报表', out.read_text(encoding='utf-8'))

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
            # 空串 -> NULL；b 列是纯数字，CSV 源下被 auto 推断还原成数字字面量
            self.assertIn('SELECT NULL AS [a], 1 AS [b]', out.read_text(encoding='utf-8'))

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


class TestConfigWiring(unittest.TestCase):
    """配置文件（excel2sql.ini）应作为默认值来源，命令行参数可覆盖。"""

    INI = """[output]
dialect = mysql
format = insert
table = TMP
batch_size = 7
output_dir = source
encoding = utf-8-sig
filename = {name}_out.sql

[data]
empty_as_null = true
header_row = 1
"""

    def test_config_file_drives_defaults(self):
        with _Sandbox(self) as root:
            (root / 'excel2sql.ini').write_text(self.INI, encoding='utf-8')
            src = make_csv(root, text='a,b\n,1\n')
            code, _, err = run([str(src)])
            self.assertEqual(code, EXIT_OK, err)
            out = root / 'data_out.sql'                   # filename + output_dir=source
            self.assertTrue(out.is_file(), err)
            sql = out.read_text(encoding='utf-8-sig')
            self.assertIn('INSERT INTO `TMP`', sql)        # dialect + table
            # empty_as_null（空串->NULL）+ CSV 默认 auto 推断（数字列还原成数字）
            self.assertIn('(NULL, 1)', sql)

    def test_cli_overrides_config(self):
        with _Sandbox(self) as root:
            (root / 'excel2sql.ini').write_text(self.INI, encoding='utf-8')
            src = make_csv(root, text='a\n1\n')
            out = root / 'cli.sql'
            code, _, err = run([str(src), '-o', str(out), '-d', 'sqlserver', '--format', 'union'])
            self.assertEqual(code, EXIT_OK, err)
            sql = out.read_text(encoding='utf-8')
            self.assertIn('WITH [TMP] AS', sql)            # table 仍来自配置
            self.assertIn('AS [a]', sql)                   # -d sqlserver 覆盖了配置的 mysql
            self.assertNotIn('`', sql)
            self.assertNotIn('INSERT INTO', sql)           # --format union 覆盖了 insert

    def test_explicit_config_path(self):
        with _Sandbox(self) as root:
            cfg = root / 'custom' / 'my.ini'
            cfg.parent.mkdir(parents=True)
            cfg.write_text(self.INI, encoding='utf-8')
            src = make_csv(root, text='a\n1\n')
            out = root / 'o.sql'
            code, _, err = run([str(src), '--config', str(cfg), '-o', str(out),
                                '--format', 'union'])
            self.assertEqual(code, EXIT_OK, err)
            self.assertIn('WITH', out.read_text(encoding='utf-8'))

    def test_missing_config_path_fails(self):
        with _Sandbox(self) as root:
            code, _, err = run([str(make_csv(root)), '--config', str(root / 'nope.ini')])
            self.assertEqual(code, EXIT_ERROR)
            self.assertIn('配置', err)
            self.assertIn('不存在', err)

    def test_init_config_writes_template(self):
        with _Sandbox(self) as root:
            code, stdout, err = run(['--init-config'])
            self.assertEqual(code, EXIT_OK, err)
            ini = root / 'excel2sql.ini'
            self.assertTrue(ini.is_file())
            text = ini.read_text(encoding='utf-8')
            self.assertIn('dialect = sqlserver', text)
            self.assertIn('[output]', text)

    def test_broken_value_warns_but_continues(self):
        with _Sandbox(self) as root:
            (root / 'excel2sql.ini').write_text('[ui]\nask_advanced = maybe\n', encoding='utf-8')
            code, _, err = run([str(make_csv(root)), '-o', str(root / 'o.sql')])
            self.assertEqual(code, EXIT_OK, err)
            self.assertIn('ask_advanced', err)


class TestConsoleEncoding(unittest.TestCase):
    """Windows 上 stdout 走管道时默认是 ANSI 码页（en-US CI 即 cp1252），
    中文提示不能把它打成 UnicodeEncodeError（2026-09-25 CI 真实翻车点）。"""

    def _run(self, argv):
        """在 cp1252 的 stdout/stderr 下跑一次 CLI，返回 (退出码, 原始字节)。"""
        buf = io.BytesIO()
        stream = io.TextIOWrapper(buf, encoding='cp1252', errors='strict', newline='')
        with _Sandbox(self) as root:
            src = make_csv(root, name='d.csv', text='装运方式,细分市场\n一级,消费者\n')
            old = sys.stdout, sys.stderr
            sys.stdout = sys.stderr = stream
            try:
                code = main([str(src)] + argv)
            finally:
                sys.stdout, sys.stderr = old
                stream.flush()
            produced = (root / 'o.sql').is_file()
        return code, buf.getvalue(), produced

    def test_cp1252_stdout_does_not_crash(self):
        code, raw, produced = self._run(['-o', 'o.sql'])
        self.assertEqual(code, EXIT_OK)
        self.assertTrue(produced)
        # 管道场景应切到 UTF-8，中文照常输出
        self.assertIn('行 x', raw.decode('utf-8'))

    def test_help_survives_cp1252(self):
        """--help 的正文是中文，也必须能打出来。"""
        with self.assertRaises(SystemExit) as ctx:
            self._run(['--help'])
        self.assertEqual(ctx.exception.code, 0)


if __name__ == '__main__':
    unittest.main()
