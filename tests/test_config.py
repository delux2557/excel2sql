# -*- coding: utf-8 -*-
import datetime
import tempfile
import unittest
from pathlib import Path

from excel2sql import config


class TestDefaults(unittest.TestCase):
    def test_builtin_defaults(self):
        s = config.Settings()
        self.assertEqual(s.dialect, 'sqlserver')
        self.assertEqual(s.format, 'union')
        self.assertEqual(s.wrap, 'cte')
        self.assertEqual(s.encoding, 'utf-8')
        self.assertEqual(s.header_row, 'auto')
        self.assertFalse(s.ask_advanced)
        self.assertFalse(s.copy_clipboard)

    def test_progress_hint(self):
        self.assertIn('sqlserver', config.Settings().progress_hint())


class TestOutputPath(unittest.TestCase):
    SRC = Path('/data/orders.xlsx')

    def test_source_mode(self):
        s = config.Settings(output_dir='source')
        self.assertEqual(s.default_output(self.SRC, 'Sheet1'),
                         Path('/data/orders_Sheet1_hardcode.sql'))

    def test_sub_mode_is_default(self):
        s = config.Settings()
        self.assertEqual(s.output_dir, 'sub')
        self.assertEqual(s.default_output(self.SRC, '订单'),
                         Path('/data/excel2sql-out/orders_订单_hardcode.sql'))

    def test_absolute_dir_mode(self):
        s = config.Settings(output_dir='/tmp/out')
        self.assertEqual(s.default_output(self.SRC, 'S'),
                         Path('/tmp/out/orders_S_hardcode.sql'))

    def test_relative_dir_mode_resolves_against_source(self):
        s = config.Settings(output_dir='sql')
        self.assertEqual(s.default_output(self.SRC, 'S'),
                         Path('/data/sql/orders_S_hardcode.sql'))

    def test_filename_template(self):
        s = config.Settings(output_dir='source', filename='{date}-{name}-{sheet}.sql')
        p = s.default_output(self.SRC, 'S')
        self.assertTrue(p.name.endswith('-orders-S.sql'))
        self.assertTrue(p.name.startswith(datetime.date.today().isoformat()))

    def test_sheet_name_sanitized(self):
        s = config.Settings(output_dir='source', filename='{sheet}.sql')
        self.assertEqual(s.default_output(self.SRC, 'a/b:c*').name, 'a_b_c.sql')

    def test_suffix_appended_when_missing(self):
        s = config.Settings(output_dir='source', filename='out')
        self.assertTrue(s.default_output(self.SRC, 'S').name.endswith('.sql'))


class TestUniquePath(unittest.TestCase):
    def test_no_collision(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'a.sql'
            self.assertEqual(config.unique_path(p), p)

    def test_adds_counter(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'a.sql'
            p.write_text('x', encoding='utf-8')
            self.assertEqual(config.unique_path(p).name, 'a-2.sql')
            (Path(d) / 'a-2.sql').write_text('x', encoding='utf-8')
            self.assertEqual(config.unique_path(p).name, 'a-3.sql')


class TestLoad(unittest.TestCase):
    def test_no_config_found(self):
        with tempfile.TemporaryDirectory() as d:
            settings, path, warnings = config.load(cwd=Path(d))
            # 若用户级 ~/.excel2sql/config.ini 不存在，则完全使用内置默认值
            if path is None:
                self.assertEqual(settings, config.Settings())
            self.assertEqual(warnings, [])

    def test_reads_values_and_types(self):
        with tempfile.TemporaryDirectory() as d:
            ini = Path(d) / 'excel2sql.ini'
            ini.write_text('[output]\ndialect = oracle\nbatch_size = 42\n\n'
                           '[ui]\nask_advanced = YES\ncopy_clipboard = off\n',
                           encoding='utf-8')
            settings, path, warnings = config.load(cwd=Path(d))
            self.assertEqual(path, ini)
            self.assertEqual(settings.dialect, 'oracle')
            self.assertEqual(settings.batch_size, 42)
            self.assertIs(settings.ask_advanced, True)
            self.assertIs(settings.copy_clipboard, False)
            self.assertEqual(warnings, [])

    def test_unknown_keys_ignored(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / 'excel2sql.ini').write_text('[x]\nwhatever = 1\n', encoding='utf-8')
            settings, _, warnings = config.load(cwd=Path(d))
            self.assertEqual(settings.dialect, 'sqlserver')
            self.assertEqual(warnings, [])

    def test_bad_bool_warns(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / 'excel2sql.ini').write_text('[ui]\ncopy_clipboard = maybe\n',
                                                   encoding='utf-8')
            _, _, warnings = config.load(cwd=Path(d))
            self.assertEqual(len(warnings), 1)
            self.assertIn('copy_clipboard', warnings[0])

    def test_bad_int_warns_and_keeps_default(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d) / 'excel2sql.ini').write_text('[output]\nbatch_size = abc\n', encoding='utf-8')
            settings, _, warnings = config.load(cwd=Path(d))
            self.assertEqual(settings.batch_size, 500)
            self.assertEqual(len(warnings), 1)

    def test_explicit_path_missing_raises(self):
        with self.assertRaises(config.ConfigError):
            config.load(explicit='/definitely/not/here.ini')

    def test_broken_ini_raises(self):
        with tempfile.TemporaryDirectory() as d:
            bad = Path(d) / 'bad.ini'
            bad.write_text('this is not = = ini\n[[[', encoding='utf-8')
            with self.assertRaises(config.ConfigError):
                config.load(explicit=str(bad))


class TestTemplate(unittest.TestCase):
    def test_write_template(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'sub' / 'excel2sql.ini'
            written = config.write_template(p)
            self.assertEqual(written, p)
            text = p.read_text(encoding='utf-8')
            for token in ('[output]', '[data]', '[ui]', 'dialect', 'output_dir', 'header_row'):
                self.assertIn(token, text)

    def test_template_is_loadable_and_matches_defaults(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'excel2sql.ini'
            config.write_template(p)
            settings, _, warnings = config.load(explicit=str(p))
            self.assertEqual(warnings, [])
            self.assertEqual(settings, config.Settings())      # 模板值 == 内置默认值

    def test_does_not_overwrite_by_default(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'excel2sql.ini'
            p.write_text('keep me', encoding='utf-8')
            self.assertIsNone(config.write_template(p))
            self.assertEqual(p.read_text(encoding='utf-8'), 'keep me')


if __name__ == '__main__':
    unittest.main()
