# -*- coding: utf-8 -*-
import datetime
import tempfile
import unittest
from pathlib import Path

from excel2sql.reader import (ReadError, probe_sheets, read_sheet_rows, read_sheets, scan_dir)

openpyxl = None
try:
    import openpyxl  # noqa: F401
except ImportError:                                   # pragma: no cover
    pass


def write_csv(path: Path, text: str, encoding: str = 'utf-8') -> Path:
    path.write_text(text, encoding=encoding)
    return path


def write_xlsx(path: Path, sheets) -> Path:
    """sheets: [(name, rows)]，rows 为二维列表。"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets:
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(list(row))
    wb.save(path)
    return path


class TestCsv(unittest.TestCase):
    def test_basic(self):
        with tempfile.TemporaryDirectory() as d:
            p = write_csv(Path(d) / 'a.csv', 'a,b\n1,2\n')
            sheet = read_sheets(p)[0]
            self.assertEqual(sheet.rows[0], ['a', 'b'])
            self.assertEqual(sheet.rows[1], ['1', '2'])
            self.assertEqual(sheet.name, 'a')

    def test_gbk_encoding(self):
        with tempfile.TemporaryDirectory() as d:
            p = write_csv(Path(d) / 'gbk.csv', '装运方式,细分市场\n一级,消费者\n', encoding='gbk')
            sheet = read_sheets(p)[0]
            self.assertEqual(sheet.rows[0], ['装运方式', '细分市场'])

    def test_utf8_bom(self):
        with tempfile.TemporaryDirectory() as d:
            p = write_csv(Path(d) / 'bom.csv', 'a,b\n1,2\n', encoding='utf-8-sig')
            self.assertEqual(read_sheets(p)[0].rows[0], ['a', 'b'])   # BOM 不该留在首列

    def test_semicolon_detected(self):
        with tempfile.TemporaryDirectory() as d:
            p = write_csv(Path(d) / 'semi.csv', 'a;b;c\n1;2;3\n')
            sheet = read_sheets(p)[0]
            self.assertEqual(sheet.rows[0], ['a', 'b', 'c'])

    def test_tab_detected(self):
        with tempfile.TemporaryDirectory() as d:
            p = write_csv(Path(d) / 'tab.csv', 'a\tb\n1\t2\n')
            self.assertEqual(read_sheets(p)[0].rows[0], ['a', 'b'])

    def test_explicit_delimiter_wins(self):
        with tempfile.TemporaryDirectory() as d:
            p = write_csv(Path(d) / 'pipe.csv', 'a|b\n1|2\n')
            sheet = read_sheets(p, delimiter='|')[0]
            self.assertEqual(sheet.rows[1], ['1', '2'])

    def test_probe_gives_shape(self):
        with tempfile.TemporaryDirectory() as d:
            p = write_csv(Path(d) / 'a.csv', 'a,b\n1,2\n3,4\n')
            info = probe_sheets(p)[0]
            self.assertEqual((info.name, info.nrows, info.ncols), ('a', 3, 2))
            self.assertEqual(info.shape_text(), '3 行 x 2 列')

    def test_read_single_sheet_by_name(self):
        with tempfile.TemporaryDirectory() as d:
            p = write_csv(Path(d) / 'a.csv', 'a,b\n1,2\n')
            self.assertEqual(read_sheet_rows(p, 'a').nrows, 2)
            with self.assertRaises(ReadError):
                read_sheet_rows(p, 'nope')


@unittest.skipIf(openpyxl is None, '需要 openpyxl')
class TestExcel(unittest.TestCase):
    ROWS = [
        ['行 ID', '装运方式', '细分市场', '订购日期'],
        [40098, '一级', 2, datetime.datetime(2024, 11, 11, 0, 0, 0)],
        [26341, '二级', None, datetime.datetime(2024, 11, 13, 0, 0, 0)],
    ]

    def _book(self, d):
        return write_xlsx(Path(d) / 'wb.xlsx', [('数据', self.ROWS), ('字典', [['k', 'v'], ['a', '1']])])

    def test_read_all_sheets(self):
        with tempfile.TemporaryDirectory() as d:
            sheets = read_sheets(self._book(d))
            self.assertEqual([s.name for s in sheets], ['数据', '字典'])
            self.assertEqual(sheets[0].rows[0], ['行 ID', '装运方式', '细分市场', '订购日期'])
            self.assertEqual(sheets[0].rows[1][3], datetime.datetime(2024, 11, 11, 0, 0, 0))

    def test_empty_cell_is_none(self):
        with tempfile.TemporaryDirectory() as d:
            sheet = [s for s in read_sheets(self._book(d)) if s.name == '数据'][0]
            self.assertIsNone(sheet.rows[2][2])
            self.assertEqual(sheet.rows[2][1], '二级')

    def test_probe_is_cheap_and_lists_sheets(self):
        with tempfile.TemporaryDirectory() as d:
            infos = probe_sheets(self._book(d))
            self.assertEqual([i.name for i in infos], ['数据', '字典'])
            self.assertGreaterEqual(infos[0].nrows, 1)

    def test_read_only_selected_sheet(self):
        with tempfile.TemporaryDirectory() as d:
            sheet = read_sheet_rows(self._book(d), '字典')
            self.assertEqual(sheet.name, '字典')
            self.assertEqual(sheet.nrows, 2)

    def test_missing_sheet_raises_with_hint(self):
        with tempfile.TemporaryDirectory() as d:
            with self.assertRaises(ReadError) as ctx:
                read_sheet_rows(self._book(d), '不存在')
            self.assertIn('字典', str(ctx.exception))


class TestErrors(unittest.TestCase):
    def test_missing_file(self):
        with self.assertRaises(ReadError):
            read_sheets('definitely-not-here.xlsx')

    def test_unsupported_extension(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'a.txt'
            p.write_text('x', encoding='utf-8')
            with self.assertRaises(ReadError) as ctx:
                read_sheets(p)
            self.assertIn('不支持', str(ctx.exception))


class TestScanDir(unittest.TestCase):
    def test_lists_supported_and_skips_temp(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            (root / 'a.csv').write_text('a\n1\n', encoding='utf-8')
            (root / 'b.xlsx').write_bytes(b'PK\x03\x04')
            (root / '~$c.xlsx').write_bytes(b'PK\x03\x04')     # Excel 打开时的临时文件
            (root / 'd.txt').write_text('x', encoding='utf-8')
            (root / 'sub').mkdir()
            names = [p.name for p in scan_dir(root)]
            self.assertEqual(names, ['a.csv', 'b.xlsx'])

    def test_not_a_dir(self):
        self.assertEqual(scan_dir('does-not-exist'), [])


if __name__ == '__main__':
    unittest.main()
