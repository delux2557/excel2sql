# -*- coding: utf-8 -*-
import datetime
import unittest

from excel2sql import headers


class TestCheck(unittest.TestCase):
    DATA = [[1, 2, 3]]

    def test_clean_header_passes(self):
        v = headers.check(['a', 'b', 'c'], self.DATA)
        self.assertTrue(v.ok, v.issues)

    def test_blank_column(self):
        v = headers.check(['a', None, 'c'], self.DATA)
        self.assertFalse(v.ok)
        self.assertTrue(any('为空' in m for m in v.issues))

    def test_whitespace_only_column(self):
        v = headers.check(['a', '   ', 'c'], self.DATA)
        self.assertFalse(v.ok)

    def test_duplicate_columns_case_insensitive(self):
        v = headers.check(['A', 'a'], self.DATA)
        self.assertFalse(v.ok)
        self.assertTrue(any('重复' in m for m in v.issues))

    def test_header_looks_like_data(self):
        v = headers.check([1, 2, 3], self.DATA)
        self.assertTrue(any('不像字段名' in m for m in v.issues))

    def test_header_dates_look_like_data(self):
        v = headers.check([datetime.date(2026, 1, 1), '2026/02/03'], self.DATA)
        self.assertTrue(any('不像字段名' in m for m in v.issues))

    def test_bool_header_is_not_treated_as_number(self):
        """评审小问题 3：bool 是 int 子类，不应被当成数字。"""
        v = headers.check([True, False], self.DATA)
        self.assertFalse(any('不像字段名' in m for m in v.issues))

    def test_no_data_rows(self):
        v = headers.check(['a', 'b'], [])
        self.assertFalse(v.ok)
        self.assertTrue(any('没有' in m for m in v.issues))

    def test_empty_header(self):
        v = headers.check([], self.DATA)
        self.assertFalse(v.ok)


class TestRepair(unittest.TestCase):
    def test_blank_columns_named_by_position(self):
        self.assertEqual(headers.repair([None, None]), ['col_1', 'col_2'])

    def test_duplicates_get_suffix(self):
        self.assertEqual(headers.repair(['a', 'a', 'a']), ['a', 'a_2', 'a_3'])

    def test_repair_result_is_always_unique(self):
        """评审小问题 1：a, a, a_2 修完不能再产生新的重复。"""
        fixed = headers.repair(['a', 'a', 'a_2'])
        self.assertEqual(len(set(n.lower() for n in fixed)), len(fixed))
        self.assertEqual(fixed, ['a', 'a_2', 'a_2_2'])

    def test_repair_result_never_empty(self):
        """评审小问题 2：不应出现 col_2_2 这类来源不明的名字。"""
        fixed = headers.repair([None, '', 'x'])
        self.assertEqual(fixed[0], 'col_1')
        self.assertEqual(fixed[1], 'col_2')
        self.assertTrue(all(n.strip() for n in fixed))

    def test_keeps_numeric_header_as_text(self):
        self.assertEqual(headers.repair([2024, 2025]), ['2024', '2025'])

    def test_strips_whitespace(self):
        self.assertEqual(headers.repair(['  a  ', 'b']), ['a', 'b'])


class TestDetect(unittest.TestCase):
    """表头行自动识别（交互/批处理都用它给出行号建议）。"""

    def test_header_on_first_row(self):
        rows = [['序号', '装运方式', '数量'], [1, 'A-1', 25], [2, '二级', 30]]
        row, reason = headers.detect(rows)
        self.assertEqual(row, 1)
        self.assertIn('第 1 行', reason)

    def test_skips_title_and_blank_rows(self):
        rows = [['销售报表'], [], ['行ID', '数量'], ['1', '20'], ['2', '30']]
        row, reason = headers.detect(rows)
        self.assertEqual(row, 3)
        self.assertIn('第 3 行', reason)

    def test_numeric_only_header_falls_back(self):
        rows = [[1, 2, 3], [4, 5, 6]]
        row, reason = headers.detect(rows)
        self.assertEqual(row, 1)
        self.assertIn('默认第 1 行', reason)

    def test_empty_sheet(self):
        self.assertEqual(headers.detect([])[0], 1)

    def test_score_prefers_text_with_data_below(self):
        header = ['订单号', '客户', '金额']
        data = ['CA-2014-1', '张三', 100.5]
        self.assertGreater(headers.score_header_row(header, data),
                           headers.score_header_row(data, ['CA-2014-2', '李四', 88.0]))

    def test_score_blank_row_is_very_low(self):
        self.assertLess(headers.score_header_row([None, None], ['a', 'b']), 0)


class TestLooksLikeValue(unittest.TestCase):
    def test_numbers_and_dates(self):
        self.assertTrue(headers.looks_like_value(12))
        self.assertTrue(headers.looks_like_value(3.14))
        self.assertTrue(headers.looks_like_value('007'))
        self.assertTrue(headers.looks_like_value('2026-09-08'))
        self.assertTrue(headers.looks_like_value('2026/09/08 10:30:00'))

    def test_text_and_bool(self):
        self.assertFalse(headers.looks_like_value('订单号'))
        self.assertFalse(headers.looks_like_value('一级'))
        self.assertFalse(headers.looks_like_value(True))
        self.assertFalse(headers.looks_like_value('2026年'))

    def test_is_blank(self):
        self.assertTrue(headers.is_blank(None))
        self.assertTrue(headers.is_blank('   '))
        self.assertFalse(headers.is_blank(0))
        self.assertFalse(headers.is_blank('0'))


if __name__ == '__main__':
    unittest.main()
