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
        rows = [['行 ID', '装运方式', '细分市场'], [1, '一级', 2], [2, '二级', 9]]
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
        self.assertFalse(headers.looks_like_value('AB-100151402'))
        self.assertFalse(headers.looks_like_value(True))
        self.assertFalse(headers.looks_like_value('2026年'))


class TestCleanName(unittest.TestCase):
    """列名清洗：不可见字符删除、空白折叠成单空格。

    起因（2026-10-14 实测）：列名里的换行会把产物头注释劈成两行，后半行失去
    `--` 前缀被 SQL 引擎当语句执行 —— 产物直接 `syntax error`。
    """

    def test_newline_becomes_single_space(self):
        self.assertEqual(headers.clean_name('订\n单'), '订 单')
        self.assertEqual(headers.clean_name('订单\r\n编号'), '订单 编号')
        self.assertEqual(headers.clean_name('单\r价'), '单 价')

    def test_tab_and_vertical_whitespace_become_space(self):
        self.assertEqual(headers.clean_name('a\tb'), 'a b')
        self.assertEqual(headers.clean_name('a\vb'), 'a b')
        self.assertEqual(headers.clean_name('a\fb'), 'a b')

    def test_surrounding_whitespace_is_trimmed(self):
        self.assertEqual(headers.clean_name(' 数量 '), '数量')
        self.assertEqual(headers.clean_name('\n金额\n'), '金额')
        self.assertEqual(headers.clean_name('\t备注\t'), '备注')

    def test_unicode_spaces_become_plain_space(self):
        """NBSP / 全角空格也是空白，但肉眼与普通空格无异 —— 必须归一。"""
        self.assertEqual(headers.clean_name('数量\u00a0(件)'), '数量 (件)')
        self.assertEqual(headers.clean_name('前\u3000后'), '前 后')
        self.assertEqual(headers.clean_name('a\u2028b'), 'a b')

    def test_runs_of_whitespace_collapse_to_one(self):
        self.assertEqual(headers.clean_name('多   空格'), '多 空格')
        self.assertEqual(headers.clean_name('a\n\n  b'), 'a b')

    def test_zero_width_chars_are_deleted_not_spaced(self):
        """零宽字符不是空白（`isspace()` 为 False），是排版噪声 —— 直接删。

        若折成空格，`x\\u200by` 会变成 `x y`（凭空多一个词），删掉才是 `xy`。
        """
        self.assertEqual(headers.clean_name('x\u200by'), 'xy')
        self.assertEqual(headers.clean_name('\u200b'), '')
        self.assertEqual(headers.clean_name('\ufeff'), '')       # BOM
        self.assertEqual(headers.clean_name('\u00ad'), '')       # 软连字符

    def test_control_chars_are_deleted(self):
        self.assertEqual(headers.clean_name('a\x00b'), 'ab')
        self.assertEqual(headers.clean_name('a\x07b'), 'ab')
        self.assertEqual(headers.clean_name('a\x7fb'), 'ab')

    def test_emoji_presentation_selector_is_kept(self):
        """变体选择符修饰前一个字符的呈现方式，属于**内容**，不能删
        （删了会把 🗃️ 变成单色字形）。"""
        self.assertEqual(headers.clean_name('\U0001f5c3\ufe0f'), '\U0001f5c3\ufe0f')

    def test_clean_input_is_untouched(self):
        for name in ('单价(元)', '2024 年', '订单编号', 'a b', '客户 ID', 'col_2'):
            with self.subTest(name=name):
                self.assertEqual(headers.clean_name(name), name)

    def test_result_never_contains_line_breaks(self):
        dirty = ['订\n单', 'a\r\nb', 'a\tb', 'a\vb', 'a\fb', 'x\u0085y', 'a\u2028b']
        for name in dirty:
            with self.subTest(name=name):
                cleaned = headers.clean_name(name)
                self.assertFalse(any(c in cleaned for c in '\r\n\t\v\f\u0085\u2028\u2029'),
                                 '{!r} -> {!r}'.format(name, cleaned))

    def test_is_idempotent(self):
        for name in ('订\n单', ' 数量 ', '\u200b', 'a\u00a0b', 'x\u200by'):
            with self.subTest(name=name):
                once = headers.clean_name(name)
                self.assertEqual(headers.clean_name(once), once)

    def test_none_and_non_strings(self):
        self.assertEqual(headers.clean_name(None), '')
        self.assertEqual(headers.clean_name(2024), '2024')
        self.assertEqual(headers.clean_name(3.5), '3.5')


class TestChangedNames(unittest.TestCase):
    """清洗必须"看得见" —— 静默修正和静默走偏只差一步。"""

    def test_reports_before_and_after(self):
        out = headers.changed_names(['订\n单', '数量'], ['订 单', '数量'])
        self.assertEqual(len(out), 1)
        self.assertIn('第1列', out[0])
        self.assertIn('订 单', out[0])

    def test_silent_when_nothing_changed(self):
        self.assertEqual(headers.changed_names(['a', 'b'], ['a', 'b']), [])

    def test_truncates_and_counts_the_rest(self):
        raw = ['a\n', 'b\n', 'c\n', 'd\n', 'e\n', 'f\n', 'g\n']
        cleaned = ['a', 'b', 'c', 'd', 'e', 'f', 'g']
        out = headers.changed_names(raw, cleaned, limit=3)
        self.assertEqual(len(out), 4)
        self.assertIn('另有 4 处', out[-1])

    def test_none_counts_as_change(self):
        self.assertEqual(len(headers.changed_names([None], ['col_1'])), 1)


class TestIsBlank(unittest.TestCase):
    def test_none_and_whitespace(self):
        self.assertTrue(headers.is_blank(None))
        self.assertTrue(headers.is_blank('   '))
        self.assertTrue(headers.is_blank('\n'))
        self.assertTrue(headers.is_blank('\u00a0'))      # NBSP
        self.assertTrue(headers.is_blank('\u3000'))      # 全角空格

    def test_zero_width_counts_as_blank(self):
        """★ 修复点：零宽字符/BOM 原先被判为"非空"，于是列名变成看不见的字符，
        既得不到 col_N 补名，也不会被报成空列（2026-10-14 实测）。"""
        for v in ('\u200b', '\ufeff', '\u200c', '\u00ad', '\x00'):
            with self.subTest(value=repr(v)):
                self.assertTrue(headers.is_blank(v))

    def test_zero_and_false_are_not_blank(self):
        self.assertFalse(headers.is_blank(0))
        self.assertFalse(headers.is_blank('0'))
        self.assertFalse(headers.is_blank(False))


if __name__ == '__main__':
    unittest.main()
