# -*- coding: utf-8 -*-
"""2026-09-25 VM 端到端实测发现的缺陷 —— 回归用例。

每个类对应实测报告里的一个问题（报告与夹具在私有目录 `数据库测试/`，不随仓库分发）。
这些用例的作用是：**任何一次改动只要让其中任何一个失败，就说明老问题复活了。**

★ 这里全部是「退出码为 0 的静默丢数据」，所以断言必须落到**行数 / 列数 / 列名**上，
  只看退出码是抓不住的 —— 这正是当初它们能溜过去的原因。
"""
import io
import os
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from unittest import mock

from excel2sql import headers, sqlgen
from excel2sql.cli import EXIT_OK, EXIT_UNCLEAN, main, split_header
from excel2sql.reader import Sheet

openpyxl = None
try:
    import openpyxl  # noqa: F401
except ImportError:                                   # pragma: no cover
    pass


def run(argv):
    out, err = io.StringIO(), io.StringIO()
    with redirect_stdout(out), redirect_stderr(err):
        code = main(argv)
    return code, out.getvalue(), err.getvalue()


class _Dir:
    """临时目录 + 阻断 input() 的非交互环境（触发交互会直接失败而不是挂住）。"""

    def __enter__(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self._cwd = Path.cwd()
        os.chdir(str(self.root))              # 顺带隔离配置文件查找
        self._patch = mock.patch('builtins.input', side_effect=AssertionError(
            '非交互模式不应调用 input()'))
        self._patch.start()
        return self.root

    def __exit__(self, *exc):
        self._patch.stop()
        os.chdir(str(self._cwd))
        self.tmp.cleanup()
        return False


def write_csv(root, name, text):
    p = root / name
    p.write_text(text, encoding='utf-8')
    return p


def write_xlsx(root, name, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    p = root / name
    wb.save(p)
    return p


# ============================================================ 问题 1（P0）
class TestHeaderRowNeverSilentlyDropsRows(unittest.TestCase):
    """表头行含**空列名**或**纯数字列名**时，表头被选错 -> 静默丢一行数据。

    实测得分：真表头 3.75 vs 数据行 3.95。原因是"位置靠前"的偏好只有 0.05/行，
    压不住"有一个空列名"造成的 0.5×0.5=0.25 分损失，于是必然翻转。
    修复前产物是 `[OK] 1 行 x 2 列`（源文件 2 行数据），**退出码 0**。
    """

    def test_blank_column_name_in_header(self):
        rows = [['装运方式', None], ['一级', '消费者'], ['二级', '公司']]
        row, reason = headers.detect(rows)
        self.assertEqual(row, 1, reason)

    def test_numeric_column_name_in_header(self):
        # 这一例的分差高达 2.2，单靠"分差阈值"拦不住，靠的是"被跳过的行不止一个非空格"
        rows = [['装运方式', 2024], ['一级', '消费者'], ['二级', '公司']]
        row, reason = headers.detect(rows)
        self.assertEqual(row, 1, reason)

    def test_normal_header_still_row_one(self):
        rows = [['装运方式', '细分市场'], ['一级', '消费者'], ['二级', '公司']]
        self.assertEqual(headers.detect(rows)[0], 1)

    # ---- 下面两个是"保护不能过头"的对照 ----

    def test_sparse_title_row_is_still_skippable(self):
        """标题行只有 1 个非空格子 —— 仍然应该正常跳过。"""
        rows = [['某部门 2024 年度 数据导出表', None, None],
                ['装运方式', '细分市场', '销售额'],
                ['一级', '消费者', 221.98],
                ['二级', '公司', 3709.39]]
        row, reason = headers.detect(rows)
        self.assertEqual(row, 2, reason)

    def test_single_column_title_is_still_skippable(self):
        """单列表的标题行"占比"也是 100%，不能被判成"填满"。"""
        rows = [['销售报表'], [], ['行ID'], ['1'], ['2']]
        self.assertEqual(headers.detect(rows)[0], 3)

    @unittest.skipIf(openpyxl is None, '需要 openpyxl')
    def test_cli_end_to_end_keeps_all_rows(self):
        with _Dir() as root:
            src = write_xlsx(root, 'b.xlsx',
                             [['装运方式', None], ['一级', '消费者'], ['二级', '公司']])
            code, out, err = run([str(src), '--force', '-o', str(root / 'o.sql')])
            self.assertEqual(code, EXIT_OK, err)
            self.assertIn('[OK] 2 行 x 2 列', out)          # 修复前是 1 行


# ============================================================ 问题 2（P0）
class TestTailColumnOnlyDroppedWhenEmpty(unittest.TestCase):
    """表头为空白的**尾列**被整列裁掉，且不检查该列有没有数据。

    实测铁证：`末尾空列名 + 有数据` 与 `末尾空列名 + 无数据` 两份产物**逐字节相同**
    —— 说明旧实现只看表头是否为空，压根没看数据。
    """

    @staticmethod
    def _sheet(rows):
        return Sheet('s', rows, max(len(r) for r in rows))

    def test_tail_column_with_data_is_kept(self):
        sheet = self._sheet([['装运方式', '细分市场', None],
                             ['一级', '消费者', 111],
                             ['二级', '公司', 222]])
        header, rows = split_header(sheet, 1)
        self.assertEqual(len(header), 3)
        self.assertEqual([r[2] for r in rows], [111, 222])

    def test_tail_column_without_data_is_dropped(self):
        sheet = self._sheet([['装运方式', '细分市场', None],
                             ['一级', '消费者', None],
                             ['二级', '公司', None]])
        header, _ = split_header(sheet, 1)
        self.assertEqual(len(header), 2)

    def test_the_two_cases_differ_now(self):
        """核心回归：这两种输入的产物以前是完全一样的。"""
        with_data = self._sheet([['a', 'b', None], ['x', 'y', 1]])
        without = self._sheet([['a', 'b', None], ['x', 'y', None]])
        self.assertNotEqual(split_header(with_data, 1)[0],
                            split_header(without, 1)[0])

    def test_middle_blank_column_is_kept(self):
        sheet = self._sheet([['装运方式', None, '销售额'],
                             ['一级', '消费者', 221.98]])
        header, _ = split_header(sheet, 1)
        self.assertEqual(len(header), 3)

    def test_explicit_header_row_also_keeps_the_data(self):
        """旧版连 --header-row 1 都救不了（列照样被裁），这里一并守住。"""
        sheet = self._sheet([['装运方式', '细分市场', None],
                             ['一级', '消费者', 111]])
        header, _ = split_header(sheet, 1)
        self.assertEqual(len(header), 3)


# ============================================================ 问题 3（P1）
class TestCsvTypeInference(unittest.TestCase):
    """CSV 没有类型信息 -> 所有列都被字符串化，`union` 产物彻底丢类型。

    实测影响：SQL Server 落成 `nvarchar`、PG 落成 `text`，`SUM()` 三库全部报错
    （SS `Msg 8117` / PG `function sum(text) does not exist`）。
    """

    CSV = '数量,金额,订购日期\n2,221.98,2024-11-11\n9,3709.39,2024-02-05\n'

    def _sql(self, extra):
        with _Dir() as root:
            src = write_csv(root, 'n.csv', self.CSV)
            code, out, err = run([str(src)] + extra + ['-o', str(root / 'o.sql')])
            self.assertEqual(code, EXIT_OK, err)
            return (root / 'o.sql').read_text(encoding='utf-8')

    def test_default_stays_text(self):
        """默认行为不变：CSV 一律按字符串输出（保守，避免隐式转换意外）。"""
        sql = self._sql([])
        self.assertIn("N'2' AS [数量]", sql)

    def test_infer_types_restores_numeric_literals(self):
        sql = self._sql(['--infer-types'])
        self.assertIn('2 AS [数量]', sql)
        self.assertIn('221.98 AS [金额]', sql)
        self.assertNotIn("N'2'", sql)

    def test_infer_types_keeps_leading_zero_column_as_text(self):
        """'007' 转成 7 会丢信息 —— 整列保持文本，不做逐格猜测。"""
        with _Dir() as root:
            src = write_csv(root, 'z.csv', '编号\n007\n000123\n')
            code, _, err = run([str(src), '--infer-types', '-o', str(root / 'o.sql')])
            self.assertEqual(code, EXIT_OK, err)
            sql = (root / 'o.sql').read_text(encoding='utf-8')
            self.assertIn("N'007'", sql)
            self.assertIn("N'000123'", sql)

    def test_infer_types_leaves_mixed_column_alone(self):
        with _Dir() as root:
            src = write_csv(root, 'm.csv', '值\n1\nabc\n')
            code, _, err = run([str(src), '--infer-types', '-o', str(root / 'o.sql')])
            self.assertEqual(code, EXIT_OK, err)
            sql = (root / 'o.sql').read_text(encoding='utf-8')
            self.assertIn("N'1'", sql)


class TestCoerceNumericColumns(unittest.TestCase):
    """`--infer-types` 的核心函数（纯函数，边界单独钉死）。"""

    def test_converts_whole_numeric_column(self):
        rows = sqlgen.coerce_numeric_columns([['2', '221.98'], ['9', '3709.39']])
        self.assertEqual(rows, [[2, 221.98], [9, 3709.39]])

    def test_leading_zero_stays_text(self):
        rows = sqlgen.coerce_numeric_columns([['007'], ['000123']])
        self.assertEqual(rows, [['007'], ['000123']])

    def test_mixed_column_untouched(self):
        rows = sqlgen.coerce_numeric_columns([['1'], ['abc']])
        self.assertEqual(rows, [['1'], ['abc']])

    def test_nan_and_inf_untouched(self):
        rows = sqlgen.coerce_numeric_columns([['NaN'], ['inf']])
        self.assertEqual(rows, [['NaN'], ['inf']])

    def test_blank_becomes_null_in_numeric_column(self):
        """数字列里的空白 -> NULL（空字符串塞不进数字列）。"""
        rows = sqlgen.coerce_numeric_columns([['1'], [''], ['3']])
        self.assertEqual(rows, [[1], [None], [3]])

    def test_non_text_values_untouched(self):
        rows = sqlgen.coerce_numeric_columns([[1], [2]])
        self.assertEqual(rows, [[1], [2]])


# ============================================================ 新增开关
class TestStrictHeader(unittest.TestCase):
    """`--strict-header`：让脚本/CI 在"表头被自动跳过"时 fail fast，而不是默默接受。"""

    TITLED = '销售报表\n\n行ID,数量\n1,20\n'

    def test_rejects_when_auto_detect_skips_rows(self):
        with _Dir() as root:
            src = write_csv(root, 't.csv', self.TITLED)
            code, _, err = run([str(src), '--strict-header', '-o', str(root / 'o.sql')])
            self.assertEqual(code, EXIT_UNCLEAN, err)
            self.assertIn('--strict-header', err)

    def test_accepts_normal_header(self):
        with _Dir() as root:
            src = write_csv(root, 't.csv', '装运方式,细分市场\n一级,消费者\n')
            code, _, err = run([str(src), '--strict-header', '-o', str(root / 'o.sql')])
            self.assertEqual(code, EXIT_OK, err)

    def test_explicit_header_row_satisfies_strict(self):
        with _Dir() as root:
            src = write_csv(root, 't.csv', self.TITLED)
            code, _, err = run([str(src), '--strict-header', '--header-row', '3',
                                '-o', str(root / 'o.sql')])
            self.assertEqual(code, EXIT_OK, err)


if __name__ == '__main__':                              # pragma: no cover
    unittest.main()
