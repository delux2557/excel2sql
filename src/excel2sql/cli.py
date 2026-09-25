"""命令行入口：交互式向导 + 可脚本化的非交互模式。"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List, Optional, Sequence

from . import __version__
from .clipboard import copy_text
from .dialects import DIALECTS, NUMBERED, resolve
from .headers import check as check_header, repair
from .reader import (ReadError, Sheet, probe_sheets, read_sheet_rows, scan_dir)
from .sqlgen import SqlGenError, SqlOptions, UNION_ROW_WARN, build_sql, prepare, write_sql

EXIT_OK, EXIT_ERROR, EXIT_UNCLEAN = 0, 1, 2
# 超过这个行数，硬编码 SQL 基本已经不适合，除非用户明确指定
HARD_LIMIT_ROWS = 200_000

BANNER = """{sep}
  excel2sql  |  Excel / CSV  ->  硬编码 SQL（UNION ALL / INSERT）
{sep}""".format(sep='=' * 60)


# ------------------------------------------------------------------ 输入助手
def ask(prompt: str, default: str = '') -> str:
    hint = ' [{}]'.format(default) if default else ''
    answer = input('{}{}: '.format(prompt, hint)).strip()
    return answer or default


def pick(items: Sequence, title: str, render, allow_path: bool = False):
    """编号选择：返回 int 下标；allow_path 时可能返回 ('file'|'dir', 路径)。"""
    while True:
        print('\n' + title)
        for i, it in enumerate(items, 1):
            print('  {:>2}. {}'.format(i, render(it)))
        raw = input('输入编号{} (q 退出): '.format(' 或直接粘贴路径' if allow_path else '')).strip()
        if raw.lower() in ('q', 'quit', 'exit'):
            raise KeyboardInterrupt
        if allow_path and ('\\' in raw or '/' in raw or ':' in raw):
            p = Path(raw.strip('"'))
            if p.is_dir():
                return ('dir', p)
            if p.is_file():
                return ('file', p)
            print('  ! 路径不存在：{}'.format(p))
            continue
        if raw.isdigit() and 1 <= int(raw) <= len(items):
            return int(raw) - 1
        print('  ! 请输入 1~{} 之间的编号'.format(len(items)))


def _default_out(src: Path, sheet_name: str) -> Path:
    import re
    safe = re.sub(r'[^\w\u4e00-\u9fff]+', '_', str(sheet_name)).strip('_') or 'sheet'
    return src.with_name('{}_{}_hardcode.sql'.format(src.stem, safe))


# ------------------------------------------------------------------ 交互流程
def interactive(opts: argparse.Namespace) -> int:
    cwd = Path.cwd()
    print(BANNER)

    files = scan_dir(cwd)
    while not files:
        print('\n当前目录没有 Excel/CSV：{}'.format(cwd))
        raw = input('请输入文件夹路径 (q 退出): ').strip().strip('"')
        if raw.lower() in ('q', 'quit'):
            return EXIT_OK
        d = Path(raw)
        if not d.is_dir():
            print('  ! 目录不存在，请重试')
            continue
        cwd, files = d, scan_dir(d)

    choice = pick(files, '发现 {} 个文件（目录：{}）'.format(len(files), cwd),
                  lambda p: '{}   ({:,} KB)'.format(p.name, p.stat().st_size // 1024),
                  allow_path=True)
    if isinstance(choice, tuple):
        kind, path = choice
        if kind == 'dir':
            sub = scan_dir(path)
            if not sub:
                print('该目录下没有 Excel/CSV：{}'.format(path))
                return EXIT_ERROR
            target = sub[pick(sub, '目录：{}'.format(path), lambda p: p.name)]
        else:
            target = path
    else:
        target = files[choice]

    print('\n读取：{}'.format(target))
    try:
        infos = probe_sheets(target)                    # 只读元信息，大表也很快
    except ReadError as e:
        print('读取失败：{}'.format(e))
        return EXIT_ERROR

    if len(infos) == 1:
        sheet_name = infos[0].name
        print('  唯一 sheet：{}  ({})'.format(sheet_name, infos[0].shape_text()))
    else:
        sheet_name = infos[pick(infos, '共 {} 个 sheet，请选择'.format(len(infos)),
                                lambda s: '{}   ({})'.format(s.name, s.shape_text()))].name

    print('  读取数据中…（几万行的 Excel 可能要十几秒）')
    try:
        sheet = read_sheet_rows(target, sheet_name, delimiter=opts.delimiter)
    except ReadError as e:
        print('读取失败：{}'.format(e))
        return EXIT_ERROR

    row = opts.header_row
    raw = ask('表头在第几行', str(row))
    row = int(raw) if raw.isdigit() and int(raw) >= 1 else row
    header, rows = split_header(sheet, row)
    if header is None:
        print('表头行 {} 超出范围（该表仅 {} 行），无法转换'.format(row, sheet.nrows))
        return EXIT_ERROR

    verdict = check_header(header, rows)
    if verdict.ok:
        print('  表头校验通过：{} 列 / {} 行数据'.format(len(header), len(rows)))
    else:
        print('\n[表格不规范] 检测到 {} 个问题：'.format(len(verdict.issues)))
        for i, msg in enumerate(verdict.issues, 1):
            print('  {}. {}'.format(i, msg))
        if not rows:
            print('\n>>> 无法转换：表头下方没有数据。')
            return EXIT_ERROR
        if ask('仍要继续？（自动用 col_N 补全空列名、重复列名加 _2 后缀）y/N', 'N').lower() != 'y':
            print('>>> 已取消。建议先整理表头：首行字段名、不空不重复。')
            return EXIT_OK
        header = repair(header)

    print('\nSQL 方言： ' + '   '.join('{}={}'.format(k, DIALECTS[k].name)
                                     for k in ('sqlserver', 'mysql', 'oracle', 'postgresql')))
    dialect = resolve(ask('选择', '1') or '1')

    shape = '1=CTE 包裹(可直接跑)  2=纯 UNION ALL 块  3=INSERT INTO ... VALUES'
    fmt_choice = ask('输出形式  ' + shape, '1')
    if fmt_choice == '3':
        fmt, wrap = 'insert', 'plain'
    else:
        fmt, wrap = 'union', ('plain' if fmt_choice == '2' else 'cte')

    if fmt == 'union' and len(rows) > UNION_ROW_WARN:
        print('  ! {} 行用 UNION ALL 会生成很大的 SQL，建议改用 3=INSERT；也可继续。'.format(len(rows)))

    table = ask('内联表名', 'HARDCODE') if fmt == 'insert' or wrap == 'cte' else 'HARDCODE'
    empty_as_null = ask('空字符串按 NULL 处理？y/N', 'N').lower() == 'y'

    options = SqlOptions(dialect=dialect, table=table or 'HARDCODE', fmt=fmt, wrap=wrap,
                         empty_as_null=empty_as_null, batch_size=opts.batch_size)

    out = Path(ask('\n输出文件路径', str(_default_out(target, sheet.name))).strip('"'))
    encoding = ask('文件编码（SSMS 老版本中文乱码时用 utf-8-sig）', 'utf-8')
    try:
        n = write_sql(out, header, rows, options, encoding=encoding)
    except SqlGenError as e:
        print('>>> 无法转换：{}'.format(e))
        return EXIT_ERROR

    print('\n[OK] {} 行 x {} 列 -> {}   ({:,} KB)'.format(
        n, len(header), out, max(1, out.stat().st_size // 1024)))
    preview_head(header, rows, options)
    if ask('\n复制到剪贴板？(y/N)', 'N').lower() == 'y':
        text = build_sql(header, rows, options)
        print('已复制 {} 字符'.format(len(text)) if copy_text(text)
              else '  ! 复制失败（系统未提供剪贴板命令）')
    return EXIT_OK


def split_header(sheet: Sheet, row_index: int):
    """按 1 基行号切出表头与数据，并裁掉右侧多余空列。"""
    if row_index < 1 or row_index > sheet.nrows:
        return None, []
    header = list(sheet.rows[row_index - 1])
    while len(header) > 1 and (header[-1] is None or str(header[-1]).strip() == ''):
        header.pop()
    rows = [list(r[:len(header)]) for r in sheet.rows[row_index:]]
    return header, rows


def preview_head(header, rows, options: SqlOptions, lines: int = 2) -> None:
    """打印前几条语句（不含注释与 CTE 尾巴）。"""
    print('预览：')
    text = build_sql(header, rows[:max(lines, 1)], options)
    body = [l for l in text.split('\n')
            if l and not l.startswith('--') and not l.startswith('WITH ')
            and not l.startswith('SELECT *') and l not in (')', 'UNION ALL')]
    for line in body[:lines]:
        print('  ' + line[:170] + (' ...' if len(line) > 170 else ''))


# ------------------------------------------------------------------ 批处理
def batch(args: argparse.Namespace) -> int:
    """非交互模式：任何情况都不调用 input()，出错直接返回非 0。"""
    target = Path(args.file)
    try:
        infos = probe_sheets(target)                    # 先拿 sheet 清单，避免为报错而全量读取
    except ReadError as e:
        print('读取失败：{}'.format(e), file=sys.stderr)
        return EXIT_ERROR

    names = [i.name for i in infos]
    if args.sheet:
        if args.sheet not in names:
            print('找不到 sheet：{}（可选：{}）'.format(args.sheet, ', '.join(names)), file=sys.stderr)
            return EXIT_ERROR
        sheet_name = args.sheet
    elif len(names) == 1:
        sheet_name = names[0]
    else:
        print('该文件有 {} 个 sheet，请用 -s 指定：{}'.format(len(names), ', '.join(names)),
              file=sys.stderr)
        return EXIT_ERROR

    try:
        sheet = read_sheet_rows(target, sheet_name,
                                encoding=args.input_encoding, delimiter=args.delimiter)
    except ReadError as e:
        print('读取失败：{}'.format(e), file=sys.stderr)
        return EXIT_ERROR

    header, rows = split_header(sheet, args.header_row)
    if header is None:
        print('--header-row {} 超出范围（该表仅 {} 行）'.format(args.header_row, sheet.nrows),
              file=sys.stderr)
        return EXIT_ERROR

    verdict = check_header(header, rows)
    if not verdict.ok:
        print('[表格不规范] {}'.format('；'.join(verdict.issues)), file=sys.stderr)
        if not rows:
            print('无法转换：表头下方没有数据。', file=sys.stderr)
            return EXIT_ERROR
        if not args.force:
            print('加 --force 可自动修复表头（col_N / 去重后缀）后继续。', file=sys.stderr)
            return EXIT_UNCLEAN
        header = repair(header)

    if len(rows) > HARD_LIMIT_ROWS:
        print('数据 {} 行，超过硬编码 SQL 的合理上限（{}），建议改用数据库原生导入。'
              .format(len(rows), HARD_LIMIT_ROWS), file=sys.stderr)
        return EXIT_ERROR
    if args.fmt == 'union' and len(rows) > UNION_ROW_WARN:
        print('提示：{} 行用 UNION ALL 生成的 SQL 很大，--format insert 通常更快。'
              .format(len(rows)), file=sys.stderr)

    options = SqlOptions(dialect=resolve(args.dialect), table=args.table, fmt=args.fmt,
                         wrap=args.wrap, empty_as_null=args.empty_as_null,
                         all_string=args.all_string, batch_size=args.batch_size)
    out = Path(args.out) if args.out else _default_out(target, sheet.name)
    try:
        n = write_sql(out, header, rows, options, encoding=args.output_encoding)
    except SqlGenError as e:
        print('无法转换：{}'.format(e), file=sys.stderr)
        return EXIT_ERROR

    print('[OK] {} 行 x {} 列 -> {} ({:,} KB)'.format(n, len(header), out,
                                                     max(1, out.stat().st_size // 1024)))
    return EXIT_OK


# ------------------------------------------------------------------ 参数
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='excel2sql',
        description='把 Excel / CSV 的每一行转成硬编码 SQL（UNION ALL 内联表或 INSERT VALUES）。',
        epilog='不带 FILE 时进入交互向导；带 FILE 时完全非交互，适合批处理与 CI。',
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('file', nargs='?', help='源文件（.xlsx / .xlsm / .xls / .csv）')
    p.add_argument('-s', '--sheet', help='工作表名（多 sheet 时必须指定）')
    p.add_argument('-o', '--out', help='输出 SQL 路径（默认 <文件名>_<sheet>_hardcode.sql）')
    p.add_argument('-d', '--dialect', default='sqlserver', metavar='NAME',
                   help='方言：sqlserver(默认) / mysql / oracle / postgresql，也可用 1~4')
    p.add_argument('--header-row', type=int, default=1, metavar='N', help='表头所在行号，从 1 开始')
    p.add_argument('--format', dest='fmt', choices=('union', 'insert'), default='union',
                   help='输出格式：union=UNION ALL 内联表（默认），insert=INSERT INTO ... VALUES')
    p.add_argument('--wrap', choices=('cte', 'plain'), default='cte',
                   help='union 模式：cte=外面包 WITH ... AS（默认），plain=纯 UNION ALL 块')
    p.add_argument('--table', default='HARDCODE', help='内联表名（默认 HARDCODE）')
    p.add_argument('--empty-as-null', action='store_true', help='空字符串按 NULL 输出')
    p.add_argument('--all-string', action='store_true', help='所有列强制按字符串输出')
    p.add_argument('--batch-size', type=int, default=500, metavar='N',
                   help='insert 模式每批行数（默认 500）')
    p.add_argument('--encoding', dest='output_encoding', default='utf-8',
                   help='输出文件编码（老版 SSMS 中文乱码时用 utf-8-sig）')
    p.add_argument('--input-encoding', default=None, help='CSV 输入编码（默认自动尝试）')
    p.add_argument('--delimiter', default=None, help='CSV 分隔符（默认自动探测）')
    p.add_argument('--force', action='store_true', help='表头不规范时自动修复并继续（非交互模式）')
    p.add_argument('-V', '--version', action='version', version='excel2sql {}'.format(__version__))
    return p


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    if args.header_row < 1:
        print('--header-row 必须 >= 1（现在给的是 {}）'.format(args.header_row), file=sys.stderr)
        return EXIT_ERROR
    if args.batch_size < 1:
        print('--batch-size 必须 >= 1', file=sys.stderr)
        return EXIT_ERROR
    try:
        if args.file is None:
            return interactive(args)
        if not Path(args.file).is_file():
            print('文件不存在：{}'.format(args.file), file=sys.stderr)
            return EXIT_ERROR
        return batch(args)
    except KeyboardInterrupt:
        print('\n已中断')
        return 130
    except EOFError:
        print('\n输入结束。非交互环境请改用：excel2sql <文件> [-s sheet] [-o out.sql]', file=sys.stderr)
        return EXIT_ERROR
    except RecursionError:
        print('文件层级异常，无法解析', file=sys.stderr)
        return EXIT_ERROR


if __name__ == '__main__':          # pragma: no cover
    sys.exit(main())
