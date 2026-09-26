"""命令行入口：交互式向导 + 可脚本化的非交互模式。

交互原则（v0.3）：
- 能自动推断的一律不问：目录里只有一个文件、工作簿只有一个 sheet
- 表头行自动识别 + 打印预览，用户回车确认或输入行号修正
- 方言/格式/表名/编码/输出目录/剪贴板等默认值来自 excel2sql.ini，
  交互时不再逐项询问（除非配置 ask_advanced = true）
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

from . import __version__
from .clipboard import copy_text
from .config import (ConfigError, INFER_AUTO, INFER_OFF, INFER_ON, Settings,
                     unique_path, write_template)
from .config import load as load_settings
from .dialects import DIALECTS, UnknownDialect, resolve
from .dialects import ORDER as DIALECT_ORDER
from .headers import check as check_header
from .headers import detect as detect_header, is_blank, repair
from .reader import (ReadError, Sheet, has_type_info, probe_sheets,
                     read_sheet_rows, scan_dir)
from .sqlgen import SqlGenError, SqlOptions, UNION_ROW_WARN, build_sql, write_sql

EXIT_OK, EXIT_ERROR, EXIT_UNCLEAN = 0, 1, 2
# 超过这个行数，硬编码 SQL 基本已经不适合
HARD_LIMIT_ROWS = 200_000

BANNER = """{sep}
  excel2sql {ver}  |  Excel / CSV  ->  硬编码 SQL
{sep}""".format(sep='=' * 58, ver=__version__)


# ------------------------------------------------------------------ 控制台编码
def setup_console() -> None:
    """让中文提示在任何控制台/管道下都不会抛 UnicodeEncodeError。

    Windows 上 stdout 不是终端时（CI、重定向到文件、`| more`）用的是 ANSI 码页，
    en-US 环境即 cp1252，打印「行 x 列」这类中文字符会直接崩。策略：

    * 管道/重定向 → 改用 UTF-8（数据工具该有的默认，且 CI 日志本身按 UTF-8 渲染）
    * 交互式终端 → 保留 Python 已探测到的编码（`chcp 65001` 时即 UTF-8，
      中文 Windows 的 cp936 也能显示中文），只把错误处理放宽，确保永不崩
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            encoding = (getattr(stream, 'encoding', '') or '').lower()
            if encoding.replace('-', '') != 'utf8' and not stream.isatty():
                stream.reconfigure(encoding='utf-8', errors='replace')
            else:
                stream.reconfigure(errors='replace')
        except (AttributeError, OSError, ValueError):
            pass                    # 被包装过的流（如 pytest 捕获）就保持原样


# ------------------------------------------------------------------ 输入助手
def ask(prompt: str, default: str = '') -> str:
    hint = ' [{}]'.format(default) if default else ''
    return input('{}{}: '.format(prompt, hint)).strip() or default


def pick(items: Sequence, title: str, render, allow_path: bool = False):
    """编号选择。返回下标；allow_path 时也可能返回 ('file'|'dir', Path)。"""
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


def pick_single_or_list(items: Sequence, title: str, render, hint: str = '回车直接使用'):
    """只有一项时不再要求输入编号，回车即用；多项时正常选择。"""
    if len(items) == 1:
        print('\n{}\n  {}  （{}）'.format(title, render(items[0]), hint))
        raw = input('回车继续，或粘贴其它路径，q 退出: ').strip()
        if raw.lower() in ('q', 'quit', 'exit'):
            raise KeyboardInterrupt
        if raw:
            p = Path(raw.strip('"'))
            if p.is_dir():
                return ('dir', p)
            if p.is_file():
                return ('file', p)
            print('  ! 路径不存在，改用 {}'.format(items[0]))
        return 0
    return pick(items, title, render, allow_path=True)


def ask_dialect(default_key: str) -> str:
    print('\nSQL 方言：')
    for i, key in enumerate(DIALECT_ORDER, 1):
        mark = '    ← 当前配置' if key == default_key else ''
        print('  {}. {}{}'.format(i, DIALECTS[key].name, mark))
    default_index = str(DIALECT_ORDER.index(default_key) + 1) if default_key in DIALECT_ORDER else '1'
    while True:
        raw = ask('选择（编号，或直接输名字如 pg）', default_index)
        if raw.isdigit() and 1 <= int(raw) <= len(DIALECT_ORDER):
            return DIALECT_ORDER[int(raw) - 1]
        try:
            return resolve(raw).key          # 也允许直接输名字，如 mysql
        except UnknownDialect as e:
            # 交互里打错字不该崩，也不该悄悄用默认方言 —— 说清楚再问一次
            print('  ! {}'.format(e))


def ask_format(default_fmt: str, default_wrap: str) -> Tuple[str, str]:
    default = '3' if default_fmt == 'insert' else ('2' if default_wrap == 'plain' else '1')
    print('\n输出形式：')
    print('  1. CTE 包裹    WITH ... AS ( ... ) SELECT *   —— 可直接执行')
    print('  2. 纯 UNION ALL 块                            —— 方便嵌进已有 SQL')
    print('  3. INSERT INTO ... VALUES                     —— 行数多时更合适')
    raw = ask('选择', default)
    if raw == '3':
        return 'insert', 'plain'
    return 'union', ('plain' if raw == '2' else 'cte')


# ------------------------------------------------------------------ 表头
def _column_all_blank(rows: Sequence[Sequence[object]], idx: int) -> bool:
    """该列在给出的所有行里是否全为空。"""
    for r in rows:
        if idx < len(r) and not is_blank(r[idx]):
            return False
    return True


def split_header(sheet: Sheet, row_index: int) -> Tuple[Optional[List], List[List]]:
    """按 1 基行号切出表头与数据，并裁掉右侧多余空列。

    ★ 只裁「表头为空 **且** 该列数据全空」的尾列。
      旧实现只看表头是否为空 —— 于是"表头留空、下面却有数据"的尾列被**整列丢掉**
      （连数据一起，且不告警、退出码 0）。2026-09-25 实测确认。
      表头空但有数据的列必须保留，列名交由 check()/repair() 补成 col_N。
    """
    if row_index < 1 or row_index > sheet.nrows:
        return None, []
    bottom = [list(r) for r in sheet.rows[row_index:]]
    header = list(sheet.rows[row_index - 1])
    while len(header) > 1 and is_blank(header[-1]) \
            and _column_all_blank(bottom, len(header) - 1):
        header.pop()
    rows = [list(r[:len(header)]) for r in bottom]
    return header, rows


def _cells(text_row: Sequence, limit: int = 7, width: int = 14) -> str:
    out = []
    for v in list(text_row)[:limit]:
        s = '' if v is None else str(v).replace('\n', '⏎')
        out.append(s if len(s) <= width else s[:width] + '…')
    if len(text_row) > limit:
        out.append('…(+{} 列)'.format(len(text_row) - limit))
    return ' | '.join(out)


def choose_header_row(sheet: Sheet, suggested: int) -> int:
    """打印预览让用户确认表头行，输入行号可修正。"""
    row = suggested if 1 <= suggested <= sheet.nrows else 1
    while True:
        header, rows = split_header(sheet, row)
        print('\n表头预览（← 标记的是当前选中的表头行）：')
        start = max(1, row - 3)
        end = min(sheet.nrows, row + 2)
        for i in range(start, end + 1):
            mark = '   ← 表头' if i == row else ''
            print('  {:>4} | {}{}'.format(i, _cells(sheet.rows[i - 1]), mark))
        if sheet.nrows > end:
            print('  {:>4} | …（共 {} 行）'.format('…', sheet.nrows))
        raw = ask('回车确认，或输入其它行号修正', str(row))
        if raw.isdigit() and 1 <= int(raw) <= sheet.nrows:
            return int(raw)
        print('  ! 请输入 1~{} 之间的行号'.format(sheet.nrows))


# ------------------------------------------------------------------ 选项合并
def check_dialect(args: argparse.Namespace, settings: Settings) -> Optional[str]:
    """校验**将会生效**的那个方言名，返回错误信息（None 表示没问题）。

    只校验会真正用到的那个：命令行给了 `-d` 就以它为准，配置里的旧值不再追究
    （那是另一回事，不该拦住一次明确指定了方言的调用）。

    ★ 校验放在**读文件之前**：51,290 行的表读进来才发现方言拼错，白等几十秒。

    ★ 用 `is not None` 而不是真假判断来区分「没给 -d」与「给了 -d 但值为空」：
      shell 里写 `-d "$DIALECT"` 而变量为空，恰恰是最容易静默用错方言的场景。
    """
    if args.dialect is not None:
        source, value = '--dialect', args.dialect
    else:
        source, value = '配置项 dialect', settings.dialect
    try:
        resolve(value)
    except UnknownDialect as e:
        return '{} {}'.format(source, e)
    return None


def resolve_infer_types(args: argparse.Namespace, settings: Settings,
                        source_path=None) -> bool:
    """把 auto / on / off 解析成确定的布尔。

    `auto` 的含义是「只在**需要且安全**时才推断」：
    - **需要**：只有 CSV 这类无类型信息的输入才需要 —— 读出来每一格都是 str，
      不推断的话所有列都会变成字符串，`SUM()` 直接报错；
      xlsx/xls 的单元格自带类型，写成文本的单元格是用户有意的选择，不动。
    - **安全**：交给 coerce_numeric_columns 把关（整列可无损解析才转）。
    """
    if args.infer_types is not None:
        return args.infer_types                       # 命令行显式开关优先级最高
    mode = str(settings.infer_types or INFER_AUTO).strip().lower()
    if mode == INFER_ON:
        return True
    if mode == INFER_OFF:
        return False
    if source_path is None:
        return False                                  # 问不到来源就保守处理
    return not has_type_info(source_path)


def merge_options(args: argparse.Namespace, settings: Settings,
                  source_path=None) -> SqlOptions:
    """命令行参数 > 配置文件 > 内置默认。"""
    # is not None：`-d ''` 要当成「给了一个非法值」报错，不能当成「没给」而悄悄用配置值
    dialect = resolve(args.dialect).key if args.dialect is not None else settings.dialect
    fmt = args.fmt or settings.format
    wrap = args.wrap or settings.wrap
    empty_as_null = settings.empty_as_null if args.empty_as_null is None else args.empty_as_null
    all_string = settings.all_string if args.all_string is None else args.all_string
    return SqlOptions(
        dialect=resolve(dialect),
        table=args.table or settings.table,
        fmt=fmt,
        wrap=wrap if fmt == 'union' else 'plain',
        empty_as_null=empty_as_null,
        all_string=all_string,
        infer_types=resolve_infer_types(args, settings, source_path),
        batch_size=args.batch_size or settings.batch_size,
    )


def resolve_header_row(args: argparse.Namespace, settings: Settings) -> Optional[int]:
    """返回 None 表示需要自动识别。"""
    raw = args.header_row if args.header_row is not None else settings.header_row
    if raw is None:
        return None
    raw = str(raw).strip().lower()
    if raw in ('auto', '', 'none'):
        return None
    if raw.isdigit() and int(raw) >= 1:
        return int(raw)
    raise ConfigError('header_row / --header-row 只能是 auto 或 >=1 的行号，收到：{}'.format(raw))


# ------------------------------------------------------------------ 交互流程
def interactive(args: argparse.Namespace, settings: Settings, config_path) -> int:
    print(BANNER)
    print('配置：{}'.format('{}  （{}）'.format(config_path, settings.progress_hint())
                          if config_path else '使用内置默认值（{}）'.format(settings.progress_hint())))
    if config_path is None:
        print('提示：在数据目录放一个 excel2sql.ini 就能固化方言/格式/输出目录等设置。')

    # 1) 选文件
    cwd = Path.cwd()
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

    choice = pick_single_or_list(
        files, '发现 {} 个文件（目录：{}）'.format(len(files), cwd),
        lambda p: '{}   ({:,} KB)'.format(p.name, p.stat().st_size // 1024))
    if isinstance(choice, tuple):
        kind, path = choice
        if kind == 'dir':
            sub = scan_dir(path)
            if not sub:
                print('该目录下没有 Excel/CSV：{}'.format(path))
                return EXIT_ERROR
            target = sub[pick_single_or_list(sub, '目录：{}'.format(path), lambda p: p.name)]
        else:
            target = path
    else:
        target = files[choice]

    # 2) 选 sheet（只有一个时自动使用）
    print('\n读取：{}'.format(target))
    try:
        infos = probe_sheets(target)
    except ReadError as e:
        print('读取失败：{}'.format(e))
        return EXIT_ERROR

    if len(infos) == 1:
        sheet_name = infos[0].name
        print('  唯一 sheet：{}  ({})，自动使用'.format(sheet_name, infos[0].shape_text()))
    else:
        sheet_name = infos[pick(infos, '共 {} 个 sheet，请选择'.format(len(infos)),
                                lambda s: '{}   ({})'.format(s.name, s.shape_text()))].name

    print('  读取数据中…（几万行的 Excel 可能要十几秒）')
    try:
        sheet = read_sheet_rows(target, sheet_name, delimiter=args.delimiter)
    except ReadError as e:
        print('读取失败：{}'.format(e))
        return EXIT_ERROR

    # 3) 表头行：自动识别 + 预览确认
    try:
        wanted = resolve_header_row(args, settings)
    except ConfigError as e:
        print('配置错误：{}'.format(e))
        return EXIT_ERROR

    if wanted:
        header_row = wanted
        print('  表头行：使用配置/参数指定的第 {} 行'.format(header_row))
    else:
        suggested, reason = detect_header(sheet.rows)
        print('  表头识别：{}'.format(reason))
        header_row = choose_header_row(sheet, suggested)

    header, rows = split_header(sheet, header_row)
    if header is None:
        print('表头行 {} 超出范围（该表仅 {} 行），无法转换'.format(header_row, sheet.nrows))
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

    # 4) 输出选项：默认全部取配置，只有 ask_advanced 时才逐项询问
    if settings.ask_advanced:
        settings.dialect = ask_dialect(settings.dialect)
        settings.format, settings.wrap = ask_format(settings.format, settings.wrap)
        settings.table = ask('内联表名', settings.table) or settings.table
        settings.empty_as_null = ask('空字符串按 NULL 处理？y/N',
                                     'y' if settings.empty_as_null else 'N').lower() == 'y'
        settings.encoding = ask('输出编码（老版 SSMS 乱码用 utf-8-sig）', settings.encoding)
        settings.copy_clipboard = ask('生成后复制到剪贴板？y/N',
                                      'y' if settings.copy_clipboard else 'N').lower() == 'y'

    try:
        options = merge_options(args, settings, target)
    except (ConfigError, SqlGenError) as e:
        print('配置错误：{}'.format(e))
        return EXIT_ERROR

    # 行数保护：硬编码 SQL 不适合太大的表（与非交互模式保持一致）
    if len(rows) > HARD_LIMIT_ROWS:
        print('\n数据 {} 行，超过硬编码 SQL 的合理上限（{}）。'.format(len(rows), HARD_LIMIT_ROWS))
        if ask('仍要继续？y/N', 'N').lower() != 'y':
            print('>>> 已取消。建议改用数据库原生导入：BULK INSERT / LOAD DATA / COPY。')
            return EXIT_OK
    elif options.fmt == 'union' and len(rows) > UNION_ROW_WARN:
        print('  提示：{} 行用 UNION ALL 生成的 SQL 很大，把配置改成 format = insert 通常更快。'
              .format(len(rows)))

    # 5) 输出路径（配置化：默认在源文件目录下建子目录，同名自动防冲突）
    default_out = settings.default_output(target, sheet.name)
    if args.out:
        out = Path(args.out)            # 显式指定就按原样写（可覆盖）
    elif settings.ask_advanced:
        out = Path(ask('\n输出文件路径', str(default_out)).strip('"'))
    elif settings.overwrite:
        out = default_out
    else:
        out = unique_path(default_out)

    try:
        n = _write(out, header, rows, options, settings)
    except (SqlGenError, ConfigError) as e:
        print('>>> 无法转换：{}'.format(e))
        return EXIT_ERROR

    print('\n[OK] {} 行 x {} 列 -> {}   ({:,} KB)'.format(
        n, len(header), out, max(1, out.stat().st_size // 1024)))
    _preview(header, rows, options)

    if settings.copy_clipboard:
        text = build_sql(header, rows, options)
        print('已复制到剪贴板（{} 字符）'.format(len(text)) if copy_text(text)
              else '  ! 复制失败（系统未提供剪贴板命令）')
    return EXIT_OK


def _write(out: Path, header, rows, options: SqlOptions, settings: Settings) -> int:
    encoding = settings.encoding
    return write_sql(out, header, rows, options, encoding=encoding)


def _preview(header, rows, options: SqlOptions, lines: int = 2) -> None:
    print('预览：')
    text = build_sql(header, rows[:max(lines, 1)], options)
    body = [l for l in text.split('\n')
            if l and not l.startswith('--') and not l.startswith('WITH ')
            and not l.startswith('SELECT *') and l not in (')', 'UNION ALL')]
    for line in body[:lines]:
        print('  ' + line[:170] + (' ...' if len(line) > 170 else ''))


# ------------------------------------------------------------------ 批处理
def batch(args: argparse.Namespace, settings: Settings) -> int:
    """非交互模式：任何情况都不调用 input()，出错直接返回非 0。"""
    target = Path(args.file)
    try:
        wanted_header = resolve_header_row(args, settings)
    except ConfigError as e:
        print(e, file=sys.stderr)
        return EXIT_ERROR

    try:
        infos = probe_sheets(target)
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

    if wanted_header:
        header_row = wanted_header
    else:
        header_row, reason = detect_header(sheet.rows)      # 自动识别：无需交互
        print('表头自动识别：{}'.format(reason), file=sys.stderr)
        if args.strict_header and header_row != 1:
            print('--strict-header 要求表头在第 1 行，但自动识别跳过了前 {} 行；'
                  '请人工确认后加 --header-row {} 明确指定。'.format(header_row - 1, header_row),
                  file=sys.stderr)
            return EXIT_UNCLEAN

    header, rows = split_header(sheet, header_row)
    if header is None:
        print('表头行 {} 超出范围（该表仅 {} 行）'.format(header_row, sheet.nrows), file=sys.stderr)
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
    if (args.fmt or settings.format) == 'union' and len(rows) > UNION_ROW_WARN:
        print('提示：{} 行用 UNION ALL 生成的 SQL 很大，--format insert 通常更快。'
              .format(len(rows)), file=sys.stderr)

    try:
        options = merge_options(args, settings, target)
    except (ConfigError, SqlGenError) as e:
        print('配置错误：{}'.format(e), file=sys.stderr)
        return EXIT_ERROR

    if args.out:
        out = Path(args.out)                    # 显式指定就按原样写（可覆盖）
    else:
        out = settings.default_output(target, sheet.name)
        if not settings.overwrite:
            out = unique_path(out)
    try:
        n = _write(out, header, rows, options, settings)
    except (SqlGenError, ConfigError) as e:
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
        epilog='不带 FILE 时进入交互向导；带 FILE 时完全非交互。\n'
               '默认值来自 excel2sql.ini（数据目录或 ~/.excel2sql/config.ini），命令行参数优先级最高。',
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('file', nargs='?', help='源文件（.xlsx / .xlsm / .xls / .csv）')
    p.add_argument('-s', '--sheet', help='工作表名（多 sheet 时必须指定）')
    p.add_argument('-o', '--out', help='输出 SQL 路径（默认按配置 output_dir 生成）')
    p.add_argument('-d', '--dialect', metavar='NAME',
                   help='方言：{}，也可用 1~{}（无法识别即报错退出，不会静默回退）'.format(
                       ' / '.join(DIALECT_ORDER), len(DIALECT_ORDER)))
    p.add_argument('--header-row', metavar='N', default=None,
                   help='表头行号；auto（默认）表示自动识别')
    p.add_argument('--format', dest='fmt', choices=('union', 'insert'), help='输出格式')
    p.add_argument('--wrap', choices=('cte', 'plain'), help='union 模式是否用 CTE 包裹')
    p.add_argument('--table', help='内联表名')
    neg = p.add_argument_group('布尔开关（不传则用配置）')
    neg.add_argument('--empty-as-null', dest='empty_as_null', action='store_true', default=None,
                     help='空字符串按 NULL 输出')
    neg.add_argument('--no-empty-as-null', dest='empty_as_null', action='store_false',
                     help='空字符串保留为空字符串')
    neg.add_argument('--all-string', dest='all_string', action='store_true', default=None,
                     help='所有列强制按字符串输出')
    neg.add_argument('--no-all-string', dest='all_string', action='store_false',
                     help='按列推断类型')
    neg.add_argument('--infer-types', dest='infer_types', action='store_true', default=None,
                     help='强制推断类型：整列都能无损解析成数字时按数字输出（Excel 源也照做）')
    neg.add_argument('--no-infer-types', dest='infer_types', action='store_false',
                     help='关闭推断：CSV 的列一律按字符串输出')
    neg.add_argument('--copy-clipboard', dest='copy_clipboard', action='store_true', default=None,
                     help='生成后复制到剪贴板')
    neg.add_argument('--no-copy-clipboard', dest='copy_clipboard', action='store_false',
                     help='不复制到剪贴板')
    p.add_argument('--batch-size', type=int, metavar='N', help='insert 模式每批行数')
    p.add_argument('--encoding', dest='output_encoding', help='输出文件编码（如 utf-8-sig）')
    p.add_argument('--input-encoding', help='CSV 输入编码（默认自动尝试）')
    p.add_argument('--delimiter', help='CSV 分隔符（默认自动探测）')
    p.add_argument('--force', action='store_true', help='表头不规范时自动修复并继续（非交互模式）')
    p.add_argument('--strict-header', action='store_true',
                   help='要求表头必须在第 1 行：自动识别若需要跳过行就直接报错退出（退出码 2）')
    p.add_argument('--config', help='指定配置文件路径')
    p.add_argument('--init-config', action='store_true',
                   help='在当前目录生成 excel2sql.ini 模板后退出')
    p.add_argument('-V', '--version', action='version', version='excel2sql {}'.format(__version__))
    return p


def main(argv: Optional[Sequence[str]] = None) -> int:
    setup_console()
    args = build_parser().parse_args(argv)

    if args.init_config:
        path = Path.cwd() / 'excel2sql.ini'
        written = write_template(path, overwrite=True)
        print('已生成配置模板：{}'.format(written))
        return EXIT_OK

    try:
        settings, config_path, warnings = load_settings(args.config)
    except ConfigError as e:
        print('配置错误：{}'.format(e), file=sys.stderr)
        return EXIT_ERROR
    for w in warnings:
        print('[配置] {}'.format(w), file=sys.stderr)

    # 命令行开关覆盖配置里的布尔项
    if args.copy_clipboard is not None:
        settings.copy_clipboard = args.copy_clipboard
    if args.output_encoding:
        settings.encoding = args.output_encoding
    if args.batch_size:
        settings.batch_size = args.batch_size

    if args.batch_size is not None and args.batch_size < 1:
        print('--batch-size 必须 >= 1', file=sys.stderr)
        return EXIT_ERROR
    if args.header_row is not None:
        raw = str(args.header_row).strip().lower()
        if raw not in ('auto', ''):
            try:
                if int(raw) < 1:
                    raise ValueError
            except ValueError:
                print('--header-row 必须是 >= 1 的行号，或 auto', file=sys.stderr)
                return EXIT_ERROR

    bad_dialect = check_dialect(args, settings)
    if bad_dialect:
        print(bad_dialect, file=sys.stderr)
        return EXIT_ERROR

    try:
        if args.file is None:
            return interactive(args, settings, config_path)
        if not Path(args.file).is_file():
            print('文件不存在：{}'.format(args.file), file=sys.stderr)
            return EXIT_ERROR
        return batch(args, settings)
    except KeyboardInterrupt:
        print('\n已中断')
        return 130
    except EOFError:
        print('\n输入结束。非交互环境请改用：excel2sql <文件> [-s sheet] [-o out.sql]', file=sys.stderr)
        return EXIT_ERROR


if __name__ == '__main__':          # pragma: no cover
    sys.exit(main())
