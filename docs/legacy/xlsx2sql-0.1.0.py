#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx2sql - Excel/CSV 转硬编码 SQL（UNION ALL 内联表）交互式命令行工具

流程：扫描当前目录 -> 编号选文件 -> 选 sheet -> 校验表头 -> 选方言 -> 生成 SQL
用法：
    python xlsx2sql.py                       # 交互式（也可双击 xlsx2sql.bat）
    python xlsx2sql.py <文件> [-s sheet] [-o out.sql] [-d 1] [--header-row N] [--plain]
依赖：openpyxl（读 .xls 另需 pandas+xlrd）
"""
import sys
import re
import csv
import datetime
import subprocess
import tempfile
import argparse
from pathlib import Path

try:
    sys.stdout.reconfigure(errors='replace')
except Exception:
    pass

ALL_EXT = {'.xlsx', '.xlsm', '.xls', '.csv'}

# 方言配置：ql/qr=标识符引号  pre=字符串前缀  cat=拼接方式  dpre/dsuf=日期包装  chr=换行函数  dual
DIALECTS = {
    '1': dict(name='SQL Server', ql='[', qr=']', pre='N', cat='plus',
              dpre='', dsuf='', chr='CHAR(10)', dual=''),
    '2': dict(name='MySQL',      ql='`', qr='`', pre='', cat='func',
              dpre='', dsuf='', chr='CHAR(10)', dual=''),
    '3': dict(name='Oracle',     ql='"', qr='"', pre='', cat='pipe',
              dpre='TO_DATE(', dsuf=",'YYYY-MM-DD HH24:MI:SS')", chr='CHR(10)', dual=' FROM dual'),
    '4': dict(name='PostgreSQL', ql='"', qr='"', pre='', cat='pipe',
              dpre='', dsuf='', chr='CHR(10)', dual=''),
}


# ---------------- 读取 ----------------
def read_sheets(path):
    """返回 [(sheet名, [[单元格,...], ...]), ...]"""
    ext = path.suffix.lower()
    if ext == '.csv':
        for enc in ('utf-8-sig', 'gbk', 'utf-8'):
            try:
                with open(path, newline='', encoding=enc) as f:
                    return [(path.stem, [r for r in csv.reader(f)])]
            except UnicodeDecodeError:
                continue
        raise RuntimeError('CSV 编码无法识别（已试 utf-8/gbk）')

    if ext == '.xls':
        try:
            import pandas as pd
        except ImportError:
            raise RuntimeError('.xls 需要 pandas+xlrd，请先另存为 .xlsx')
        dfs = pd.read_excel(path, sheet_name=None, header=None)
        return [(k, [['' if pd.isna(c) else c for c in row] for row in df.values.tolist()])
                for k, df in dfs.items()]

    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        return [(ws.title, [list(r) for r in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]
    finally:
        wb.close()


def scan_dir(d):
    return sorted(p for p in Path(d).iterdir()
                  if p.is_file() and p.suffix.lower() in ALL_EXT and not p.name.startswith('~$'))


def safe_name(s):
    return re.sub(r'[^\w\u4e00-\u9fff]+', '_', str(s)).strip('_') or 'sheet'


# ---------------- 交互 ----------------
def ask(prompt, default=''):
    s = input('{} {}: '.format(prompt, '[{}]'.format(default) if default else '')).strip()
    return s or default


def pick(items, title, render, allow_path=False):
    while True:
        print('\n' + title)
        for i, it in enumerate(items, 1):
            print('  {:>2}. {}'.format(i, render(it)))
        s = input('输入编号{} (q 退出): '.format(' 或直接粘贴路径' if allow_path else '')).strip()
        if s.lower() in ('q', 'quit', 'exit'):
            sys.exit(0)
        if allow_path and ('\\' in s or '/' in s or ':' in s):
            p = Path(s.strip('"'))
            if p.is_dir():
                return 'DIR|' + str(p)
            if p.is_file():
                return 'FILE|' + str(p)
            print('  ! 路径不存在：{}'.format(p))
            continue
        if s.isdigit() and 1 <= int(s) <= len(items):
            return int(s)
        print('  ! 请输入 1~{} 之间的编号'.format(len(items)))


# ---------------- 表头校验 ----------------
def check_header(hdr, data_rows):
    """返回 (是否规范, [问题描述...], 修复后的表头)"""
    issues, fixed = [], list(hdr)
    if not fixed:
        return False, ['表头行为空或整行无内容'], fixed
    if not data_rows:
        issues.append('表头下方没有任何数据行')

    holes = [i for i, h in enumerate(fixed) if h is None or str(h).strip() == '']
    if holes:
        issues.append('表头第 {} 列为空（共 {} 处）'.format(
            ','.join(str(i + 1) for i in holes[:8]), len(holes)))

    seen, dup = set(), []
    for i, h in enumerate(fixed):
        k = str(h).strip()
        if k in seen:
            dup.append(i)
        seen.add(k)
    if dup:
        issues.append('列名重复：{}'.format(', '.join(str(fixed[i]) for i in dup[:8])))

    def looks_like_value(v):
        if isinstance(v, (int, float, datetime.datetime, datetime.date)):
            return True
        s = str(v).strip()
        return bool(re.fullmatch(r'-?\d+(\.\d+)?', s)) or bool(
            re.fullmatch(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?)?', s))

    if fixed and all(looks_like_value(v) for v in fixed):
        issues.append('表头行全是数字/日期，不像字段名（表头可能不在这一行）')

    # 修复方案：空列名 -> col_N；重复名 -> 名_2
    for i in holes:
        fixed[i] = 'col_{}'.format(i + 1)
    cnt = {}
    for i in dup:
        base = str(fixed[i]).strip() or 'col_{}'.format(i + 1)
        cnt[base] = cnt.get(base, 1) + 1
        fixed[i] = '{}_{}'.format(base, cnt[base])
    return (not issues), issues, fixed


# ---------------- 生成 ----------------
def _interleave(segs, ch):
    out = []
    for i, s in enumerate(segs):
        if i:
            out.append(ch)
        out.append(s)
    return out


def lit(v, force_str, dia):
    if v is None:
        return 'NULL'
    if isinstance(v, bool):
        return "'1'" if v else "'0'"
    if isinstance(v, datetime.datetime):
        s = v.strftime('%Y-%m-%d %H:%M:%S')
        return "{}{}'{}'{}".format(dia['dpre'], dia['pre'], s.replace("'", "''"), dia['dsuf'])
    if isinstance(v, datetime.date):
        s = v.strftime('%Y-%m-%d')
        return "{}{}'{}'{}".format(dia['dpre'], dia['pre'], s.replace("'", "''"), dia['dsuf'])
    if isinstance(v, (int, float)) and not force_str:
        return str(v)
    s = str(v).replace('\r\n', '\n').replace('\r', '\n')
    segs = [dia['pre'] + "'" + p.replace("'", "''") + "'" for p in s.split('\n')]
    if len(segs) == 1:
        return segs[0]
    parts = _interleave(segs, dia['chr'])
    if dia['cat'] == 'func':                       # MySQL 用 CONCAT
        return 'CONCAT({})'.format(', '.join(parts))
    return (' + ' if dia['cat'] == 'plus' else ' || ').join(parts)


def build_sql(hdr, rows, dia, table_name='HARDCODE', wrap='cte'):
    """返回 (sql文本, 生成行数)"""
    ncols = len(hdr)
    body_rows = [r for r in rows if not all(c is None or str(c).strip() == '' for c in r)]
    if not body_rows:
        raise RuntimeError('没有有效数据行')

    force = [any(i < len(r) and isinstance(r[i], str) and r[i] != '' for r in body_rows)
             for i in range(ncols)]
    forced = [h for h, f in zip(hdr, force) if f]

    lines = []
    for r in body_rows:
        cols = ', '.join('{} AS {}{}{}'.format(
            lit(r[i] if i < len(r) else None, force[i], dia),
            dia['ql'], h, dia['qr']) for i, h in enumerate(hdr))
        lines.append('SELECT ' + cols + dia['dual'])

    body = '\nUNION ALL\n'.join(lines)
    head = ['-- 由 xlsx2sql.py 生成：{} 行 x {} 列'.format(len(lines), ncols),
            '-- 方言：{}'.format(dia['name'])]
    if forced:
        head.append('-- 混类型列已统一为字符串（UNION ALL 同列类型须一致）：'
                    + ', '.join(str(x) for x in forced))

    if wrap == 'plain':
        return '\n'.join(head) + '\n' + body + '\n', len(lines)
    return ('\n'.join(head) + '\nWITH {} AS (\n{}\n)\nSELECT * FROM {};\n'
            .format(table_name, body, table_name)), len(lines)


# ---------------- 剪贴板 ----------------
def copy_clipboard(text):
    try:
        with tempfile.NamedTemporaryFile('w', suffix='.txt', delete=False, encoding='utf-8') as f:
            f.write(text)
            tmp = f.name
        subprocess.run(['powershell', '-NoProfile', '-Command',
                        'Get-Content -Raw -Encoding UTF8 "{}" | Set-Clipboard'.format(tmp)],
                       check=True, timeout=60)
        Path(tmp).unlink(missing_ok=True)
        return True
    except Exception as e:
        print('  ! 复制失败：{}'.format(e))
        return False


# ---------------- 主流程 ----------------
def main():
    ap = argparse.ArgumentParser(description='Excel/CSV -> 硬编码 UNION ALL SQL')
    ap.add_argument('file', nargs='?', help='源文件；留空则交互式扫描当前目录')
    ap.add_argument('-s', '--sheet')
    ap.add_argument('-o', '--out')
    ap.add_argument('-d', '--dialect', default='1', choices=list(DIALECTS))
    ap.add_argument('--header-row', type=int, default=1)
    ap.add_argument('--plain', action='store_true', help='输出纯 UNION ALL 块（不加 CTE 包裹）')
    args = ap.parse_args()
    interactive = args.file is None

    print('=' * 58)
    print('  xlsx2sql  |  Excel/CSV  ->  硬编码 UNION ALL SQL')
    print('=' * 58)

    # 1) 选文件
    if not interactive:
        target = Path(args.file)
        if not target.is_file():
            print('文件不存在：{}'.format(target))
            return 1
    else:
        cwd = Path.cwd()
        files = scan_dir(cwd)
        while not files:
            print('\n当前目录没有 Excel/CSV：{}'.format(cwd))
            s = input('请输入文件夹路径 (q 退出): ').strip().strip('"')
            if s.lower() in ('q', 'quit'):
                return 0
            d = Path(s)
            if not d.is_dir():
                print('  ! 目录不存在，请重试')
                continue
            cwd, files = d, scan_dir(d)

        res = pick(files, '发现 {} 个文件（目录：{}）'.format(len(files), cwd),
                   lambda p: '{}   ({:,} KB)'.format(p.name, p.stat().st_size // 1024),
                   allow_path=True)
        if isinstance(res, str):
            kind, val = res.split('|', 1)
            if kind == 'DIR':
                sub = scan_dir(val)
                if not sub:
                    print('该目录下没有 Excel/CSV：{}'.format(val))
                    return 1
                target = sub[pick(sub, '目录：{}'.format(val), lambda p: p.name) - 1]
            else:
                target = Path(val)
        else:
            target = files[res - 1]

    # 2) 读取 sheet
    print('\n读取：{}'.format(target))
    try:
        sheets = read_sheets(target)
    except Exception as e:
        print('读取失败：{}'.format(e))
        return 1

    # 3) 选 sheet
    if args.sheet:
        hit = next((x for x in sheets if x[0] == args.sheet), None)
        if hit is None:
            print('找不到 sheet：{}（可选：{}）'.format(args.sheet, ', '.join(s[0] for s in sheets)))
            return 1
        sname, rows = hit
    elif len(sheets) == 1:
        sname, rows = sheets[0]
        print('  唯一 sheet：{}'.format(sname))
    else:
        sname, rows = sheets[pick(sheets, '共 {} 个 sheet，请选择'.format(len(sheets)),
                                  lambda it: '{}   ({} 行 x {} 列)'.format(
                                      it[0], len(it[1]), max((len(x) for x in it[1]), default=0))) - 1]

    # 4) 表头行 + 校验
    hr = args.header_row
    if interactive:
        v = ask('表头在第几行', str(hr))
        hr = int(v) if v.isdigit() and int(v) >= 1 else hr
    if hr > len(rows):
        print('表头行 {} 超出范围（该文件仅 {} 行），无法转换'.format(hr, len(rows)))
        return 1

    hdr_raw = list(rows[hr - 1])
    while len(hdr_raw) > 1 and (hdr_raw[-1] is None or str(hdr_raw[-1]).strip() == ''):
        hdr_raw.pop()
    data = [r[:len(hdr_raw)] for r in rows[hr:]]

    ok, issues, hdr = check_header(hdr_raw, data)
    if ok:
        print('  表头校验通过：{} 列 / {} 行数据'.format(len(hdr), len(data)))
    else:
        print('\n[表格不规范] 检测到 {} 个问题：'.format(len(issues)))
        for i, m in enumerate(issues, 1):
            print('  {}. {}'.format(i, m))
        if not data:
            print('\n>>> 无法转换：表头下方没有数据。')
            return 1
        if ask('仍要继续？（自动用 col_N 补全空列名、重复列名加 _2 后缀）y/N', 'N').lower() != 'y':
            print('>>> 已取消。建议先在 Excel 中整理：首行放字段名、不空不重复。')
            return 0
    if len(data) > 2000:
        print('  ! 数据 {} 行，UNION ALL 链较长，注意数据库解析开销'.format(len(data)))

    # 5) 方言 / 输出形式
    if interactive:
        print('\nSQL 方言： ' + '   '.join('{}={}'.format(k, v['name']) for k, v in DIALECTS.items()))
        dk = ask('选择', '1')
        dia = DIALECTS.get(dk, DIALECTS['1'])
        wrap = 'plain' if ask('输出形式 1=CTE 包裹(可直接跑)  2=纯 UNION ALL 块', '1') == '2' else 'cte'
        tname = safe_name(ask('内联表名', 'HARDCODE')) if wrap == 'cte' else 'HARDCODE'
    else:
        dia, wrap, tname = DIALECTS[args.dialect], ('plain' if args.plain else 'cte'), 'HARDCODE'

    # 6) 生成并写出
    try:
        sql, nrows = build_sql(hdr, data, dia, tname, wrap)
    except RuntimeError as e:
        print('>>> 无法转换：{}'.format(e))
        return 1

    default_out = target.with_name('{}_{}_hardcode.sql'.format(target.stem, safe_name(sname)))
    if args.out:
        out = Path(args.out)
    else:
        out = Path(ask('\n输出文件路径', str(default_out)).strip('"'))
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(sql, encoding='utf-8')

    print('\n[OK] {} 行 x {} 列 -> {}   ({:,} KB)'.format(
        nrows, len(hdr), out, max(1, out.stat().st_size // 1024)))
    print('预览：')
    for l in [x for x in sql.split('\n') if x.startswith('SELECT ') and not x.startswith('SELECT *')][:2]:
        print('  ' + l[:170] + (' ...' if len(l) > 170 else ''))

    if interactive and ask('\n复制到剪贴板？(y/N)', 'N').lower() == 'y':
        if copy_clipboard(sql):
            print('已复制 {} 字符'.format(len(sql)))
    return 0


if __name__ == '__main__':
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print('\n已中断')
        sys.exit(130)
    except EOFError:
        print('\n输入结束（非交互环境请改用：python xlsx2sql.py <文件>）')
        sys.exit(1)
