"""表头校验与修复。

把「表不规范」的判断从主流程里独立出来，便于单元测试与复用。
"""
from __future__ import annotations

import datetime
import re
from dataclasses import dataclass
from typing import List, Sequence

_NUMBER_RE = re.compile(r'-?\d+(\.\d+)?')
_DATE_RE = re.compile(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?)?')

# 表头行里"看起来像数据"的判断上限：超过这个比例就认为表头不在该行
LIKE_DATA_RATIO = 1.0


def is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == '')


def looks_like_value(v) -> bool:
    """单元格内容是否像"数据"而不是"字段名"。"""
    if isinstance(v, bool):
        return False                    # bool 是 int 子类，需先排除，否则 Y/N 列会被误判
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, (datetime.datetime, datetime.date)):
        return True
    s = str(v).strip()
    return bool(_NUMBER_RE.fullmatch(s)) or bool(_DATE_RE.fullmatch(s))


@dataclass
class HeaderCheck:
    ok: bool
    issues: List[str]
    names: List[str]                # 原始表头（修复请另调 repair()）


def repair(names: Sequence[object]) -> List[str]:
    """生成全局唯一、非空的列名：空列 -> col_N；重名 -> 名字_2、名字_3……"""
    out: List[str] = []
    seen = set()
    for i, raw in enumerate(names):
        base = str(raw).strip() if raw is not None else ''
        if not base:
            base = 'col_{}'.format(i + 1)
        name = base
        n = 1
        while name.lower() in seen:      # 循环到真正不冲突（避免 a, a_2, a 修出新的重复）
            n += 1
            name = '{}_{}'.format(base, n)
        seen.add(name.lower())
        out.append(name)
    return out


def check(names: Sequence[object], data_rows: Sequence[Sequence[object]]) -> HeaderCheck:
    """校验表头，返回问题清单。不修改入参。"""
    issues: List[str] = []
    names = list(names)

    if not names:
        return HeaderCheck(False, ['表头行为空或整行无内容'], [])

    if not data_rows:
        issues.append('表头下方没有任何数据行')

    holes = [i for i, h in enumerate(names) if is_blank(h)]
    if holes:
        issues.append('表头第 {} 列为空（共 {} 处）'.format(
            ','.join(str(i + 1) for i in holes[:8]), len(holes)))

    seen = set()
    dup = []
    for i, h in enumerate(names):
        k = str(h).strip().lower()
        if k in seen:
            dup.append(i)
        seen.add(k)
    if dup:
        issues.append('列名重复：{}'.format(', '.join(str(names[i]) for i in dup[:8])))

    like_values = sum(1 for v in names if looks_like_value(v))
    if like_values / len(names) >= LIKE_DATA_RATIO:
        issues.append('表头行全是数字/日期，不像字段名（表头可能不在这一行）')

    return HeaderCheck(not issues, issues, list(names))
