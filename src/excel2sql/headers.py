"""表头校验与修复。

把「表不规范」的判断从主流程里独立出来，便于单元测试与复用。
"""
from __future__ import annotations

import datetime
import re
from dataclasses import dataclass
from typing import List, Sequence, Tuple

_NUMBER_RE = re.compile(r'-?\d+(\.\d+)?')
_DATE_RE = re.compile(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?)?')

# 表头行里"看起来像数据"的判断上限：超过这个比例就认为表头不在该行
LIKE_DATA_RATIO = 1.0
# 自动识别表头时，低于这个分数就认为没有明显表头（纯数字行大约只有 1.0 分）
MIN_SCORE = 1.5
# 候选表头行与第 1 行的分差小于此值时，保守选第 1 行。
# 实测（2026-09-25）：表头里出现**一个空列名**就会让 filled 掉 0.5×0.5=0.25，
# 而"位置靠前"的偏好只有 0.05/行 —— 分差远小于这个扣分幅度时必然翻转，
# 结果是**静默丢一行数据**（退出码仍为 0）。宁可保守，也不悄悄丢数据。
AMBIGUOUS_MARGIN = 0.3
# 允许被跳过的行"有多空"。表头之上的标题/说明行通常只有**一个**非空单元格
# （`['销售报表']`、`['某部门 2024 年度数据导出表', None, None]`），
# 而被误判成"标题行"的**真表头**是填满的。
# 实测（2026-09-25）：表头 `装运方式 | 2024` 因为列名是数字，得分被压到 1.75，
# 而数据行 `一级 | 消费者` 拿到 3.95 —— 分差 2.2，单靠阈值拦不住；
# 但它有 2 个非空格子，一眼就不是标题。
# ★ 这里用"非空单元格**个数**"而不是"占比"：单列表的标题行（`['销售报表']`）
#   占比同样是 100%，用占比会把合法的"标题 + 表头"结构误伤。
SKIPPED_MAX_CELLS = 1


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


def score_header_row(row: Sequence[object], next_row: Sequence[object]) -> float:
    """给"这一行像不像表头"打分：越高越像。

    表头的典型特征：几乎全是文本、互不重复、下方紧跟数据行；
    数据行则相反（数字/日期占比高）。
    """
    cells = [v for v in row if not is_blank(v)]
    if not cells:
        return -10.0

    filled = len(cells) / max(1, len(row))
    value_like = sum(1 for v in cells if looks_like_value(v)) / len(cells)
    text_like = 1.0 - value_like
    unique = len({str(v).strip().lower() for v in cells}) / len(cells)
    short = sum(1 for v in cells if isinstance(v, str) and 0 < len(v) <= 40) / len(cells)

    below = [v for v in next_row if not is_blank(v)] if next_row else []
    next_is_data = (sum(1 for v in below if looks_like_value(v)) / len(below)) if below else 0.0

    return (2.0 * text_like + 1.0 * unique + 0.5 * short
            + 1.5 * next_is_data + 0.5 * filled - 2.0 * value_like)


def nonblank_count(row: Sequence[object]) -> int:
    """非空单元格个数。标题/说明行通常是 1（往往只有第一格有内容）。"""
    return sum(1 for v in row if not is_blank(v))


def detect(rows: Sequence[Sequence[object]], max_scan: int = 15) -> Tuple[int, str]:
    """自动识别表头行。

    返回 (1 基行号, 说明)。识别不出来时退回第 1 行并说明原因。

    两条保护（都是 2026-09-25 实测踩出来的）：
    1. **分差过小时保守选第 1 行**（见 AMBIGUOUS_MARGIN）——
       表头里有一个空列名就足以把第 1 行压下去；
    2. **被跳过的行只要不止一个非空格就不跳**（见 SKIPPED_MAX_CELLS）——
       标题/说明行总是稀疏的，而"表头被误判成标题"的那种行是填满的。
    两条都宁可保守，也不静默丢数据。
    """
    if not rows:
        return 1, '工作表为空'
    limit = min(len(rows), max(1, max_scan))
    scores: List[float] = []
    best_row, best_score = 1, None
    for idx in range(limit):
        nxt = rows[idx + 1] if idx + 1 < len(rows) else []
        score = score_header_row(rows[idx], nxt) - 0.05 * idx
        scores.append(score)
        if best_score is None or score > best_score:
            best_row, best_score = idx + 1, score

    if best_score is None or best_score < MIN_SCORE:
        return 1, '没有明显像表头的行（前 {} 行都更像数据），默认第 1 行'.format(limit)

    if best_row > 1:
        margin = best_score - scores[0]
        # 被跳过的行若"不止一个非空格"，就不像标题/说明行 —— 可能是真表头，不跳
        dense = [i for i in range(1, best_row)
                 if nonblank_count(rows[i - 1]) > SKIPPED_MAX_CELLS]
        if margin < AMBIGUOUS_MARGIN or dense:
            why = ('第 {} 行只高 {:.2f} 分，证据不足'.format(best_row, margin)
                   if margin < AMBIGUOUS_MARGIN
                   else '第 {} 行填得较满，不像标题/说明行'
                        .format('、'.join(str(i) for i in dense[:5])))
            return 1, (
                '第 1 行最像表头（{}，保守取第 1 行；若表头确实在第 {} 行，'
                '请用 --header-row {} 明确指定）'.format(why, best_row, best_row))

        return best_row, '跳过前 {} 行，第 {} 行最像表头'.format(best_row - 1, best_row)

    return 1, '第 1 行最像表头（文本列名 + 下方是数据）'


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
