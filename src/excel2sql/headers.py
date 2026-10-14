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


# 「不可见但不属于空白」的字符：`strip()` 去不掉、肉眼也看不见。
# 留着有两个后果（2026-10-14 实测）：
#   1. 列名变成"看不见的名字"，产物里是 `[]`，谁也不知道那列叫什么；
#   2. `is_blank()` 认不出它（`str.isspace()` 对它们返回 False），于是 `col_N`
#      补名逻辑不触发 —— 一个"看着没名字"的列表头其实被当成了正常列名。
# 这类字符在任何语言里都不是列名该有的内容，直接删除。
_INVISIBLE = (
    '\u00ad'                                        # 软连字符
    '\u200b\u200c\u200d\u2060'                      # 零宽空格 / 零宽不连字 / 零宽连字 / 词连接符
    '\u200e\u200f\u202a\u202b\u202c\u202d\u202e'    # 方向标记与方向嵌入（不可见，且会扰乱显示顺序）
    '\u3164'                                        # 谚文填充符（常被当空格用，但 isspace() 不认）
    '\ufeff'                                        # BOM / 零宽不换行空格
)
_DELETE = dict.fromkeys(map(ord, _INVISIBLE))
# C0/C1 控制字符里**不是空白**的那些（NUL、BEL、DEL…）同样删掉；
# 是空白的（\t \n \v \f \r、NEL）留给 split() 折叠成空格 —— 它们表示"这里原本有分隔"。
# ★ 变体选择符 U+FE0E/U+FE0F 与 emoji 序列不在删除范围内：它们修饰的是前面那个字符的
#   呈现方式，属于**有意的内容**，删掉会把 🗃️ 变成单色字形。
_DELETE.update(dict.fromkeys(
    c for c in list(range(0x00, 0x20)) + list(range(0x7f, 0xa0)) if not chr(c).isspace()))


def clean_name(v) -> str:
    """把列名里的不可见噪声归一：删掉零宽/控制字符，再把任意空白折叠成单个空格。

    为什么是"单空格"而不是"直接删掉"（2026-10-14 定）：
    换行往往表示"这里原本有个分隔"，`订\\n单` 保留成 `订 单` 比粘成 `订单` 更保守；
    列名带空格在 SQL 里完全合法（工具本来就会按方言加引号）。

    只用于**表头/列名**。数据格里的换行是合法的多行文本，必须原样保留
    （有 `CHAR(10)` 拼接机制专门支持），绝不能走到这里来。
    """
    if v is None:
        return ''
    s = str(v).translate(_DELETE)
    return ' '.join(s.split())          # 任意空白 -> 单空格，顺带去掉首尾


def changed_names(raw: Sequence[object], cleaned: Sequence[object],
                  limit: int = 5) -> List[str]:
    """列出被清洗过的列名，形如 `[第2列] '订\\n单' → '订 单'`（供调用方提示用户）。

    ★ 清洗必须"看得见"：静默修正和静默走偏只差一步 —— 用户有权知道自己的表头被改过。

    比较的是「原始文本」与「清洗结果」，不能用 clean_name 两边都洗一遍再比
    （那样恒相等，等于什么都没比出来）。
    """
    pairs = []
    for i, (a, b) in enumerate(zip(raw, cleaned)):
        original = '' if a is None else str(a)
        if original == str(b):
            continue
        pairs.append('[第{}列] {!r} → {!r}'.format(i + 1, a, b))

    if len(pairs) <= limit:
        return pairs
    return pairs[:limit] + ['…（另有 {} 处）'.format(len(pairs) - limit)]


def is_blank(v) -> bool:
    """单元格是否为空。

    ★ 用 `clean_name()` 判空而不是 `str.strip()`：零宽字符、BOM 这类"看不见"的内容
      在 `strip()` 眼里是非空的，于是它们会被当成**有名字的列**，
      既得不到 `col_N` 补名，也不会被当成空列报出来（2026-10-14 实测）。
    """
    if v is None:
        return True
    if isinstance(v, str):
        return clean_name(v) == ''
    return False


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
        base = clean_name(raw)
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
        k = clean_name(h).lower()
        if not k:
            continue                    # 空列名已由上面的"列为空"报出，别在这里复述
        if k in seen:
            dup.append(i)
        seen.add(k)
    if dup:
        issues.append('列名重复：{}'.format(', '.join(str(names[i]) for i in dup[:8])))

    like_values = sum(1 for v in names if looks_like_value(v))
    if like_values / len(names) >= LIKE_DATA_RATIO:
        issues.append('表头行全是数字/日期，不像字段名（表头可能不在这一行）')

    return HeaderCheck(not issues, issues, list(names))
