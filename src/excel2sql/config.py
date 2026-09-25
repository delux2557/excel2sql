"""配置文件（ssh config 风格）：把常用选择固化成默认值，不必每次交互输入。

查找顺序（找到即用，后者可被前者覆盖）：
    1. --config 指定的文件
    2. ./excel2sql.ini            （放在数据目录，项目级）
    3. ~/.excel2sql/config.ini    （用户级）
    4. 内置默认值

命令行参数优先级最高，会覆盖配置文件。
"""
from __future__ import annotations

import configparser
import datetime
from dataclasses import dataclass, fields
from pathlib import Path
from typing import List, Optional, Tuple

CONFIG_NAME = 'excel2sql.ini'
USER_CONFIG = Path.home() / '.excel2sql' / 'config.ini'
DEFAULT_FILENAME = '{name}_{sheet}_hardcode.sql'
TRUE_WORDS = {'1', 'true', 'yes', 'y', 'on'}


class ConfigError(Exception):
    """配置文件无法读取。"""


@dataclass
class Settings:
    """可持久化的默认值。字段名即 ini 键名。"""

    # [output]
    dialect: str = 'sqlserver'          # sqlserver | mysql | oracle | postgresql
    format: str = 'union'               # union | insert
    wrap: str = 'cte'                   # cte | plain
    table: str = 'HARDCODE'
    batch_size: int = 500
    output_dir: str = 'sub'             # source | sub | <路径>
    output_subdir: str = 'excel2sql-out'
    overwrite: bool = False             # False = 同名时自动加 -2/-3 后缀
    encoding: str = 'utf-8'             # utf-8 | utf-8-sig
    filename: str = DEFAULT_FILENAME    # 支持 {name} {sheet} {date}

    # [data]
    empty_as_null: bool = False
    all_string: bool = False
    # CSV 等无类型信息的输入：整列都是数字时按数字输出（否则 CSV 所有列都会被字符串化）
    infer_types: bool = False
    header_row: str = 'auto'            # auto | 1 | 2 ...

    # [ui]
    ask_advanced: bool = False          # True = 逐项询问方言/格式/编码等
    copy_clipboard: bool = False

    # ---------------------------------------------------------------- 派生
    def default_output(self, src: Path, sheet: str) -> Path:
        """按配置算出默认输出路径（不负责防冲突）。"""
        import re

        safe_sheet = re.sub(r'[^\w\u4e00-\u9fff]+', '_', str(sheet)).strip('_') or 'sheet'
        name = (self.filename or DEFAULT_FILENAME).format(
            name=src.stem, sheet=safe_sheet,
            date=datetime.date.today().isoformat())
        if not name.lower().endswith(('.sql', '.txt')):
            name += '.sql'

        mode = (self.output_dir or 'source').strip()
        if mode in ('', 'source', 'same'):
            folder = src.parent
        elif mode == 'sub':
            folder = src.parent / (self.output_subdir or 'excel2sql-out')
        else:
            p = Path(mode).expanduser()
            folder = p if p.is_absolute() else (src.parent / p)
        return folder / name

    def progress_hint(self) -> str:
        """一行摘要，用于交互时的『当前配置』提示。"""
        return ('{} / {} / {} / 表名 {}{}{}'.format(
            self.dialect, self.format, self.wrap, self.table,
            ' / 空串=NULL' if self.empty_as_null else '',
            ' / 编码 ' + self.encoding if self.encoding != 'utf-8' else ''))


# ---------------------------------------------------------------- 读写
def _as_bool(value: str, key: str, warnings: List[str]) -> bool:
    v = str(value).strip().lower()
    if v in TRUE_WORDS:
        return True
    if v in ('0', 'false', 'no', 'n', 'off', ''):
        return False
    warnings.append('配置项 {}={} 不是布尔值，已按 false 处理'.format(key, value))
    return False


def _apply(settings: Settings, path: Path, warnings: List[str]) -> None:
    parser = configparser.ConfigParser(interpolation=None)
    try:
        parser.read(path, encoding='utf-8')
    except (configparser.Error, OSError, UnicodeDecodeError) as e:
        raise ConfigError('配置文件解析失败：{}（{}）'.format(path, e))

    values = {}
    for section in parser.sections():
        for key, value in parser.items(section):
            values[key.strip().lower()] = value.strip()

    for f in fields(settings):
        if f.name not in values:
            continue
        raw = values[f.name]
        try:
            if f.type is bool or isinstance(getattr(settings, f.name), bool):
                setattr(settings, f.name, _as_bool(raw, f.name, warnings))
            elif isinstance(getattr(settings, f.name), int):
                setattr(settings, f.name, int(raw))
            else:
                setattr(settings, f.name, raw)
        except (TypeError, ValueError):
            warnings.append('配置项 {}={} 无法解析，已用默认值 {}'.format(
                f.name, raw, getattr(settings, f.name)))


def search_paths(cwd: Optional[Path] = None) -> List[Path]:
    return [Path(cwd or Path.cwd()) / CONFIG_NAME, USER_CONFIG]


def load(explicit: Optional[str] = None, cwd: Optional[Path] = None
         ) -> Tuple[Settings, Optional[Path], List[str]]:
    """返回 (设置, 命中的配置文件, 警告列表)。"""
    settings, warnings = Settings(), []
    if explicit:
        p = Path(explicit).expanduser()
        if not p.is_file():
            raise ConfigError('指定了 --config，但文件不存在：{}'.format(p))
        _apply(settings, p, warnings)
        return settings, p, warnings

    for candidate in search_paths(cwd):
        if candidate.is_file():
            _apply(settings, candidate, warnings)
            return settings, candidate, warnings
    return settings, None, warnings


def unique_path(path: Path) -> Path:
    """防冲突：同名文件已存在时追加 -2、-3……"""
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    for i in range(2, 1000):
        candidate = path.with_name('{}-{}{}'.format(stem, i, suffix))
        if not candidate.exists():
            return candidate
    raise ConfigError('同名文件过多，请手工指定输出路径：{}'.format(path))


TEMPLATE = """# excel2sql 配置文件
#
# 位置（任选其一，靠前的优先）：
#   1) 与数据同目录的 excel2sql.ini      —— 项目级
#   2) ~/.excel2sql/config.ini           —— 用户级
# 命令行参数优先级最高，会覆盖本文件。
# 布尔值写 true / false。

[output]
# 目标数据库方言: sqlserver | mysql | oracle | postgresql
dialect = sqlserver
# 输出格式: union = UNION ALL 内联表, insert = INSERT INTO ... VALUES
format = union
# union 模式: cte = 外面包 WITH ... AS (...), plain = 纯 UNION ALL 块
wrap = cte
# 内联表名
table = HARDCODE
# insert 模式每批行数
batch_size = 500

# 输出到哪里:
#   source = 与源文件同目录
#   sub    = 源文件目录下新建子目录（名字见 output_subdir）
#   <路径> = 指定目录（相对路径按源文件目录解析）
output_dir = sub
output_subdir = excel2sql-out
# 同名文件已存在时: false = 自动加 -2/-3 后缀, true = 直接覆盖
overwrite = false
# 输出文件编码: utf-8 | utf-8-sig（老版 SSMS 中文乱码时用后者）
encoding = utf-8
# 输出文件名，可用占位符 {name} 源文件名, {sheet} 工作表名, {date} 当天日期
filename = {name}_{sheet}_hardcode.sql

[data]
# 空字符串是否当成 NULL
empty_as_null = false
# 所有列强制按字符串输出
all_string = false
# CSV 等无类型信息的输入：整列都是数字时按数字输出
# （CSV 读出来全是文本，不开这项的话所有列都会按字符串输出）
# 注意：带前导零的列（如 "007"）不会被转换，以免丢信息
infer_types = false
# 表头行号: auto = 自动识别（交互里会打印预览让你确认）, 或写 1 / 2 / ...
header_row = auto

[ui]
# true = 交互时逐项询问方言/格式/表名/编码/剪贴板；false = 直接用本文件设置
ask_advanced = false
# 生成后自动复制到剪贴板
copy_clipboard = false
"""


def write_template(path: Path, overwrite: bool = False) -> Optional[Path]:
    """写一份带注释的配置模板；已存在且未指定覆盖时返回 None。"""
    if path.exists() and not overwrite:
        return None
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(TEMPLATE, encoding='utf-8')
    return path
