"""excel2sql —— 把 Excel / CSV 的每一行转成硬编码 SQL。

对外主要接口：
    excel2sql.reader.read_sheets      读取工作簿
    excel2sql.sqlgen.SqlOptions       生成选项
    excel2sql.sqlgen.build_sql        生成 SQL 文本
    excel2sql.sqlgen.write_sql        流式写文件
    excel2sql.cli.main                命令行入口
"""
from __future__ import annotations

__version__ = '0.3.0'
__all__ = ['__version__']
