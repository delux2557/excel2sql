"""测试包：把 src/ 加入 import 路径，使「未安装包」时也能跑测试。

（unittest discover 要求 start_dir 可导入，因此这里保留 __init__.py。）
"""
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parents[1] / 'src'
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))
