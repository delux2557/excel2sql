"""跨平台复制到剪贴板（Windows / macOS / Linux）。"""
from __future__ import annotations

import platform
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import List, Optional

TIMEOUT = 60


def _linux_command() -> Optional[List[str]]:
    for candidate in (['wl-copy'], ['xclip', '-selection', 'clipboard'], ['xsel', '--clipboard', '--input']):
        if shutil.which(candidate[0]):
            return candidate
    return None


def copy_text(text: str) -> bool:
    """把文本放进系统剪贴板。任何失败都返回 False，不抛异常。"""
    system = platform.system()
    try:
        if system == 'Windows':
            return _copy_windows(text)
        if system == 'Darwin':
            subprocess.run(['pbcopy'], input=text.encode('utf-8'), check=True,
                           timeout=TIMEOUT, capture_output=True)
            return True
        cmd = _linux_command()
        if not cmd:
            return False
        subprocess.run(cmd, input=text.encode('utf-8'), check=True,
                       timeout=TIMEOUT, capture_output=True)
        return True
    except Exception:
        return False


def _copy_windows(text: str) -> bool:
    """PowerShell 的 Set-Clipboard 走管道时编码不可控，这里用临时文件 + -Encoding UTF8。"""
    tmp = None
    try:
        with tempfile.NamedTemporaryFile('w', suffix='.txt', delete=False, encoding='utf-8') as f:
            f.write(text)
            tmp = f.name
        safe = tmp.replace('"', '`"')
        subprocess.run(
            ['powershell', '-NoProfile', '-NonInteractive', '-Command',
             'Get-Content -Raw -Encoding UTF8 -LiteralPath "{}" | Set-Clipboard'.format(safe)],
            check=True, timeout=TIMEOUT, capture_output=True)
        return True
    finally:
        if tmp:
            Path(tmp).unlink(missing_ok=True)          # 异常路径也清理，避免残留
