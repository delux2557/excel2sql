@echo off
chcp 65001 >nul
setlocal

rem ============================================================
rem  excel2sql - 双击启动交互向导（扫描“当前目录”的 Excel/CSV）
rem  无需安装：直接用仓库内 src/ 的源码运行，兼容 embed 版 Python
rem  找不到 python 时的三种配置方式见文末 :nopython 提示
rem ============================================================

set "HERE=%~dp0"
set "PY="

rem --- 1) 本目录下的 python-path.txt（一行，写 python.exe 完整路径）---
if exist "%HERE%python-path.txt" set /p PY=<"%HERE%python-path.txt"
if defined PY if not exist "%PY%" set "PY="

rem --- 2) 环境变量 EXCEL2SQL_PYTHON ---
if not defined PY if defined EXCEL2SQL_PYTHON if exist "%EXCEL2SQL_PYTHON%" set "PY=%EXCEL2SQL_PYTHON%"

rem --- 3) 仓库内的虚拟环境 ---
if not defined PY if exist "%HERE%.venv\Scripts\python.exe" set "PY=%HERE%.venv\Scripts\python.exe"

rem --- 4) PATH 中的 python（跳过 Microsoft Store 的占位程序）---
if not defined PY for /f "delims=" %%i in ('where python 2^>nul ^| findstr /v /i "WindowsApps"') do if not defined PY set "PY=%%i"

rem --- 5) py 启动器 ---
if not defined PY for /f "delims=" %%i in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do if not defined PY set "PY=%%i"

rem --- 6) 常见安装位置（含本机自定义安装目录）---
if not defined PY for /d %%d in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if not defined PY if exist "%%~fd\python.exe" set "PY=%%~fd\python.exe"
if not defined PY for /d %%d in ("%ProgramFiles%\Python3*") do if not defined PY if exist "%%~fd\python.exe" set "PY=%%~fd\python.exe"
if not defined PY for /d %%d in ("%ProgramFiles(x86)%\Python3*") do if not defined PY if exist "%%~fd\python.exe" set "PY=%%~fd\python.exe"
if not defined PY for /d %%d in ("C:\Python3*") do if not defined PY if exist "%%~fd\python.exe" set "PY=%%~fd\python.exe"
if not defined PY for /d %%d in ("C:\Python3*") do if not defined PY if exist "%%~fd\python.exe" set "PY=%%~fd\python.exe"
if not defined PY for /d %%d in ("C:\Python3*") do if not defined PY if exist "%%~fd\python.exe" set "PY=%%~fd\python.exe"

if not defined PY goto :nopython

rem --- 版本检查：需要 Python 3.9+ ---
"%PY%" -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 9)" 2>nul
if errorlevel 9 goto :oldpython

"%PY%" -c "import sys,runpy; sys.path.insert(0, r'%HERE%src'); runpy.run_module('excel2sql', run_name='__main__')" %*
set "RC=%ERRORLEVEL%"
echo.
pause
exit /b %RC%

:nopython
echo.
echo [!] 没有找到 python.exe（需要 Python 3.9 或更高版本）。
echo.
echo     任选一种方式指定 python，然后重新双击本文件：
echo       1. 在本文件所在目录新建 python-path.txt，里面写一行 python.exe 的完整路径
echo          例：C:\Python312\python.exe
echo          （记事本“另存为”时编码选 ANSI，避免带上 BOM）
echo       2. 设置环境变量 EXCEL2SQL_PYTHON 指向 python.exe
echo       3. 把 python 加入 PATH；或在本目录执行：python -m venv .venv
echo.
echo     还没装 python？https://www.python.org/downloads/
echo     装完记得执行一次：pip install openpyxl
echo.
pause
exit /b 1

:oldpython
echo.
echo [!] 找到的 python 版本低于 3.9，excel2sql 需要 3.9 及以上。
echo     当前使用：%PY%
echo     请修改 python-path.txt 或 EXCEL2SQL_PYTHON 指向新版本。
echo.
pause
exit /b 1
