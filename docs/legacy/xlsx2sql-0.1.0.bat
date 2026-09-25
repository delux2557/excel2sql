@echo off
chcp 65001 >nul
setlocal
rem 双击运行：扫描当前目录的 Excel/CSV，交互式生成硬编码 SQL
rem 若提示找不到 python，把下面这一行改成你的 python.exe 完整路径
set "PY=python"

where %PY% >nul 2>nul
if errorlevel 1 set "PY=C:\Python312\python.exe"

"%PY%" "%~dp0xlsx2sql.py" %*
echo.
pause
