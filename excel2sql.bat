@echo off
chcp 65001 >nul
setlocal

rem ===========================================================
rem  excel2sql - double-click launcher (interactive wizard)
rem  Scans Excel/CSV in the CURRENT directory and generates SQL.
rem  If python is not found, set PY to the full path of python.exe
rem  Works with embeddable Python too (does NOT rely on PYTHONPATH).
rem ===========================================================
set "PY=python"

where %PY% >nul 2>nul
if errorlevel 1 set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not exist "%PY%" (
    echo [!] python not found. Edit this file and set PY to python.exe path.
    pause
    exit /b 1
)

"%PY%" -c "import sys,runpy; sys.path.insert(0, r'%~dp0src'); runpy.run_module('excel2sql', run_name='__main__')" %*
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" if not "%RC%"=="2" echo [exit code %RC%]
echo.
pause
exit /b %RC%
