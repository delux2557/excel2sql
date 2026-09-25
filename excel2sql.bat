@echo off
chcp 65001 >nul
setlocal

rem ============================================================
rem  excel2sql - double-click launcher (interactive wizard)
rem
rem  Scans Excel/CSV in the CURRENT directory and generates SQL.
rem  No install needed: runs src/ from this folder via runpy.
rem
rem  NOTE: this file is intentionally ASCII-only. Chinese text in a
rem  .bat read under a mismatched code page can shift the parser and
rem  make it execute fragments of the next line. All user-facing
rem  messages are printed by Python instead.
rem ============================================================

set "HERE=%~dp0"
set "PY="

rem --- 0) hard-code python.exe here if you prefer (uncomment + edit) ---
rem set "PY=C:\Python312\python.exe"

rem --- 1) python-path.txt next to this file (single line: python.exe path) ---
if exist "%HERE%python-path.txt" for /f "usebackq delims=" %%i in ("%HERE%python-path.txt") do if not defined PY set "PY=%%i"
if defined PY if not exist "%PY%" set "PY="

rem --- 2) environment variable EXCEL2SQL_PYTHON ---
if not defined PY if defined EXCEL2SQL_PYTHON set "PY=%EXCEL2SQL_PYTHON%"
if defined PY if not exist "%PY%" set "PY="

rem --- 3) virtualenv inside this repo ---
if not defined PY if exist "%HERE%.venv\Scripts\python.exe" set "PY=%HERE%.venv\Scripts\python.exe"

rem --- 4) common installation folders ---
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
if not defined PY if exist "C:\Python312\python.exe" set "PY=C:\Python312\python.exe"
if not defined PY if exist "C:\Python311\python.exe" set "PY=C:\Python311\python.exe"
if not defined PY if exist "C:\Python312\python.exe" set "PY=C:\Python312\python.exe"

rem --- 5) python on PATH (skip the Microsoft Store stub) ---
if not defined PY for /f "delims=" %%i in ('where python 2^>nul ^| findstr /i /v WindowsApps') do if not defined PY set "PY=%%i"

rem --- 6) py launcher ---
if not defined PY for /f "delims=" %%i in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do if not defined PY set "PY=%%i"

if not defined PY goto nopython

rem --- need Python 3.9 or newer ---
"%PY%" -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 9)" >nul 2>nul
if errorlevel 9 goto oldpython

rem --- run excel2sql from this folder (no install, no PYTHONPATH) ---
"%PY%" -c "import sys,runpy; sys.path.insert(0, r'%HERE%src'); runpy.run_module('excel2sql', run_name='__main__')" %*
set RC=%ERRORLEVEL%
echo.
pause
exit /b %RC%

:nopython
echo.
echo [!] python.exe not found (Python 3.9+ required).
echo.
echo     Pick ONE of these, then double-click this file again:
echo       1. create python-path.txt in this folder containing one line:
echo          the full path to python.exe
echo       2. set environment variable EXCEL2SQL_PYTHON to that path
echo       3. add python to PATH, or run here:  python -m venv .venv
echo.
echo     Python download: https://www.python.org/downloads/
echo     After install, run once:  pip install openpyxl
echo.
pause
exit /b 1

:oldpython
echo.
echo [!] The python found is older than 3.9; excel2sql needs 3.9+.
echo     in use: %PY%
echo     Fix python-path.txt or EXCEL2SQL_PYTHON to point at a newer one.
echo.
pause
exit /b 1
