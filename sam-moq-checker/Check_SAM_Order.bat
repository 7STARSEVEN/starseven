@echo off
title SAM MOQ Order Checker
cd /d "%~dp0"

rem ---- find Python ----
set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY ( python -c "import sys" >nul 2>nul && set "PY=python" )
if not defined PY (
  echo.
  echo  Python is not installed on this computer.
  echo  Install it once from python.org - tick "Add python.exe to PATH" during install.
  echo  Then double-click this file again.
  echo.
  pause
  exit /b 1
)

rem ---- first run only: install the two free Excel readers (needs internet once) ----
%PY% -c "import openpyxl, xlrd, pypdf" >nul 2>nul
if errorlevel 1 (
  echo  First run: installing openpyxl, xlrd and pypdf - needs internet, only once...
  %PY% -m pip install --user --quiet openpyxl xlrd pypdf
)

rem ---- run the check (an order file dragged onto this icon is used directly) ----
%PY% "%~dp0sam_moq_check.py" %*
echo.
pause
