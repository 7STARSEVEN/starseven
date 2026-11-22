@echo off
setlocal
cd /d "%~dp0"
title Canara passbook cash report

REM ---------------------------------------------------------------
REM  If your e-passbook PDFs open with a password, type it after =
REM  (no spaces), e.g.   set PDF_PASSWORD=ABCD1234
REM ---------------------------------------------------------------
set PDF_PASSWORD=

set PYTHONIOENCODING=utf-8
set PY=python
where python >nul 2>nul || set PY=py
%PY% --version >nul 2>nul
if errorlevel 1 (
    echo Python is not installed. Install it from python.org ^(tick "Add Python to PATH"^) and run this again.
    pause
    exit /b 1
)

%PY% -c "import pdfplumber, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Installing pdfplumber and openpyxl - one time only...
    if exist wheels (
        %PY% -m pip install --no-index --find-links wheels -r requirements.txt
    ) else (
        %PY% -m pip install -r requirements.txt
    )
)

if not exist pdfs mkdir pdfs
%PY% passbook_report.py %*
echo.

REM open the newest report
for /f "delims=" %%F in ('dir /b /o-d "output\Cash_Report_*.xlsx" 2^>nul') do (
    start "" "output\%%F"
    goto :done
)
:done
pause
