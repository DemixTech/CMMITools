@echo off
REM Audit CAS Plan - Windows Batch Script (v1.1.0)
REM Validates _FieldMap against actual sheet content.
REM TARGET is read from cas-project-config.json (files.target).

setlocal enabledelayedexpansion

echo ============================================================
echo Audit CAS Plan (v1.1.0)
echo ============================================================
echo.

REM Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8+ from https://python.org
    pause
    exit /b 1
)

REM Check for openpyxl
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing required package: openpyxl
    pip install openpyxl
)

REM ---- Read TARGET from cas-project-config.json -----------------------------
REM This script lives at cas_helper\skills\skill-audit-cas-plan\, so the
REM project config is two levels up.
set "PROJECT_CONFIG=%~dp0..\..\cas-project-config.json"

if not exist "%PROJECT_CONFIG%" (
    echo ERROR: cas-project-config.json not found at "%PROJECT_CONFIG%"
    pause
    exit /b 1
)

for /f "usebackq delims=" %%I in (`python -c "import json,sys; print(json.load(open(r'%PROJECT_CONFIG%','r',encoding='utf-8'))['files']['target'])"`) do set "TARGET=%%I"

if "%TARGET%"=="" (
    echo ERROR: Could not read files.target from cas-project-config.json
    pause
    exit /b 1
)

REM Derive REPORT location: same folder as TARGET, audit_report_YYYYMMDD.txt
for %%F in ("%TARGET%") do set "TARGET_DIR=%%~dpF"
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value 2^>nul') do set "DT=%%I"
if "%DT%"=="" (
    REM fallback if wmic unavailable: use %date% (locale-dependent)
    set "DATE_TAG=%date:~-4,4%%date:~-10,2%%date:~-7,2%"
) else (
    set "DATE_TAG=%DT:~0,8%"
)
set "REPORT=%TARGET_DIR%audit_report_%DATE_TAG%.txt"

echo.
echo File:   %TARGET%
echo Report: %REPORT%
echo.

REM Ask for confirmation
set /p confirm="Run audit? (Y/N): "
if /i not "%confirm%"=="Y" (
    echo Operation cancelled.
    pause
    exit /b 0
)

REM Run the audit script
python "%~dp0audit_cas_plan.py" --file "%TARGET%" --report "%REPORT%"

if errorlevel 1 (
    echo.
    echo ERROR: Audit failed. Please check the error messages above.
) else (
    echo.
    echo SUCCESS: Audit completed.
    echo Report saved to: %REPORT%
    echo.
    echo Check Column G in _FieldMap sheet for results.
)

echo.
pause
endlocal
