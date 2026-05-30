@echo off
REM Setup Base CAS Plan - Windows Batch Script (v2.1.0)
REM Uses _FieldMap from target file.
REM SOURCE and TARGET are read from cas-project-config.json (files.source / files.target).

setlocal enabledelayedexpansion

echo ============================================================
echo Setup Base CAS Plan (v2.1.0)
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

REM ---- Read SOURCE and TARGET from cas-project-config.json ------------------
REM This script lives at cas_helper\skills\skill-setup-base-cas-plan\, so the
REM project config is two levels up.
set "PROJECT_CONFIG=%~dp0..\..\cas-project-config.json"

if not exist "%PROJECT_CONFIG%" (
    echo ERROR: cas-project-config.json not found at "%PROJECT_CONFIG%"
    pause
    exit /b 1
)

for /f "usebackq delims=" %%I in (`python -c "import json; print(json.load(open(r'%PROJECT_CONFIG%','r',encoding='utf-8'))['files']['source'])"`) do set "SOURCE=%%I"
for /f "usebackq delims=" %%I in (`python -c "import json; print(json.load(open(r'%PROJECT_CONFIG%','r',encoding='utf-8'))['files']['target'])"`) do set "TARGET=%%I"

if "%SOURCE%"=="" (
    echo ERROR: Could not read files.source from cas-project-config.json
    pause
    exit /b 1
)
if "%TARGET%"=="" (
    echo ERROR: Could not read files.target from cas-project-config.json
    pause
    exit /b 1
)

REM Derive REPORT location: same folder as TARGET, setup_report_YYYYMMDD.txt
for %%F in ("%TARGET%") do set "TARGET_DIR=%%~dpF"
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value 2^>nul') do set "DT=%%I"
if "%DT%"=="" (
    set "DATE_TAG=%date:~-4,4%%date:~-10,2%%date:~-7,2%"
) else (
    set "DATE_TAG=%DT:~0,8%"
)
set "REPORT=%TARGET_DIR%setup_report_%DATE_TAG%.txt"

echo.
echo Source: %SOURCE%
echo Target: %TARGET%
echo Report: %REPORT%
echo.

REM Ask for confirmation
set /p confirm="Proceed with setup? (Y/N): "
if /i not "%confirm%"=="Y" (
    echo Operation cancelled.
    pause
    exit /b 0
)

REM Run the setup script with backup and report
python "%~dp0setup_cas_plan.py" --source "%SOURCE%" --target "%TARGET%" --backup --report "%REPORT%"

if errorlevel 1 (
    echo.
    echo ERROR: Setup failed. Please check the error messages above.
) else (
    echo.
    echo SUCCESS: CAS Plan setup completed.
    echo Report saved to: %REPORT%
    echo.
    echo Review the source file for yellow/red marked anomalies.
)

echo.
pause
endlocal
