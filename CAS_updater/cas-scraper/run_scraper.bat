@echo off
echo ========================================
echo CAS Online Form Field Scraper
echo ========================================
echo.

REM Check if Node.js is installed
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: Node.js is not installed or not in PATH
    echo Please install Node.js from https://nodejs.org/
    pause
    exit /b 1
)

REM Check for credentials
if "%CAS_EMAIL%"=="" (
    set /p CAS_EMAIL="Enter your CAS email: "
)
if "%CAS_PASSWORD%"=="" (
    set /p CAS_PASSWORD="Enter your CAS password: "
)

REM Check if node_modules exists
if not exist "node_modules" (
    echo Installing dependencies...
    call npm install
    if %errorlevel% neq 0 (
        echo ERROR: npm install failed
        pause
        exit /b 1
    )
    echo.
)

REM Run the scraper
echo Starting scraper...
echo.
call npx ts-node scraper.ts

echo.
echo ========================================
echo Scraping complete!
echo Check cas_form_fields.json for results
echo ========================================
pause
