@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM ==========================================================
REM NKU passport generator launcher for Windows
REM Put this BAT file in the same folder with:
REM - generate_passports_v7_Onln.py
REM - passport_template_v2.docx
REM - google_service_account.json for online mode
REM - one .xlsx file for local mode
REM ==========================================================

cd /d "%~dp0"

set "SCRIPT_NAME=generate_passports_v7_Onln.py"
set "TEMPLATE_NAME=passport_template_v2.docx"

echo ==============================================
echo  NKU Passport Generator
echo ==============================================
echo.
echo Work folder:
echo %CD%
echo.

if not exist "%SCRIPT_NAME%" (
    echo ERROR: %SCRIPT_NAME% not found near this BAT file.
    echo Put this BAT file in the same folder with the Python script.
    echo.
    pause
    exit /b 1
)

if not exist "%TEMPLATE_NAME%" (
    echo ERROR: %TEMPLATE_NAME% not found near this BAT file.
    echo Put the Word template near the Python script.
    echo.
    pause
    exit /b 1
)

set "PYTHON_CMD="

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=python"
    )
)

if "%PYTHON_CMD%"=="" (
    echo ERROR: Python was not found.
    echo Install Python 3 and enable Add Python to PATH during installation.
    echo.
    pause
    exit /b 1
)

echo Python command: %PYTHON_CMD%
echo.

if not exist ".venv\Scripts\python.exe" (
    echo .venv was not found.
    echo Creating .venv in this folder...
    %PYTHON_CMD% -m venv .venv

    if errorlevel 1 (
        echo.
        echo ERROR: Could not create .venv.
        echo Check Python installation.
        echo.
        pause
        exit /b 1
    )
)

set "VENV_PY=.venv\Scripts\python.exe"

echo Upgrading pip...
"%VENV_PY%" -m pip install --upgrade pip

echo.
echo Installing/checking dependencies...
"%VENV_PY%" -m pip install python-docx openpyxl gspread google-auth

if errorlevel 1 (
    echo.
    echo ERROR: Could not install dependencies.
    echo Check internet connection and Python/pip installation.
    echo.
    pause
    exit /b 1
)

echo.
echo Select mode:
echo 1 - Online Google Sheets
echo 2 - Local Excel file near the script
echo.

:choose_mode
set /p MODE="Enter 1 or 2 and press Enter: "

if "%MODE%"=="1" goto online_mode
if "%MODE%"=="2" goto local_mode

echo Wrong input. Enter 1 or 2.
goto choose_mode

:online_mode
echo.
if not exist "google_service_account.json" (
    echo ERROR: google_service_account.json not found near this BAT file.
    echo This file is required for Google Sheets online mode.
    echo.
    pause
    exit /b 1
)

echo Paste Google Sheet URL.
echo Example:
echo https://docs.google.com/spreadsheets/d/.../edit?gid=...
echo.
set /p SHEET_URL="URL: "

if "%SHEET_URL%"=="" (
    echo.
    echo ERROR: URL is empty.
    echo.
    pause
    exit /b 1
)

echo.
echo Starting online mode...
echo.
"%VENV_PY%" "%SCRIPT_NAME%" "%SHEET_URL%"
goto finish

:local_mode
echo.
echo Starting local Excel mode.
echo The script will search for an Excel file near itself.
echo.
"%VENV_PY%" "%SCRIPT_NAME%"
goto finish

:finish
echo.
echo Done.
pause
exit /b 0
