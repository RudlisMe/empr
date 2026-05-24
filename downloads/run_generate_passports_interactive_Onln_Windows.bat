@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion

REM ==========================================================
REM Интерактивный запуск генератора паспортов НКУ для Windows
REM Файл должен лежать в той же папке, что и:
REM - generate_passports_v7_Onln.py
REM - passport_template_v2.docx
REM - google_service_account.json для онлайн-режима
REM - Excel-файл для локального режима, если запускается локально
REM ==========================================================

cd /d "%~dp0"

echo ==============================================
echo  Генератор паспортов НКУ
echo ==============================================
echo.
echo Папка запуска:
echo %CD%
echo.

if not exist "generate_passports_v7_Onln.py" (
    echo ОШИБКА: generate_passports_v7_Onln.py не найден рядом с этим BAT-файлом.
    echo Положите BAT-файл в папку со скриптом.
    echo.
    pause
    exit /b 1
)

if not exist "passport_template_v2.docx" (
    echo ОШИБКА: passport_template_v2.docx не найден рядом с этим BAT-файлом.
    echo Положите Word-шаблон рядом со скриптом.
    echo.
    pause
    exit /b 1
)

REM --- Поиск Python ---
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
    echo ОШИБКА: Python не найден.
    echo Установите Python 3 и добавьте его в PATH.
    echo.
    pause
    exit /b 1
)

echo Найден Python: %PYTHON_CMD%
echo.

REM --- Создание .venv при необходимости ---
if not exist ".venv\Scripts\python.exe" (
    echo Виртуальное окружение .venv не найдено.
    echo Создаю .venv...
    %PYTHON_CMD% -m venv .venv

    if errorlevel 1 (
        echo.
        echo ОШИБКА: не удалось создать виртуальное окружение .venv.
        pause
        exit /b 1
    )
)

set "VENV_PY=.venv\Scripts\python.exe"

echo Обновляю pip...
"%VENV_PY%" -m pip install --upgrade pip

echo.
echo Проверка и установка зависимостей...

set "CHECK_DEPS_FILE=%TEMP%\check_nku_passport_deps.py"

> "%CHECK_DEPS_FILE%" echo import importlib
>> "%CHECK_DEPS_FILE%" echo import subprocess
>> "%CHECK_DEPS_FILE%" echo import sys
>> "%CHECK_DEPS_FILE%" echo.
>> "%CHECK_DEPS_FILE%" echo packages = {
>> "%CHECK_DEPS_FILE%" echo     "docx": "python-docx",
>> "%CHECK_DEPS_FILE%" echo     "openpyxl": "openpyxl",
>> "%CHECK_DEPS_FILE%" echo     "gspread": "gspread",
>> "%CHECK_DEPS_FILE%" echo     "google.oauth2": "google-auth",
>> "%CHECK_DEPS_FILE%" echo }
>> "%CHECK_DEPS_FILE%" echo.
>> "%CHECK_DEPS_FILE%" echo missing = []
>> "%CHECK_DEPS_FILE%" echo.
>> "%CHECK_DEPS_FILE%" echo for import_name, pip_name in packages.items():
>> "%CHECK_DEPS_FILE%" echo     try:
>> "%CHECK_DEPS_FILE%" echo         importlib.import_module(import_name)
>> "%CHECK_DEPS_FILE%" echo     except ImportError:
>> "%CHECK_DEPS_FILE%" echo         missing.append(pip_name)
>> "%CHECK_DEPS_FILE%" echo.
>> "%CHECK_DEPS_FILE%" echo if missing:
>> "%CHECK_DEPS_FILE%" echo     print("Не хватает зависимостей: " + ", ".join(missing))
>> "%CHECK_DEPS_FILE%" echo     print("Устанавливаю...")
>> "%CHECK_DEPS_FILE%" echo     subprocess.check_call([sys.executable, "-m", "pip", "install"] + missing)
>> "%CHECK_DEPS_FILE%" echo else:
>> "%CHECK_DEPS_FILE%" echo     print("Все зависимости найдены.")

"%VENV_PY%" "%CHECK_DEPS_FILE%"

if errorlevel 1 (
    echo.
    echo ОШИБКА: не удалось проверить или установить зависимости.
    if exist "%CHECK_DEPS_FILE%" del "%CHECK_DEPS_FILE%" >nul 2>nul
    pause
    exit /b 1
)

if exist "%CHECK_DEPS_FILE%" del "%CHECK_DEPS_FILE%" >nul 2>nul

echo.
echo Выберите режим работы:
echo 1 - Онлайн Google Таблица
echo 2 - Локальный Excel-файл рядом со скриптом
echo.

:choose_mode
set /p MODE="Введите 1 или 2 и нажмите Enter: "

if "%MODE%"=="1" goto online_mode
if "%MODE%"=="2" goto local_mode

echo Неверный ввод. Нужно ввести 1 или 2.
goto choose_mode

:online_mode
echo.
if not exist "google_service_account.json" (
    echo ОШИБКА: google_service_account.json не найден рядом с этим BAT-файлом.
    echo Он нужен для онлайн-режима Google Sheets.
    echo.
    pause
    exit /b 1
)

echo Вставьте ссылку на Google Таблицу.
echo Пример:
echo https://docs.google.com/spreadsheets/d/.../edit?gid=...
echo.
set /p SHEET_URL="Ссылка: "

if "%SHEET_URL%"=="" (
    echo.
    echo ОШИБКА: ссылка не введена.
    pause
    exit /b 1
)

echo.
echo Запускаю онлайн-режим...
echo.
"%VENV_PY%" generate_passports_v7_Onln.py "%SHEET_URL%"
goto finish

:local_mode
echo.
echo Запускаю локальный режим.
echo Скрипт будет искать Excel-файл рядом с собой, как раньше.
echo.
"%VENV_PY%" generate_passports_v7_Onln.py
goto finish

:finish
echo.
echo Работа завершена.
pause
exit /b 0
