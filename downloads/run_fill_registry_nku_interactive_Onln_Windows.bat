@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion

REM ==========================================================
REM Интерактивный запуск заполнения реестра НКУ для Windows
REM Можно переносить в любую папку вместе с:
REM - fill_registry_nku_Onln_v9.py
REM - google_service_account.json для онлайн-режима
REM - Excel-файлом .xlsx для локального режима
REM
REM Если .venv нет — BAT сам создаст виртуальное окружение
REM и установит зависимости.
REM ==========================================================

cd /d "%~dp0"

set "SCRIPT_NAME=fill_registry_nku_Onln_v9.py"

echo ==============================================
echo  Заполнение реестра НКУ
echo ==============================================
echo.
echo Папка запуска:
echo %CD%
echo.

if not exist "%SCRIPT_NAME%" (
    echo ОШИБКА: %SCRIPT_NAME% не найден рядом с этим BAT-файлом.
    echo Положите BAT-файл в папку со скриптом.
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
    echo Скачать Python можно с официального сайта python.org.
    echo При установке отметьте галочку "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

echo Найден Python: %PYTHON_CMD%
echo.

REM --- Создание .venv при необходимости ---
if not exist ".venv\Scripts\python.exe" (
    echo Виртуальное окружение .venv не найдено.
    echo Создаю .venv в этой папке...
    %PYTHON_CMD% -m venv .venv

    if errorlevel 1 (
        echo.
        echo ОШИБКА: не удалось создать виртуальное окружение .venv.
        echo Проверьте, что Python установлен корректно.
        echo.
        pause
        exit /b 1
    )
)

set "VENV_PY=.venv\Scripts\python.exe"

echo Обновляю pip...
"%VENV_PY%" -m pip install --upgrade pip

if errorlevel 1 (
    echo.
    echo ВНИМАНИЕ: pip не удалось обновить, но запуск продолжится.
    echo.
)

echo.
echo Проверка и установка зависимостей...

set "CHECK_DEPS_FILE=%TEMP%\check_fill_registry_nku_deps.py"

> "%CHECK_DEPS_FILE%" echo import importlib
>> "%CHECK_DEPS_FILE%" echo import subprocess
>> "%CHECK_DEPS_FILE%" echo import sys
>> "%CHECK_DEPS_FILE%" echo.
>> "%CHECK_DEPS_FILE%" echo packages = {
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
    echo.
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
    echo.
    pause
    exit /b 1
)

echo.
echo Запускаю онлайн-режим...
echo.
"%VENV_PY%" "%SCRIPT_NAME%" "%SHEET_URL%" --credentials google_service_account.json
goto finish

:local_mode
echo.
echo Запускаю локальный режим.
echo Скрипт сам найдёт один Excel-файл .xlsx рядом с собой.
echo Если рядом несколько .xlsx, скрипт остановится и покажет список.
echo.
"%VENV_PY%" "%SCRIPT_NAME%"
goto finish

:finish
echo.
echo Работа завершена.
pause
exit /b 0
