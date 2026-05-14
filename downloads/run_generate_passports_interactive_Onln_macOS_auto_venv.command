#!/bin/bash

# Интерактивный запуск генератора паспортов НКУ для macOS
# Можно переносить в другую папку вместе с:
# - generate_passports_v7_Onln.py
# - passport_template_v2.docx
# - google_service_account.json для онлайн-режима
# - Excel-файлом для локального режима
#
# Если .venv нет — файл сам создаст виртуальное окружение и установит зависимости.

cd "$(dirname "$0")"

echo "=============================================="
echo " Генератор паспортов НКУ"
echo "=============================================="
echo ""
echo "Папка запуска:"
pwd
echo ""

if [ ! -f "generate_passports_v7_Onln.py" ]; then
    echo "ОШИБКА: generate_passports_v7_Onln.py не найден рядом с этим файлом."
    echo "Положи этот .command файл в папку со скриптом."
    echo ""
    read -p "Нажми Enter для закрытия..."
    exit 1
fi

if [ ! -f "passport_template_v2.docx" ]; then
    echo "ОШИБКА: passport_template_v2.docx не найден рядом с этим файлом."
    echo "Положи Word-шаблон passport_template_v2.docx рядом со скриптом."
    echo ""
    read -p "Нажми Enter для закрытия..."
    exit 1
fi

# --- Поиск Python ---
PYTHON_CMD=""

if command -v python3 >/dev/null 2>&1; then
    PYTHON_CMD="python3"
elif command -v python >/dev/null 2>&1; then
    PYTHON_CMD="python"
else
    echo "ОШИБКА: Python не найден."
    echo "Установи Python 3 и попробуй снова."
    echo ""
    read -p "Нажми Enter для закрытия..."
    exit 1
fi

echo "Найден Python: $PYTHON_CMD"
echo ""

# --- Создание .venv при необходимости ---
if [ ! -f ".venv/bin/python" ]; then
    echo "Виртуальное окружение .venv не найдено."
    echo "Создаю .venv в этой папке..."
    "$PYTHON_CMD" -m venv .venv

    if [ $? -ne 0 ]; then
        echo ""
        echo "ОШИБКА: не удалось создать виртуальное окружение .venv."
        echo "Возможно, Python установлен некорректно или нет модуля venv."
        echo ""
        read -p "Нажми Enter для закрытия..."
        exit 1
    fi
fi

VENV_PY=".venv/bin/python"

echo "Обновляю pip..."
"$VENV_PY" -m pip install --upgrade pip

echo ""
echo "Проверка и установка зависимостей..."
"$VENV_PY" - <<'PY'
import importlib
import subprocess
import sys

packages = {
    "docx": "python-docx",
    "openpyxl": "openpyxl",
    "gspread": "gspread",
    "google.oauth2": "google-auth",
}

missing = []

for import_name, pip_name in packages.items():
    try:
        importlib.import_module(import_name)
    except ImportError:
        missing.append(pip_name)

if missing:
    print("Не хватает зависимостей:", ", ".join(missing))
    print("Устанавливаю...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
else:
    print("Все зависимости найдены.")
PY

if [ $? -ne 0 ]; then
    echo ""
    echo "ОШИБКА: не удалось проверить или установить зависимости."
    echo ""
    read -p "Нажми Enter для закрытия..."
    exit 1
fi

echo ""
echo "Выбери режим работы:"
echo "1 — Онлайн Google Таблица"
echo "2 — Локальный Excel-файл рядом со скриптом"
echo ""

while true; do
    read -p "Введи 1 или 2 и нажми Enter: " MODE

    case "$MODE" in
        1)
            echo ""
            if [ ! -f "google_service_account.json" ]; then
                echo "ОШИБКА: google_service_account.json не найден рядом с этим файлом."
                echo "Он нужен для онлайн-режима Google Sheets."
                echo ""
                read -p "Нажми Enter для закрытия..."
                exit 1
            fi

            echo "Вставь ссылку на Google Таблицу."
            echo "Пример:"
            echo "https://docs.google.com/spreadsheets/d/.../edit?gid=..."
            echo ""
            read -p "Ссылка: " SHEET_URL

            if [ -z "$SHEET_URL" ]; then
                echo ""
                echo "ОШИБКА: ссылка не введена."
                read -p "Нажми Enter для закрытия..."
                exit 1
            fi

            echo ""
            echo "Запускаю онлайн-режим..."
            echo ""
            "$VENV_PY" generate_passports_v7_Onln.py "$SHEET_URL"
            break
            ;;

        2)
            echo ""
            echo "Запускаю локальный режим."
            echo "Скрипт будет искать Excel-файл рядом с собой, как раньше."
            echo ""
            "$VENV_PY" generate_passports_v7_Onln.py
            break
            ;;

        *)
            echo "Неверный ввод. Нужно ввести 1 или 2."
            ;;
    esac
done

echo ""
echo "Работа завершена."
read -p "Нажми Enter для закрытия..."
