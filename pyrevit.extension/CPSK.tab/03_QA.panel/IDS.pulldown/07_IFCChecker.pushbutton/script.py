# -*- coding: utf-8 -*-
"""
IFC Checker - проверка IFC моделей по требованиям IDS.
Запускает внешний скрипт ifc_checker с GUI интерфейсом.
"""

__title__ = "IFC\nChecker"
__author__ = "CPSK"

import os
import sys
import subprocess

# Добавляем lib в путь для импорта
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Импорт модулей из lib
from cpsk_notify import show_error, show_warning, show_info
from cpsk_auth import require_auth
from cpsk_config import require_environment, get_venv_python, get_clean_env
from cpsk_logger import Logger

# Проверка авторизации
if not require_auth():
    sys.exit()

# Проверка окружения
if not require_environment():
    sys.exit()

# Инициализация логгера
SCRIPT_NAME = "IFCChecker"
Logger.init(SCRIPT_NAME)
Logger.info(SCRIPT_NAME, "Скрипт запущен")

# Путь к скрипту ifc_checker
REPO_ROOT = os.path.dirname(EXTENSION_DIR)
IFC_CHECKER_SCRIPT = os.path.join(REPO_ROOT, "ifc_checker", "ifc_checker_script", "main.py")


def main():
    """Запустить IFC Checker."""
    Logger.log_separator(SCRIPT_NAME, "Запуск IFC Checker")

    # Проверяем Python
    python_path = get_venv_python()
    if not python_path or not os.path.exists(python_path):
        Logger.error(SCRIPT_NAME, "Python окружение не найдено!")
        show_error(
            "Ошибка",
            "Python окружение не найдено!",
            details="Перейдите в Settings -> Окружение и установите окружение."
        )
        return

    Logger.info(SCRIPT_NAME, "Python: {}".format(python_path))

    # Проверяем скрипт
    if not os.path.exists(IFC_CHECKER_SCRIPT):
        Logger.error(SCRIPT_NAME, "Скрипт не найден: {}".format(IFC_CHECKER_SCRIPT))
        show_error(
            "Ошибка",
            "Скрипт IFC Checker не найден!",
            details="Путь: {}".format(IFC_CHECKER_SCRIPT)
        )
        return

    Logger.info(SCRIPT_NAME, "Скрипт: {}".format(IFC_CHECKER_SCRIPT))

    # Запускаем скрипт в отдельном процессе
    try:
        # Очищаем переменные окружения IronPython
        clean_env = get_clean_env()

        Logger.info(SCRIPT_NAME, "Запуск subprocess...")

        # Запускаем без ожидания завершения (GUI приложение)
        process = subprocess.Popen(
            [python_path, IFC_CHECKER_SCRIPT],
            env=clean_env,
            cwd=os.path.dirname(IFC_CHECKER_SCRIPT)
        )

        Logger.info(SCRIPT_NAME, "Процесс запущен, PID: {}".format(process.pid))
        Logger.info(SCRIPT_NAME, "Лог: {}".format(Logger.get_log_path()))

        show_info(
            "IFC Checker",
            "Приложение запущено",
            details="Окно IFC Checker откроется отдельно."
        )

    except Exception as e:
        Logger.error(SCRIPT_NAME, "Ошибка запуска: {}".format(str(e)))
        show_error(
            "Ошибка",
            "Не удалось запустить IFC Checker",
            details=str(e)
        )


if __name__ == "__main__":
    main()
