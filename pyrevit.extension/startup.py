# -*- coding: utf-8 -*-
"""
CPSK Tools - Startup Script
Проверяет окружение при запуске и показывает всплывающее уведомление.
"""

import os
import sys

# Добавляем lib в путь
EXTENSION_DIR = os.path.dirname(os.path.abspath(__file__))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Импорты после добавления lib в путь
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import Timer

try:
    from cpsk_config import check_environment, get_setting, VENV_BASE_DIR
except ImportError:
    check_environment = None
    get_setting = None
    VENV_BASE_DIR = ""

try:
    from cpsk_notify import show_toast
except ImportError:
    show_toast = None


def check_and_notify():
    """Проверить окружение и показать уведомление."""
    # Проверяем настройки
    if get_setting is None or show_toast is None:
        return

    # Проверяем нужно ли показывать уведомление
    show_notification = get_setting("notifications.show_startup_check", True)
    if not show_notification:
        return

    # Проверяем окружение
    if check_environment is None:
        return

    try:
        result = check_environment()

        if result["is_ready"]:
            # Формируем детали
            details = """Python окружение успешно инициализировано.

Версия: {}
Путь: {}

Все инструменты CPSK готовы к работе.""".format(
                result.get("venv_version", "N/A"),
                VENV_BASE_DIR
            )

            show_toast(
                "CPSK Tools",
                "Окружение готово к работе. {}".format(result.get("venv_version", "")),
                details=details,
                notification_type="success",
                auto_close=4
            )
        else:
            show_env_warning = get_setting("notifications.show_env_warnings", True)
            if show_env_warning:
                errors = result.get("errors", [])
                error_msg = errors[0] if errors else "Требуется настройка"

                details = """Python окружение не настроено или повреждено.

Проблема: {}

Как исправить:
1. Перейдите в панель CPSK на ленте Revit
2. Откройте Settings → Окружение
3. Нажмите "Установить окружение"

Путь установки: {}""".format(
                    "\n".join(errors) if errors else "Неизвестная ошибка",
                    VENV_BASE_DIR
                )

                show_toast(
                    "CPSK Tools - Внимание",
                    "{}. Откройте Settings → Окружение".format(error_msg),
                    details=details,
                    notification_type="warning",
                    auto_close=8
                )
    except Exception:
        pass  # Игнорируем ошибки


# === MAIN - выполняется при загрузке extension ===
if __name__ == "__main__" or True:  # True чтобы выполнялось при импорте
    # Запускаем проверку асинхронно через небольшую задержку
    try:
        startup_timer = Timer()
        startup_timer.Interval = 2000  # 2 секунды после загрузки
        startup_timer.Tick += lambda s, e: (startup_timer.Stop(), check_and_notify())
        startup_timer.Start()
    except Exception:
        pass
