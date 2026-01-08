#! python3
# -*- coding: utf-8 -*-
"""
Проверка обновлений CPSK Tools через GitHub API.
"""

__title__ = "Проверить\nобновления"
__author__ = "CPSK"

import os
import sys
import json
import codecs

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
REPO_DIR = os.path.dirname(EXTENSION_DIR)  # Корень репозитория
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from cpsk_notify import show_info, show_success, show_warning, show_error

# Конфигурация GitHub
GITHUB_USER = "fesworkscience"
GITHUB_REPO = "pyrevit_rocket"
GITHUB_API_URL = "https://api.github.com/repos/{}/{}/releases/latest".format(GITHUB_USER, GITHUB_REPO)


def get_current_version():
    """Получить текущую версию из version.yaml."""
    version_file = os.path.join(REPO_DIR, "version.yaml")
    try:
        with codecs.open(version_file, 'r', 'utf-8') as f:
            content = f.read()
            # Простой парсинг YAML: version: 1.0.5
            for line in content.split('\n'):
                line = line.strip()
                if line.startswith('version:'):
                    return line.split(':', 1)[1].strip()
        return "0.0.0"
    except Exception:
        return "0.0.0"


def parse_version(version_str):
    """Парсить строку версии в кортеж чисел."""
    # Убираем 'v' в начале если есть
    version_str = version_str.lstrip('v').lstrip('V')
    try:
        parts = version_str.split('.')
        return tuple(int(p) for p in parts)
    except:
        return (0, 0, 0)


def check_for_updates():
    """
    Проверить наличие обновлений на GitHub.

    Returns:
        tuple: (has_update, latest_version, download_url, release_name)
    """
    try:
        # Python 2/3 совместимость для HTTP запросов
        try:
            # Python 2
            import urllib2
            request = urllib2.Request(GITHUB_API_URL)
            request.add_header('User-Agent', 'CPSK-Tools-Updater')
            response = urllib2.urlopen(request, timeout=10)
            data = json.load(response)
        except ImportError:
            # Python 3
            import urllib.request
            request = urllib.request.Request(GITHUB_API_URL)
            request.add_header('User-Agent', 'CPSK-Tools-Updater')
            response = urllib.request.urlopen(request, timeout=10)
            data = json.loads(response.read().decode('utf-8'))

        latest_version = data.get("tag_name", "v0.0.0").lstrip('v')
        release_name = data.get("name", "")

        # Получить ссылку на скачивание
        download_url = ""
        assets = data.get("assets", [])
        for asset in assets:
            name = asset.get("name", "")
            if name.endswith(".msi") or name.endswith(".zip"):
                download_url = asset.get("browser_download_url", "")
                break

        # Если нет assets, используем страницу релиза
        if not download_url:
            download_url = data.get("html_url", "https://github.com/{}/{}/releases".format(GITHUB_USER, GITHUB_REPO))

        current_version = get_current_version()

        # Сравнить версии
        current_tuple = parse_version(current_version)
        latest_tuple = parse_version(latest_version)

        has_update = latest_tuple > current_tuple

        return has_update, latest_version, download_url, release_name, current_version

    except Exception as e:
        return None, None, None, None, str(e)


def main():
    """Основная функция."""
    # Сразу делаем запрос (занимает 1-2 сек)
    has_update, latest_version, download_url, release_name, current_or_error = check_for_updates()

    if has_update is None:
        # Ошибка при проверке
        show_error(
            "Ошибка проверки",
            "Не удалось проверить обновления",
            details="Ошибка: {}\n\nПроверьте подключение к интернету.".format(current_or_error)
        )
        return

    if has_update:
        # Есть обновление
        details = "Текущая версия: {}\nНовая версия: {}\n\nСсылка для скачивания:\n{}".format(
            current_or_error,
            latest_version,
            download_url
        )
        if release_name:
            details = "Релиз: {}\n\n{}".format(release_name, details)

        show_warning(
            "Доступно обновление",
            "Доступна новая версия: {}".format(latest_version),
            details=details
        )
    else:
        # Версия актуальна
        show_success(
            "Версия актуальна",
            "У вас установлена последняя версия: {}".format(current_or_error)
        )


if __name__ == "__main__":
    main()
