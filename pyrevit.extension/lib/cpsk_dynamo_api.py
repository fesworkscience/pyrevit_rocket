# -*- coding: utf-8 -*-
"""
CPSK Dynamo Scripts API Client
API клиент для загрузки Dynamo скриптов с сервера.

Использование:
    from cpsk_dynamo_api import DynamoApiClient

    client = DynamoApiClient()

    # Получить список скриптов
    success, scripts, error = client.get_scripts()

    # Получить список секций
    success, sections, error = client.get_sections()

    # Скачать скрипт
    success, file_path, error = client.download_script(script_info, save_folder)
"""

import os
import sys
import re
import urllib2
import ssl

from cpsk_auth import AuthService, _create_ssl_context

# Импортируем config из корня проекта
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)
from config import API_BASE_URL, API_ROCKETREVIT_URL


def _url_decode(s):
    """Декодировать URL-encoded строку."""
    if not s:
        return s
    try:
        # Простая реализация URL decode для IronPython
        result = s
        # Декодируем %XX последовательности
        i = 0
        decoded = []
        while i < len(result):
            if result[i] == '%' and i + 2 < len(result):
                try:
                    hex_val = result[i+1:i+3]
                    decoded.append(chr(int(hex_val, 16)))
                    i += 3
                except ValueError:
                    decoded.append(result[i])
                    i += 1
            else:
                decoded.append(result[i])
                i += 1
        # Декодируем UTF-8 байты
        byte_str = ''.join(decoded)
        try:
            return byte_str.decode('utf-8')
        except (UnicodeDecodeError, AttributeError):
            return byte_str
    except Exception:
        return s


def _parse_json_array(json_str):
    """
    Простой парсер JSON массива объектов.
    Возвращает список словарей.
    """
    results = []
    if not json_str or not json_str.strip().startswith('['):
        return results

    # Убираем внешние скобки
    content = json_str.strip()[1:-1].strip()
    if not content:
        return results

    # Разбиваем на объекты (по "},")
    depth = 0
    current = []
    for char in content:
        if char == '{':
            depth += 1
        elif char == '}':
            depth -= 1
        current.append(char)
        if depth == 0 and char == '}':
            obj_str = ''.join(current).strip()
            if obj_str.startswith(','):
                obj_str = obj_str[1:].strip()
            if obj_str:
                obj = _parse_json_object(obj_str)
                if obj:
                    results.append(obj)
            current = []

    return results


def _parse_json_object(json_str):
    """
    Простой парсер JSON объекта.
    Возвращает словарь.
    """
    result = {}
    if not json_str or not json_str.strip().startswith('{'):
        return result

    content = json_str.strip()[1:-1].strip()
    if not content:
        return result

    # Паттерны для извлечения значений
    # Строки: "key": "value" или "key": null
    str_pattern = r'"([^"]+)"\s*:\s*(?:"([^"]*)"|(null)|(true)|(false)|(\d+(?:\.\d+)?))'

    for match in re.finditer(str_pattern, content):
        key = match.group(1)
        if match.group(2) is not None:
            result[key] = match.group(2)
        elif match.group(3):  # null
            result[key] = None
        elif match.group(4):  # true
            result[key] = True
        elif match.group(5):  # false
            result[key] = False
        elif match.group(6):  # number
            num_str = match.group(6)
            if '.' in num_str:
                result[key] = float(num_str)
            else:
                result[key] = int(num_str)

    return result


def _get_safe_filename(filename):
    """Получить безопасное имя файла."""
    if not filename:
        return "unknown_script.dyn"

    # Декодируем URL
    decoded = _url_decode(filename)

    # Получаем только имя файла (без пути)
    name = os.path.basename(decoded)

    # Заменяем недопустимые символы
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')

    # Ограничиваем длину
    if len(name) > 100:
        name = name[:100]

    # Добавляем расширение если нет
    if not name.lower().endswith('.dyn'):
        name += '.dyn'

    return name


class DynamoApiClient(object):
    """Клиент API для работы с Dynamo скриптами."""

    def __init__(self, base_url=None):
        """
        Инициализация клиента.

        Args:
            base_url: Базовый URL API (по умолчанию API_ROCKETREVIT_URL)
        """
        self.base_url = base_url or API_ROCKETREVIT_URL

    def _make_request(self, method, endpoint, require_auth=True):
        """
        Выполнить HTTP запрос.

        Args:
            method: HTTP метод (GET)
            endpoint: Эндпоинт API
            require_auth: Требовать авторизацию

        Returns:
            tuple: (success, response_body, error_message)
        """
        if require_auth and not AuthService.is_authenticated():
            return (False, None, "Требуется авторизация")

        url = self.base_url + endpoint
        if not url.endswith("/") and "?" not in url:
            url += "/"

        try:
            request = urllib2.Request(url)
            request.add_header("Accept", "application/json")
            request.add_header("User-Agent", "CPSK-pyRevit/1.0")

            # Добавляем токен авторизации
            token = AuthService.get_token()
            if token:
                request.add_header("Authorization", "Bearer {}".format(token))

            # Выполняем запрос
            ssl_context = _create_ssl_context()
            if ssl_context:
                response = urllib2.urlopen(request, context=ssl_context, timeout=60)
            else:
                response = urllib2.urlopen(request, timeout=60)

            response_body = response.read().decode("utf-8")
            return (True, response_body, None)

        except urllib2.HTTPError as e:
            error_body = ""
            try:
                error_body = e.read().decode("utf-8")
            except Exception:
                pass

            if e.code == 401:
                return (False, error_body, "Сессия истекла. Войдите заново.")
            elif e.code == 403:
                return (False, error_body, "Доступ запрещён")
            elif e.code == 404:
                return (False, error_body, "Ресурс не найден (404)")
            else:
                return (False, error_body, "Ошибка сервера: {}".format(e.code))

        except urllib2.URLError as e:
            reason = str(e.reason) if hasattr(e, "reason") else str(e)
            return (False, None, "Ошибка сети: {}".format(reason))

        except Exception as e:
            return (False, None, "Ошибка: {}".format(str(e)))

    def get_scripts(self):
        """
        Получить список Dynamo скриптов с сервера.

        Returns:
            tuple: (success, scripts_list, error_message)
                scripts_list содержит словари с ключами:
                - id: ID скрипта
                - file_input: путь к файлу на сервере
                - section_name: название секции
                - author_username: автор
                - comment: описание
                - is_approved: одобрен ли
                - date_uploaded: дата загрузки
                - download_link: ссылка для скачивания
        """
        success, body, error = self._make_request("GET", "/dynamo-scripts/shared_files/")

        if not success:
            return (False, None, error)

        # Парсим JSON массив
        scripts = _parse_json_array(body)

        # Добавляем вычисляемые поля
        for script in scripts:
            file_input = script.get("file_input", "")
            script["filename"] = _get_safe_filename(file_input)

            # Формируем download_link если его нет
            if not script.get("download_link") and file_input:
                script["download_link"] = "{}{}".format(API_BASE_URL, file_input)

        return (True, scripts, None)

    def get_sections(self):
        """
        Получить список секций (категорий).

        Returns:
            tuple: (success, sections_list, error_message)
                sections_list содержит словари с ключами:
                - id: ID секции
                - name: название
                - description: описание
        """
        success, body, error = self._make_request("GET", "/sections/")

        if not success:
            return (False, None, error)

        # Парсим JSON массив
        sections = _parse_json_array(body)

        return (True, sections, None)

    def download_script(self, script_info, save_folder):
        """
        Скачать Dynamo скрипт.

        Args:
            script_info: словарь с информацией о скрипте (должен содержать download_link или file_input)
            save_folder: папка для сохранения

        Returns:
            tuple: (success, file_path, error_message)
        """
        # Получаем URL для скачивания
        download_url = script_info.get("download_link")
        if not download_url:
            file_input = script_info.get("file_input")
            if file_input:
                download_url = "{}{}".format(API_BASE_URL, file_input)
            else:
                return (False, None, "Нет ссылки для скачивания")

        # Создаём папку если нет
        if not os.path.exists(save_folder):
            try:
                os.makedirs(save_folder)
            except Exception as e:
                return (False, None, "Ошибка создания папки: {}".format(str(e)))

        # Получаем имя файла
        filename = script_info.get("filename")
        if not filename:
            filename = _get_safe_filename(script_info.get("file_input", ""))

        save_path = os.path.join(save_folder, filename)

        try:
            request = urllib2.Request(download_url)
            request.add_header("User-Agent", "CPSK-pyRevit/1.0")

            # Добавляем токен
            token = AuthService.get_token()
            if token:
                request.add_header("Authorization", "Bearer {}".format(token))

            ssl_context = _create_ssl_context()
            if ssl_context:
                response = urllib2.urlopen(request, context=ssl_context, timeout=300)
            else:
                response = urllib2.urlopen(request, timeout=300)

            # Сохраняем файл
            with open(save_path, "wb") as f:
                f.write(response.read())

            return (True, save_path, None)

        except urllib2.HTTPError as e:
            return (False, None, "Ошибка скачивания: {}".format(e.code))

        except Exception as e:
            return (False, None, "Ошибка: {}".format(str(e)))


# Путь к папке для скачанных скриптов
def get_downloaded_scripts_folder():
    """Получить путь к папке для скачанных скриптов."""
    appdata = os.environ.get("APPDATA", "")
    if not appdata:
        appdata = os.path.expanduser("~")
    return os.path.join(appdata, "CPSK", "dynamo_scripts_downloaded")
