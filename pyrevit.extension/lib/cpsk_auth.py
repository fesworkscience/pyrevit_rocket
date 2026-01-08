# -*- coding: utf-8 -*-
"""
CPSK Authentication Module
Модуль аутентификации для работы с API сервером.

Использование:
    from cpsk_auth import AuthService

    # Проверка статуса
    if AuthService.is_authenticated():
        print("Пользователь: " + AuthService.get_username())

    # Вход
    success, error, details = AuthService.login("user@email.com", "password")

    # Выход
    AuthService.logout()

    # Получить токен для API запросов
    token = AuthService.get_token()
"""

import os
import re
import urllib2
import ssl

# Импортируем cpsk_config для работы с настройками
from cpsk_config import get_setting, set_setting

# API настройки
API_BASE_URL = "https://api-cpsk-superapp.gip.su"
API_TOKEN_ENDPOINT = "/api/rocketrevit/token/"

# Кэш токена в памяти
_token_cache = {
    "token": None,
    "username": None
}


def _escape_json(s):
    """Экранировать строку для JSON."""
    if not s:
        return ""
    return s.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n').replace('\r', '\\r').replace('\t', '\\t')


def _parse_json_value(json_str, key):
    """Извлечь значение из JSON строки по ключу."""
    pattern = r'"{}"\s*:\s*"([^"]*)"'.format(re.escape(key))
    match = re.search(pattern, json_str)
    if match:
        return match.group(1)
    return None


def _create_ssl_context():
    """Создать SSL контекст для HTTPS запросов."""
    try:
        # Пробуем создать контекст без проверки сертификата
        # (для корпоративных сетей с прокси)
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        return context
    except Exception:
        return None


def _format_request_details(url, method, status_code, response_body, request_body=None, error=None):
    """
    Форматировать детали запроса для отладки.

    Args:
        url: URL запроса
        method: HTTP метод (GET, POST, etc.)
        status_code: Код ответа (или None)
        response_body: Тело ответа (или None)
        request_body: Тело запроса (или None)
        error: Сообщение об ошибке (или None)

    Returns:
        str: Форматированная строка с деталями
    """
    lines = []
    lines.append("=" * 50)
    lines.append("ДЕТАЛИ ЗАПРОСА")
    lines.append("=" * 50)
    lines.append("")
    lines.append("URL: {}".format(url))
    lines.append("Метод: {}".format(method))

    if status_code is not None:
        lines.append("Код ответа: {}".format(status_code))

    if error:
        lines.append("")
        lines.append("--- ОШИБКА ---")
        lines.append(str(error))

    if request_body:
        lines.append("")
        lines.append("--- ТЕЛО ЗАПРОСА ---")
        lines.append(str(request_body))

    if response_body:
        lines.append("")
        lines.append("--- ОТВЕТ СЕРВЕРА ---")
        # Ограничиваем длину для читаемости
        if len(response_body) > 2000:
            lines.append(response_body[:2000] + "\n... (обрезано)")
        else:
            lines.append(str(response_body))

    lines.append("")
    lines.append("=" * 50)

    return "\n".join(lines)


class AuthService(object):
    """Сервис аутентификации."""

    @staticmethod
    def is_authenticated():
        """Проверить, авторизован ли пользователь."""
        token = AuthService.get_token()
        return token is not None and len(token) > 0

    @staticmethod
    def get_token():
        """Получить токен авторизации."""
        global _token_cache

        # Сначала проверяем кэш
        if _token_cache["token"]:
            return _token_cache["token"]

        # Загружаем из конфига (секция auth)
        token = get_setting("auth.token", "")
        if token:
            _token_cache["token"] = token
            _token_cache["username"] = get_setting("auth.email", "")

        return _token_cache["token"]

    @staticmethod
    def get_username():
        """Получить имя пользователя."""
        global _token_cache

        if _token_cache["username"]:
            return _token_cache["username"]

        return get_setting("auth.email", "")

    @staticmethod
    def login(username, password):
        """
        Выполнить вход.

        Args:
            username: Логин (email)
            password: Пароль

        Returns:
            tuple: (success: bool, error_message: str или None, details: str или None)
                   details содержит детальную информацию для отладки
        """
        global _token_cache

        url = API_BASE_URL + API_TOKEN_ENDPOINT

        try:
            # Формируем JSON запрос
            request_body = '{{"username":"{}","password":"{}"}}'.format(
                _escape_json(username),
                _escape_json(password)
            )

            # Создаём запрос
            request = urllib2.Request(url)
            request.add_header('Content-Type', 'application/json')
            request.add_header('Accept', 'application/json')
            request.add_header('User-Agent', 'CPSK-pyRevit/1.0')
            request.add_data(request_body.encode('utf-8'))

            # Выполняем запрос
            ssl_context = _create_ssl_context()
            if ssl_context:
                response = urllib2.urlopen(request, context=ssl_context, timeout=30)
            else:
                response = urllib2.urlopen(request, timeout=30)

            response_code = response.getcode()
            response_body = response.read().decode('utf-8')

            # Парсим ответ - ищем токен
            token = _parse_json_value(response_body, "access")

            if token:
                # Сохраняем в кэш
                _token_cache["token"] = token
                _token_cache["username"] = username

                # Сохраняем в конфиг (секция auth)
                set_setting("auth.token", token)
                set_setting("auth.email", username)

                return (True, None, None)
            else:
                details = _format_request_details(
                    url, "POST", response_code, response_body, request_body
                )
                return (False, "Токен не найден в ответе сервера", details)

        except urllib2.HTTPError as e:
            error_body = ""
            try:
                error_body = e.read().decode('utf-8')
            except Exception:
                error_body = "(не удалось прочитать тело ответа)"

            details = _format_request_details(
                url, "POST", e.code, error_body,
                '{{"username":"{}","password":"***"}}'.format(_escape_json(username))
            )

            if e.code == 401:
                # Парсим сообщение об ошибке
                detail = _parse_json_value(error_body, "detail")
                if detail:
                    return (False, detail, details)
                return (False, "Неверный логин или пароль", details)
            elif e.code == 400:
                return (False, "Неверный формат запроса", details)
            elif e.code == 405:
                return (False, "Метод не разрешён (405). Проверьте URL эндпоинта.", details)
            elif e.code == 404:
                return (False, "Эндпоинт не найден (404). Проверьте URL.", details)
            else:
                return (False, "Ошибка сервера: {}".format(e.code), details)

        except urllib2.URLError as e:
            reason = str(e.reason) if hasattr(e, 'reason') else str(e)
            details = _format_request_details(url, "POST", None, None, None, reason)

            if 'SSL' in reason or 'certificate' in reason.lower():
                return (False, "Ошибка SSL соединения. Проверьте настройки сети.", details)
            elif 'timeout' in reason.lower():
                return (False, "Превышено время ожидания ответа от сервера", details)
            elif 'connection' in reason.lower():
                return (False, "Не удается подключиться к серверу. Проверьте интернет.", details)
            else:
                return (False, "Ошибка сети: {}".format(reason), details)

        except Exception as e:
            details = _format_request_details(url, "POST", None, None, None, str(e))
            return (False, "Ошибка: {}".format(str(e)), details)

    @staticmethod
    def logout():
        """Выполнить выход."""
        global _token_cache

        # Очищаем кэш
        _token_cache["token"] = None
        _token_cache["username"] = None

        # Очищаем в конфиге (секция auth)
        set_setting("auth.token", "")
        set_setting("auth.email", "")

    @staticmethod
    def get_auth_header():
        """
        Получить заголовок авторизации для HTTP запросов.

        Returns:
            dict: {"Authorization": "Bearer <token>"} или пустой dict
        """
        token = AuthService.get_token()
        if token:
            return {"Authorization": "Bearer {}".format(token)}
        return {}


def require_auth(silent=False):
    """
    Проверить авторизацию и показать ошибку если не авторизован.

    Использование в начале скрипта:
        from cpsk_auth import require_auth
        if not require_auth():
            import sys
            sys.exit()

    Args:
        silent: Если True, не показывать окно ошибки (только вернуть False)

    Returns:
        bool: True если авторизован, False если нет
    """
    if AuthService.is_authenticated():
        return True

    if not silent:
        details = """Для работы с инструментами CPSK необходима авторизация.

Как войти в систему:
1. Найдите панель "CPSK" на ленте Revit
2. Нажмите кнопку "Вход"
3. Введите логин (email) и пароль
4. Нажмите "Войти"

После успешного входа все инструменты станут доступны.

Если у вас нет учётной записи, обратитесь к администратору."""

        try:
            from cpsk_notify import show_warning
            show_warning(
                "Требуется авторизация",
                "Для использования этой функции необходимо войти в систему.",
                details=details
            )
        except Exception:
            # Если cpsk_notify недоступен, используем простой MessageBox
            try:
                from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon
                MessageBox.Show(
                    "Для использования этой функции необходимо войти в систему.\n\n"
                    "Нажмите кнопку 'Вход' на панели CPSK.",
                    "Требуется авторизация",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
            except Exception:
                pass

    return False


class ApiClient(object):
    """Клиент для работы с API."""

    def __init__(self, base_url=None):
        """
        Инициализация клиента.

        Args:
            base_url: Базовый URL API (по умолчанию API_BASE_URL)
        """
        self.base_url = base_url or API_BASE_URL

    def get(self, endpoint, require_auth=True):
        """
        Выполнить GET запрос.

        Args:
            endpoint: Эндпоинт API (начинается с /)
            require_auth: Требовать авторизацию

        Returns:
            tuple: (success: bool, data: dict/list или error_message: str)
        """
        if require_auth and not AuthService.is_authenticated():
            return (False, "Требуется авторизация")

        try:
            url = self.base_url + endpoint
            request = urllib2.Request(url)
            request.add_header('Accept', 'application/json')
            request.add_header('User-Agent', 'CPSK-pyRevit/1.0')

            # Добавляем токен если есть
            token = AuthService.get_token()
            if token:
                request.add_header('Authorization', 'Bearer {}'.format(token))

            ssl_context = _create_ssl_context()
            if ssl_context:
                response = urllib2.urlopen(request, context=ssl_context, timeout=60)
            else:
                response = urllib2.urlopen(request, timeout=60)

            response_body = response.read().decode('utf-8')

            # Простой парсинг JSON (для списков и объектов)
            # В IronPython нет json модуля, используем регулярки
            return (True, response_body)

        except urllib2.HTTPError as e:
            if e.code == 401:
                return (False, "Сессия истекла. Войдите заново.")
            elif e.code == 403:
                return (False, "Доступ запрещён")
            elif e.code == 404:
                return (False, "Ресурс не найден")
            else:
                return (False, "Ошибка сервера: {}".format(e.code))
        except Exception as e:
            return (False, "Ошибка: {}".format(str(e)))

    def download_file(self, url):
        """
        Скачать файл.

        Args:
            url: URL файла

        Returns:
            tuple: (success: bool, bytes или error_message: str)
        """
        try:
            request = urllib2.Request(url)
            request.add_header('User-Agent', 'CPSK-pyRevit/1.0')

            token = AuthService.get_token()
            if token:
                request.add_header('Authorization', 'Bearer {}'.format(token))

            ssl_context = _create_ssl_context()
            if ssl_context:
                response = urllib2.urlopen(request, context=ssl_context, timeout=300)
            else:
                response = urllib2.urlopen(request, timeout=300)

            return (True, response.read())

        except Exception as e:
            return (False, "Ошибка загрузки: {}".format(str(e)))
