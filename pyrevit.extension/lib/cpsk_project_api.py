# -*- coding: utf-8 -*-
"""
CPSK Project API Client
API клиент для работы с проектами на сервере.

Использование:
    from cpsk_project_api import ProjectApiClient

    client = ProjectApiClient()

    # Проверка регистрации
    is_registered, project_id, project_data = client.check_registration(project_guid)

    # Регистрация проекта
    success, result, error = client.register_project(project_data)

    # Отправка статистики
    success, result, error = client.send_statistics(project_id, statistics_data)

    # Регистрация элемента
    success, result, error = client.register_element(element_data)
"""

import os
import re
import urllib2
import ssl

from cpsk_auth import AuthService, _create_ssl_context, _escape_json


# API настройки
API_BASE_URL = "https://api-cpsk-superapp.gip.su/api/rocketrevit"


def _parse_json_value(json_str, key):
    """Извлечь значение из JSON строки по ключу (строковое значение)."""
    pattern = r'"{}"\\s*:\\s*"([^"]*)"'.format(re.escape(key))
    match = re.search(pattern, json_str)
    if match:
        return match.group(1)
    return None


def _parse_json_int(json_str, key):
    """Извлечь целое число из JSON строки по ключу."""
    pattern = r'"{}"\\s*:\\s*(\\d+)'.format(re.escape(key))
    match = re.search(pattern, json_str)
    if match:
        return int(match.group(1))
    return None


def _parse_json_float(json_str, key):
    """Извлечь число с плавающей точкой из JSON строки по ключу."""
    pattern = r'"{}"\\s*:\\s*([\\d.]+)'.format(re.escape(key))
    match = re.search(pattern, json_str)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            return None
    return None


def _parse_json_bool(json_str, key):
    """Извлечь булево значение из JSON строки по ключу."""
    pattern = r'"{}"\\s*:\\s*(true|false)'.format(re.escape(key))
    match = re.search(pattern, json_str)
    if match:
        return match.group(1) == "true"
    return None


def _build_json_object(data):
    """
    Собрать JSON объект из словаря.
    Поддерживает строки, числа, bool, None и вложенные dict.
    """
    parts = []
    for key, value in data.items():
        if value is None:
            parts.append('"{}":null'.format(key))
        elif isinstance(value, bool):
            parts.append('"{}":'.format(key) + ("true" if value else "false"))
        elif isinstance(value, (int, float)):
            parts.append('"{}":'.format(key) + str(value))
        elif isinstance(value, dict):
            parts.append('"{}":'.format(key) + _build_json_object(value))
        elif isinstance(value, list):
            list_items = []
            for item in value:
                if isinstance(item, dict):
                    list_items.append(_build_json_object(item))
                elif isinstance(item, bool):
                    list_items.append("true" if item else "false")
                elif isinstance(item, (int, float)):
                    list_items.append(str(item))
                elif item is None:
                    list_items.append("null")
                else:
                    list_items.append('"{}"'.format(_escape_json(str(item))))
            parts.append('"{}":'.format(key) + "[" + ",".join(list_items) + "]")
        else:
            parts.append('"{}":"{}"'.format(key, _escape_json(str(value))))
    return "{" + ",".join(parts) + "}"


class ProjectApiClient(object):
    """Клиент API для работы с проектами."""

    def __init__(self, base_url=None):
        """
        Инициализация клиента.

        Args:
            base_url: Базовый URL API (по умолчанию API_BASE_URL)
        """
        self.base_url = base_url or API_BASE_URL

    def _make_request(self, method, endpoint, data=None, require_auth=True):
        """
        Выполнить HTTP запрос.

        Args:
            method: HTTP метод (GET, POST)
            endpoint: Эндпоинт API
            data: Данные для отправки (dict)
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
            request.add_header("Content-Type", "application/json")
            request.add_header("User-Agent", "CPSK-pyRevit/1.0")

            # Добавляем токен авторизации
            token = AuthService.get_token()
            if token:
                request.add_header("Authorization", "Bearer {}".format(token))

            # Устанавливаем метод и данные
            if method == "POST" and data:
                json_data = _build_json_object(data)
                request.add_data(json_data.encode("utf-8"))
            elif method == "GET":
                request.get_method = lambda: "GET"

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
                error_body = "(не удалось прочитать тело ответа)"

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

    def check_registration(self, project_guid):
        """
        Проверить регистрацию проекта.

        Args:
            project_guid: GUID проекта (UniqueId)

        Returns:
            tuple: (is_registered, project_id, project_data или error_message)
        """
        endpoint = "/projects/check_registration/?project_guid={}".format(
            urllib2.quote(project_guid)
        )

        success, body, error = self._make_request("GET", endpoint)

        if not success:
            return (False, None, error or body)

        # Парсим ответ
        is_registered = _parse_json_bool(body, "is_registered")

        if is_registered:
            # Извлекаем данные проекта
            project_id = _parse_json_int(body, "id")
            project_name = _parse_json_value(body, "project_name")

            project_data = {
                "id": project_id,
                "project_name": project_name,
                "project_guid": _parse_json_value(body, "project_guid"),
            }
            return (True, project_id, project_data)
        else:
            return (False, None, None)

    def register_project(self, project_data):
        """
        Зарегистрировать проект.

        Args:
            project_data: dict с данными проекта:
                - project_guid: GUID проекта
                - project_name: Имя проекта
                - file_path: Путь к файлу
                - revit_version: Версия Revit
                - machine_name: Имя машины

        Returns:
            tuple: (success, result_data или None, error_message или None)
        """
        success, body, error = self._make_request("POST", "/projects/", project_data)

        if not success:
            return (False, None, error or body)

        # Парсим ответ
        # Ответ может быть в формате {"message": "...", "project": {...}, "created": true/false}
        created = _parse_json_bool(body, "created")
        message = _parse_json_value(body, "message")

        # Извлекаем ID проекта из вложенного объекта project
        # Ищем паттерн "project":{..."id":123...}
        project_match = re.search(r'"project"\s*:\s*\{[^}]*"id"\s*:\s*(\d+)', body)
        project_id = int(project_match.group(1)) if project_match else None

        # Также ищем имя проекта
        project_name_match = re.search(
            r'"project"\s*:\s*\{[^}]*"project_name"\s*:\s*"([^"]*)"', body
        )
        project_name = project_name_match.group(1) if project_name_match else None

        result_data = {
            "id": project_id,
            "project_name": project_name,
            "message": message,
            "created": created,
        }

        return (True, result_data, None)

    def send_statistics(self, project_id, statistics_data):
        """
        Отправить статистику проекта.

        Args:
            project_id: ID проекта на сервере
            statistics_data: dict со статистикой:
                - total_elements: Всего элементов
                - total_parameters: Всего параметров
                - filled_parameters: Заполненных параметров
                - additional_data: Дополнительные данные (dict)

        Returns:
            tuple: (success, result_data или None, error_message или None)
        """
        endpoint = "/projects/{}/register_statistics/".format(project_id)
        success, body, error = self._make_request("POST", endpoint, statistics_data)

        if not success:
            return (False, None, error or body)

        # Парсим ответ
        # Может быть в формате {"statistics": {...}} или напрямую {...}
        total_elements = _parse_json_int(body, "total_elements")
        total_parameters = _parse_json_int(body, "total_parameters")
        filled_parameters = _parse_json_int(body, "filled_parameters")
        fill_rate = _parse_json_float(body, "fill_rate")

        result_data = {
            "total_elements": total_elements,
            "total_parameters": total_parameters,
            "filled_parameters": filled_parameters,
            "fill_rate": fill_rate,
        }

        return (True, result_data, None)

    def register_element(self, element_data):
        """
        Зарегистрировать элемент.

        Args:
            element_data: dict с данными элемента:
                - project: ID проекта
                - unique_id: UniqueId элемента
                - element_id: ElementId элемента
                - category: Категория
                - family_name: Имя семейства
                - type_name: Имя типа
                - last_event: Последнее событие
                - event_timestamp: Время события (строка ISO)
                - event_payload: Дополнительные данные (dict)

        Returns:
            tuple: (success, result_data или None, error_message или None)
        """
        success, body, error = self._make_request("POST", "/elements/", element_data)

        if not success:
            return (False, None, error or body)

        # Парсим ответ
        element_id = _parse_json_int(body, "id")
        unique_id = _parse_json_value(body, "unique_id")

        result_data = {
            "id": element_id,
            "unique_id": unique_id,
        }

        return (True, result_data, None)

    def get_project_elements(self, project_id):
        """
        Получить элементы проекта.

        Args:
            project_id: ID проекта на сервере

        Returns:
            tuple: (success, elements_list или None, error_message или None)
        """
        endpoint = "/projects/{}/elements/".format(project_id)
        success, body, error = self._make_request("GET", endpoint)

        if not success:
            return (False, None, error or body)

        # Парсим массив элементов
        # Упрощённый парсинг - возвращаем raw body для дальнейшей обработки
        return (True, body, None)
