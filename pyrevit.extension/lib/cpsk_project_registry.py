# -*- coding: utf-8 -*-
"""
CPSK Project Registry Service
Локальный реестр проектов для хранения ID проектов.

Использование:
    from cpsk_project_registry import ProjectRegistry

    # Регистрация проекта
    ProjectRegistry.register(project_guid, project_id)

    # Получение ID проекта
    project_id = ProjectRegistry.get_project_id(project_guid)

    # Проверка регистрации
    if ProjectRegistry.is_registered(project_guid):
        ...

    # Удаление регистрации
    ProjectRegistry.unregister(project_guid)
"""

import os
import codecs
import re


def _get_registry_path():
    """Получить путь к файлу реестра."""
    # Используем LOCALAPPDATA для согласованности с cpsk_settings.yaml
    localappdata = os.environ.get("LOCALAPPDATA", "")
    if not localappdata:
        localappdata = os.path.expanduser("~")

    registry_dir = os.path.join(localappdata, "CPSK")
    if not os.path.exists(registry_dir):
        os.makedirs(registry_dir)

    return os.path.join(registry_dir, "project_registry.json")


def _load_registry():
    """Загрузить реестр из файла."""
    registry_path = _get_registry_path()

    if not os.path.exists(registry_path):
        return {}

    try:
        with codecs.open(registry_path, "r", "utf-8") as f:
            content = f.read().strip()

        if not content or content == "{}":
            return {}

        # Простой парсинг JSON: {"guid1": 123, "guid2": 456}
        registry = {}

        # Убираем внешние скобки
        content = content.strip()
        if content.startswith("{") and content.endswith("}"):
            content = content[1:-1].strip()

        if not content:
            return {}

        # Ищем пары "ключ": значение
        pattern = r'"([^"]+)"\s*:\s*(\d+)'
        matches = re.findall(pattern, content)

        for guid, project_id in matches:
            registry[guid] = int(project_id)

        return registry

    except Exception:
        return {}


def _save_registry(registry):
    """Сохранить реестр в файл."""
    registry_path = _get_registry_path()

    try:
        # Формируем JSON
        if not registry:
            content = "{}"
        else:
            pairs = []
            for guid, project_id in registry.items():
                pairs.append('"{}": {}'.format(guid, project_id))
            content = "{\n  " + ",\n  ".join(pairs) + "\n}"

        with codecs.open(registry_path, "w", "utf-8") as f:
            f.write(content)

        return True

    except Exception:
        return False


class ProjectRegistry(object):
    """Сервис локального реестра проектов."""

    _cache = None

    @classmethod
    def _get_registry(cls):
        """Получить реестр (с кэшированием)."""
        if cls._cache is None:
            cls._cache = _load_registry()
        return cls._cache

    @classmethod
    def _invalidate_cache(cls):
        """Сбросить кэш."""
        cls._cache = None

    @classmethod
    def register(cls, project_guid, project_id):
        """
        Зарегистрировать проект.

        Args:
            project_guid: GUID проекта (UniqueId)
            project_id: ID проекта на сервере
        """
        registry = cls._get_registry()
        registry[project_guid] = int(project_id)
        _save_registry(registry)
        cls._cache = registry

    @classmethod
    def get_project_id(cls, project_guid):
        """
        Получить ID проекта по GUID.

        Args:
            project_guid: GUID проекта (UniqueId)

        Returns:
            int: ID проекта или 0 если не найден
        """
        registry = cls._get_registry()
        return registry.get(project_guid, 0)

    @classmethod
    def is_registered(cls, project_guid):
        """
        Проверить, зарегистрирован ли проект.

        Args:
            project_guid: GUID проекта (UniqueId)

        Returns:
            bool: True если проект зарегистрирован
        """
        return cls.get_project_id(project_guid) > 0

    @classmethod
    def unregister(cls, project_guid):
        """
        Удалить регистрацию проекта.

        Args:
            project_guid: GUID проекта (UniqueId)
        """
        registry = cls._get_registry()
        if project_guid in registry:
            del registry[project_guid]
            _save_registry(registry)
            cls._cache = registry

    @classmethod
    def get_all(cls):
        """
        Получить все зарегистрированные проекты.

        Returns:
            dict: Словарь {project_guid: project_id}
        """
        return cls._get_registry().copy()

    @classmethod
    def clear(cls):
        """Очистить весь реестр."""
        _save_registry({})
        cls._cache = {}

    @classmethod
    def reload(cls):
        """Перезагрузить реестр из файла."""
        cls._invalidate_cache()
        return cls._get_registry()
