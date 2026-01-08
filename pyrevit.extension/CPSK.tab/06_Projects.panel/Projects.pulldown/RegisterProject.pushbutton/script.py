# -*- coding: utf-8 -*-
"""Регистрация проекта на сервере CPSK."""

__title__ = "Регистрация\nпроекта"
__author__ = "CPSK"

import os
import sys

# Добавляем lib в путь для импорта
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(
    os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
    ),
    "lib",
)
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Проверка авторизации
from cpsk_auth import require_auth

if not require_auth():
    sys.exit()

from cpsk_notify import show_error, show_warning, show_success, show_info
from cpsk_project_api import ProjectApiClient
from cpsk_project_registry import ProjectRegistry

from pyrevit import revit

# Получаем документ
doc = revit.doc
uidoc = revit.uidoc
app = doc.Application

# Проверяем, что это не документ семейства
if doc.IsFamilyDocument:
    show_warning(
        "Ошибка",
        "Эта команда не может быть выполнена в документе семейства.",
    )
    sys.exit()

# Собираем данные о проекте
project_guid = doc.ProjectInformation.UniqueId
project_name = doc.Title
file_path = doc.PathName

if not file_path:
    file_path = "Не сохранен"

revit_version = app.VersionNumber
machine_name = os.environ.get("COMPUTERNAME", "Unknown")

# Создаем клиент API
client = ProjectApiClient()

# Формируем данные для отправки
project_data = {
    "project_guid": project_guid,
    "project_name": project_name,
    "file_path": file_path,
    "revit_version": revit_version,
    "machine_name": machine_name,
}

# Отправляем запрос на регистрацию
success, result, error = client.register_project(project_data)

if success:
    # Сохраняем ID проекта в локальный реестр
    project_id = result.get("id")
    if project_id:
        ProjectRegistry.register(project_guid, project_id)

    # Проверяем, был ли проект создан заново или обновлен
    created = result.get("created", True)
    message = result.get("message", "")

    if created:
        show_success(
            "Проект зарегистрирован",
            "Проект '{}' успешно зарегистрирован!".format(project_name),
            details="ID проекта: {}\nВерсия Revit: {}\nМашина: {}".format(
                project_id, revit_version, machine_name
            ),
        )
    else:
        # Проект уже был зарегистрирован ранее
        if message and "уже существует" in message:
            show_warning(
                "Проект уже зарегистрирован",
                "Проект '{}' уже был зарегистрирован ранее и обновлен.".format(
                    project_name
                ),
                details="ID проекта: {}\nВерсия Revit: {}\nМашина: {}".format(
                    project_id, revit_version, machine_name
                ),
            )
        else:
            show_success(
                "Проект обновлен",
                "Данные проекта '{}' обновлены.".format(project_name),
                details="ID проекта: {}\nВерсия Revit: {}\nМашина: {}".format(
                    project_id, revit_version, machine_name
                ),
            )
else:
    show_error(
        "Ошибка регистрации",
        "Не удалось зарегистрировать проект",
        details=error,
        blocking=True,
    )
