# -*- coding: utf-8 -*-
"""Отправка статистики проекта на сервер CPSK."""

__title__ = "Статистика\nпроекта"
__author__ = "CPSK"

import os
import sys
import datetime

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
from Autodesk.Revit.DB import FilteredElementCollector

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

# Проверяем, зарегистрирован ли проект локально
project_guid = doc.ProjectInformation.UniqueId
project_id = ProjectRegistry.get_project_id(project_guid)

# Создаем клиент API
client = ProjectApiClient()

# Если проект не зарегистрирован локально, проверяем на сервере
if project_id == 0:
    is_registered, server_project_id, project_data = client.check_registration(
        project_guid
    )

    if is_registered and server_project_id:
        # Сохраняем в локальный реестр
        ProjectRegistry.register(project_guid, server_project_id)
        project_id = server_project_id
    else:
        show_warning(
            "Проект не зарегистрирован",
            "Сначала зарегистрируйте проект с помощью команды 'Регистрация проекта'.",
        )
        sys.exit()


def collect_project_statistics():
    """Собрать статистику проекта."""
    # Получаем все элементы модели
    collector = FilteredElementCollector(doc)
    all_elements = list(collector.WhereElementIsNotElementType().ToElements())

    total_elements = len(all_elements)
    total_parameters = 0
    filled_parameters = 0

    # Статистика по категориям
    category_stats = {}

    for element in all_elements:
        # Подсчет по категориям
        if element.Category:
            category_name = element.Category.Name
        else:
            category_name = "Без категории"

        if category_name in category_stats:
            category_stats[category_name] += 1
        else:
            category_stats[category_name] = 1

        # Подсчет параметров
        for param in element.Parameters:
            total_parameters += 1

            try:
                if param.HasValue:
                    value_string = param.AsValueString()
                    if value_string and len(value_string) > 0:
                        filled_parameters += 1
            except Exception:
                # Некоторые параметры не конвертируются в строку
                pass

    # Формируем дополнительные данные
    additional_data = {
        "categories_count": len(category_stats),
        "document_title": doc.Title,
        "analysis_timestamp": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
    }

    # Добавляем статистику по основным категориям
    main_categories = [
        "Стены",
        "Перекрытия",
        "Крыши",
        "Двери",
        "Окна",
        "Колонны",
        "Балки",
        "Лестницы",
        "Ограждения",
        "Помещения",
        "Уровни",
    ]

    for cat in main_categories:
        if cat in category_stats:
            key = cat.lower() + "_count"
            additional_data[key] = category_stats[cat]

    return {
        "total_elements": total_elements,
        "total_parameters": total_parameters,
        "filled_parameters": filled_parameters,
        "additional_data": additional_data,
    }, category_stats


# Собираем статистику
statistics_data, category_stats = collect_project_statistics()

# Отправляем на сервер
success, result, error = client.send_statistics(project_id, statistics_data)

if success:
    total_elements = result.get("total_elements") or statistics_data["total_elements"]
    total_parameters = (
        result.get("total_parameters") or statistics_data["total_parameters"]
    )
    filled_parameters = (
        result.get("filled_parameters") or statistics_data["filled_parameters"]
    )
    fill_rate = result.get("fill_rate")

    if fill_rate is None and total_parameters > 0:
        fill_rate = (float(filled_parameters) / total_parameters) * 100

    details = "Элементов: {:,}\n".format(total_elements)
    details += "Параметров: {:,}\n".format(total_parameters)
    details += "Заполнено: {:,}".format(filled_parameters)
    if fill_rate is not None:
        details += " ({:.1f}%)".format(fill_rate)
    details += "\n\nКатегорий: {}\n".format(len(category_stats))

    # Добавляем топ-5 категорий
    sorted_categories = sorted(
        category_stats.items(), key=lambda x: x[1], reverse=True
    )[:5]
    if sorted_categories:
        details += "\nТоп-5 категорий:\n"
        for cat_name, count in sorted_categories:
            details += "  {} - {:,}\n".format(cat_name, count)

    show_success(
        "Статистика отправлена",
        "Статистика проекта успешно отправлена на сервер!",
        details=details,
    )
else:
    # Проверяем, не истекла ли регистрация
    if error and "не зарегистрирован" in error.lower():
        show_warning(
            "Проект не зарегистрирован",
            "Сначала зарегистрируйте проект с помощью команды 'Регистрация проекта'.",
            details=error,
        )
    else:
        show_error(
            "Ошибка отправки",
            "Не удалось отправить статистику",
            details=error,
            blocking=True,
        )
