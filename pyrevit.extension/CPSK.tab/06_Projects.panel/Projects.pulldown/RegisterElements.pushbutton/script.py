# -*- coding: utf-8 -*-
"""Регистрация элементов проекта на сервере CPSK."""

__title__ = "Регистрация\nэлементов"
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

from cpsk_notify import show_error, show_warning, show_success, show_info, show_confirm
from cpsk_project_api import ProjectApiClient
from cpsk_project_registry import ProjectRegistry

from pyrevit import revit, forms
from Autodesk.Revit.DB import FilteredElementCollector, FamilyInstance, ElementId
from Autodesk.Revit.UI.Selection import ObjectType

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

# Проверяем, зарегистрирован ли проект
project_guid = doc.ProjectInformation.UniqueId
project_id = ProjectRegistry.get_project_id(project_guid)

if project_id == 0:
    show_warning(
        "Проект не зарегистрирован",
        "Сначала зарегистрируйте проект с помощью команды 'Регистрация проекта'.",
    )
    sys.exit()

# Предлагаем выбор способа выбора элементов
options = [
    "Выделенные элементы",
    "Все элементы модели",
    "Выбрать элементы",
]

selected_option = forms.SelectFromList.show(
    options,
    title="Регистрация элементов",
    button_name="Выбрать",
    multiselect=False,
)

if not selected_option:
    sys.exit()

elements_to_register = []

if selected_option == "Выделенные элементы":
    # Получаем выделенные элементы
    selection = uidoc.Selection.GetElementIds()
    if selection.Count == 0:
        show_warning("Предупреждение", "Нет выделенных элементов.")
        sys.exit()

    for eid in selection:
        element = doc.GetElement(eid)
        if element:
            elements_to_register.append(element)

elif selected_option == "Все элементы модели":
    # Получаем все элементы модели
    collector = FilteredElementCollector(doc)
    elements_to_register = list(collector.WhereElementIsNotElementType().ToElements())

elif selected_option == "Выбрать элементы":
    # Позволяем пользователю выбрать элементы
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, "Выберите элементы для регистрации"
        )
        for ref in refs:
            element = doc.GetElement(ref.ElementId)
            if element:
                elements_to_register.append(element)
    except Exception:
        # Пользователь отменил выбор
        sys.exit()

if len(elements_to_register) == 0:
    show_warning("Предупреждение", "Нет элементов для регистрации.")
    sys.exit()

# Подтверждение
confirm = show_confirm(
    "Подтверждение",
    "Зарегистрировать {} элементов?".format(len(elements_to_register)),
    details="Это может занять некоторое время для большого количества элементов.",
)

if not confirm:
    sys.exit()


def get_element_level_id(element):
    """Получить ID уровня элемента."""
    try:
        level_id = element.LevelId
        if level_id and level_id != ElementId.InvalidElementId:
            return str(level_id.IntegerValue)
    except Exception:
        pass
    return None


def get_element_workset_id(element):
    """Получить ID рабочего набора элемента."""
    try:
        from Autodesk.Revit.DB import WorksetId

        workset_id = element.WorksetId
        if workset_id and workset_id != WorksetId.InvalidWorksetId:
            return str(workset_id.IntegerValue)
    except Exception:
        pass
    return None


def get_element_phase_id(element):
    """Получить ID фазы создания элемента."""
    try:
        phase_id = element.CreatedPhaseId
        if phase_id and phase_id != ElementId.InvalidElementId:
            return str(phase_id.IntegerValue)
    except Exception:
        pass
    return None


def create_element_data(element, proj_id):
    """Создать данные для регистрации элемента."""
    category = "Без категории"
    if element.Category:
        category = element.Category.Name

    family_name = ""
    type_name = ""

    # Получаем информацию о семействе и типе
    if isinstance(element, FamilyInstance):
        symbol = element.Symbol
        if symbol:
            family = symbol.Family
            if family:
                family_name = family.Name
            type_name = symbol.Name
    else:
        type_id = element.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            elem_type = doc.GetElement(type_id)
            if elem_type:
                type_name = elem_type.Name
                try:
                    family_name = elem_type.FamilyName
                except Exception:
                    family_name = category

    if not family_name:
        family_name = category
    if not type_name:
        type_name = "Стандартный"

    # Формируем timestamp
    now = datetime.datetime.now()
    timestamp = now.strftime("%Y-%m-%dT%H:%M:%S")

    # Формируем payload
    event_payload = {
        "action": "created",
        "user": os.environ.get("USERNAME", "Unknown"),
        "machine": os.environ.get("COMPUTERNAME", "Unknown"),
        "revit_version": str(app.VersionNumber),
        "level_id": get_element_level_id(element),
        "workset_id": get_element_workset_id(element),
        "phase_id": get_element_phase_id(element),
        "parameters_count": element.Parameters.Size,
    }

    return {
        "project": proj_id,
        "unique_id": element.UniqueId,
        "element_id": str(element.Id.IntegerValue),
        "category": category,
        "family_name": family_name,
        "type_name": type_name,
        "last_event": "created",
        "event_timestamp": timestamp,
        "event_payload": event_payload,
    }


# Создаем клиент API
client = ProjectApiClient()

# Регистрируем элементы
success_count = 0
error_count = 0
errors = []

for element in elements_to_register:
    try:
        element_data = create_element_data(element, project_id)
        success, result, error = client.register_element(element_data)

        if success:
            success_count += 1
        else:
            error_count += 1
            errors.append("Элемент {}: {}".format(element.Id, error))

    except Exception as ex:
        error_count += 1
        errors.append("Элемент {}: {}".format(element.Id, str(ex)))

# Показываем результат
if error_count == 0:
    show_success(
        "Регистрация завершена",
        "Успешно зарегистрировано {} элементов.".format(success_count),
    )
else:
    details = "Успешно: {}\nОшибок: {}\n\n".format(success_count, error_count)
    if errors:
        # Показываем первые 10 ошибок
        first_errors = errors[:10]
        details += "Первые ошибки:\n" + "\n".join(first_errors)
        if len(errors) > 10:
            details += "\n... и ещё {} ошибок".format(len(errors) - 10)

    if success_count > 0:
        show_warning(
            "Регистрация завершена с ошибками",
            "Зарегистрировано {} из {} элементов.".format(
                success_count, success_count + error_count
            ),
            details=details,
        )
    else:
        show_error(
            "Ошибка регистрации",
            "Не удалось зарегистрировать элементы.",
            details=details,
            blocking=True,
        )
