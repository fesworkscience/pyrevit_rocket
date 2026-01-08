#! python3
# -*- coding: utf-8 -*-
"""
Обновление спецификации для отображения эскизов.
Добавляет или показывает поле "Изображение" в спецификации.
"""

__title__ = "Обновить\nспецификацию"
__author__ = "CPSK"

import os
import sys

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Проверка авторизации
from cpsk_auth import require_auth
if not require_auth():
    sys.exit()

from cpsk_notify import show_error, show_warning, show_success, show_info

from pyrevit import revit

from Autodesk.Revit.DB import ViewSchedule, Transaction

doc = revit.doc
uidoc = revit.uidoc


def find_image_field(schedule_def):
    """Найти поле изображения в спецификации."""
    for i in range(schedule_def.GetFieldCount()):
        field = schedule_def.GetField(i)
        name = field.GetName()
        if name in ["Изображение", "Image", "Эскиз", "Рисунок"]:
            return field
    return None


def find_schedulable_image_field(schedule_def):
    """Найти добавляемое поле изображения."""
    schedulable_fields = schedule_def.GetSchedulableFields()
    for sf in schedulable_fields:
        name = sf.GetName(doc)
        if name in ["Изображение", "Image", "Эскиз", "Рисунок"]:
            return sf
    return None


def main():
    """Основная функция."""
    # Получить активную спецификацию
    schedule = doc.ActiveView
    if not isinstance(schedule, ViewSchedule):
        # Попробовать получить из выделения
        selected_ids = uidoc.Selection.GetElementIds()
        if selected_ids.Count > 0:
            elem = doc.GetElement(list(selected_ids)[0])
            if isinstance(elem, ViewSchedule):
                schedule = elem

    if not isinstance(schedule, ViewSchedule):
        show_error("Ошибка", "Пожалуйста, откройте спецификацию или выберите её в браузере проекта.")
        return

    # Проверить префикс
    if not schedule.Name.startswith("CPSK_VD_"):
        show_warning(
            "Внимание",
            "Выбранная спецификация не имеет префикса CPSK_VD_",
            details="Спецификация: {}\n\nПродолжить обновление?".format(schedule.Name)
        )
        # В WinForms диалоге нужно было бы запросить подтверждение,
        # но cpsk_notify не поддерживает Yes/No. Продолжаем.

    schedule_def = schedule.Definition

    # Найти или добавить поле изображения
    with revit.Transaction("Обновить спецификацию"):
        image_field = find_image_field(schedule_def)

        if image_field is None:
            # Попробовать добавить поле
            schedulable_field = find_schedulable_image_field(schedule_def)
            if schedulable_field:
                try:
                    image_field = schedule_def.AddField(schedulable_field)
                    show_success("Успех", "Поле изображения добавлено в спецификацию.")
                except Exception as ex:
                    show_error("Ошибка", "Не удалось добавить поле изображения: {}".format(str(ex)))
                    return
            else:
                show_warning(
                    "Предупреждение",
                    "Не удалось найти поле для изображений.",
                    details="Убедитесь, что в семействах есть параметр изображения."
                )
                return
        else:
            # Поле существует, проверить видимость
            if image_field.IsHidden:
                image_field.IsHidden = False
                show_success("Успех", "Поле изображения сделано видимым.")
            else:
                show_info("Информация", "Поле изображения уже настроено и видимо в спецификации.")
                return

        # Настроить параметры отображения
        if image_field:
            try:
                format_options = image_field.GetFormatOptions()
                format_options.UseDefault = False
                image_field.SetFormatOptions(format_options)
                image_field.ColumnHeading = "Эскиз"
            except Exception:
                pass

        # Показать заголовки
        schedule_def.ShowHeaders = True

    # Обновить вид
    uidoc.RefreshActiveView()

    show_success(
        "Завершено",
        "Спецификация '{}' обновлена для отображения эскизов.".format(schedule.Name),
        details="Убедитесь, что в семействах загружены изображения эскизов."
    )


if __name__ == "__main__":
    main()
