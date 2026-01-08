# -*- coding: utf-8 -*-
"""
Показ скрытых строк в спецификациях CPSK_VD_.
Удаляет фильтры скрытия (##HIDE_ALL_ROWS##) из всех спецификаций CPSK_VD_.
"""

__title__ = "Показать\nстроки"
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

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, Transaction,
    ScheduleFilterType
)

doc = revit.doc


def remove_hide_filters(schedule):
    """Удалить фильтры скрытия из спецификации."""
    try:
        schedule_def = schedule.Definition
        filters = schedule_def.GetFilters()

        # Найти индексы фильтров скрытия
        filter_indices = []
        for i in range(len(filters)):
            try:
                f = filters[i]
                if f.FilterType == ScheduleFilterType.Equal:
                    try:
                        value = f.GetStringValue()
                        if value == "##HIDE_ALL_ROWS##":
                            filter_indices.append(i)
                    except Exception:
                        # Не строковый фильтр
                        pass
            except Exception:
                continue

        if not filter_indices:
            return 0  # Нет фильтров для удаления

        # Удалить фильтры в обратном порядке
        for i in reversed(filter_indices):
            schedule_def.RemoveFilter(i)

        # Обновить спецификацию
        schedule.RefreshData()

        return len(filter_indices)

    except Exception as ex:
        raise ex


def main():
    """Основная функция."""
    # Найти все спецификации CPSK_VD_
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    cpsk_schedules = [s for s in collector if s.Name.startswith("CPSK_VD_")]
    cpsk_schedules.sort(key=lambda x: x.Name)

    if not cpsk_schedules:
        show_info("Информация", "В проекте не найдено спецификаций с префиксом CPSK_VD_")
        return

    # Показать все строки во всех спецификациях CPSK_VD_
    with revit.Transaction("Показать все строки в спецификациях CPSK_VD_"):
        restored_count = 0
        errors = []

        for schedule in cpsk_schedules:
            try:
                removed = remove_hide_filters(schedule)
                if removed > 0:
                    restored_count += 1
            except Exception as ex:
                errors.append("{}: {}".format(schedule.Name, str(ex)))

        # Показать результат
        if restored_count > 0:
            message = "Строки восстановлены в {} спецификациях CPSK_VD_".format(restored_count)
            if errors:
                details = "Ошибки:\n" + "\n".join(errors)
                show_warning("Частично выполнено", message, details=details)
            else:
                show_success("Успех", message)
        else:
            if errors:
                show_error("Ошибка", "Не удалось восстановить строки",
                          details="\n".join(errors))
            else:
                show_info("Информация", "Все строки в спецификациях CPSK_VD_ уже видимы")


if __name__ == "__main__":
    main()
