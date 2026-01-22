# -*- coding: utf-8 -*-
"""
Раскраска SLAM - применение цветов к DirectShape на текущем виде.

Читает цвет из марки элемента в формате:
SLAM_0.3m_RGB(176,144,144) -> применяет цвет RGB(176,144,144)
"""

__title__ = "Раскрасить\nSLAM"
__author__ = "CPSK"

import clr
import re
import sys
import os

clr.AddReference('System.Windows.Forms')

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from cpsk_notify import show_error, show_warning, show_success, show_info
from cpsk_auth import require_auth

if not require_auth():
    sys.exit()

from pyrevit import revit
from Autodesk.Revit.DB import (
    FilteredElementCollector, DirectShape, BuiltInCategory,
    BuiltInParameter, OverrideGraphicSettings, Transaction
)
from Autodesk.Revit.DB import Color as RevitColor

doc = revit.doc
uidoc = revit.uidoc


def parse_rgb_from_mark(mark):
    """
    Извлекает RGB цвет из марки элемента.

    Форматы:
        SLAM_0.3m_RGB(176,144,144)
        SLAM_1.5-2.0m_RGB(80,112,144)
        RGB(255,0,0)

    Args:
        mark: строка марки элемента

    Returns:
        tuple: (r, g, b) или None если не найдено
    """
    if not mark:
        return None

    # Ищем паттерн RGB(r,g,b)
    pattern = r'RGB\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)'
    match = re.search(pattern, mark)

    if match:
        r = int(match.group(1))
        g = int(match.group(2))
        b = int(match.group(3))
        return (r, g, b)

    return None


def apply_color_override(view, element_id, r, g, b):
    """
    Применяет цвет к элементу через Override Graphics на виде.

    Args:
        view: вид Revit
        element_id: ID элемента
        r, g, b: компоненты цвета 0-255
    """
    ogs = OverrideGraphicSettings()
    color = RevitColor(int(r), int(g), int(b))

    # Устанавливаем цвет линий проекции
    ogs.SetProjectionLineColor(color)

    # Применяем к элементу на виде
    view.SetElementOverrides(element_id, ogs)


def get_slam_directshapes(doc):
    """
    Получает все DirectShape элементы SLAM (с маркой содержащей RGB).

    Returns:
        list: [(element, mark, (r,g,b)), ...]
    """
    result = []

    collector = FilteredElementCollector(doc).OfClass(DirectShape)

    for ds in collector:
        mark_param = ds.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if not mark_param:
            continue

        mark = mark_param.AsString()
        if not mark:
            continue

        # Проверяем что это SLAM элемент с цветом
        if "SLAM" not in mark:
            continue

        rgb = parse_rgb_from_mark(mark)
        if rgb:
            result.append((ds, mark, rgb))

    return result


def main():
    """Основная функция."""
    active_view = uidoc.ActiveView

    if active_view is None:
        show_error("Ошибка", "Нет активного вида")
        return

    # Получаем SLAM элементы
    slam_elements = get_slam_directshapes(doc)

    if not slam_elements:
        show_warning(
            "Нет элементов",
            "Не найдены DirectShape элементы SLAM с цветовой информацией.",
            details="Ожидаемый формат марки:\nSLAM_0.3m_RGB(176,144,144)"
        )
        return

    # Группируем по цветам для статистики
    color_stats = {}
    for elem, mark, rgb in slam_elements:
        if rgb not in color_stats:
            color_stats[rgb] = 0
        color_stats[rgb] += 1

    # Применяем цвета
    with Transaction(doc, "Раскраска SLAM") as t:
        t.Start()

        colored_count = 0
        for elem, mark, rgb in slam_elements:
            try:
                apply_color_override(active_view, elem.Id, rgb[0], rgb[1], rgb[2])
                colored_count += 1
            except Exception:
                continue

        t.Commit()

    # Формируем отчёт
    log_lines = [
        "Вид: {}".format(active_view.Name),
        "",
        "Раскрашено элементов: {}".format(colored_count),
        "Уникальных цветов: {}".format(len(color_stats)),
        "",
        "Цвета:"
    ]

    for rgb, count in sorted(color_stats.items(), key=lambda x: -x[1]):
        log_lines.append("  RGB({},{},{}) - {} элементов".format(
            rgb[0], rgb[1], rgb[2], count
        ))

    show_success(
        "Раскраска завершена",
        "Раскрашено {} элементов на виде '{}'".format(colored_count, active_view.Name),
        details="\n".join(log_lines)
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        show_error("Ошибка", "Ошибка выполнения скрипта", details=str(e))
