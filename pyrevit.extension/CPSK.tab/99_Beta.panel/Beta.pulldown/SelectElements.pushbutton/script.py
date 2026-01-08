# -*- coding: utf-8 -*-
"""Select and count elements by category."""

__title__ = "Select\nElements"
__author__ = "CPSK"

import os
import sys

# Добавляем lib в путь для импорта cpsk_auth
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Проверка авторизации
from cpsk_auth import require_auth
from cpsk_notify import show_info, show_warning
if not require_auth():
    sys.exit()

from pyrevit import revit, forms
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

doc = revit.doc

# Category selection
categories = {
    "Walls": BuiltInCategory.OST_Walls,
    "Floors": BuiltInCategory.OST_Floors,
    "Columns": BuiltInCategory.OST_StructuralColumns,
    "Beams": BuiltInCategory.OST_StructuralFraming,
    "Foundations": BuiltInCategory.OST_StructuralFoundation,
}

selected = forms.SelectFromList.show(
    sorted(categories.keys()),
    title="Select Category",
    multiselect=False
)

if selected:
    cat = categories[selected]
    collector = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType()
    elements = list(collector.ToElements())

    if elements:
        from cpsk_selection import select_elements
        select_elements(elements)
        show_info("Выбор", "Выбрано {} элементов".format(len(elements)))
    else:
        show_warning("Выбор", "Элементы не найдены")
