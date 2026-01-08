#! python3
# -*- coding: utf-8 -*-
"""Show parameters of selected elements."""

__title__ = "Get\nParams"
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
from cpsk_notify import show_warning
if not require_auth():
    sys.exit()

from pyrevit import revit, script

output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

selection = uidoc.Selection.GetElementIds()

if not selection:
    show_warning("Выбор", "Сначала выберите элементы")
else:
    for eid in selection:
        element = doc.GetElement(eid)
        output.print_md("## Element: {} (ID: {})".format(
            element.Name if hasattr(element, 'Name') else element.Id,
            element.Id
        ))

        output.print_md("### Parameters:")

        params = []
        for param in element.Parameters:
            name = param.Definition.Name
            value = None

            if param.StorageType.ToString() == "String":
                value = param.AsString()
            elif param.StorageType.ToString() == "Integer":
                value = param.AsInteger()
            elif param.StorageType.ToString() == "Double":
                value = round(param.AsDouble() * 304.8, 2)  # to mm
            elif param.StorageType.ToString() == "ElementId":
                value = param.AsElementId().IntegerValue

            params.append((name, value))

        params.sort(key=lambda x: x[0])

        for name, value in params:
            if value is not None:
                output.print_md("- **{}**: {}".format(name, value))

        output.print_md("---")
