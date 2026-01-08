#! python3
# -*- coding: utf-8 -*-
"""Show bounding box info for selected elements."""

__title__ = "Bounding\nBox"
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
    def to_mm(feet):
        return round(feet * 304.8, 2)

    for eid in selection:
        element = doc.GetElement(eid)
        bbox = element.get_BoundingBox(doc.ActiveView)

        output.print_md("## Element ID: {}".format(element.Id))

        if bbox:
            min_pt = bbox.Min
            max_pt = bbox.Max

            width = to_mm(max_pt.X - min_pt.X)
            depth = to_mm(max_pt.Y - min_pt.Y)
            height = to_mm(max_pt.Z - min_pt.Z)

            output.print_md("### Bounding Box (mm):")
            output.print_md("- **Min**: ({}, {}, {})".format(
                to_mm(min_pt.X), to_mm(min_pt.Y), to_mm(min_pt.Z)
            ))
            output.print_md("- **Max**: ({}, {}, {})".format(
                to_mm(max_pt.X), to_mm(max_pt.Y), to_mm(max_pt.Z)
            ))
            output.print_md("- **Size**: {} x {} x {} mm".format(width, depth, height))
        else:
            output.print_md("*No bounding box available*")

        output.print_md("---")
