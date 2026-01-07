# -*- coding: utf-8 -*-
"""Show parameters of selected elements."""

__title__ = "Get\nParams"
__author__ = "CPSK"

from pyrevit import revit, forms, script

output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

selection = uidoc.Selection.GetElementIds()

if not selection:
    forms.alert("Please select elements first")
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
