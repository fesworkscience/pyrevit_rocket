# -*- coding: utf-8 -*-
"""Select and count elements by category."""

__title__ = "Select\nElements"
__author__ = "CPSK"

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
        forms.alert("Selected {} elements".format(len(elements)))
    else:
        forms.alert("No elements found")
