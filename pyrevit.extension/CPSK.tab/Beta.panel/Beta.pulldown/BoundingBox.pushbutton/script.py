# -*- coding: utf-8 -*-
"""Show bounding box info for selected elements."""

__title__ = "Bounding\nBox"
__author__ = "CPSK"

from pyrevit import revit, forms, script

output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

selection = uidoc.Selection.GetElementIds()

if not selection:
    forms.alert("Please select elements first")
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
