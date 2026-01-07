# -*- coding: utf-8 -*-
"""
Selection utilities for Revit elements.
"""

from Autodesk.Revit.DB import ElementId
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List

from cpsk_utils import get_doc, get_uidoc

def get_selected_elements():
    """Get currently selected elements."""
    doc = get_doc()
    uidoc = get_uidoc()
    selection = uidoc.Selection.GetElementIds()
    return [doc.GetElement(eid) for eid in selection]

def select_elements(elements):
    """Select elements in Revit UI."""
    uidoc = get_uidoc()

    if not elements:
        return

    element_ids = []
    for el in elements:
        if isinstance(el, ElementId):
            element_ids.append(el)
        else:
            element_ids.append(el.Id)

    id_list = List[ElementId](element_ids)
    uidoc.Selection.SetElementIds(id_list)

def pick_element(message="Select element"):
    """Prompt user to pick single element."""
    doc = get_doc()
    uidoc = get_uidoc()

    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, message)
        return doc.GetElement(ref.ElementId)
    except:
        return None

def pick_elements(message="Select elements"):
    """Prompt user to pick multiple elements."""
    doc = get_doc()
    uidoc = get_uidoc()

    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, message)
        return [doc.GetElement(ref.ElementId) for ref in refs]
    except:
        return []
