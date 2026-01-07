# -*- coding: utf-8 -*-
"""
Core utilities for Revit API access.
"""

from Autodesk.Revit.DB import *
from pyrevit import revit, DB, UI

def get_doc():
    """Get active Revit document."""
    return revit.doc

def get_uidoc():
    """Get active UI document."""
    return revit.uidoc

def get_app():
    """Get Revit application."""
    return revit.app

def get_active_view():
    """Get active view."""
    return get_doc().ActiveView

def collect_elements(category=None, of_class=None, view_id=None):
    """Collect elements from document."""
    doc = get_doc()

    if view_id:
        collector = FilteredElementCollector(doc, view_id)
    else:
        collector = FilteredElementCollector(doc)

    if of_class:
        collector = collector.OfClass(of_class)

    if category:
        if isinstance(category, str):
            category = getattr(BuiltInCategory, category, None)
        if category:
            collector = collector.OfCategory(category)

    return list(collector.ToElements())

def feet_to_mm(feet):
    """Convert feet to millimeters."""
    return feet * 304.8

def mm_to_feet(mm):
    """Convert millimeters to feet."""
    return mm / 304.8
