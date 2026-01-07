# -*- coding: utf-8 -*-
"""
Geometry utilities for Revit elements.
"""

from Autodesk.Revit.DB import Options, ViewDetailLevel

from cpsk_utils import get_active_view

def get_bounding_box(element, view=None):
    """Get element bounding box."""
    if view is None:
        view = get_active_view()
    return element.get_BoundingBox(view)

def get_location(element):
    """Get element location point or curve."""
    loc = element.Location

    if not loc:
        return None

    if hasattr(loc, 'Point'):
        return loc.Point
    elif hasattr(loc, 'Curve'):
        return loc.Curve

    return None

def get_geometry(element, detail_level=ViewDetailLevel.Fine):
    """Get element geometry."""
    opts = Options()
    opts.DetailLevel = detail_level
    opts.ComputeReferences = True
    return element.get_Geometry(opts)

def get_center(bbox):
    """Get center point of bounding box."""
    if not bbox:
        return None
    return (bbox.Min + bbox.Max) / 2
