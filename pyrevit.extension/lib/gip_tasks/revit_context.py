# -*- coding: utf-8 -*-
try:
    from Autodesk.Revit.DB import LocationPoint
except Exception:
    LocationPoint = None

MM_PER_FOOT = 304.8


def xyz_to_dict(point):
    if point is None:
        return {}
    return {"x": point.X * MM_PER_FOOT, "y": point.Y * MM_PER_FOOT, "z": point.Z * MM_PER_FOOT}


def active_view_info(doc):
    view = doc.ActiveView
    level = ""
    try:
        if view.GenLevel:
            level = view.GenLevel.Name
    except Exception:
        pass
    return {"id": view.Id.IntegerValue, "name": view.Name, "view_type": str(view.ViewType), "level": level}


def marker_info(marker):
    family_name = ""
    type_name = ""
    try:
        family_name = marker.Symbol.Family.Name
        type_name = marker.Symbol.Name
    except Exception:
        pass
    location = {}
    try:
        if marker.Location:
            location = xyz_to_dict(marker.Location.Point)
    except Exception:
        pass
    return {
        "element_id": marker.Id.IntegerValue,
        "unique_id": getattr(marker, "UniqueId", ""),
        "family_name": family_name,
        "type_name": type_name,
        "location": location,
    }


def collect(uidoc, marker=None):
    doc = uidoc.Document
    data = {
        "document_title": doc.Title,
        "document_path": getattr(doc, "PathName", ""),
        "revit_username": doc.Application.Username,
        "active_view": active_view_info(doc),
    }
    if marker is not None:
        data["marker"] = marker_info(marker)
    return data

