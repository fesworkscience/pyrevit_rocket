# -*- coding: utf-8 -*-
"""
Parameter utilities for Revit elements.
"""

from Autodesk.Revit.DB import StorageType, BuiltInParameter

def get_param(element, param_name):
    """Get parameter value from element."""
    param = None

    if isinstance(param_name, BuiltInParameter):
        param = element.get_Parameter(param_name)
    else:
        param = element.LookupParameter(param_name)

    if not param:
        return None

    storage_type = param.StorageType

    if storage_type == StorageType.String:
        return param.AsString()
    elif storage_type == StorageType.Integer:
        return param.AsInteger()
    elif storage_type == StorageType.Double:
        return param.AsDouble()
    elif storage_type == StorageType.ElementId:
        return param.AsElementId()

    return None

def set_param(element, param_name, value):
    """Set parameter value on element."""
    param = None

    if isinstance(param_name, BuiltInParameter):
        param = element.get_Parameter(param_name)
    else:
        param = element.LookupParameter(param_name)

    if not param or param.IsReadOnly:
        return False

    try:
        param.Set(value)
        return True
    except:
        return False

def get_all_params(element):
    """Get all parameters from element as dict."""
    result = {}

    for param in element.Parameters:
        name = param.Definition.Name
        result[name] = get_param(element, name)

    return result
