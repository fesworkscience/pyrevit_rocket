# coding: utf-8
import os

from Autodesk.Revit.DB import IFCExportOptions, View3D

from cpsk_logger import Logger


def sanitize_filename(text):
    invalid = '<>:"/\\|?*'
    result = text or "gipvision_export"
    for ch in invalid:
        result = result.replace(ch, "_")
    return result.strip(" .") or "gipvision_export"


def export_active_document_to_ifc(doc, export_folder, file_prefix):
    if not os.path.exists(export_folder):
        os.makedirs(export_folder)

    prefix = sanitize_filename(file_prefix)
    base_name = prefix

    # Avoid collisions: prefix.ifc, prefix_1.ifc, prefix_2.ifc, ...
    index = 0
    while True:
        suffix = "" if index == 0 else "_{0}".format(index)
        candidate = "{0}{1}.ifc".format(base_name, suffix)
        candidate_path = os.path.join(export_folder, candidate)
        if not os.path.exists(candidate_path):
            file_name_no_ext = os.path.splitext(candidate)[0]
            break
        index += 1

    options = IFCExportOptions()
    ok = doc.Export(export_folder, file_name_no_ext, options)
    if not ok:
        raise Exception("Revit IFC export returned False.")

    final_path = os.path.join(export_folder, file_name_no_ext + ".ifc")
    if not os.path.exists(final_path):
        raise Exception("IFC file not found after export: {0}".format(final_path))

    return final_path


def describe_view3d_state(view):
    if view is None:
        return {
            "view_name": "",
            "view_id": "",
            "is_template": False,
            "is_section_box_active": False
        }

    section_box_active = False
    try:
        section_box_active = bool(view.IsSectionBoxActive)
    except Exception:
        section_box_active = False

    return {
        "view_name": getattr(view, "Name", ""),
        "view_id": getattr(getattr(view, "Id", None), "IntegerValue", ""),
        "is_template": bool(getattr(view, "IsTemplate", False)),
        "is_section_box_active": section_box_active
    }


def export_current_3d_view_to_ifc(doc, view, export_folder, file_prefix, logger_name):
    if not doc:
        raise Exception("Active Revit document is not available.")
    if view is None:
        raise Exception("Active Revit view is not available.")
    if not isinstance(view, View3D):
        raise Exception("Active view must be a 3D view for GIP Vision IFC export.")
    if view.IsTemplate:
        raise Exception("Template 3D view cannot be exported to IFC.")

    Logger.log_separator(logger_name, "GIP Vision IFC export from active 3D view")
    Logger.info(logger_name, "Document title: {0}".format(getattr(doc, "Title", "")))
    Logger.info(logger_name, "Document path: {0}".format(getattr(doc, "PathName", "")))
    Logger.data(logger_name, "Active 3D view", describe_view3d_state(view))

    if not export_folder:
        raise Exception("Export folder is empty.")

    if not os.path.exists(export_folder):
        Logger.info(logger_name, "Creating export folder: {0}".format(export_folder))
        os.makedirs(export_folder)
    else:
        Logger.info(logger_name, "Using export folder: {0}".format(export_folder))

    prefix = sanitize_filename(file_prefix)
    Logger.info(logger_name, "Requested IFC file prefix: {0}".format(file_prefix))
    Logger.info(logger_name, "Sanitized IFC file prefix: {0}".format(prefix))
    base_name = prefix

    index = 0
    while True:
        suffix = "" if index == 0 else "_{0}".format(index)
        candidate = "{0}{1}.ifc".format(base_name, suffix)
        candidate_path = os.path.join(export_folder, candidate)
        if not os.path.exists(candidate_path):
            file_name_no_ext = os.path.splitext(candidate)[0]
            Logger.info(logger_name, "Resolved IFC output path: {0}".format(candidate_path))
            break
        Logger.warning(logger_name, "IFC file already exists, trying next name: {0}".format(candidate_path))
        index += 1

    options = IFCExportOptions()
    options.FilterViewId = view.Id
    options.AddOption("VisibleElementsOfCurrentView", "true")
    options.AddOption("UseActiveViewGeometry", "true")
    options.AddOption("ExportRoomsInView", "true")

    Logger.data(logger_name, "IFC export options", {
        "FilterViewId": view.Id.IntegerValue,
        "VisibleElementsOfCurrentView": "true",
        "UseActiveViewGeometry": "true",
        "ExportRoomsInView": "true"
    })

    Logger.info(logger_name, "Calling Revit Document.Export for current 3D view IFC export")
    ok = doc.Export(export_folder, file_name_no_ext, options)
    Logger.info(logger_name, "Document.Export returned: {0}".format(ok))
    if not ok:
        raise Exception("Revit IFC export returned False for current 3D view.")

    final_path = os.path.join(export_folder, file_name_no_ext + ".ifc")
    if not os.path.exists(final_path):
        raise Exception("IFC file not found after export: {0}".format(final_path))

    Logger.file_saved(logger_name, final_path, "IFC export from active 3D view")
    try:
        Logger.info(logger_name, "IFC file size: {0} bytes".format(os.path.getsize(final_path)))
    except Exception:
        Logger.warning(logger_name, "Unable to read exported IFC file size.")

    return final_path
