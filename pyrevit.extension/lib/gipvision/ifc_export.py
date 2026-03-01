# coding: utf-8
import os

from Autodesk.Revit.DB import IFCExportOptions


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
