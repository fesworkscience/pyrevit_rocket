# -*- coding: utf-8 -*-
"""Check Materials Category - debug script for parameter binding."""

__title__ = "Check\nMaterials"
__author__ = "CPSK"

from pyrevit import revit, script
from Autodesk.Revit.DB import BuiltInCategory, Category, FilteredElementCollector, Material
import System

# Import logger
from cpsk_logger import Logger

SCRIPT_NAME = "CheckMaterials"

doc = revit.doc
output = script.get_output()

# Initialize logger
Logger.init(SCRIPT_NAME)
Logger.info(SCRIPT_NAME, "Script started")

# === 1. Search Materials category by IDs ===
Logger.log_separator(SCRIPT_NAME, "SEARCH MATERIAL CATEGORIES BY ID")

material_ids = [
    ("OST_Materials", -2000500),
    ("OST_MaterialQuantities", -2000600),
    ("OST_ProjectInformation", -2003101),
]

for name, cat_id in material_ids:
    try:
        bic = BuiltInCategory(cat_id)
        cat = Category.GetCategory(doc, bic)
        if cat:
            allows = cat.AllowsBoundParameters
            Logger.info(SCRIPT_NAME, "{} | {} | ID: {} | AllowsBoundParameters={}".format(
                cat.Name, name, cat_id, allows
            ))
        else:
            Logger.warning(SCRIPT_NAME, "{} (ID: {}) - category not found".format(name, cat_id))
    except Exception as e:
        Logger.error(SCRIPT_NAME, "{} (ID: {}) - error: {}".format(name, cat_id, str(e)))
        continue

# === 2. Search categories with "Material" in BIC name ===
Logger.log_separator(SCRIPT_NAME, "ALL CATEGORIES WITH MATERIAL IN BIC")

material_related = []

for bic in System.Enum.GetValues(BuiltInCategory):
    try:
        if str(bic) == "INVALID" or int(bic) == -1:
            continue

        bic_name = str(bic)
        if "MATERIAL" in bic_name.upper():
            cat = Category.GetCategory(doc, bic)
            if cat:
                allows = cat.AllowsBoundParameters
                Logger.info(SCRIPT_NAME, "{} | {} | ID: {} | AllowsBoundParameters={}".format(
                    cat.Name, bic_name, int(bic), allows
                ))
                material_related.append({
                    'name': cat.Name,
                    'bic_name': bic_name,
                    'cat_id': int(bic),
                    'allows': allows
                })
            else:
                Logger.debug(SCRIPT_NAME, "{} (ID: {}) - not found".format(bic_name, int(bic)))
    except Exception as e:
        Logger.debug(SCRIPT_NAME, "Error for {}: {}".format(str(bic), str(e)))
        continue

Logger.info(SCRIPT_NAME, "Found Material categories: {}".format(len(material_related)))

# === 3. Check materials in project ===
Logger.log_separator(SCRIPT_NAME, "MATERIALS IN PROJECT")

try:
    materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
    Logger.info(SCRIPT_NAME, "Materials in project: {}".format(len(materials)))

    if materials:
        mat = materials[0]
        Logger.info(SCRIPT_NAME, "Example: {}".format(mat.Name))

        mat_cat = mat.Category
        if mat_cat:
            Logger.info(SCRIPT_NAME, "  Category: {} (ID: {})".format(mat_cat.Name, mat_cat.Id.IntegerValue))
            Logger.info(SCRIPT_NAME, "  AllowsBoundParameters: {}".format(mat_cat.AllowsBoundParameters))
            bic_id = mat_cat.Id.IntegerValue
            Logger.info(SCRIPT_NAME, "  BuiltInCategory ID: {}".format(bic_id))
        else:
            Logger.warning(SCRIPT_NAME, "  Material has no category")

except Exception as e:
    Logger.error(SCRIPT_NAME, "Error getting materials: {}".format(str(e)))
    raise

# === 4. Check OST_Materials directly ===
Logger.log_separator(SCRIPT_NAME, "CHECK OST_Materials DIRECTLY")

try:
    mat_cat = Category.GetCategory(doc, BuiltInCategory.OST_Materials)
    if mat_cat:
        Logger.info(SCRIPT_NAME, "OST_Materials found!")
        Logger.info(SCRIPT_NAME, "  Name: {}".format(mat_cat.Name))
        Logger.info(SCRIPT_NAME, "  Id: {}".format(mat_cat.Id.IntegerValue))
        Logger.info(SCRIPT_NAME, "  AllowsBoundParameters: {}".format(mat_cat.AllowsBoundParameters))
    else:
        Logger.warning(SCRIPT_NAME, "OST_Materials returned None")
except Exception as e:
    Logger.error(SCRIPT_NAME, "OST_Materials error: {}".format(str(e)))
    raise

# === 5. Summary ===
Logger.log_separator(SCRIPT_NAME, "SUMMARY")

allows_count = sum(1 for m in material_related if m['allows'])
Logger.info(SCRIPT_NAME, "Material categories with AllowsBoundParameters=True: {}".format(allows_count))
Logger.info(SCRIPT_NAME, "Material categories with AllowsBoundParameters=False: {}".format(len(material_related) - allows_count))

if allows_count > 0:
    Logger.info(SCRIPT_NAME, "Categories supporting binding:")
    for m in material_related:
        if m['allows']:
            Logger.info(SCRIPT_NAME, "  + {} | {} | ID: {}".format(m['name'], m['bic_name'], m['cat_id']))

Logger.info(SCRIPT_NAME, "Script finished")
Logger.info(SCRIPT_NAME, "Log: {}".format(Logger.get_log_path()))
