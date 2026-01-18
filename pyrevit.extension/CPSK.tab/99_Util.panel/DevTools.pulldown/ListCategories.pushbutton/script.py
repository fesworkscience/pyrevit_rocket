# -*- coding: utf-8 -*-
"""
List Revit Categories - output all available categories with detailed info.
Includes AllowsBoundParameters property for debugging parameter binding issues.
"""

__title__ = "List\nCategories"
__author__ = "CPSK"

from pyrevit import revit, script
from Autodesk.Revit.DB import BuiltInCategory, Category
import System

# Import logger
try:
    from cpsk_logger import Logger
    HAS_LOGGER = True
except ImportError as e:
    HAS_LOGGER = False
    raise

SCRIPT_NAME = "ListCategories"

doc = revit.doc
output = script.get_output()

# Initialize logger
if HAS_LOGGER:
    Logger.init(SCRIPT_NAME)
    Logger.info(SCRIPT_NAME, "Скрипт запущен")

output.print_md("# Revit Categories - Detailed Info")
output.print_md("---")

# Get all BuiltInCategory enum values with detailed info
all_categories = []
allows_binding = []
not_allows_binding = []

for bic in System.Enum.GetValues(BuiltInCategory):
    try:
        # Skip INVALID
        if str(bic) == "INVALID" or int(bic) == -1:
            continue

        # Get category
        cat = Category.GetCategory(doc, bic)
        if cat:
            # Get all properties
            name = cat.Name
            bic_name = str(bic)
            cat_id = int(bic)
            allows_params = cat.AllowsBoundParameters

            cat_info = {
                'name': name,
                'bic_name': bic_name,
                'cat_id': cat_id,
                'allows_params': allows_params
            }
            all_categories.append(cat_info)

            if allows_params:
                allows_binding.append(cat_info)
            else:
                not_allows_binding.append(cat_info)
    except Exception as e:
        if HAS_LOGGER:
            Logger.debug(SCRIPT_NAME, "Ошибка для {}: {}".format(str(bic), str(e)))
        continue

# Sort by name
all_categories.sort(key=lambda x: x['name'])
allows_binding.sort(key=lambda x: x['name'])
not_allows_binding.sort(key=lambda x: x['name'])

# Log summary
if HAS_LOGGER:
    Logger.info(SCRIPT_NAME, "Всего категорий: {}".format(len(all_categories)))
    Logger.info(SCRIPT_NAME, "Допускают привязку параметров: {}".format(len(allows_binding)))
    Logger.info(SCRIPT_NAME, "НЕ допускают привязку параметров: {}".format(len(not_allows_binding)))

# Output summary
output.print_md("## Summary")
output.print_md("- **Total categories:** {}".format(len(all_categories)))
output.print_md("- **AllowsBoundParameters = True:** {}".format(len(allows_binding)))
output.print_md("- **AllowsBoundParameters = False:** {}".format(len(not_allows_binding)))
output.print_md("")

# === Categories that DON'T allow bound parameters ===
output.print_md("---")
output.print_md("## Categories that DON'T allow bound parameters ({})".format(len(not_allows_binding)))
output.print_md("These categories cannot have shared parameters bound to them:")
output.print_md("")

if HAS_LOGGER:
    Logger.log_separator(SCRIPT_NAME, "КАТЕГОРИИ БЕЗ ПРИВЯЗКИ ПАРАМЕТРОВ")

for cat_info in not_allows_binding:
    line = "- **{}** | `{}` | ID: {}".format(
        cat_info['name'],
        cat_info['bic_name'],
        cat_info['cat_id']
    )
    output.print_md(line)
    if HAS_LOGGER:
        Logger.info(SCRIPT_NAME, "{} | {} | ID: {}".format(
            cat_info['name'],
            cat_info['bic_name'],
            cat_info['cat_id']
        ))

output.print_md("")

# === Check specific categories we're interested in ===
output.print_md("---")
output.print_md("## Specific Categories Check (IFC-related)")

# List of categories we use in IFC mapping
target_categories = [
    ("OST_ProjectInformation", -2000201),
    ("OST_Rooms", -2000160),
    ("OST_MEPSpaces", -2000350),
    ("OST_Areas", -2000003),
    ("OST_Zones", -2000356),
    ("OST_Walls", -2000011),
    ("OST_Floors", -2000032),
    ("OST_Roofs", -2000035),
    ("OST_StructuralColumns", -2001330),
    ("OST_Doors", -2000023),
    ("OST_Windows", -2000014),
    ("OST_Stairs", -2000120),
    ("OST_MechanicalEquipment", -2001140),
    ("OST_Furniture", -2000080),
    ("OST_CurtainWallPanels", -2000170),
]

if HAS_LOGGER:
    Logger.log_separator(SCRIPT_NAME, "ПРОВЕРКА ЦЕЛЕВЫХ КАТЕГОРИЙ")

output.print_md("")
output.print_md("| Категория | BuiltInCategory | ID | AllowsBoundParameters |")
output.print_md("|-----------|-----------------|----|-----------------------|")

for bic_name, cat_id in target_categories:
    bic = None
    try:
        bic = getattr(BuiltInCategory, bic_name.replace("OST_", ""))
    except AttributeError:
        # Try with full name
        for b in System.Enum.GetValues(BuiltInCategory):
            if str(b) == bic_name:
                bic = b
                break
        continue

    if bic:
        cat = Category.GetCategory(doc, bic)
        if cat:
            allows = cat.AllowsBoundParameters
            allows_str = "✓ Yes" if allows else "✗ NO"
            output.print_md("| {} | {} | {} | {} |".format(
                cat.Name, bic_name, cat_id, allows_str
            ))
            if HAS_LOGGER:
                Logger.info(SCRIPT_NAME, "{} | {} | {} | AllowsBinding={}".format(
                    cat.Name, bic_name, cat_id, allows
                ))
        else:
            output.print_md("| ? | {} | {} | Category not found |".format(bic_name, cat_id))
            if HAS_LOGGER:
                Logger.warning(SCRIPT_NAME, "Категория не найдена: {} (ID: {})".format(bic_name, cat_id))
    else:
        output.print_md("| ? | {} | {} | BIC not found |".format(bic_name, cat_id))
        if HAS_LOGGER:
            Logger.warning(SCRIPT_NAME, "BuiltInCategory не найден: {}".format(bic_name))

output.print_md("")

# === Categories that ALLOW bound parameters ===
output.print_md("---")
output.print_md("## Categories that ALLOW bound parameters ({})".format(len(allows_binding)))
output.print_md("These can be used for shared parameter binding:")
output.print_md("")

if HAS_LOGGER:
    Logger.log_separator(SCRIPT_NAME, "КАТЕГОРИИ С ПРИВЯЗКОЙ ПАРАМЕТРОВ")

# Group by first letter for readability
from collections import defaultdict
grouped = defaultdict(list)
for cat_info in allows_binding:
    first_letter = cat_info['name'][0].upper() if cat_info['name'] else '?'
    grouped[first_letter].append(cat_info)

for letter in sorted(grouped.keys()):
    output.print_md("### {}".format(letter))
    for cat_info in grouped[letter]:
        output.print_md("- {} | `{}` | ID: {}".format(
            cat_info['name'],
            cat_info['bic_name'],
            cat_info['cat_id']
        ))
        if HAS_LOGGER:
            Logger.debug(SCRIPT_NAME, "{} | {} | ID: {}".format(
                cat_info['name'],
                cat_info['bic_name'],
                cat_info['cat_id']
            ))
    output.print_md("")

# === Output for mapping ===
output.print_md("---")
output.print_md("## Python dict for IFC mapping (only categories that allow binding)")
output.print_md("```python")
output.print_md("IFC_TO_REVIT_CATEGORY_IDS = {")
for cat_info in allows_binding[:20]:  # First 20 as example
    output.print_md('    "IFCXXX": [("{}", {})],  # {}'.format(
        cat_info['name'],
        cat_info['cat_id'],
        cat_info['bic_name']
    ))
output.print_md("    # ... and {} more".format(len(allows_binding) - 20))
output.print_md("}")
output.print_md("```")

if HAS_LOGGER:
    Logger.info(SCRIPT_NAME, "Скрипт завершён")
    Logger.info(SCRIPT_NAME, "Лог сохранён в: {}".format(Logger.get_log_path()))
    output.print_md("")
    output.print_md("---")
    output.print_md("**Log saved to:** `{}`".format(Logger.get_log_path()))
