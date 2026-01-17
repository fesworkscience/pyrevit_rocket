# -*- coding: utf-8 -*-
"""
Check Materials Category - проверка категории Материалы для привязки параметров.
Исследует возможности привязки параметров к материалам в Revit.
Результаты записываются в лог.
"""

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
Logger.info(SCRIPT_NAME, "Скрипт запущен")

# === 1. Поиск категории Materials по разным ID ===
Logger.log_separator(SCRIPT_NAME, "ПОИСК КАТЕГОРИЙ МАТЕРИАЛОВ ПО ID")

# Известные ID которые могут быть связаны с материалами
material_ids = [
    ("OST_Materials", -2000500),
    ("OST_MaterialQuantities", -2000600),
    ("OST_ProjectInformation", -2003101),  # для сравнения
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
            Logger.warning(SCRIPT_NAME, "{} (ID: {}) - категория не найдена в документе".format(name, cat_id))
    except Exception as e:
        Logger.error(SCRIPT_NAME, "{} (ID: {}) - ошибка: {}".format(name, cat_id, str(e)))

# === 2. Поиск всех категорий содержащих "Material" в имени ===
Logger.log_separator(SCRIPT_NAME, "ВСЕ КАТЕГОРИИ С 'MATERIAL' В BIC")

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
                Logger.debug(SCRIPT_NAME, "{} (ID: {}) - категория не найдена в документе".format(bic_name, int(bic)))
    except Exception as e:
        Logger.debug(SCRIPT_NAME, "Ошибка для {}: {}".format(str(bic), str(e)))

Logger.info(SCRIPT_NAME, "Найдено категорий с Material: {}".format(len(material_related)))

# === 3. Проверка существующих материалов в проекте ===
Logger.log_separator(SCRIPT_NAME, "МАТЕРИАЛЫ В ПРОЕКТЕ")

try:
    materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
    Logger.info(SCRIPT_NAME, "Найдено материалов в проекте: {}".format(len(materials)))

    # Проверяем первый материал для понимания структуры
    if materials:
        mat = materials[0]
        Logger.info(SCRIPT_NAME, "Пример материала: {}".format(mat.Name))

        mat_cat = mat.Category
        if mat_cat:
            Logger.info(SCRIPT_NAME, "  Категория: {} (ID: {})".format(mat_cat.Name, mat_cat.Id.IntegerValue))
            Logger.info(SCRIPT_NAME, "  AllowsBoundParameters: {}".format(mat_cat.AllowsBoundParameters))

            # Проверяем BuiltInCategory для этой категории
            try:
                bic_id = mat_cat.Id.IntegerValue
                Logger.info(SCRIPT_NAME, "  BuiltInCategory ID: {}".format(bic_id))
            except:
                pass
        else:
            Logger.warning(SCRIPT_NAME, "  У материала нет категории")

except Exception as e:
    Logger.error(SCRIPT_NAME, "Ошибка получения материалов: {}".format(str(e)))

# === 4. Проверка через BuiltInCategory напрямую ===
Logger.log_separator(SCRIPT_NAME, "ПРОВЕРКА OST_Materials НАПРЯМУЮ")

try:
    # Пробуем получить категорию Materials
    mat_cat = Category.GetCategory(doc, BuiltInCategory.OST_Materials)
    if mat_cat:
        Logger.info(SCRIPT_NAME, "OST_Materials найдена!")
        Logger.info(SCRIPT_NAME, "  Name: {}".format(mat_cat.Name))
        Logger.info(SCRIPT_NAME, "  Id: {}".format(mat_cat.Id.IntegerValue))
        Logger.info(SCRIPT_NAME, "  AllowsBoundParameters: {}".format(mat_cat.AllowsBoundParameters))
    else:
        Logger.warning(SCRIPT_NAME, "OST_Materials вернул None")
except Exception as e:
    Logger.error(SCRIPT_NAME, "OST_Materials ошибка: {}".format(str(e)))

# === 5. Итоги ===
Logger.log_separator(SCRIPT_NAME, "ИТОГИ")

# Подсчет категорий которые поддерживают привязку
allows_count = sum(1 for m in material_related if m['allows'])
Logger.info(SCRIPT_NAME, "Категорий Material с AllowsBoundParameters=True: {}".format(allows_count))
Logger.info(SCRIPT_NAME, "Категорий Material с AllowsBoundParameters=False: {}".format(len(material_related) - allows_count))

if allows_count > 0:
    Logger.info(SCRIPT_NAME, "Категории поддерживающие привязку:")
    for m in material_related:
        if m['allows']:
            Logger.info(SCRIPT_NAME, "  + {} | {} | ID: {}".format(m['name'], m['bic_name'], m['cat_id']))

Logger.info(SCRIPT_NAME, "Скрипт завершён")
Logger.info(SCRIPT_NAME, "Лог: {}".format(Logger.get_log_path()))



