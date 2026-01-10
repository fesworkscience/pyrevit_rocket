# -*- coding: utf-8 -*-
"""Автоматическая генерация разрезов и планов."""

__title__ = "Generate\nViews"
__author__ = "CPSK"

# 1. Сначала import clr и стандартные модули
import clr
import os
import sys
import math
from datetime import datetime

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

# 2. System и WinForms импорты
import System
from System.Windows.Forms import (
    Form, Label, Button, Panel, CheckedListBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, AnchorStyles, CheckBox, Padding,
    GroupBox, RadioButton, TextBox, NumericUpDown,
    ListBox, SelectionMode, ComboBox, OpenFileDialog,
    SaveFileDialog
)
from System.Drawing import Point, Size, Font, FontStyle, Color

# 3. Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# 4. Проверка авторизации
from cpsk_auth import require_auth
from cpsk_notify import show_error, show_success, show_warning, show_info
if not require_auth():
    sys.exit()

# 5. pyrevit и Revit API
from pyrevit import revit

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Transaction,
    BoundingBoxXYZ,
    XYZ,
    ViewSection,
    ViewPlan,
    View3D,
    ViewFamilyType,
    ElementId,
    Transform,
    Line,
    Outline,
    BoundingBoxIntersectsFilter,
    BoundingBoxIsInsideFilter,
    LogicalOrFilter,
    RevitLinkInstance,
    CategorySet,
    InstanceBinding,
    ExternalDefinitionCreationOptions,
    Category,
    SpecTypeId,
    GroupTypeId,
    ViewDetailLevel,
    Level,
    ViewOrientation3D
)

doc = revit.doc
app = doc.Application


# Константы
PARAM_NAME = "CPSK_AUTO"
STEP_METERS = 0.5
STEP_FEET = STEP_METERS / 0.3048  # Конвертация в футы
SHARED_PARAM_FILENAME = "CPSK_SharedParameters.txt"
PARAM_GROUP_NAME = "CPSK"


def get_default_shared_param_path():
    """Получить путь по умолчанию для файла общих параметров."""
    if doc.PathName:
        return os.path.join(os.path.dirname(doc.PathName), SHARED_PARAM_FILENAME)
    return os.path.join(os.path.expanduser("~"), "Documents", SHARED_PARAM_FILENAME)


def create_shared_param_file(filepath):
    """Создать новый файл общих параметров."""
    import codecs
    try:
        with codecs.open(filepath, 'w', 'utf-8') as f:
            f.write("# This is a Revit shared parameter file.\n")
            f.write("# Do not edit manually.\n")
            f.write("*META\tVERSION\tMINVERSION\n")
            f.write("META\t2\t1\n")
            f.write("*GROUP\tID\tNAME\n")
            f.write("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\n")
        return True
    except Exception as e:
        return False


def get_or_create_shared_param_file(filepath):
    """Получить или создать файл общих параметров."""
    if not os.path.exists(filepath):
        if not create_shared_param_file(filepath):
            return None

    try:
        app.SharedParametersFilename = filepath
        return app.OpenSharedParameterFile()
    except:
        return None


def get_or_create_param_group(shared_file, group_name):
    """Получить или создать группу параметров."""
    for group in shared_file.Groups:
        if group.Name == group_name:
            return group
    return shared_file.Groups.Create(group_name)


def get_or_create_param_definition(param_group, param_name):
    """Получить или создать определение параметра."""
    for definition in param_group.Definitions:
        if definition.Name == param_name:
            return definition

    # Создаём новый параметр (Text/String)
    try:
        options = ExternalDefinitionCreationOptions(param_name, SpecTypeId.String.Text)
    except:
        from Autodesk.Revit.DB import ParameterType
        options = ExternalDefinitionCreationOptions(param_name, ParameterType.Text)

    options.Description = "Тег автоматической генерации (разрезы, виды и т.д.)"
    options.UserModifiable = True
    options.Visible = True

    return param_group.Definitions.Create(options)


def add_param_to_views(param_definition):
    """Добавить параметр к категории Виды."""
    try:
        # Получаем категорию Views
        views_category = Category.GetCategory(doc, BuiltInCategory.OST_Views)
        if views_category is None:
            return False

        # Создаём CategorySet
        cat_set = app.Create.NewCategorySet()
        cat_set.Insert(views_category)

        # Создаём InstanceBinding
        binding = app.Create.NewInstanceBinding(cat_set)

        # Добавляем в проект
        binding_map = doc.ParameterBindings
        return binding_map.Insert(param_definition, binding, GroupTypeId.Data)
    except Exception as e:
        return False


def create_cpsk_auto_param(shared_param_path=None):
    """Создать параметр CPSK_AUTO для видов."""
    if shared_param_path is None:
        shared_param_path = get_default_shared_param_path()

    # Получаем или создаём файл параметров
    shared_file = get_or_create_shared_param_file(shared_param_path)
    if shared_file is None:
        return False, "Не удалось открыть/создать файл параметров"

    # Получаем или создаём группу
    param_group = get_or_create_param_group(shared_file, PARAM_GROUP_NAME)
    if param_group is None:
        return False, "Не удалось создать группу параметров"

    # Получаем или создаём определение
    param_def = get_or_create_param_definition(param_group, PARAM_NAME)
    if param_def is None:
        return False, "Не удалось создать определение параметра"

    # Добавляем к видам
    if add_param_to_views(param_def):
        return True, "Параметр {} успешно добавлен к видам".format(PARAM_NAME)
    else:
        return False, "Не удалось добавить параметр к категории Виды"


def get_all_model_elements(include_links=False):
    """Получить все элементы модели."""
    elements = []

    # Категории модели (не аннотации, не виды)
    model_categories = [
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_StructuralFoundation,
        BuiltInCategory.OST_Floors,
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_Roofs,
        BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_Columns,
        BuiltInCategory.OST_Stairs,
        BuiltInCategory.OST_Ramps,
        BuiltInCategory.OST_Railings,
        BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_Windows,
        BuiltInCategory.OST_Furniture,
        BuiltInCategory.OST_Casework,
        BuiltInCategory.OST_Ceilings,
        BuiltInCategory.OST_CurtainWallPanels,
        BuiltInCategory.OST_CurtainWallMullions,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_PlumbingFixtures,
        BuiltInCategory.OST_ElectricalEquipment,
        BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_LightingFixtures,
        BuiltInCategory.OST_SpecialityEquipment,
        BuiltInCategory.OST_Mass,
        BuiltInCategory.OST_Parts,
        BuiltInCategory.OST_StructuralTruss,
        BuiltInCategory.OST_StructuralStiffener,
        BuiltInCategory.OST_StructConnections,
    ]

    for cat in model_categories:
        try:
            collector = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType()
            elements.extend(list(collector))
        except:
            pass

    # Элементы из связанных файлов
    if include_links:
        links = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
        for link in links:
            try:
                link_doc = link.GetLinkDocument()
                if link_doc is None:
                    continue

                link_transform = link.GetTotalTransform()

                for cat in model_categories:
                    try:
                        link_collector = FilteredElementCollector(link_doc).OfCategory(cat).WhereElementIsNotElementType()
                        for elem in link_collector:
                            elements.append((elem, link_transform))
                    except:
                        pass
            except:
                pass

    return elements


def get_building_bounding_box(elements):
    """Получить BoundingBox всего здания."""
    min_x, min_y, min_z = float('inf'), float('inf'), float('inf')
    max_x, max_y, max_z = float('-inf'), float('-inf'), float('-inf')

    for item in elements:
        try:
            if isinstance(item, tuple):
                elem, transform = item
                bbox = elem.get_BoundingBox(None)
                if bbox:
                    # Трансформируем точки для связанных файлов
                    min_pt = transform.OfPoint(bbox.Min)
                    max_pt = transform.OfPoint(bbox.Max)
                else:
                    continue
            else:
                bbox = item.get_BoundingBox(None)
                if bbox is None:
                    continue
                min_pt = bbox.Min
                max_pt = bbox.Max

            min_x = min(min_x, min_pt.X, max_pt.X)
            min_y = min(min_y, min_pt.Y, max_pt.Y)
            min_z = min(min_z, min_pt.Z, max_pt.Z)
            max_x = max(max_x, min_pt.X, max_pt.X)
            max_y = max(max_y, min_pt.Y, max_pt.Y)
            max_z = max(max_z, min_pt.Z, max_pt.Z)
        except:
            pass

    if min_x == float('inf'):
        return None

    result = BoundingBoxXYZ()
    result.Min = XYZ(min_x, min_y, min_z)
    result.Max = XYZ(max_x, max_y, max_z)
    return result


def get_project_north_angle():
    """Получить угол проектного севера."""
    try:
        project_location = doc.ActiveProjectLocation
        project_position = project_location.GetProjectPosition(XYZ.Zero)
        return project_position.Angle  # В радианах
    except:
        return 0.0


def get_section_view_family_type():
    """Получить тип семейства для разреза."""
    collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)

    for vft in collector:
        try:
            if vft.ViewFamily.ToString() == "Section":
                return vft
        except:
            pass

    return None


def get_floor_plan_view_family_type():
    """Получить тип семейства для плана этажа."""
    collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)

    for vft in collector:
        try:
            if vft.ViewFamily.ToString() == "FloorPlan":
                return vft
        except:
            pass

    return None


def get_3d_view_family_type():
    """Получить тип семейства для 3D вида."""
    collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)

    for vft in collector:
        try:
            if vft.ViewFamily.ToString() == "ThreeDimensional":
                return vft
        except:
            pass

    return None


def get_elements_in_z_range(z_min, z_max, all_elements):
    """Получить элементы в заданном диапазоне высот."""
    element_ids = set()

    for item in all_elements:
        try:
            if isinstance(item, tuple):
                elem, transform = item
                elem_bbox = elem.get_BoundingBox(None)
                if elem_bbox:
                    min_pt = transform.OfPoint(elem_bbox.Min)
                    max_pt = transform.OfPoint(elem_bbox.Max)
                    elem_z_min = min(min_pt.Z, max_pt.Z)
                    elem_z_max = max(min_pt.Z, max_pt.Z)
                else:
                    continue
            else:
                elem_bbox = item.get_BoundingBox(None)
                if elem_bbox is None:
                    continue
                elem_z_min = elem_bbox.Min.Z
                elem_z_max = elem_bbox.Max.Z
                elem = item

            # Проверяем пересечение по Z
            if elem_z_min <= z_max and elem_z_max >= z_min:
                element_ids.add(elem.Id.IntegerValue)
        except:
            pass

    return element_ids


def get_elements_in_section(section_bbox, all_elements):
    """Получить элементы, попадающие в область разреза."""
    element_ids = set()

    for item in all_elements:
        try:
            if isinstance(item, tuple):
                elem, transform = item
                elem_bbox = elem.get_BoundingBox(None)
                if elem_bbox:
                    # Трансформируем для связей
                    min_pt = transform.OfPoint(elem_bbox.Min)
                    max_pt = transform.OfPoint(elem_bbox.Max)
                    elem_min = XYZ(min(min_pt.X, max_pt.X), min(min_pt.Y, max_pt.Y), min(min_pt.Z, max_pt.Z))
                    elem_max = XYZ(max(min_pt.X, max_pt.X), max(min_pt.Y, max_pt.Y), max(min_pt.Z, max_pt.Z))
                else:
                    continue
            else:
                elem_bbox = item.get_BoundingBox(None)
                if elem_bbox is None:
                    continue
                elem_min = elem_bbox.Min
                elem_max = elem_bbox.Max
                elem = item

            # Проверяем пересечение с областью разреза
            if boxes_intersect(section_bbox.Min, section_bbox.Max, elem_min, elem_max):
                element_ids.add(elem.Id.IntegerValue)
        except:
            pass

    return element_ids


def boxes_intersect(min1, max1, min2, max2):
    """Проверить пересечение двух BoundingBox."""
    return (min1.X <= max2.X and max1.X >= min2.X and
            min1.Y <= max2.Y and max1.Y >= min2.Y and
            min1.Z <= max2.Z and max1.Z >= min2.Z)


def create_section_bbox(origin, direction, right, up, width, height, depth):
    """Создать BoundingBox для разреза."""
    bbox = BoundingBoxXYZ()

    # Трансформация для разреза
    transform = Transform.Identity
    transform.Origin = origin
    transform.BasisX = right
    transform.BasisY = up
    transform.BasisZ = direction

    bbox.Transform = transform

    # Размеры (в локальных координатах)
    bbox.Min = XYZ(-width / 2, -0.5, 0)  # -0.5 фута снизу
    bbox.Max = XYZ(width / 2, height + 0.5, depth)  # +0.5 фута сверху

    return bbox


def check_cpsk_auto_param():
    """Проверить наличие параметра CPSK_AUTO на видах."""
    # Пробуем найти параметр на любом виде
    # Сначала пробуем разрезы
    sections = FilteredElementCollector(doc).OfClass(ViewSection).ToElements()
    if sections:
        param = sections[0].LookupParameter(PARAM_NAME)
        if param is not None:
            return True

    # Пробуем планы
    plans = FilteredElementCollector(doc).OfClass(ViewPlan).ToElements()
    if plans:
        param = plans[0].LookupParameter(PARAM_NAME)
        if param is not None:
            return True

    # Пробуем 3D виды
    views3d = FilteredElementCollector(doc).OfClass(View3D).ToElements()
    if views3d:
        param = views3d[0].LookupParameter(PARAM_NAME)
        if param is not None:
            return True

    # Нет видов для проверки - предполагаем что параметр не создан
    # Если видов нет, параметр точно нужно создать
    return False


def get_existing_generations():
    """Получить существующие генерации (разрезы, планы и 3D виды)."""
    generations = {}

    # Разрезы
    sections = FilteredElementCollector(doc).OfClass(ViewSection).ToElements()
    for section in sections:
        try:
            param = section.LookupParameter(PARAM_NAME)
            if param and param.HasValue:
                tag = param.AsString()
                if tag:
                    if tag not in generations:
                        generations[tag] = []
                    generations[tag].append(section)
        except:
            pass

    # Планы
    plans = FilteredElementCollector(doc).OfClass(ViewPlan).ToElements()
    for plan in plans:
        try:
            param = plan.LookupParameter(PARAM_NAME)
            if param and param.HasValue:
                tag = param.AsString()
                if tag:
                    if tag not in generations:
                        generations[tag] = []
                    generations[tag].append(plan)
        except:
            pass

    # 3D виды
    views3d = FilteredElementCollector(doc).OfClass(View3D).ToElements()
    for view3d in views3d:
        try:
            param = view3d.LookupParameter(PARAM_NAME)
            if param and param.HasValue:
                tag = param.AsString()
                if tag:
                    if tag not in generations:
                        generations[tag] = []
                    generations[tag].append(view3d)
        except:
            pass

    return generations


def delete_views_by_tag(tag):
    """Удалить виды (разрезы, планы и 3D виды) по тегу."""
    deleted = 0

    # Удаляем разрезы
    sections = FilteredElementCollector(doc).OfClass(ViewSection).ToElements()
    for section in sections:
        try:
            param = section.LookupParameter(PARAM_NAME)
            if param and param.HasValue and param.AsString() == tag:
                doc.Delete(section.Id)
                deleted += 1
        except:
            pass

    # Удаляем 3D виды
    views3d = FilteredElementCollector(doc).OfClass(View3D).ToElements()
    for view3d in views3d:
        try:
            param = view3d.LookupParameter(PARAM_NAME)
            if param and param.HasValue and param.AsString() == tag:
                doc.Delete(view3d.Id)
                deleted += 1
        except:
            pass

    # Удаляем планы и связанные уровни
    plans = FilteredElementCollector(doc).OfClass(ViewPlan).ToElements()
    levels_to_delete = []
    for plan in plans:
        try:
            param = plan.LookupParameter(PARAM_NAME)
            if param and param.HasValue and param.AsString() == tag:
                # Запоминаем уровень для удаления
                level_id = plan.GenLevel.Id
                levels_to_delete.append(level_id)
                doc.Delete(plan.Id)
                deleted += 1
        except:
            pass

    # Удаляем уровни (созданные для планов)
    for level_id in levels_to_delete:
        try:
            doc.Delete(level_id)
        except:
            pass

    return deleted


def generate_sections(direction_type, step_feet, include_links, tag, depth_feet=None, full_depth=False):
    """Генерировать разрезы.

    Args:
        direction_type: 'longitudinal', 'transverse', или 'both'
        step_feet: шаг в футах
        include_links: учитывать связанные файлы
        tag: тег для маркировки разрезов
        depth_feet: глубина разреза в футах (если None - используется step_feet)
        full_depth: если True - глубина на всю ширину/длину здания
    """
    # Получаем все элементы
    all_elements = get_all_model_elements(include_links)

    if not all_elements:
        show_warning("Нет элементов", "В модели не найдено элементов для создания разрезов")
        return 0

    # Получаем BoundingBox здания
    building_bbox = get_building_bounding_box(all_elements)
    if building_bbox is None:
        show_warning("Ошибка", "Не удалось определить границы здания")
        return 0

    # Получаем угол проектного севера
    north_angle = get_project_north_angle()

    # Направления (с учётом проектного севера)
    north_dir = XYZ(math.sin(north_angle), math.cos(north_angle), 0)
    east_dir = XYZ(math.cos(north_angle), -math.sin(north_angle), 0)
    up_dir = XYZ.BasisZ

    # Получаем тип разреза
    section_type = get_section_view_family_type()
    if section_type is None:
        show_error("Ошибка", "Не найден тип семейства для разрезов")
        return 0

    # Размеры здания
    width_ns = abs((building_bbox.Max - building_bbox.Min).DotProduct(north_dir))
    width_ew = abs((building_bbox.Max - building_bbox.Min).DotProduct(east_dir))
    height = building_bbox.Max.Z - building_bbox.Min.Z

    # Центр здания
    center = XYZ(
        (building_bbox.Min.X + building_bbox.Max.X) / 2,
        (building_bbox.Min.Y + building_bbox.Max.Y) / 2,
        building_bbox.Min.Z
    )

    created_sections = []

    # Продольные разрезы (смотрят на север, идут с юга на север)
    if direction_type in ["longitudinal", "both"]:
        # Определяем глубину для продольных разрезов
        if full_depth:
            section_depth = width_ns + 2  # На всю длину + запас
        elif depth_feet:
            section_depth = depth_feet
        else:
            section_depth = step_feet

        # Начинаем за зданием (с юга)
        start_offset = -width_ns / 2 - step_feet
        end_offset = width_ns / 2 + step_feet

        current_offset = start_offset
        prev_elements = set()
        section_num = 1

        while current_offset <= end_offset:
            # Позиция разреза
            origin = center + north_dir * current_offset

            # BoundingBox для проверки элементов
            section_check_min = XYZ(
                origin.X - east_dir.X * width_ew / 2 - north_dir.X * step_feet / 2,
                origin.Y - east_dir.Y * width_ew / 2 - north_dir.Y * step_feet / 2,
                building_bbox.Min.Z
            )
            section_check_max = XYZ(
                origin.X + east_dir.X * width_ew / 2 + north_dir.X * step_feet / 2,
                origin.Y + east_dir.Y * width_ew / 2 + north_dir.Y * step_feet / 2,
                building_bbox.Max.Z
            )

            check_bbox = BoundingBoxXYZ()
            check_bbox.Min = section_check_min
            check_bbox.Max = section_check_max

            # Получаем элементы в этой области
            current_elements = get_elements_in_section(check_bbox, all_elements)

            # Проверяем, есть ли НОВЫЕ элементы
            new_elements = current_elements - prev_elements

            if new_elements:
                # Создаём разрез
                section_bbox = create_section_bbox(
                    origin=origin,
                    direction=north_dir,
                    right=east_dir,
                    up=up_dir,
                    width=width_ew + 2,  # +2 фута запас
                    height=height,
                    depth=section_depth
                )

                try:
                    section = ViewSection.CreateSection(doc, section_type.Id, section_bbox)
                    section.Name = "{}_L_{:03d}".format(tag, section_num)

                    # Устанавливаем высокую детализацию
                    section.DetailLevel = ViewDetailLevel.Fine

                    # Устанавливаем параметр
                    param = section.LookupParameter(PARAM_NAME)
                    if param:
                        param.Set(tag)

                    created_sections.append(section)
                    section_num += 1
                except Exception as e:
                    pass

                prev_elements = current_elements.copy()

            current_offset += step_feet

    # Поперечные разрезы (смотрят на восток, идут с запада на восток)
    if direction_type in ["transverse", "both"]:
        # Определяем глубину для поперечных разрезов
        if full_depth:
            section_depth = width_ew + 2  # На всю ширину + запас
        elif depth_feet:
            section_depth = depth_feet
        else:
            section_depth = step_feet

        start_offset = -width_ew / 2 - step_feet
        end_offset = width_ew / 2 + step_feet

        current_offset = start_offset
        prev_elements = set()
        section_num = 1

        while current_offset <= end_offset:
            # Позиция разреза
            origin = center + east_dir * current_offset

            # BoundingBox для проверки элементов
            section_check_min = XYZ(
                origin.X - north_dir.X * width_ns / 2 - east_dir.X * step_feet / 2,
                origin.Y - north_dir.Y * width_ns / 2 - east_dir.Y * step_feet / 2,
                building_bbox.Min.Z
            )
            section_check_max = XYZ(
                origin.X + north_dir.X * width_ns / 2 + east_dir.X * step_feet / 2,
                origin.Y + north_dir.Y * width_ns / 2 + east_dir.Y * step_feet / 2,
                building_bbox.Max.Z
            )

            check_bbox = BoundingBoxXYZ()
            check_bbox.Min = section_check_min
            check_bbox.Max = section_check_max

            # Получаем элементы в этой области
            current_elements = get_elements_in_section(check_bbox, all_elements)

            # Проверяем, есть ли НОВЫЕ элементы
            new_elements = current_elements - prev_elements

            if new_elements:
                # Создаём разрез
                section_bbox = create_section_bbox(
                    origin=origin,
                    direction=east_dir,
                    right=north_dir.Negate(),
                    up=up_dir,
                    width=width_ns + 2,
                    height=height,
                    depth=section_depth
                )

                try:
                    section = ViewSection.CreateSection(doc, section_type.Id, section_bbox)
                    section.Name = "{}_T_{:03d}".format(tag, section_num)

                    # Устанавливаем высокую детализацию
                    section.DetailLevel = ViewDetailLevel.Fine

                    param = section.LookupParameter(PARAM_NAME)
                    if param:
                        param.Set(tag)

                    created_sections.append(section)
                    section_num += 1
                except Exception as e:
                    pass

                prev_elements = current_elements.copy()

            current_offset += step_feet

    return len(created_sections)


def generate_plans(step_feet, include_links, tag):
    """Генерировать планы с заданным шагом по высоте.

    Планы создаются только если на этой высоте есть НОВЫЕ элементы.
    """
    # Получаем все элементы
    all_elements = get_all_model_elements(include_links)

    if not all_elements:
        show_warning("Нет элементов", "В модели не найдено элементов для создания планов")
        return 0

    # Получаем BoundingBox здания
    building_bbox = get_building_bounding_box(all_elements)
    if building_bbox is None:
        show_warning("Ошибка", "Не удалось определить границы здания")
        return 0

    # Получаем тип плана
    plan_type = get_floor_plan_view_family_type()
    if plan_type is None:
        show_error("Ошибка", "Не найден тип семейства для планов этажей")
        return 0

    # Диапазон высот
    z_min = building_bbox.Min.Z
    z_max = building_bbox.Max.Z

    created_plans = []
    prev_elements = set()
    plan_num = 1
    current_z = z_min

    # Создаём или находим уровни для каждого плана
    while current_z <= z_max:
        # Получаем элементы в этом диапазоне высот
        current_elements = get_elements_in_z_range(current_z, current_z + step_feet, all_elements)

        # Проверяем, есть ли НОВЫЕ элементы
        new_elements = current_elements - prev_elements

        if new_elements:
            try:
                # Создаём временный уровень
                level_name = "{}_P_{:03d}".format(tag, plan_num)
                level = Level.Create(doc, current_z + step_feet / 2)
                level.Name = level_name

                # Создаём план на этом уровне
                plan = ViewPlan.Create(doc, plan_type.Id, level.Id)
                plan.Name = level_name

                # Устанавливаем высокую детализацию
                plan.DetailLevel = ViewDetailLevel.Fine

                # Устанавливаем параметр
                param = plan.LookupParameter(PARAM_NAME)
                if param:
                    param.Set(tag)

                created_plans.append(plan)
                plan_num += 1
            except Exception as e:
                pass

            prev_elements = current_elements.copy()

        current_z += step_feet

    return len(created_plans)


class GenerateSectionsForm(Form):
    """Форма генерации разрезов."""

    def __init__(self):
        self.setup_form()
        self.load_existing_generations()

    def setup_form(self):
        self.Text = "Генератор разрезов и планов"
        self.Width = 580
        self.Height = 720
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MinimumSize = Size(560, 700)

        # === Группа существующих генераций ===
        grp_existing = GroupBox()
        grp_existing.Text = "Существующие генерации"
        grp_existing.Location = Point(10, 10)
        grp_existing.Size = Size(540, 150)
        grp_existing.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        self.lst_generations = ListBox()
        self.lst_generations.Location = Point(10, 20)
        self.lst_generations.Size = Size(410, 90)
        self.lst_generations.SelectionMode = SelectionMode.MultiExtended
        self.lst_generations.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        btn_delete = Button()
        btn_delete.Text = "Удалить\nвыбранные"
        btn_delete.Location = Point(430, 20)
        btn_delete.Size = Size(100, 45)
        btn_delete.Click += self.on_delete

        btn_refresh = Button()
        btn_refresh.Text = "Обновить"
        btn_refresh.Location = Point(430, 70)
        btn_refresh.Size = Size(100, 25)
        btn_refresh.Click += self.on_refresh

        self.lbl_total = Label()
        self.lbl_total.Text = "Всего: 0 видов"
        self.lbl_total.Location = Point(10, 115)
        self.lbl_total.Width = 200

        grp_existing.Controls.Add(self.lst_generations)
        grp_existing.Controls.Add(btn_delete)
        grp_existing.Controls.Add(btn_refresh)
        grp_existing.Controls.Add(self.lbl_total)

        # === Группа разрезов ===
        grp_sections = GroupBox()
        grp_sections.Text = "Разрезы"
        grp_sections.Location = Point(10, 165)
        grp_sections.Size = Size(540, 200)
        grp_sections.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        # Галочка генерации разрезов
        self.chk_sections = CheckBox()
        self.chk_sections.Text = "Генерировать разрезы"
        self.chk_sections.Location = Point(10, 20)
        self.chk_sections.Width = 180
        self.chk_sections.Checked = True
        self.chk_sections.CheckedChanged += self.on_sections_check_changed

        # Направление
        lbl_direction = Label()
        lbl_direction.Text = "Направление:"
        lbl_direction.Location = Point(10, 50)
        lbl_direction.Width = 90

        self.rb_longitudinal = RadioButton()
        self.rb_longitudinal.Text = "Продольное (С-Ю)"
        self.rb_longitudinal.Location = Point(105, 48)
        self.rb_longitudinal.Width = 135

        self.rb_transverse = RadioButton()
        self.rb_transverse.Text = "Поперечное (В-З)"
        self.rb_transverse.Location = Point(245, 48)
        self.rb_transverse.Width = 135

        self.rb_both = RadioButton()
        self.rb_both.Text = "Оба"
        self.rb_both.Location = Point(385, 48)
        self.rb_both.Width = 60
        self.rb_both.Checked = True

        # Шаг разрезов
        lbl_step = Label()
        lbl_step.Text = "Шаг (м):"
        lbl_step.Location = Point(10, 80)
        lbl_step.Width = 90

        self.txt_step = TextBox()
        self.txt_step.Text = "0.5"
        self.txt_step.Location = Point(105, 77)
        self.txt_step.Width = 60

        # Глубина разреза
        lbl_depth = Label()
        lbl_depth.Text = "Глубина (м):"
        lbl_depth.Location = Point(180, 80)
        lbl_depth.Width = 80

        self.txt_depth = TextBox()
        self.txt_depth.Text = "1.0"
        self.txt_depth.Location = Point(265, 77)
        self.txt_depth.Width = 60

        self.chk_full_depth = CheckBox()
        self.chk_full_depth.Text = "На всю длину/ширину здания"
        self.chk_full_depth.Location = Point(340, 77)
        self.chk_full_depth.Width = 190
        self.chk_full_depth.CheckedChanged += self.on_full_depth_changed

        # Связанные файлы
        self.chk_links = CheckBox()
        self.chk_links.Text = "Учитывать элементы в связанных файлах"
        self.chk_links.Location = Point(10, 110)
        self.chk_links.Width = 300
        self.chk_links.Checked = False

        # 3D виды с углов (4 галочки)
        lbl_3d_views = Label()
        lbl_3d_views.Text = "3D виды:"
        lbl_3d_views.Location = Point(10, 140)
        lbl_3d_views.Width = 90

        self.chk_3d_ne = CheckBox()
        self.chk_3d_ne.Text = "СВ"
        self.chk_3d_ne.Location = Point(105, 138)
        self.chk_3d_ne.Width = 50

        self.chk_3d_nw = CheckBox()
        self.chk_3d_nw.Text = "СЗ"
        self.chk_3d_nw.Location = Point(160, 138)
        self.chk_3d_nw.Width = 50

        self.chk_3d_se = CheckBox()
        self.chk_3d_se.Text = "ЮВ"
        self.chk_3d_se.Location = Point(215, 138)
        self.chk_3d_se.Width = 50

        self.chk_3d_sw = CheckBox()
        self.chk_3d_sw.Text = "ЮЗ"
        self.chk_3d_sw.Location = Point(270, 138)
        self.chk_3d_sw.Width = 50

        # Кнопка генерации разрезов
        self.btn_generate_sections = Button()
        self.btn_generate_sections.Text = "Сгенерировать разрезы"
        self.btn_generate_sections.Location = Point(10, 165)
        self.btn_generate_sections.Size = Size(180, 28)
        self.btn_generate_sections.Click += self.on_generate_sections

        grp_sections.Controls.Add(self.chk_sections)
        grp_sections.Controls.Add(lbl_direction)
        grp_sections.Controls.Add(self.rb_longitudinal)
        grp_sections.Controls.Add(self.rb_transverse)
        grp_sections.Controls.Add(self.rb_both)
        grp_sections.Controls.Add(lbl_step)
        grp_sections.Controls.Add(self.txt_step)
        grp_sections.Controls.Add(lbl_depth)
        grp_sections.Controls.Add(self.txt_depth)
        grp_sections.Controls.Add(self.chk_full_depth)
        grp_sections.Controls.Add(self.chk_links)
        grp_sections.Controls.Add(lbl_3d_views)
        grp_sections.Controls.Add(self.chk_3d_ne)
        grp_sections.Controls.Add(self.chk_3d_nw)
        grp_sections.Controls.Add(self.chk_3d_se)
        grp_sections.Controls.Add(self.chk_3d_sw)
        grp_sections.Controls.Add(self.btn_generate_sections)

        # === Группа планов ===
        grp_plans = GroupBox()
        grp_plans.Text = "Планы"
        grp_plans.Location = Point(10, 370)
        grp_plans.Size = Size(540, 120)
        grp_plans.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        # Галочка генерации планов
        self.chk_plans = CheckBox()
        self.chk_plans.Text = "Генерировать планы"
        self.chk_plans.Location = Point(10, 20)
        self.chk_plans.Width = 180
        self.chk_plans.CheckedChanged += self.on_plans_check_changed

        # Шаг планов
        lbl_plan_step = Label()
        lbl_plan_step.Text = "Шаг по высоте (м):"
        lbl_plan_step.Location = Point(10, 50)
        lbl_plan_step.Width = 120

        self.txt_plan_step = TextBox()
        self.txt_plan_step.Text = "0.3"
        self.txt_plan_step.Location = Point(135, 47)
        self.txt_plan_step.Width = 60
        self.txt_plan_step.Enabled = False

        lbl_plan_note = Label()
        lbl_plan_note.Text = "(создаются только при появлении новых элементов)"
        lbl_plan_note.Location = Point(205, 50)
        lbl_plan_note.Width = 320
        lbl_plan_note.ForeColor = Color.Gray

        # Кнопка генерации планов
        self.btn_generate_plans = Button()
        self.btn_generate_plans.Text = "Сгенерировать планы"
        self.btn_generate_plans.Location = Point(10, 80)
        self.btn_generate_plans.Size = Size(180, 28)
        self.btn_generate_plans.Click += self.on_generate_plans
        self.btn_generate_plans.Enabled = False

        grp_plans.Controls.Add(self.chk_plans)
        grp_plans.Controls.Add(lbl_plan_step)
        grp_plans.Controls.Add(self.txt_plan_step)
        grp_plans.Controls.Add(lbl_plan_note)
        grp_plans.Controls.Add(self.btn_generate_plans)

        # === Группа параметра ===
        grp_param = GroupBox()
        grp_param.Text = "Параметр автогенерации"
        grp_param.Location = Point(10, 495)
        grp_param.Size = Size(540, 80)
        grp_param.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        # Информация о параметре
        self.param_exists = check_cpsk_auto_param()
        param_status = "найден" if self.param_exists else "НЕ НАЙДЕН"
        param_color = Color.Green if self.param_exists else Color.Red

        self.lbl_param = Label()
        self.lbl_param.Text = "Параметр {}: {}".format(PARAM_NAME, param_status)
        self.lbl_param.Location = Point(10, 25)
        self.lbl_param.Width = 280
        self.lbl_param.ForeColor = param_color

        # Кнопка создания параметра
        self.btn_create_param = Button()
        self.btn_create_param.Text = "Создать параметр"
        self.btn_create_param.Location = Point(10, 48)
        self.btn_create_param.Size = Size(140, 25)
        self.btn_create_param.Click += self.on_create_param
        self.btn_create_param.Visible = not self.param_exists

        grp_param.Controls.Add(self.lbl_param)
        grp_param.Controls.Add(self.btn_create_param)

        # === Статус и кнопка закрыть ===
        self.lbl_status = Label()
        self.lbl_status.Text = "Готов к работе"
        self.lbl_status.Location = Point(10, 585)
        self.lbl_status.Size = Size(400, 40)
        self.lbl_status.ForeColor = Color.Gray
        self.lbl_status.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(460, 640)
        btn_close.Size = Size(100, 30)
        btn_close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        btn_close.Click += self.on_close

        self.Controls.Add(grp_existing)
        self.Controls.Add(grp_sections)
        self.Controls.Add(grp_plans)
        self.Controls.Add(grp_param)
        self.Controls.Add(self.lbl_status)
        self.Controls.Add(btn_close)

        # Устанавливаем начальное состояние кнопок
        self.update_buttons_state()

    def update_buttons_state(self):
        """Обновить состояние кнопок в зависимости от наличия параметра."""
        self.btn_generate_sections.Enabled = self.param_exists and self.chk_sections.Checked
        self.btn_generate_plans.Enabled = self.param_exists and self.chk_plans.Checked

    def load_existing_generations(self):
        """Загрузить существующие генерации."""
        self.lst_generations.Items.Clear()
        self.generations = get_existing_generations()

        total = 0
        for tag, views in sorted(self.generations.items(), reverse=True):
            count = len(views)
            total += count
            self.lst_generations.Items.Add("{} ({} видов)".format(tag, count))

        self.lbl_total.Text = "Всего: {} видов".format(total)

    def on_delete(self, sender, args):
        """Удалить выбранные генерации."""
        if self.lst_generations.SelectedItems.Count == 0:
            show_warning("Выбор", "Выберите генерации для удаления")
            return

        tags_to_delete = []
        for item in self.lst_generations.SelectedItems:
            tag = item.split(" (")[0]
            tags_to_delete.append(tag)

        total_deleted = 0
        with revit.Transaction("CPSK: Удалить виды"):
            for tag in tags_to_delete:
                deleted = delete_views_by_tag(tag)
                total_deleted += deleted

        show_success("Удалено", "Удалено {} видов".format(total_deleted))
        self.load_existing_generations()

    def on_refresh(self, sender, args):
        """Обновить список."""
        self.load_existing_generations()

    def on_create_param(self, sender, args):
        """Создать параметр CPSK_AUTO."""
        with revit.Transaction("CPSK: Создать параметр CPSK_AUTO"):
            success, message = create_cpsk_auto_param()

        if success:
            show_success("Параметр создан", message)
            # Обновляем UI
            self.param_exists = True
            self.lbl_param.Text = "Параметр {}: найден".format(PARAM_NAME)
            self.lbl_param.ForeColor = Color.Green
            self.btn_create_param.Visible = False
            self.update_buttons_state()
        else:
            show_error("Ошибка", message)

    def on_sections_check_changed(self, sender, args):
        """Обработчик изменения галочки разрезов."""
        self.update_buttons_state()

    def on_plans_check_changed(self, sender, args):
        """Обработчик изменения галочки планов."""
        self.txt_plan_step.Enabled = self.chk_plans.Checked
        self.update_buttons_state()

    def on_full_depth_changed(self, sender, args):
        """Обработчик изменения галочки полной глубины."""
        self.txt_depth.Enabled = not self.chk_full_depth.Checked

    def on_generate_sections(self, sender, args):
        """Сгенерировать разрезы."""
        # Определяем направление
        if self.rb_longitudinal.Checked:
            direction = "longitudinal"
        elif self.rb_transverse.Checked:
            direction = "transverse"
        else:
            direction = "both"

        # Шаг
        try:
            step_m = float(self.txt_step.Text.replace(",", "."))
            step_feet = step_m / 0.3048
        except:
            show_error("Ошибка", "Некорректное значение шага")
            return

        # Глубина
        full_depth = self.chk_full_depth.Checked
        depth_feet = None
        if not full_depth:
            try:
                depth_m = float(self.txt_depth.Text.replace(",", "."))
                depth_feet = depth_m / 0.3048
            except:
                show_error("Ошибка", "Некорректное значение глубины")
                return

        include_links = self.chk_links.Checked

        # Генерируем уникальный тег
        tag = "SEC_{}".format(datetime.now().strftime("%Y%m%d_%H%M%S"))

        self.lbl_status.Text = "Генерация разрезов..."
        self.lbl_status.ForeColor = Color.Blue
        self.Refresh()

        total_count = 0

        with revit.Transaction("CPSK: Генерация разрезов"):
            # Генерируем разрезы
            count = generate_sections(direction, step_feet, include_links, tag, depth_feet, full_depth)
            total_count += count

            # Генерируем 3D виды
            views_3d_count = self.generate_3d_views(tag, include_links)
            total_count += views_3d_count

        if total_count > 0:
            msg = "Создано {} видов".format(total_count)
            if views_3d_count > 0:
                msg = "Создано {} разрезов, {} 3D видов".format(count, views_3d_count)
            show_success("Готово", msg, details="Тег: {}".format(tag))
            self.lbl_status.Text = msg
            self.lbl_status.ForeColor = Color.Green
        else:
            self.lbl_status.Text = "Разрезы не созданы"
            self.lbl_status.ForeColor = Color.Red

        self.load_existing_generations()

    def generate_3d_views(self, tag, include_links):
        """Генерировать 3D аксонометрические виды с углов здания."""
        views_to_create = []
        if self.chk_3d_ne.Checked:
            views_to_create.append(("NE", "СВ"))
        if self.chk_3d_nw.Checked:
            views_to_create.append(("NW", "СЗ"))
        if self.chk_3d_se.Checked:
            views_to_create.append(("SE", "ЮВ"))
        if self.chk_3d_sw.Checked:
            views_to_create.append(("SW", "ЮЗ"))

        if not views_to_create:
            return 0

        # Получаем элементы и BoundingBox
        all_elements = get_all_model_elements(include_links)
        if not all_elements:
            show_error("3D виды", "Нет элементов для создания 3D видов")
            return 0

        building_bbox = get_building_bounding_box(all_elements)
        if building_bbox is None:
            show_error("3D виды", "Не удалось определить границы здания")
            return 0

        # Получаем тип 3D вида
        view3d_type = get_3d_view_family_type()
        if view3d_type is None:
            show_error("Ошибка", "Не найден тип семейства для 3D видов")
            return 0

        north_angle = get_project_north_angle()

        # Центр здания
        center = XYZ(
            (building_bbox.Min.X + building_bbox.Max.X) / 2,
            (building_bbox.Min.Y + building_bbox.Max.Y) / 2,
            (building_bbox.Min.Z + building_bbox.Max.Z) / 2
        )

        created = 0
        errors = []

        for view_code, view_name in views_to_create:
            try:
                # Определяем горизонтальный угол для каждого угла
                # (относительно проектного севера)
                if view_code == "NE":
                    horizontal_angle = north_angle + math.radians(45)
                elif view_code == "NW":
                    horizontal_angle = north_angle + math.radians(315)
                elif view_code == "SE":
                    horizontal_angle = north_angle + math.radians(135)
                else:  # SW
                    horizontal_angle = north_angle + math.radians(225)

                # Горизонтальное направление от камеры к зданию
                dir_x = math.sin(horizontal_angle)
                dir_y = math.cos(horizontal_angle)

                # Позиция глаза - сзади-сверху от здания
                # Камера на расстоянии и выше центра
                eye_distance = 50  # футов по горизонтали
                eye_height = 30    # футов выше центра
                eye = XYZ(
                    center.X - dir_x * eye_distance,
                    center.Y - dir_y * eye_distance,
                    center.Z + eye_height
                )

                # Направление взгляда - от глаза к центру здания
                forward = (center - eye).Normalize()

                # Вычисляем правильный up-вектор, перпендикулярный forward
                # Сначала вычисляем right-вектор (перпендикулярен forward и мировому Z)
                world_up = XYZ.BasisZ
                right = forward.CrossProduct(world_up).Normalize()

                # Теперь up-вектор перпендикулярен и forward и right
                up = right.CrossProduct(forward).Normalize()

                # Создаём ориентацию 3D вида
                orientation = ViewOrientation3D(eye, up, forward)

                # Создаём 3D вид
                view3d = View3D.CreateIsometric(doc, view3d_type.Id)
                view3d.SetOrientation(orientation)
                view3d.Name = "{}_3D_{}".format(tag, view_code)
                view3d.DetailLevel = ViewDetailLevel.Fine

                # Устанавливаем параметр
                param = view3d.LookupParameter(PARAM_NAME)
                if param:
                    param.Set(tag)

                created += 1
            except Exception as e:
                errors.append("{}: {}".format(view_code, str(e)))

        if errors:
            show_error("Ошибки 3D видов", "Не удалось создать виды", details="\n".join(errors))

        return created

    def on_generate_plans(self, sender, args):
        """Сгенерировать планы."""
        # Шаг
        try:
            step_m = float(self.txt_plan_step.Text.replace(",", "."))
            step_feet = step_m / 0.3048
        except:
            show_error("Ошибка", "Некорректное значение шага")
            return

        include_links = self.chk_links.Checked

        # Генерируем уникальный тег
        tag = "PLN_{}".format(datetime.now().strftime("%Y%m%d_%H%M%S"))

        self.lbl_status.Text = "Генерация планов..."
        self.lbl_status.ForeColor = Color.Blue
        self.Refresh()

        with revit.Transaction("CPSK: Генерация планов"):
            count = generate_plans(step_feet, include_links, tag)

        if count > 0:
            show_success("Готово", "Создано {} планов".format(count), details="Тег: {}".format(tag))
            self.lbl_status.Text = "Создано {} планов с тегом {}".format(count, tag)
            self.lbl_status.ForeColor = Color.Green
        else:
            self.lbl_status.Text = "Планы не созданы"
            self.lbl_status.ForeColor = Color.Red

        self.load_existing_generations()

    def on_close(self, sender, args):
        """Закрыть форму."""
        self.Close()


def main():
    """Основная функция."""
    # Проверяем параметр
    if not check_cpsk_auto_param():
        show_warning(
            "Параметр не найден",
            "Параметр {} не найден на видах.\nДобавьте shared параметр к категории 'Виды' перед использованием.".format(PARAM_NAME)
        )

    form = GenerateSectionsForm()
    form.ShowDialog()


if __name__ == "__main__":
    main()
