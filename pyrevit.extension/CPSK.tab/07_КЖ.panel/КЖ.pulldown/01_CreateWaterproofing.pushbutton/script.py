# -*- coding: utf-8 -*-
"""
Создание гидроизоляции на выбранных поверхностях ЖБ конструкций.
- Вертикальные грани → Стена (Wall)
- Горизонтальные грани → Крыша (Roof)
"""

__title__ = "Создать\nгидроизоляцию"
__author__ = "CPSK"

import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, Panel, ComboBox, TextBox, CheckBox,
    ListBox, CheckedListBox, NumericUpDown, RadioButton, GroupBox,
    DockStyle, FormStartPosition, FormBorderStyle, SelectionMode,
    DialogResult, MessageBoxButtons, MessageBoxIcon, Padding,
    FormWindowState, Application
)
from System.Drawing import Point, Size, Font, FontStyle

from pyrevit import revit

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    Transaction, ElementId, XYZ, Line, Reference, Face, UV,
    BuiltInParameter, WallType, WallKind, Wall, Level,
    RoofType, FootPrintRoof, ModelCurveArray, CurveArray,
    PlanarFace, Options, Outline, BoundingBoxIntersectsFilter,
    Solid, GeometryInstance, ViewDetailLevel, Plane,
    BooleanOperationsType, BooleanOperationsUtils
)
from Autodesk.Revit.DB.Structure import StructuralType
from System.Collections.Generic import List
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

import clr

# Добавляем lib в путь для импорта
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Импорт модулей из lib
from cpsk_notify import show_error, show_info, show_success, show_warning
from cpsk_auth import require_auth

# Проверка авторизации
if not require_auth():
    sys.exit()

# Настройки
doc = revit.doc
uidoc = revit.uidoc


class FaceSelectionFilter(ISelectionFilter):
    """Фильтр для выбора только граней."""

    def AllowElement(self, elem):
        return True

    def AllowReference(self, reference, position):
        try:
            geom_obj = doc.GetElement(reference).GetGeometryObjectFromReference(reference)
            return isinstance(geom_obj, Face)
        except Exception as e:
            show_warning("Фильтр", "Ошибка проверки грани", details=str(e))
            return False


def mm_to_feet(mm):
    """Перевод мм в футы."""
    return mm / 304.8


def feet_to_mm(feet):
    """Перевод футов в мм."""
    return feet * 304.8


def calc_polygon_area(points):
    """
    Вычислить площадь полигона по формуле Shoelace (площадь Гаусса).
    Работает для плоских полигонов (проецирует на XY или XZ в зависимости от ориентации).

    Args:
        points: список XYZ точек

    Returns:
        площадь в квадратных футах
    """
    if len(points) < 3:
        return 0.0

    # Определяем ориентацию по первым 3 точкам
    # Если Z почти одинаковый - горизонтальный полигон (проекция на XY)
    # Иначе - вертикальный (проекция на плоскость наибольшего размаха)
    z_vals = [p.Z for p in points]
    z_range = max(z_vals) - min(z_vals)

    x_vals = [p.X for p in points]
    y_vals = [p.Y for p in points]
    x_range = max(x_vals) - min(x_vals)
    y_range = max(y_vals) - min(y_vals)

    # Выбираем две координаты с наибольшим размахом для проекции
    if z_range < 0.01:
        # Горизонтальный полигон - проекция на XY
        coords = [(p.X, p.Y) for p in points]
    elif x_range < y_range:
        # Вертикальный, вытянут по Y - проекция на YZ
        coords = [(p.Y, p.Z) for p in points]
    else:
        # Вертикальный, вытянут по X - проекция на XZ
        coords = [(p.X, p.Z) for p in points]

    # Формула Shoelace
    n = len(coords)
    area = 0.0
    for i in range(n):
        j = (i + 1) % n
        area += coords[i][0] * coords[j][1]
        area -= coords[j][0] * coords[i][1]

    return abs(area) / 2.0


def sq_feet_to_sq_m(sq_feet):
    """Перевод квадратных футов в квадратные метры."""
    return sq_feet * 0.092903


def get_wall_types():
    """Получить все типы стен (базовые)."""
    collector = FilteredElementCollector(doc).OfClass(WallType)
    wall_types = []

    for wt in collector:
        try:
            if wt.Kind == WallKind.Basic:
                wall_types.append(wt)
        except Exception as e:
            show_warning("Внимание", "Ошибка обработки типа стены", details=str(e))

    wall_types.sort(key=lambda x: x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "")
    return wall_types


def get_roof_types():
    """Получить все типы крыш."""
    collector = FilteredElementCollector(doc).OfClass(RoofType)
    roof_types = list(collector)

    roof_types.sort(key=lambda x: x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() if x.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM) else "")
    return roof_types


def get_levels():
    """Получить все уровни, отсортированные по высоте."""
    collector = FilteredElementCollector(doc).OfClass(Level)
    levels = list(collector)
    levels.sort(key=lambda x: x.Elevation)
    return levels


def get_nearest_level(elevation):
    """Получить ближайший уровень к заданной отметке."""
    levels = get_levels()
    if not levels:
        return None

    nearest = levels[0]
    min_diff = abs(elevation - nearest.Elevation)

    for level in levels:
        diff = abs(elevation - level.Elevation)
        if diff < min_diff:
            min_diff = diff
            nearest = level

    return nearest


def get_level_below(elevation):
    """Получить уровень ниже заданной отметки."""
    levels = get_levels()
    if not levels:
        return None

    level_below = None
    for level in levels:
        if level.Elevation <= elevation:
            level_below = level
        else:
            break

    return level_below if level_below else levels[0]


def get_all_categories_for_hiding():
    """Получить категории элементов для временного скрытия."""
    categories = []

    category_ids = [
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_Floors,
        BuiltInCategory.OST_StructuralFoundation,
        BuiltInCategory.OST_Roofs,
        BuiltInCategory.OST_Ceilings,
        BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_Mass,
    ]

    for cat_id in category_ids:
        try:
            cat = doc.Settings.Categories.get_Item(cat_id)
            if cat and cat.AllowsVisibilityControl:
                categories.append((cat.Name, cat.Id))
        except Exception as e:
            show_warning("Внимание", "Ошибка получения категории", details=str(e))

    categories.sort(key=lambda x: x[0])
    return categories


def get_element_solid(element):
    """Получить Solid геометрию элемента."""
    solids = []
    try:
        opt = Options()
        opt.ComputeReferences = True
        opt.DetailLevel = ViewDetailLevel.Fine

        geom = element.get_Geometry(opt)
        if geom is None:
            return solids

        for geom_obj in geom:
            if isinstance(geom_obj, Solid):
                if geom_obj.Volume > 0:
                    solids.append(geom_obj)
            elif isinstance(geom_obj, GeometryInstance):
                inst_geom = geom_obj.GetInstanceGeometry()
                if inst_geom:
                    for inst_obj in inst_geom:
                        if isinstance(inst_obj, Solid):
                            if inst_obj.Volume > 0:
                                solids.append(inst_obj)
    except Exception as geom_err:
        # Ожидаемо для элементов без геометрии
        show_warning("Геометрия", "Не удалось получить геометрию элемента", details=str(geom_err), blocking=False, auto_close=1)
    return solids


def get_solid_contour_at_z(solid, z_level, tolerance=0.5):
    """
    Получить контур пересечения Solid с горизонтальной плоскостью на уровне z_level.
    Ищет ближайшую горизонтальную грань к z_level.

    Returns:
        list of XYZ points или None
    """
    try:
        # Собираем ВСЕ горизонтальные грани с их Z-уровнями
        # НЕ используем GetBoundingBox() - он может вернуть локальные координаты!
        horizontal_faces = []

        for face in solid.Faces:
            if not isinstance(face, PlanarFace):
                continue

            normal = face.FaceNormal
            # Горизонтальная грань (нормаль вверх или вниз)
            if abs(normal.Z) < 0.7:
                continue

            edge_loops = face.EdgeLoops
            if edge_loops.Size == 0:
                continue

            outer_loop = edge_loops.get_Item(0)
            points = []
            face_z = None

            for edge in outer_loop:
                pt = edge.AsCurve().GetEndPoint(0)
                points.append(pt)
                if face_z is None:
                    face_z = pt.Z

            if face_z is None or len(points) < 3:
                continue

            # Сохраняем грань с расстоянием до целевого уровня
            dist = abs(face_z - z_level)
            horizontal_faces.append({
                'points': points,
                'z': face_z,
                'dist': dist
            })

        if not horizontal_faces:
            return None

        # Сортируем по расстоянию до целевого Z
        horizontal_faces.sort(key=lambda f: f['dist'])

        # Берём ближайшую грань, если она в пределах допуска
        best = horizontal_faces[0]
        if best['dist'] <= tolerance:
            return best['points']

        return None

    except Exception as contour_err:
        show_warning("Контур", "Ошибка получения контура", details=str(contour_err), blocking=False, auto_close=1)
        return None


def find_intersecting_elements_contours(face_ref, source_element_id, face_z, tolerance=0.5):
    """
    Найти все элементы, пересекающиеся с гранью, и получить их контуры на уровне грани.

    Args:
        face_ref: Reference на грань
        source_element_id: ID исходного элемента (чтобы исключить его)
        face_z: Z-координата грани
        tolerance: допуск для поиска пересечений

    Returns:
        list of (element_name, contour_points)
    """
    results = []

    try:
        # Получаем грань и её bounding box
        elem = doc.GetElement(face_ref.ElementId)
        face = elem.GetGeometryObjectFromReference(face_ref)

        if not isinstance(face, PlanarFace):
            return results

        # Получаем bounding box грани
        edge_loops = face.EdgeLoops
        if edge_loops.Size == 0:
            return results

        outer_loop = edge_loops.get_Item(0)
        points = []
        for edge in outer_loop:
            points.append(edge.AsCurve().GetEndPoint(0))

        if len(points) < 3:
            return results

        x_vals = [p.X for p in points]
        y_vals = [p.Y for p in points]

        min_pt = XYZ(min(x_vals) - tolerance, min(y_vals) - tolerance, face_z - tolerance)
        max_pt = XYZ(max(x_vals) + tolerance, max(y_vals) + tolerance, face_z + tolerance)

        # Создаём фильтр по bounding box
        outline = Outline(min_pt, max_pt)
        bb_filter = BoundingBoxIntersectsFilter(outline)

        # Категории для поиска пересечений
        categories_to_check = [
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_Columns,
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_StructuralFoundation,
            BuiltInCategory.OST_GenericModel,
        ]

        for cat in categories_to_check:
            try:
                collector = FilteredElementCollector(doc).OfCategory(cat).WherePasses(bb_filter)
                for intersecting_elem in collector:
                    # Пропускаем исходный элемент
                    if intersecting_elem.Id == source_element_id:
                        continue

                    # Получаем solid геометрию
                    solids = get_element_solid(intersecting_elem)
                    for solid in solids:
                        contour = get_solid_contour_at_z(solid, face_z, tolerance)
                        if contour and len(contour) >= 3:
                            elem_name = intersecting_elem.Name if hasattr(intersecting_elem, 'Name') else "Элемент"
                            results.append((elem_name, contour))
            except Exception as cat_err:
                show_warning("Поиск", "Ошибка поиска в категории", details=str(cat_err), blocking=False, auto_close=1)
                continue

    except Exception as search_err:
        show_warning("Поиск", "Ошибка поиска пересечений", details=str(search_err), blocking=False, auto_close=1)

    return results


def get_solid_rect_on_vertical_plane(solid, plane_origin, plane_normal, tolerance=0.5):
    """
    Получить прямоугольник пересечения Solid с вертикальной плоскостью.

    Args:
        solid: Solid геометрия элемента
        plane_origin: XYZ - точка на плоскости грани
        plane_normal: XYZ - нормаль плоскости (в горизонтальной проекции)
        tolerance: допуск для определения близости к плоскости

    Returns:
        dict с ключами: min_z, max_z, min_along, max_along, center_along
        или None если пересечение не найдено
    """
    try:
        # Нормализуем вектор нормали (только X,Y)
        normal_2d = XYZ(plane_normal.X, plane_normal.Y, 0)
        if normal_2d.GetLength() < 0.001:
            return None
        normal_2d = normal_2d.Normalize()

        # Вектор вдоль стены (перпендикулярен нормали)
        along_wall = XYZ(-normal_2d.Y, normal_2d.X, 0)

        # Собираем все точки solid, близкие к плоскости
        points_on_plane = []

        for face in solid.Faces:
            if not isinstance(face, PlanarFace):
                continue

            edge_loops = face.EdgeLoops
            for loop_idx in range(edge_loops.Size):
                loop = edge_loops.get_Item(loop_idx)
                for edge in loop:
                    pt = edge.AsCurve().GetEndPoint(0)

                    # Проверяем расстояние до плоскости
                    vec_to_pt = XYZ(pt.X - plane_origin.X, pt.Y - plane_origin.Y, 0)
                    dist = abs(vec_to_pt.X * normal_2d.X + vec_to_pt.Y * normal_2d.Y)

                    if dist <= tolerance:
                        points_on_plane.append(pt)

        if len(points_on_plane) < 2:
            return None

        # Вычисляем координаты вдоль стены и Z
        z_values = [p.Z for p in points_on_plane]
        along_values = []
        for p in points_on_plane:
            vec = XYZ(p.X - plane_origin.X, p.Y - plane_origin.Y, 0)
            along = vec.X * along_wall.X + vec.Y * along_wall.Y
            along_values.append(along)

        return {
            'min_z': min(z_values),
            'max_z': max(z_values),
            'min_along': min(along_values),
            'max_along': max(along_values),
            'center_along': (min(along_values) + max(along_values)) / 2
        }

    except Exception as e:
        return None


def find_intersecting_elements_for_vertical_face(face_ref, source_element_id, tolerance=0.5):
    """
    Найти все элементы, пересекающиеся с вертикальной гранью.

    Args:
        face_ref: Reference на грань
        source_element_id: ID исходного элемента (чтобы исключить его)
        tolerance: допуск для поиска пересечений

    Returns:
        list of dict с ключами: name, min_z, max_z, min_along, max_along
    """
    results = []

    try:
        elem = doc.GetElement(face_ref.ElementId)
        face = elem.GetGeometryObjectFromReference(face_ref)

        if not isinstance(face, PlanarFace):
            return results

        normal = face.FaceNormal

        # Получаем точки грани для bounding box
        edge_loops = face.EdgeLoops
        if edge_loops.Size == 0:
            return results

        outer_loop = edge_loops.get_Item(0)
        points = []
        for edge in outer_loop:
            points.append(edge.AsCurve().GetEndPoint(0))

        if len(points) < 3:
            return results

        # Центр грани как origin плоскости
        center_x = sum(p.X for p in points) / len(points)
        center_y = sum(p.Y for p in points) / len(points)
        center_z = sum(p.Z for p in points) / len(points)
        plane_origin = XYZ(center_x, center_y, center_z)

        # Bounding box грани
        x_vals = [p.X for p in points]
        y_vals = [p.Y for p in points]
        z_vals = [p.Z for p in points]

        min_pt = XYZ(min(x_vals) - tolerance, min(y_vals) - tolerance, min(z_vals) - tolerance)
        max_pt = XYZ(max(x_vals) + tolerance, max(y_vals) + tolerance, max(z_vals) + tolerance)

        outline = Outline(min_pt, max_pt)
        bb_filter = BoundingBoxIntersectsFilter(outline)

        # Категории для поиска
        categories_to_check = [
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_Columns,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_StructuralFoundation,
            BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_Walls,
        ]

        for cat in categories_to_check:
            try:
                collector = FilteredElementCollector(doc).OfCategory(cat).WherePasses(bb_filter)
                for intersecting_elem in collector:
                    if intersecting_elem.Id == source_element_id:
                        continue

                    solids = get_element_solid(intersecting_elem)
                    for solid in solids:
                        rect = get_solid_rect_on_vertical_plane(solid, plane_origin, normal, tolerance)
                        if rect:
                            elem_name = intersecting_elem.Name if hasattr(intersecting_elem, 'Name') else "Элемент"
                            rect['name'] = elem_name
                            rect['plane_origin'] = plane_origin
                            rect['normal'] = normal
                            results.append(rect)
            except Exception as cat_err:
                # Ожидаемо для категорий без элементов
                show_warning("Поиск", "Пропуск категории", details=str(cat_err), blocking=False, auto_close=1)
                continue

    except Exception as search_err:
        # Ошибка поиска не критична
        show_warning("Поиск", "Ошибка поиска элементов", details=str(search_err), blocking=False, auto_close=1)

    return results


class WaterproofingForm(Form):
    """Диалог создания гидроизоляции."""

    def __init__(self, wall_types, roof_types, all_categories):
        self.selected_faces = []
        self.face_info_list = []
        self.selected_wall_type = None
        self.selected_roof_type = None
        self.categories_to_hide = []

        self.saved_wall_type_index = 0
        self.saved_roof_type_index = 0

        self.wall_types = wall_types
        self.roof_types = roof_types
        self.all_categories = all_categories

        self.setup_form()

    def restore_state(self):
        """Восстановить состояние формы после выбора граней."""
        if self.saved_wall_type_index >= 0 and self.saved_wall_type_index < self.cmb_wall_type.Items.Count:
            self.cmb_wall_type.SelectedIndex = self.saved_wall_type_index
        if self.saved_roof_type_index >= 0 and self.saved_roof_type_index < self.cmb_roof_type.Items.Count:
            self.cmb_roof_type.SelectedIndex = self.saved_roof_type_index

        self.lst_faces.Items.Clear()
        for face_info in self.face_info_list:
            self.lst_faces.Items.Add(face_info)

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Создание гидроизоляции - CPSK"
        self.Width = 550
        self.Height = 620
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # === ТИП СТЕНЫ (для вертикальных граней) ===
        lbl_wall_type = Label()
        lbl_wall_type.Text = "Тип стены (для вертикальных граней):"
        lbl_wall_type.Location = Point(15, y)
        lbl_wall_type.AutoSize = True
        lbl_wall_type.Font = Font(lbl_wall_type.Font, FontStyle.Bold)
        self.Controls.Add(lbl_wall_type)
        y += 25

        self.cmb_wall_type = ComboBox()
        self.cmb_wall_type.Location = Point(15, y)
        self.cmb_wall_type.Width = 500
        self.cmb_wall_type.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList

        for wt in self.wall_types:
            name_param = wt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            name = name_param.AsString() if name_param else "Без имени"
            self.cmb_wall_type.Items.Add(name)

        if self.cmb_wall_type.Items.Count > 0:
            self.cmb_wall_type.SelectedIndex = 0

        self.Controls.Add(self.cmb_wall_type)
        y += 35

        # === ТИП КРЫШИ (для горизонтальных граней) ===
        lbl_roof_type = Label()
        lbl_roof_type.Text = "Тип крыши (для горизонтальных граней):"
        lbl_roof_type.Location = Point(15, y)
        lbl_roof_type.AutoSize = True
        lbl_roof_type.Font = Font(lbl_roof_type.Font, FontStyle.Bold)
        self.Controls.Add(lbl_roof_type)
        y += 25

        self.cmb_roof_type = ComboBox()
        self.cmb_roof_type.Location = Point(15, y)
        self.cmb_roof_type.Width = 500
        self.cmb_roof_type.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList

        for rt in self.roof_types:
            name_param = rt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            name = name_param.AsString() if name_param else "Без имени"
            self.cmb_roof_type.Items.Add(name)

        if self.cmb_roof_type.Items.Count > 0:
            self.cmb_roof_type.SelectedIndex = 0

        self.Controls.Add(self.cmb_roof_type)
        y += 40

        # === ВЫБОР ПОВЕРХНОСТЕЙ ===
        lbl_faces = Label()
        lbl_faces.Text = "Выбранные поверхности:"
        lbl_faces.Location = Point(15, y)
        lbl_faces.AutoSize = True
        lbl_faces.Font = Font(lbl_faces.Font, FontStyle.Bold)
        self.Controls.Add(lbl_faces)
        y += 25

        self.lst_faces = ListBox()
        self.lst_faces.Location = Point(15, y)
        self.lst_faces.Size = Size(400, 120)
        self.lst_faces.SelectionMode = SelectionMode.MultiExtended
        self.Controls.Add(self.lst_faces)

        btn_add_faces = Button()
        btn_add_faces.Text = "Добавить"
        btn_add_faces.Location = Point(425, y)
        btn_add_faces.Width = 90
        btn_add_faces.Click += self.on_add_faces
        self.Controls.Add(btn_add_faces)

        btn_remove_faces = Button()
        btn_remove_faces.Text = "Удалить"
        btn_remove_faces.Location = Point(425, y + 35)
        btn_remove_faces.Width = 90
        btn_remove_faces.Click += self.on_remove_faces
        self.Controls.Add(btn_remove_faces)

        btn_clear_faces = Button()
        btn_clear_faces.Text = "Очистить"
        btn_clear_faces.Location = Point(425, y + 70)
        btn_clear_faces.Width = 90
        btn_clear_faces.Click += self.on_clear_faces
        self.Controls.Add(btn_clear_faces)

        y += 135

        # === СКРЫТИЕ КАТЕГОРИЙ ===
        lbl_hide = Label()
        lbl_hide.Text = "Временно скрыть категории (для просмотра):"
        lbl_hide.Location = Point(15, y)
        lbl_hide.AutoSize = True
        lbl_hide.Font = Font(lbl_hide.Font, FontStyle.Bold)
        self.Controls.Add(lbl_hide)
        y += 25

        self.chk_categories = CheckedListBox()
        self.chk_categories.Location = Point(15, y)
        self.chk_categories.Size = Size(500, 100)
        self.chk_categories.CheckOnClick = True

        for cat_name, cat_id in self.all_categories:
            self.chk_categories.Items.Add(cat_name)

        self.Controls.Add(self.chk_categories)
        y += 110

        btn_hide = Button()
        btn_hide.Text = "Скрыть выбранные"
        btn_hide.Location = Point(15, y)
        btn_hide.Width = 140
        btn_hide.Click += self.on_hide_categories
        self.Controls.Add(btn_hide)

        btn_show_all = Button()
        btn_show_all.Text = "Показать все"
        btn_show_all.Location = Point(165, y)
        btn_show_all.Width = 120
        btn_show_all.Click += self.on_show_all_categories
        self.Controls.Add(btn_show_all)

        y += 45

        # === КНОПКИ ===
        btn_create = Button()
        btn_create.Text = "Создать гидроизоляцию"
        btn_create.Location = Point(15, y)
        btn_create.Width = 180
        btn_create.Height = 35
        btn_create.Click += self.on_create
        self.Controls.Add(btn_create)

        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(435, y)
        btn_cancel.Width = 80
        btn_cancel.Height = 35
        btn_cancel.Click += self.on_cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

    def on_add_faces(self, sender, args):
        """Добавить поверхности."""
        self.saved_wall_type_index = self.cmb_wall_type.SelectedIndex
        self.saved_roof_type_index = self.cmb_roof_type.SelectedIndex
        self.DialogResult = DialogResult.Retry
        self.Close()

    def on_remove_faces(self, sender, args):
        """Удалить выбранные поверхности из списка."""
        indices = list(self.lst_faces.SelectedIndices)
        indices.sort(reverse=True)

        for idx in indices:
            self.selected_faces.pop(idx)
            self.face_info_list.pop(idx)
            self.lst_faces.Items.RemoveAt(idx)

    def on_clear_faces(self, sender, args):
        """Очистить список поверхностей."""
        self.selected_faces = []
        self.face_info_list = []
        self.lst_faces.Items.Clear()

    def on_hide_categories(self, sender, args):
        """Скрыть выбранные категории на активном виде."""
        if not self.chk_categories.CheckedItems.Count:
            show_warning("Внимание", "Выберите категории для скрытия")
            return

        view = doc.ActiveView
        if not view.CanCategoryBeHidden(ElementId(BuiltInCategory.OST_Walls)):
            show_error("Ошибка", "На текущем виде нельзя скрывать категории")
            return

        checked_names = [str(item) for item in self.chk_categories.CheckedItems]
        cat_ids_to_hide = []

        for cat_name, cat_id in self.all_categories:
            if cat_name in checked_names:
                cat_ids_to_hide.append(cat_id)

        if not cat_ids_to_hide:
            return

        try:
            with Transaction(doc, "Скрыть категории") as t:
                t.Start()
                for cat_id in cat_ids_to_hide:
                    if view.CanCategoryBeHidden(cat_id):
                        view.SetCategoryHidden(cat_id, True)
                t.Commit()

            show_success("Готово", "Скрыто категорий: {}".format(len(cat_ids_to_hide)))
        except Exception as e:
            show_error("Ошибка", "Не удалось скрыть категории", details=str(e))

    def on_show_all_categories(self, sender, args):
        """Показать все категории на активном виде."""
        view = doc.ActiveView

        try:
            with Transaction(doc, "Показать все категории") as t:
                t.Start()
                for cat_name, cat_id in self.all_categories:
                    try:
                        if view.CanCategoryBeHidden(cat_id):
                            view.SetCategoryHidden(cat_id, False)
                    except Exception as e:
                        show_warning("Внимание", "Не удалось показать категорию: {}".format(cat_name), details=str(e))
                t.Commit()

            show_success("Готово", "Все категории показаны")
        except Exception as e:
            show_error("Ошибка", "Не удалось показать категории", details=str(e))

    def on_create(self, sender, args):
        """Создать гидроизоляцию."""
        if self.cmb_wall_type.SelectedIndex < 0:
            show_warning("Внимание", "Выберите тип стены")
            return

        if self.cmb_roof_type.SelectedIndex < 0:
            show_warning("Внимание", "Выберите тип крыши")
            return

        if not self.selected_faces:
            show_warning("Внимание", "Добавьте хотя бы одну поверхность")
            return

        self.selected_wall_type = self.wall_types[self.cmb_wall_type.SelectedIndex]
        self.selected_roof_type = self.roof_types[self.cmb_roof_type.SelectedIndex]

        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        """Отмена."""
        self.DialogResult = DialogResult.Cancel
        self.Close()


def highlight_selected_elements(element_ids):
    """Подсветить выбранные элементы."""
    try:
        ids = List[ElementId]()
        for eid in element_ids:
            ids.Add(eid)
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        return


def pick_faces(form):
    """Выбор граней в Revit (Enter для завершения)."""
    sel_filter = FaceSelectionFilter()

    try:
        # Используем PickObjects для множественного выбора (завершается Enter)
        refs = uidoc.Selection.PickObjects(
            ObjectType.Face,
            sel_filter,
            "Выберите грани и нажмите Enter для завершения"
        )

        for ref in refs:
            elem = doc.GetElement(ref.ElementId)
            elem_name = elem.Name if hasattr(elem, 'Name') else "Элемент"
            cat_name = elem.Category.Name if elem.Category else "Без категории"

            # Определяем тип грани
            face = elem.GetGeometryObjectFromReference(ref)
            if isinstance(face, PlanarFace):
                normal = face.FaceNormal
                if abs(normal.Z) > 0.7:
                    face_type = "горизонт."
                else:
                    face_type = "вертикал."
            else:
                face_type = "непланар."

            face_info = "{} [{}] ({}) - ID:{}".format(face_type, elem_name, cat_name, ref.ElementId.IntegerValue)

            # Проверяем, не добавлена ли уже эта грань
            ref_str = ref.ConvertToStableRepresentation(doc)
            already_added = False
            for existing_ref in form.selected_faces:
                if existing_ref.ConvertToStableRepresentation(doc) == ref_str:
                    already_added = True
                    break

            if not already_added:
                form.selected_faces.append(ref)
                form.face_info_list.append(face_info)

    except OperationCanceledException:
        # ESC - просто выходим без добавления
        pass
    except Exception as e:
        show_error("Ошибка выбора", "Не удалось выбрать грани", details=str(e))


def get_face_info(face_ref):
    """
    Получить информацию о грани: тип, контуры (внешний + отверстия), высоту.

    Returns:
        dict с ключами: is_horizontal, outer_points, inner_loops, min_z, max_z, height, normal
        или None при ошибке
    """
    try:
        elem = doc.GetElement(face_ref.ElementId)
        face = elem.GetGeometryObjectFromReference(face_ref)

        if not isinstance(face, PlanarFace):
            return None

        normal = face.FaceNormal
        edge_loops = face.EdgeLoops
        if edge_loops.Size == 0:
            return None

        # Первый контур - внешний
        outer_loop = edge_loops.get_Item(0)
        outer_points = []
        for edge in outer_loop:
            outer_points.append(edge.AsCurve().GetEndPoint(0))

        if len(outer_points) < 3:
            return None

        # Остальные контуры - отверстия (внутренние)
        inner_loops = []
        for i in range(1, edge_loops.Size):
            inner_loop = edge_loops.get_Item(i)
            inner_points = []
            for edge in inner_loop:
                inner_points.append(edge.AsCurve().GetEndPoint(0))
            if len(inner_points) >= 3:
                inner_loops.append(inner_points)

        # Собираем все Z для определения высоты
        all_points = outer_points[:]
        for inner_pts in inner_loops:
            all_points.extend(inner_pts)

        z_values = [p.Z for p in all_points]
        min_z = min(z_values)
        max_z = max(z_values)

        is_horizontal = abs(normal.Z) > 0.7

        return {
            'is_horizontal': is_horizontal,
            'outer_points': outer_points,
            'points': outer_points,  # для обратной совместимости
            'inner_loops': inner_loops,
            'min_z': min_z,
            'max_z': max_z,
            'height': max_z - min_z,
            'normal': normal
        }

    except Exception as e:
        show_error("Ошибка", "Не удалось получить информацию о грани", details=str(e))
        return None


def create_wall_on_vertical_face(face_ref, wall_type):
    """
    Создать стену по вертикальной грани с отверстиями для пересекающихся элементов.

    Args:
        face_ref: Reference на грань
        wall_type: WallType

    Returns:
        (Wall, area_info, None) или (None, None, error_message)
        area_info = {'outer_area': float, 'openings': [(name, area), ...], 'net_area': float}
    """
    try:
        info = get_face_info(face_ref)
        if info is None:
            return None, None, "Не удалось получить информацию о грани"

        if info['is_horizontal']:
            return None, None, "Грань горизонтальная, используйте крышу"

        points = info['points']
        min_z = info['min_z']
        max_z = info['max_z']
        height = info['height']
        normal = info['normal']

        if height < 0.01:
            return None, None, "Слишком маленькая высота грани"

        # Ищем горизонтальные рёбра грани (где Z почти одинаковый)
        elem = doc.GetElement(face_ref.ElementId)
        face = elem.GetGeometryObjectFromReference(face_ref)
        edge_loops = face.EdgeLoops
        outer_loop = edge_loops.get_Item(0)

        # Собираем все горизонтальные рёбра на нижнем уровне
        horizontal_edges = []
        for edge in outer_loop:
            curve = edge.AsCurve()
            ep0 = curve.GetEndPoint(0)
            ep1 = curve.GetEndPoint(1)
            # Проверяем, что ребро горизонтальное (Z почти одинаковый)
            if abs(ep0.Z - ep1.Z) < 0.01:
                # Проверяем, что ребро на нижнем уровне
                avg_z = (ep0.Z + ep1.Z) / 2
                if abs(avg_z - min_z) < 0.1:
                    length = ep0.DistanceTo(ep1)
                    horizontal_edges.append((ep0, ep1, length))

        # Берём самое длинное горизонтальное ребро
        if horizontal_edges:
            horizontal_edges.sort(key=lambda x: x[2], reverse=True)
            p1, p2, _ = horizontal_edges[0]
        else:
            # Если нет горизонтальных рёбер, используем bounding box
            x_vals = [p.X for p in points]
            y_vals = [p.Y for p in points]

            dx = max(x_vals) - min(x_vals)
            dy = max(y_vals) - min(y_vals)

            if dx > dy and dx > 0.01:
                # Грань вытянута по X
                avg_y = sum(y_vals) / len(y_vals)
                p1 = XYZ(min(x_vals), avg_y, min_z)
                p2 = XYZ(max(x_vals), avg_y, min_z)
            elif dy > 0.01:
                # Грань вытянута по Y
                avg_x = sum(x_vals) / len(x_vals)
                p1 = XYZ(avg_x, min(y_vals), min_z)
                p2 = XYZ(avg_x, max(y_vals), min_z)
            else:
                return None, None, "Грань слишком маленькая для создания стены"

        if p1.DistanceTo(p2) < 0.01:
            return None, None, "Не удалось определить базовую линию грани (слишком короткая)"

        # Вычисляем площадь внешнего контура (длина × высота)
        wall_length = p1.DistanceTo(p2)
        outer_area = wall_length * height

        # Смещаем линию наружу на половину толщины стены
        wall_width = 0
        width_param = wall_type.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
        if width_param:
            wall_width = width_param.AsDouble()

        normal_2d = XYZ(normal.X, normal.Y, 0)
        if normal_2d.GetLength() > 0.001:
            normal_2d = normal_2d.Normalize()
        else:
            edge_dir = (p2 - p1).Normalize()
            normal_2d = XYZ(-edge_dir.Y, edge_dir.X, 0)

        offset = wall_width / 2
        p1_offset = XYZ(p1.X + normal_2d.X * offset, p1.Y + normal_2d.Y * offset, min_z)
        p2_offset = XYZ(p2.X + normal_2d.X * offset, p2.Y + normal_2d.Y * offset, min_z)

        base_curve = Line.CreateBound(p1_offset, p2_offset)

        # Находим уровень
        level = get_level_below(min_z)
        if level is None:
            return None, None, "В проекте нет уровней"

        base_offset_val = min_z - level.Elevation

        # Создаём стену
        wall = Wall.Create(
            doc,
            base_curve,
            wall_type.Id,
            level.Id,
            height,
            base_offset_val,
            False,  # flip
            False   # structural
        )

        if wall is None:
            return None, None, "Не удалось создать стену"

        # Ищем пересекающиеся элементы и создаём отверстия
        source_element_id = face_ref.ElementId
        intersecting = find_intersecting_elements_for_vertical_face(face_ref, source_element_id)

        # Вектор вдоль стены
        wall_dir = (p2 - p1).Normalize()

        # Создаём отверстия и собираем площади
        openings_areas = []
        for rect in intersecting:
            try:
                plane_origin = rect.get('plane_origin')
                if plane_origin is None:
                    continue

                # Вычисляем offset относительно начала стены
                vec_to_origin = XYZ(plane_origin.X - p1.X, plane_origin.Y - p1.Y, 0)
                base_along = vec_to_origin.X * wall_dir.X + vec_to_origin.Y * wall_dir.Y

                # Координаты отверстия вдоль стены
                open_min_along = base_along + rect['min_along']
                open_max_along = base_along + rect['max_along']

                # XYZ точки углов отверстия
                lower_left = XYZ(
                    p1.X + wall_dir.X * open_min_along,
                    p1.Y + wall_dir.Y * open_min_along,
                    rect['min_z']
                )
                upper_right = XYZ(
                    p1.X + wall_dir.X * open_max_along,
                    p1.Y + wall_dir.Y * open_max_along,
                    rect['max_z']
                )

                # Проверяем, что отверстие в пределах стены
                if open_min_along < -0.1 or open_max_along > wall_length + 0.1:
                    continue
                if rect['min_z'] < min_z - 0.1 or rect['max_z'] > max_z + 0.1:
                    continue

                # Создаём отверстие в стене
                doc.Create.NewOpening(wall, lower_left, upper_right)

                # Площадь отверстия (ширина × высота)
                open_width = open_max_along - open_min_along
                open_height = rect['max_z'] - rect['min_z']
                opening_area = open_width * open_height
                elem_name = rect.get('name', 'Элемент')
                openings_areas.append((elem_name, opening_area))

            except Exception as open_err:
                elem_name = rect.get('name', 'Элемент')
                show_warning("Отверстие", "Не удалось создать отверстие для: {}".format(elem_name),
                            details=str(open_err), blocking=False, auto_close=1)

        # Вычисляем чистую площадь
        total_openings_area = sum(a for _, a in openings_areas)
        net_area = outer_area - total_openings_area

        area_info = {
            'outer_area': outer_area,
            'openings': openings_areas,
            'net_area': net_area
        }

        return wall, area_info, None

    except Exception as e:
        return None, None, str(e)


def create_roof_on_horizontal_face(face_ref, roof_type):
    """
    Создать крышу по горизонтальной грани с отверстиями.

    Args:
        face_ref: Reference на грань
        roof_type: RoofType

    Returns:
        (FootPrintRoof, area_info, None) или (None, None, error_message)
        area_info = {'outer_area': float, 'openings': [(name, area), ...], 'net_area': float}
    """
    try:
        info = get_face_info(face_ref)
        if info is None:
            return None, None, "Не удалось получить информацию о грани"

        if not info['is_horizontal']:
            return None, None, "Грань вертикальная, используйте стену"

        outer_points = info['outer_points']
        inner_loops = info.get('inner_loops', [])
        face_z = info['min_z']
        normal = info['normal']

        # Вычисляем площадь внешнего контура
        outer_area = calc_polygon_area(outer_points)

        # Получаем толщину крыши для корректного позиционирования
        roof_thickness = 0
        compound = roof_type.GetCompoundStructure()
        if compound:
            roof_thickness = compound.GetWidth()

        # Создаём внешний контур крыши
        curve_array = CurveArray()
        for i in range(len(outer_points)):
            p1 = outer_points[i]
            p2 = outer_points[(i + 1) % len(outer_points)]
            pt1 = XYZ(p1.X, p1.Y, face_z)
            pt2 = XYZ(p2.X, p2.Y, face_z)
            if pt1.DistanceTo(pt2) > 0.001:
                curve_array.Append(Line.CreateBound(pt1, pt2))

        if curve_array.Size < 3:
            return None, None, "Недостаточно сегментов для контура крыши"

        # Находим уровень
        level = get_level_below(face_z)
        if level is None:
            return None, None, "В проекте нет уровней"

        # Создаём крышу по контуру
        from clr import StrongBox
        model_curves = StrongBox[ModelCurveArray](ModelCurveArray())
        roof = doc.Create.NewFootPrintRoof(
            curve_array,
            level,
            roof_type,
            model_curves
        )

        if roof is None:
            return None, None, "Не удалось создать крышу"

        # Устанавливаем смещение от уровня
        base_offset_param = roof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM)
        if base_offset_param:
            if normal.Z > 0:
                roof_offset = face_z - level.Elevation
            else:
                roof_offset = face_z - level.Elevation - roof_thickness
            base_offset_param.Set(roof_offset)

        # Собираем все контуры для отверстий:
        # 1. Внутренние контуры грани (если есть)
        # 2. Контуры пересекающихся элементов (колонны, стены и т.д.)
        all_openings = []

        # Добавляем внутренние контуры грани
        for inner_points in inner_loops:
            all_openings.append(("Внутренний контур", inner_points))

        # Ищем пересекающиеся элементы
        source_element_id = face_ref.ElementId
        intersecting = find_intersecting_elements_contours(face_ref, source_element_id, face_z)
        all_openings.extend(intersecting)

        # Создаём отверстия для всех контуров и собираем площади
        openings_areas = []
        for elem_name, opening_points in all_openings:
            try:
                opening_curves = CurveArray()
                for i in range(len(opening_points)):
                    p1 = opening_points[i]
                    p2 = opening_points[(i + 1) % len(opening_points)]
                    pt1 = XYZ(p1.X, p1.Y, face_z)
                    pt2 = XYZ(p2.X, p2.Y, face_z)
                    if pt1.DistanceTo(pt2) > 0.001:
                        opening_curves.Append(Line.CreateBound(pt1, pt2))

                if opening_curves.Size >= 3:
                    doc.Create.NewOpening(roof, opening_curves, True)
                    opening_area = calc_polygon_area(opening_points)
                    openings_areas.append((elem_name, opening_area))
            except Exception as open_err:
                show_warning("Отверстие", "Не удалось создать отверстие для: {}".format(elem_name), details=str(open_err))

        # Вычисляем чистую площадь
        total_openings_area = sum(a for _, a in openings_areas)
        net_area = outer_area - total_openings_area

        area_info = {
            'outer_area': outer_area,
            'openings': openings_areas,
            'net_area': net_area
        }

        return roof, area_info, None

    except Exception as e:
        return None, None, str(e)


def create_waterproofing_on_face(face_ref, wall_type, roof_type):
    """
    Создать гидроизоляцию на грани.
    Вертикальная грань → Стена, Горизонтальная → Крыша.

    Returns:
        (element, element_type, area_info, None) или (None, None, None, error_message)
    """
    info = get_face_info(face_ref)
    if info is None:
        return None, None, None, "Не удалось определить тип грани"

    if info['is_horizontal']:
        elem, area_info, err = create_roof_on_horizontal_face(face_ref, roof_type)
        return elem, "крыша", area_info, err
    else:
        elem, area_info, err = create_wall_on_vertical_face(face_ref, wall_type)
        return elem, "стена", area_info, err


def format_area_report(created_elements):
    """
    Сформировать детальный отчёт по площадям.

    Args:
        created_elements: список dict с ключами:
            - index: номер элемента
            - elem_type: 'стена' или 'крыша'
            - area_info: dict с outer_area, openings, net_area

    Returns:
        строка с отчётом
    """
    lines = []

    total_net_area = 0.0

    for item in created_elements:
        idx = item['index']
        elem_type = item['elem_type']
        area_info = item['area_info']

        if area_info is None:
            continue

        outer_area = area_info['outer_area']
        openings = area_info['openings']
        net_area = area_info['net_area']

        # Конвертируем в м²
        outer_m2 = sq_feet_to_sq_m(outer_area)
        net_m2 = sq_feet_to_sq_m(net_area)

        total_net_area += net_m2

        # Формируем строку с формулой
        type_name = "Горизонтальная" if elem_type == "крыша" else "Вертикальная"
        lines.append("{} поверхность {} ({}):".format(idx, type_name, elem_type))

        if openings:
            # Формируем формулу: S_контура - S_отв1 - S_отв2 - ... = S_итого
            openings_m2 = [(name, sq_feet_to_sq_m(a)) for name, a in openings]
            total_openings_m2 = sum(a for _, a in openings_m2)

            formula_parts = ["{:.3f}".format(outer_m2)]
            for name, a in openings_m2:
                formula_parts.append("{:.3f}".format(a))

            formula = " - ".join(formula_parts) + " = {:.3f} m2".format(net_m2)
            lines.append("  S = {}".format(formula))

            # Детали отверстий
            for name, a in openings_m2:
                lines.append("    - {}: {:.3f} m2".format(name, a))
        else:
            lines.append("  S = {:.3f} m2 (без отверстий)".format(outer_m2))

        lines.append("")

    # Итого
    lines.append("=" * 40)
    lines.append("ИТОГО площадь: {:.3f} m2".format(total_net_area))

    return "\n".join(lines)


# === MAIN ===
if __name__ == "__main__":
    # Получаем типы стен и крыш
    wall_types = get_wall_types()
    roof_types = get_roof_types()

    if not wall_types:
        show_error("Ошибка", "В проекте нет типов стен",
                   details="Загрузите семейства стен")
        sys.exit()

    if not roof_types:
        show_error("Ошибка", "В проекте нет типов крыш",
                   details="Загрузите семейства крыш")
        sys.exit()

    all_categories = get_all_categories_for_hiding()

    # Данные между показами формы
    saved_faces = []
    saved_face_info = []
    saved_wall_type_idx = 0
    saved_roof_type_idx = 0

    # Цикл показа формы
    while True:
        form = WaterproofingForm(wall_types, roof_types, all_categories)

        form.selected_faces = saved_faces
        form.face_info_list = saved_face_info
        form.saved_wall_type_index = saved_wall_type_idx
        form.saved_roof_type_index = saved_roof_type_idx
        form.restore_state()

        result = form.ShowDialog()

        if result == DialogResult.Retry:
            saved_faces = form.selected_faces
            saved_face_info = form.face_info_list
            saved_wall_type_idx = form.cmb_wall_type.SelectedIndex
            saved_roof_type_idx = form.cmb_roof_type.SelectedIndex

            class TempHolder:
                def __init__(self):
                    self.selected_faces = saved_faces
                    self.face_info_list = saved_face_info

            temp = TempHolder()
            pick_faces(temp)

            saved_faces = temp.selected_faces
            saved_face_info = temp.face_info_list

        elif result == DialogResult.OK:
            selected_faces = form.selected_faces
            wall_type = form.selected_wall_type
            roof_type = form.selected_roof_type

            walls_created = 0
            roofs_created = 0
            errors = []
            created_elements = []

            with Transaction(doc, "Создание гидроизоляции") as t:
                t.Start()

                for i, face_ref in enumerate(selected_faces):
                    elem, elem_type, area_info, err = create_waterproofing_on_face(face_ref, wall_type, roof_type)
                    if elem:
                        if elem_type == "стена":
                            walls_created += 1
                        else:
                            roofs_created += 1
                        # Сохраняем информацию для отчёта
                        created_elements.append({
                            'index': i + 1,
                            'elem_type': elem_type,
                            'area_info': area_info
                        })
                    else:
                        errors.append("Грань {}: {}".format(i + 1, err or "неизвестная ошибка"))

                t.Commit()

            # Результат
            total = walls_created + roofs_created
            if total > 0:
                msg = "Создано: {} стен, {} крыш".format(walls_created, roofs_created)
                # Формируем детальный отчёт
                report = format_area_report(created_elements)
                if errors:
                    report = report + "\n\nОшибки:\n" + "\n".join(errors)
                    show_warning("Частичный успех", msg, details=report)
                else:
                    show_success("Готово", msg, details=report)
            else:
                show_error("Ошибка", "Не удалось создать гидроизоляцию",
                          details="\n".join(errors) if errors else "Неизвестная ошибка")
            break

        else:
            break
