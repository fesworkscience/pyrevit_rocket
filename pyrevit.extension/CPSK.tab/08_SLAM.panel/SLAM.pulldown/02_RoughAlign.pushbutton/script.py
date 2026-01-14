# -*- coding: utf-8 -*-
"""
Грубое выравнивание - ручное выравнивание по 4 точкам.

Алгоритм:
1. Выбрать первую модель (DirectShape)
2. Указать 4 точки на ней (синие маркеры)
3. Выбрать вторую модель
4. Указать 4 точки на ней (красные маркеры)
5. Выбрать направление выравнивания
6. Применить трансформацию
"""

__title__ = "Грубое\nвыравн."
__author__ = "CPSK"

import clr
import os
import sys
import math

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, GroupBox, Panel,
    DialogResult, FormStartPosition, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Point, Size, Color, Font, FontStyle

# Добавляем lib и текущую папку в путь
SCRIPT_DIR = os.path.dirname(__file__)
PULLDOWN_DIR = os.path.dirname(SCRIPT_DIR)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)
if PULLDOWN_DIR not in sys.path:
    sys.path.insert(0, PULLDOWN_DIR)

from cpsk_notify import show_error, show_warning, show_success, show_info
from cpsk_auth import require_auth

if not require_auth():
    sys.exit()

from alignment_utils import calculate_rigid_transform, apply_transform, matrix_vector_multiply

from pyrevit import revit, script
from Autodesk.Revit.DB import (
    Transaction, XYZ, Line, ElementId,
    DirectShape, BuiltInCategory, BuiltInParameter,
    FilteredElementCollector, GeometryElement, Options,
    Transform, ViewDetailLevel
)
from Autodesk.Revit.DB import Color as RevitColor
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# Константы для действий
ACTION_NONE = 0
ACTION_SELECT_MODEL1 = 1
ACTION_PICK_POINT1_1 = 2
ACTION_PICK_POINT1_2 = 3
ACTION_PICK_POINT1_3 = 4
ACTION_PICK_POINT1_4 = 5
ACTION_SELECT_MODEL2 = 6
ACTION_PICK_POINT2_1 = 7
ACTION_PICK_POINT2_2 = 8
ACTION_PICK_POINT2_3 = 9
ACTION_PICK_POINT2_4 = 10


class DirectShapeFilter(ISelectionFilter):
    """Фильтр для выбора только DirectShape."""

    def AllowElement(self, element):
        return isinstance(element, DirectShape)

    def AllowReference(self, reference, position):
        return True


def mm_to_feet(mm):
    """Конвертация мм в футы."""
    return mm / 304.8


def feet_to_m(feet):
    """Конвертация футов в метры."""
    return feet / 3.28084


def m_to_feet(m):
    """Конвертация метров в футы."""
    return m * 3.28084


def create_marker(doc, point, color_rgb, size_mm=50):
    """
    Создать временный маркер (крестик) в точке.

    Args:
        doc: документ Revit
        point: XYZ точка
        color_rgb: (r, g, b) цвет
        size_mm: размер маркера в мм

    Returns:
        DirectShape: созданный маркер
    """
    ds = DirectShape.CreateElement(doc, ElementId(BuiltInCategory.OST_GenericModel))

    half_size = mm_to_feet(size_mm)
    curves = []

    # Крестик в 3D
    curves.append(Line.CreateBound(
        XYZ(point.X - half_size, point.Y, point.Z),
        XYZ(point.X + half_size, point.Y, point.Z)
    ))
    curves.append(Line.CreateBound(
        XYZ(point.X, point.Y - half_size, point.Z),
        XYZ(point.X, point.Y + half_size, point.Z)
    ))
    curves.append(Line.CreateBound(
        XYZ(point.X, point.Y, point.Z - half_size),
        XYZ(point.X, point.Y, point.Z + half_size)
    ))

    ds.SetShape(curves)

    # Устанавливаем метку
    mark_param = ds.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
    if mark_param:
        mark_param.Set("TEMP_MARKER")

    return ds


def delete_markers(doc, marker_ids):
    """Удалить временные маркеры."""
    for eid in marker_ids:
        try:
            doc.Delete(eid)
        except Exception:
            # Маркер уже удалён или не существует - продолжаем
            continue


def extract_geometry_points(element):
    """
    Извлечь точки из геометрии DirectShape.

    Returns:
        list: [(x, y, z), ...] в футах
    """
    points = []

    opt = Options()
    opt.DetailLevel = ViewDetailLevel.Fine

    geom = element.get_Geometry(opt)
    if geom is None:
        return points

    for geom_obj in geom:
        if hasattr(geom_obj, 'GetEndPoint'):
            # Это линия
            start = geom_obj.GetEndPoint(0)
            end = geom_obj.GetEndPoint(1)
            center = XYZ(
                (start.X + end.X) / 2,
                (start.Y + end.Y) / 2,
                (start.Z + end.Z) / 2
            )
            points.append((center.X, center.Y, center.Z))

    # Убираем дубликаты (округляем до 1 мм)
    unique = {}
    for p in points:
        key = (round(p[0], 4), round(p[1], 4), round(p[2], 4))
        if key not in unique:
            unique[key] = p

    return list(unique.values())


def apply_transform_to_directshape(doc, element, R, t):
    """
    Применить трансформацию к DirectShape.

    Args:
        doc: документ Revit
        element: DirectShape
        R: матрица поворота 3x3
        t: вектор переноса (tx, ty, tz) в футах
    """
    # Получаем текущую геометрию
    opt = Options()
    opt.DetailLevel = ViewDetailLevel.Fine
    geom = element.get_Geometry(opt)

    if geom is None:
        return False

    # Создаём Transform из R и t
    # В Revit Transform - это 4x4 матрица
    transform = Transform.Identity

    # Устанавливаем компоненты поворота
    transform.BasisX = XYZ(R[0][0], R[1][0], R[2][0])
    transform.BasisY = XYZ(R[0][1], R[1][1], R[2][1])
    transform.BasisZ = XYZ(R[0][2], R[1][2], R[2][2])

    # Устанавливаем перенос
    transform.Origin = XYZ(t[0], t[1], t[2])

    # Трансформируем каждую линию
    new_curves = []
    for geom_obj in geom:
        if hasattr(geom_obj, 'CreateTransformed'):
            transformed = geom_obj.CreateTransformed(transform)
            new_curves.append(transformed)

    if new_curves:
        element.SetShape(new_curves)
        return True

    return False


class AlignmentState:
    """Состояние выравнивания."""

    def __init__(self):
        self.model1_id = None
        self.model1_name = ""
        self.points1 = []  # XYZ точки
        self.markers1 = []  # ElementId маркеров

        self.model2_id = None
        self.model2_name = ""
        self.points2 = []
        self.markers2 = []

        self.pending_action = ACTION_NONE


class RoughAlignForm(Form):
    """Форма грубого выравнивания."""

    def __init__(self, state):
        self.state = state
        self.setup_form()

    def setup_form(self):
        self.Text = "Грубое выравнивание по точкам"
        self.Width = 500
        self.Height = 420
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # === Модель 1 (синяя) ===
        grp1 = GroupBox()
        grp1.Text = "Модель 1 (синяя)"
        grp1.Location = Point(15, y)
        grp1.Size = Size(455, 130)
        grp1.ForeColor = Color.Blue

        self.lbl_model1 = Label()
        self.lbl_model1.Text = self.state.model1_name if self.state.model1_name else "Не выбрана"
        self.lbl_model1.Location = Point(15, 22)
        self.lbl_model1.Size = Size(280, 20)
        grp1.Controls.Add(self.lbl_model1)

        self.btn_select1 = Button()
        self.btn_select1.Text = "Выбрать модель"
        self.btn_select1.Location = Point(310, 18)
        self.btn_select1.Size = Size(130, 25)
        self.btn_select1.Click += self.on_select_model1
        grp1.Controls.Add(self.btn_select1)

        # Кнопки выбора точек
        y_pts = 55
        for i in range(4):
            btn = Button()
            btn.Text = "Точка {}".format(i + 1)
            btn.Location = Point(15 + i * 110, y_pts)
            btn.Size = Size(100, 25)
            btn.Tag = i + 1  # Номер точки
            btn.Click += self.on_pick_point1
            btn.Enabled = self.state.model1_id is not None

            # Если точка уже выбрана - меняем текст
            if i < len(self.state.points1):
                btn.Text = "Точка {} OK".format(i + 1)
                btn.BackColor = Color.LightBlue

            grp1.Controls.Add(btn)
            setattr(self, "btn_pt1_{}".format(i + 1), btn)

        # Статус
        self.lbl_status1 = Label()
        self.lbl_status1.Text = "Выбрано точек: {}".format(len(self.state.points1))
        self.lbl_status1.Location = Point(15, 90)
        self.lbl_status1.Size = Size(200, 20)
        grp1.Controls.Add(self.lbl_status1)

        self.Controls.Add(grp1)

        y += 140

        # === Модель 2 (красная) ===
        grp2 = GroupBox()
        grp2.Text = "Модель 2 (красная)"
        grp2.Location = Point(15, y)
        grp2.Size = Size(455, 130)
        grp2.ForeColor = Color.Red

        self.lbl_model2 = Label()
        self.lbl_model2.Text = self.state.model2_name if self.state.model2_name else "Не выбрана"
        self.lbl_model2.Location = Point(15, 22)
        self.lbl_model2.Size = Size(280, 20)
        grp2.Controls.Add(self.lbl_model2)

        self.btn_select2 = Button()
        self.btn_select2.Text = "Выбрать модель"
        self.btn_select2.Location = Point(310, 18)
        self.btn_select2.Size = Size(130, 25)
        self.btn_select2.Click += self.on_select_model2
        grp2.Controls.Add(self.btn_select2)

        # Кнопки выбора точек
        y_pts = 55
        for i in range(4):
            btn = Button()
            btn.Text = "Точка {}".format(i + 1)
            btn.Location = Point(15 + i * 110, y_pts)
            btn.Size = Size(100, 25)
            btn.Tag = i + 1
            btn.Click += self.on_pick_point2
            btn.Enabled = self.state.model2_id is not None

            if i < len(self.state.points2):
                btn.Text = "Точка {} OK".format(i + 1)
                btn.BackColor = Color.LightCoral

            grp2.Controls.Add(btn)
            setattr(self, "btn_pt2_{}".format(i + 1), btn)

        self.lbl_status2 = Label()
        self.lbl_status2.Text = "Выбрано точек: {}".format(len(self.state.points2))
        self.lbl_status2.Location = Point(15, 90)
        self.lbl_status2.Size = Size(200, 20)
        grp2.Controls.Add(self.lbl_status2)

        self.Controls.Add(grp2)

        y += 140

        # === Кнопки выравнивания ===
        self.btn_align_blue = Button()
        self.btn_align_blue.Text = "Синюю -> к Красной"
        self.btn_align_blue.Location = Point(15, y)
        self.btn_align_blue.Size = Size(150, 35)
        self.btn_align_blue.Click += self.on_align_blue_to_red
        self.btn_align_blue.Enabled = self.can_align()
        self.btn_align_blue.ForeColor = Color.Blue
        self.Controls.Add(self.btn_align_blue)

        self.btn_align_red = Button()
        self.btn_align_red.Text = "Красную -> к Синей"
        self.btn_align_red.Location = Point(175, y)
        self.btn_align_red.Size = Size(150, 35)
        self.btn_align_red.Click += self.on_align_red_to_blue
        self.btn_align_red.Enabled = self.can_align()
        self.btn_align_red.ForeColor = Color.Red
        self.Controls.Add(self.btn_align_red)

        # Очистка
        self.btn_clear = Button()
        self.btn_clear.Text = "Очистить"
        self.btn_clear.Location = Point(335, y)
        self.btn_clear.Size = Size(65, 35)
        self.btn_clear.Click += self.on_clear
        self.Controls.Add(self.btn_clear)

        # Закрыть
        self.btn_close = Button()
        self.btn_close.Text = "Закрыть"
        self.btn_close.Location = Point(405, y)
        self.btn_close.Size = Size(65, 35)
        self.btn_close.Click += self.on_close
        self.Controls.Add(self.btn_close)

        y += 50

        # Подсказка
        lbl_hint = Label()
        lbl_hint.Text = "Выберите обе модели и укажите по 4 соответствующих точки на каждой"
        lbl_hint.Location = Point(15, y)
        lbl_hint.Size = Size(455, 20)
        lbl_hint.ForeColor = Color.Gray
        self.Controls.Add(lbl_hint)

    def can_align(self):
        """Проверить, достаточно ли данных для выравнивания."""
        return (
            self.state.model1_id is not None and
            self.state.model2_id is not None and
            len(self.state.points1) >= 3 and
            len(self.state.points2) >= 3 and
            len(self.state.points1) == len(self.state.points2)
        )

    def on_select_model1(self, sender, args):
        self.state.pending_action = ACTION_SELECT_MODEL1
        self.DialogResult = DialogResult.Retry
        self.Close()

    def on_select_model2(self, sender, args):
        self.state.pending_action = ACTION_SELECT_MODEL2
        self.DialogResult = DialogResult.Retry
        self.Close()

    def on_pick_point1(self, sender, args):
        point_num = sender.Tag
        action_map = {
            1: ACTION_PICK_POINT1_1,
            2: ACTION_PICK_POINT1_2,
            3: ACTION_PICK_POINT1_3,
            4: ACTION_PICK_POINT1_4
        }
        self.state.pending_action = action_map.get(point_num, ACTION_NONE)
        self.DialogResult = DialogResult.Retry
        self.Close()

    def on_pick_point2(self, sender, args):
        point_num = sender.Tag
        action_map = {
            1: ACTION_PICK_POINT2_1,
            2: ACTION_PICK_POINT2_2,
            3: ACTION_PICK_POINT2_3,
            4: ACTION_PICK_POINT2_4
        }
        self.state.pending_action = action_map.get(point_num, ACTION_NONE)
        self.DialogResult = DialogResult.Retry
        self.Close()

    def on_align_blue_to_red(self, sender, args):
        """Выровнять синюю модель к красной."""
        self.state.pending_action = 100  # Align blue to red
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_align_red_to_blue(self, sender, args):
        """Выровнять красную модель к синей."""
        self.state.pending_action = 101  # Align red to blue
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_clear(self, sender, args):
        """Очистить состояние."""
        self.state.pending_action = 200  # Clear
        self.DialogResult = DialogResult.Retry
        self.Close()

    def on_close(self, sender, args):
        self.state.pending_action = ACTION_NONE
        self.DialogResult = DialogResult.Cancel
        self.Close()


def main():
    """Основная функция."""

    state = AlignmentState()
    ds_filter = DirectShapeFilter()

    while True:
        form = RoughAlignForm(state)
        result = form.ShowDialog()

        if result == DialogResult.Cancel:
            # Удаляем маркеры перед выходом
            if state.markers1 or state.markers2:
                with Transaction(doc, "Удаление маркеров") as t:
                    t.Start()
                    delete_markers(doc, state.markers1)
                    delete_markers(doc, state.markers2)
                    t.Commit()
            break

        action = state.pending_action

        # === Выбор модели 1 ===
        if action == ACTION_SELECT_MODEL1:
            try:
                ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    ds_filter,
                    "Выберите первую модель (DirectShape)"
                )
                elem = doc.GetElement(ref.ElementId)
                state.model1_id = ref.ElementId
                mark = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
                state.model1_name = mark.AsString() if mark else "DirectShape"
            except OperationCanceledException:
                pass
            continue

        # === Выбор модели 2 ===
        if action == ACTION_SELECT_MODEL2:
            try:
                ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    ds_filter,
                    "Выберите вторую модель (DirectShape)"
                )
                elem = doc.GetElement(ref.ElementId)
                state.model2_id = ref.ElementId
                mark = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
                state.model2_name = mark.AsString() if mark else "DirectShape"
            except OperationCanceledException:
                pass
            continue

        # === Выбор точек модели 1 ===
        if ACTION_PICK_POINT1_1 <= action <= ACTION_PICK_POINT1_4:
            point_idx = action - ACTION_PICK_POINT1_1
            try:
                ref = uidoc.Selection.PickObject(
                    ObjectType.PointOnElement,
                    "Укажите точку {} на модели 1".format(point_idx + 1)
                )
                point = ref.GlobalPoint

                with Transaction(doc, "Создание маркера") as t:
                    t.Start()
                    marker = create_marker(doc, point, (0, 0, 255), 100)
                    t.Commit()

                # Обновляем или добавляем точку
                if point_idx < len(state.points1):
                    # Удаляем старый маркер
                    old_marker = state.markers1[point_idx]
                    with Transaction(doc, "Удаление маркера") as t:
                        t.Start()
                        delete_markers(doc, [old_marker])
                        t.Commit()
                    state.points1[point_idx] = point
                    state.markers1[point_idx] = marker.Id
                else:
                    state.points1.append(point)
                    state.markers1.append(marker.Id)

            except OperationCanceledException:
                pass
            continue

        # === Выбор точек модели 2 ===
        if ACTION_PICK_POINT2_1 <= action <= ACTION_PICK_POINT2_4:
            point_idx = action - ACTION_PICK_POINT2_1
            try:
                ref = uidoc.Selection.PickObject(
                    ObjectType.PointOnElement,
                    "Укажите точку {} на модели 2".format(point_idx + 1)
                )
                point = ref.GlobalPoint

                with Transaction(doc, "Создание маркера") as t:
                    t.Start()
                    marker = create_marker(doc, point, (255, 0, 0), 100)
                    t.Commit()

                if point_idx < len(state.points2):
                    old_marker = state.markers2[point_idx]
                    with Transaction(doc, "Удаление маркера") as t:
                        t.Start()
                        delete_markers(doc, [old_marker])
                        t.Commit()
                    state.points2[point_idx] = point
                    state.markers2[point_idx] = marker.Id
                else:
                    state.points2.append(point)
                    state.markers2.append(marker.Id)

            except OperationCanceledException:
                pass
            continue

        # === Очистка ===
        if action == 200:
            with Transaction(doc, "Очистка маркеров") as t:
                t.Start()
                delete_markers(doc, state.markers1)
                delete_markers(doc, state.markers2)
                t.Commit()
            state = AlignmentState()
            continue

        # === Выравнивание ===
        if action == 100 or action == 101:
            # Проверяем данные
            if len(state.points1) < 3 or len(state.points2) < 3:
                show_warning("Внимание", "Нужно минимум 3 точки на каждой модели")
                continue

            if len(state.points1) != len(state.points2):
                show_warning("Внимание", "Количество точек должно совпадать")
                continue

            # Конвертируем XYZ в туплы
            pts1 = [(p.X, p.Y, p.Z) for p in state.points1]
            pts2 = [(p.X, p.Y, p.Z) for p in state.points2]

            if action == 100:
                # Синюю к красной: source=pts1, target=pts2
                source_pts = pts1
                target_pts = pts2
                source_id = state.model1_id
                direction = "синей к красной"
            else:
                # Красную к синей: source=pts2, target=pts1
                source_pts = pts2
                target_pts = pts1
                source_id = state.model2_id
                direction = "красной к синей"

            # Вычисляем трансформацию
            R, t = calculate_rigid_transform(source_pts, target_pts)

            # Применяем
            elem = doc.GetElement(source_id)

            with Transaction(doc, "Выравнивание модели") as trans:
                trans.Start()

                success = apply_transform_to_directshape(doc, elem, R, t)

                # Удаляем маркеры
                delete_markers(doc, state.markers1)
                delete_markers(doc, state.markers2)

                trans.Commit()

            if success:
                show_success(
                    "Выравнивание завершено",
                    "Модель трансформирована ({})".format(direction)
                )
            else:
                show_error("Ошибка", "Не удалось применить трансформацию")

            break

    # Конец


if __name__ == "__main__":
    main()
