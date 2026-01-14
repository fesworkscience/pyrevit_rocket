# -*- coding: utf-8 -*-
"""
Fine-Tune ICP - автоматическое выравнивание методом ICP.

Алгоритм ICP (Iterative Closest Point):
1. Для каждой точки source находим ближайшую в target
2. Вычисляем оптимальную трансформацию
3. Применяем трансформацию к source
4. Повторяем до сходимости
"""

__title__ = "Fine-Tune\nICP"
__author__ = "CPSK"

import clr
import os
import sys
import math

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, GroupBox, ProgressBar, NumericUpDown,
    DialogResult, FormStartPosition, FormBorderStyle
)
from System.Drawing import Point, Size, Color

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

from alignment_utils import icp_align, apply_transform, build_spatial_grid

from pyrevit import revit, script
from Autodesk.Revit.DB import (
    Transaction, XYZ, Line, ElementId,
    DirectShape, BuiltInCategory, BuiltInParameter,
    Options, Transform, ViewDetailLevel
)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# Константы действий
ACTION_NONE = 0
ACTION_SELECT_SOURCE = 1
ACTION_SELECT_TARGET = 2


class DirectShapeFilter(ISelectionFilter):
    """Фильтр для выбора только DirectShape."""

    def AllowElement(self, element):
        return isinstance(element, DirectShape)

    def AllowReference(self, reference, position):
        return True


def extract_points_from_directshape(element):
    """
    Извлечь точки из DirectShape.

    Предполагаем, что каждые 3 линии - это крестик одной точки.
    Берём центр первой линии как координату точки.

    Returns:
        list: [(x, y, z), ...] в футах
    """
    points = []

    opt = Options()
    opt.DetailLevel = ViewDetailLevel.Fine
    geom = element.get_Geometry(opt)

    if geom is None:
        return points

    lines = []
    for geom_obj in geom:
        if hasattr(geom_obj, 'GetEndPoint'):
            lines.append(geom_obj)

    # Каждые 3 линии = 1 точка
    for i in range(0, len(lines), 3):
        if i < len(lines):
            line = lines[i]
            start = line.GetEndPoint(0)
            end = line.GetEndPoint(1)
            center = (
                (start.X + end.X) / 2,
                (start.Y + end.Y) / 2,
                (start.Z + end.Z) / 2
            )
            points.append(center)

    return points


def apply_transform_to_directshape(doc, element, R, t):
    """
    Применить трансформацию к DirectShape.
    """
    opt = Options()
    opt.DetailLevel = ViewDetailLevel.Fine
    geom = element.get_Geometry(opt)

    if geom is None:
        return False

    transform = Transform.Identity
    transform.BasisX = XYZ(R[0][0], R[1][0], R[2][0])
    transform.BasisY = XYZ(R[0][1], R[1][1], R[2][1])
    transform.BasisZ = XYZ(R[0][2], R[1][2], R[2][2])
    transform.Origin = XYZ(t[0], t[1], t[2])

    new_curves = []
    for geom_obj in geom:
        if hasattr(geom_obj, 'CreateTransformed'):
            transformed = geom_obj.CreateTransformed(transform)
            new_curves.append(transformed)

    if new_curves:
        element.SetShape(new_curves)
        return True

    return False


class ICPState:
    """Состояние ICP выравнивания."""

    def __init__(self):
        self.source_id = None
        self.source_name = ""
        self.source_points = []

        self.target_id = None
        self.target_name = ""
        self.target_points = []

        self.pending_action = ACTION_NONE


class ICPForm(Form):
    """Форма ICP выравнивания."""

    def __init__(self, state):
        self.state = state
        self.setup_form()

    def setup_form(self):
        self.Text = "Fine-Tune ICP - Автоматическое выравнивание"
        self.Width = 450
        self.Height = 380
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # === Исходная модель (Source) ===
        grp_source = GroupBox()
        grp_source.Text = "Исходная модель (будет перемещена)"
        grp_source.Location = Point(15, y)
        grp_source.Size = Size(405, 70)

        self.lbl_source = Label()
        self.lbl_source.Text = self.state.source_name if self.state.source_name else "Не выбрана"
        self.lbl_source.Location = Point(15, 25)
        self.lbl_source.Size = Size(250, 20)
        grp_source.Controls.Add(self.lbl_source)

        self.lbl_source_pts = Label()
        pts_count = len(self.state.source_points)
        self.lbl_source_pts.Text = "{} точек".format(pts_count) if pts_count else ""
        self.lbl_source_pts.Location = Point(15, 45)
        self.lbl_source_pts.Size = Size(150, 18)
        self.lbl_source_pts.ForeColor = Color.Gray
        grp_source.Controls.Add(self.lbl_source_pts)

        self.btn_source = Button()
        self.btn_source.Text = "Выбрать"
        self.btn_source.Location = Point(300, 22)
        self.btn_source.Size = Size(90, 25)
        self.btn_source.Click += self.on_select_source
        grp_source.Controls.Add(self.btn_source)

        self.Controls.Add(grp_source)

        y += 80

        # === Целевая модель (Target) ===
        grp_target = GroupBox()
        grp_target.Text = "Целевая модель (остаётся на месте)"
        grp_target.Location = Point(15, y)
        grp_target.Size = Size(405, 70)

        self.lbl_target = Label()
        self.lbl_target.Text = self.state.target_name if self.state.target_name else "Не выбрана"
        self.lbl_target.Location = Point(15, 25)
        self.lbl_target.Size = Size(250, 20)
        grp_target.Controls.Add(self.lbl_target)

        self.lbl_target_pts = Label()
        pts_count = len(self.state.target_points)
        self.lbl_target_pts.Text = "{} точек".format(pts_count) if pts_count else ""
        self.lbl_target_pts.Location = Point(15, 45)
        self.lbl_target_pts.Size = Size(150, 18)
        self.lbl_target_pts.ForeColor = Color.Gray
        grp_target.Controls.Add(self.lbl_target_pts)

        self.btn_target = Button()
        self.btn_target.Text = "Выбрать"
        self.btn_target.Location = Point(300, 22)
        self.btn_target.Size = Size(90, 25)
        self.btn_target.Click += self.on_select_target
        grp_target.Controls.Add(self.btn_target)

        self.Controls.Add(grp_target)

        y += 80

        # === Параметры ICP ===
        grp_params = GroupBox()
        grp_params.Text = "Параметры ICP"
        grp_params.Location = Point(15, y)
        grp_params.Size = Size(405, 70)

        lbl_iter = Label()
        lbl_iter.Text = "Макс. итераций:"
        lbl_iter.Location = Point(15, 28)
        lbl_iter.Size = Size(100, 20)
        grp_params.Controls.Add(lbl_iter)

        self.num_iterations = NumericUpDown()
        self.num_iterations.Location = Point(120, 25)
        self.num_iterations.Size = Size(60, 20)
        self.num_iterations.Minimum = 10
        self.num_iterations.Maximum = 200
        self.num_iterations.Value = 50
        grp_params.Controls.Add(self.num_iterations)

        lbl_dist = Label()
        lbl_dist.Text = "Макс. расстояние (м):"
        lbl_dist.Location = Point(200, 28)
        lbl_dist.Size = Size(130, 20)
        grp_params.Controls.Add(lbl_dist)

        self.num_max_dist = NumericUpDown()
        self.num_max_dist.Location = Point(335, 25)
        self.num_max_dist.Size = Size(55, 20)
        self.num_max_dist.Minimum = System.Decimal(0.1)
        self.num_max_dist.Maximum = System.Decimal(10.0)
        self.num_max_dist.Value = System.Decimal(1.0)
        self.num_max_dist.DecimalPlaces = 1
        self.num_max_dist.Increment = System.Decimal(0.1)
        grp_params.Controls.Add(self.num_max_dist)

        self.Controls.Add(grp_params)

        y += 80

        # Прогресс
        self.progress = ProgressBar()
        self.progress.Location = Point(15, y)
        self.progress.Size = Size(405, 20)
        self.progress.Minimum = 0
        self.progress.Maximum = 100
        self.Controls.Add(self.progress)

        y += 30

        self.lbl_status = Label()
        self.lbl_status.Text = ""
        self.lbl_status.Location = Point(15, y)
        self.lbl_status.Size = Size(405, 20)
        self.Controls.Add(self.lbl_status)

        y += 30

        # Кнопки
        self.btn_align = Button()
        self.btn_align.Text = "Выровнять"
        self.btn_align.Location = Point(230, y)
        self.btn_align.Size = Size(90, 30)
        self.btn_align.Click += self.on_align
        self.btn_align.Enabled = self.can_align()
        self.Controls.Add(self.btn_align)

        self.btn_close = Button()
        self.btn_close.Text = "Закрыть"
        self.btn_close.Location = Point(330, y)
        self.btn_close.Size = Size(90, 30)
        self.btn_close.Click += self.on_close
        self.Controls.Add(self.btn_close)

    def can_align(self):
        return (
            self.state.source_id is not None and
            self.state.target_id is not None and
            len(self.state.source_points) > 0 and
            len(self.state.target_points) > 0
        )

    def on_select_source(self, sender, args):
        self.state.pending_action = ACTION_SELECT_SOURCE
        self.DialogResult = DialogResult.Retry
        self.Close()

    def on_select_target(self, sender, args):
        self.state.pending_action = ACTION_SELECT_TARGET
        self.DialogResult = DialogResult.Retry
        self.Close()

    def on_align(self, sender, args):
        self.state.pending_action = 100  # Run ICP
        self.state.max_iterations = int(self.num_iterations.Value)
        self.state.max_dist = float(self.num_max_dist.Value)
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_close(self, sender, args):
        self.state.pending_action = ACTION_NONE
        self.DialogResult = DialogResult.Cancel
        self.Close()


def main():
    """Основная функция."""

    state = ICPState()
    ds_filter = DirectShapeFilter()

    while True:
        form = ICPForm(state)
        result = form.ShowDialog()

        if result == DialogResult.Cancel:
            break

        action = state.pending_action

        # === Выбор исходной модели ===
        if action == ACTION_SELECT_SOURCE:
            try:
                ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    ds_filter,
                    "Выберите исходную модель (будет перемещена)"
                )
                elem = doc.GetElement(ref.ElementId)
                state.source_id = ref.ElementId

                mark = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
                state.source_name = mark.AsString() if mark and mark.AsString() else "DirectShape"

                # Извлекаем точки
                state.source_points = extract_points_from_directshape(elem)

            except OperationCanceledException:
                pass
            continue

        # === Выбор целевой модели ===
        if action == ACTION_SELECT_TARGET:
            try:
                ref = uidoc.Selection.PickObject(
                    ObjectType.Element,
                    ds_filter,
                    "Выберите целевую модель (остаётся на месте)"
                )
                elem = doc.GetElement(ref.ElementId)
                state.target_id = ref.ElementId

                mark = elem.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
                state.target_name = mark.AsString() if mark and mark.AsString() else "DirectShape"

                state.target_points = extract_points_from_directshape(elem)

            except OperationCanceledException:
                pass
            continue

        # === Запуск ICP ===
        if action == 100:
            if not state.source_points or not state.target_points:
                show_warning("Внимание", "Не удалось извлечь точки из моделей")
                continue

            # Конвертируем футы в метры для ICP
            def feet_to_m(ft):
                return ft / 3.28084

            source_m = [(feet_to_m(p[0]), feet_to_m(p[1]), feet_to_m(p[2]))
                        for p in state.source_points]
            target_m = [(feet_to_m(p[0]), feet_to_m(p[1]), feet_to_m(p[2]))
                        for p in state.target_points]

            # Запускаем ICP
            R, t, final_error, iterations = icp_align(
                source_m, target_m,
                max_iterations=state.max_iterations,
                max_correspondence_dist=state.max_dist
            )

            # Конвертируем перенос обратно в футы
            m_to_feet = 3.28084
            t_feet = (t[0] * m_to_feet, t[1] * m_to_feet, t[2] * m_to_feet)

            # Применяем трансформацию
            elem = doc.GetElement(state.source_id)

            with Transaction(doc, "ICP выравнивание") as trans:
                trans.Start()
                success = apply_transform_to_directshape(doc, elem, R, t_feet)
                trans.Commit()

            if success:
                show_success(
                    "ICP выравнивание завершено",
                    "Итераций: {}, финальная ошибка: {:.4f} м".format(iterations, final_error)
                )
            else:
                show_error("Ошибка", "Не удалось применить трансформацию")

            break


if __name__ == "__main__":
    main()
