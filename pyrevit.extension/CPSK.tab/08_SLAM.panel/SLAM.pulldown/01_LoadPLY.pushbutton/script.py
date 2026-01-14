# -*- coding: utf-8 -*-
"""
SLAM PLY iOS - Загрузка PLY файлов из iOS LiDAR сканера.
Поддерживает binary little endian формат с цветами.
Создаёт DirectShape или точечное облако в Revit.
"""

__title__ = "SLAM PLY\niOS"
__author__ = "CPSK"

import clr
import os
import sys
import struct
import codecs

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, CheckBox, RadioButton, GroupBox,
    Panel, OpenFileDialog, ProgressBar, DialogResult,
    FormStartPosition, FormBorderStyle, TextBox,
    MessageBox, MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Point, Size, Color, Font, FontStyle

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from cpsk_notify import show_error, show_warning, show_info, show_success
from cpsk_auth import require_auth

if not require_auth():
    sys.exit()

from pyrevit import revit, script
from Autodesk.Revit.DB import (
    Transaction, XYZ, Line,
    DirectShape, ElementId, BuiltInCategory,
    ViewPlan, Level, FilteredElementCollector, ViewFamilyType,
    PlanViewRange, PlanViewPlane
)

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


def mm_to_feet(mm):
    """Конвертация мм в футы."""
    return mm / 304.8


def m_to_feet(m):
    """Конвертация метров в футы."""
    return m * 3.28084


def parse_ply_header(file_path):
    """
    Парсит заголовок PLY файла.
    Возвращает: (vertex_count, has_colors, is_binary, header_size)
    """
    vertex_count = 0
    has_colors = False
    is_binary = False
    header_size = 0

    with open(file_path, 'rb') as f:
        line = f.readline().decode('ascii', errors='ignore').strip()
        if line != 'ply':
            return None, None, None, None

        while True:
            line = f.readline().decode('ascii', errors='ignore').strip()
            header_size = f.tell()

            if line.startswith('format'):
                if 'binary_little_endian' in line:
                    is_binary = True
            elif line.startswith('element vertex'):
                vertex_count = int(line.split()[-1])
            elif line.startswith('property') and ('red' in line or 'green' in line or 'blue' in line):
                has_colors = True
            elif line == 'end_header':
                break

    return vertex_count, has_colors, is_binary, header_size


def parse_ply_binary(file_path, with_colors=True, progress_callback=None):
    """
    Парсит binary PLY файл.
    ARKit использует Y-up, Revit использует Z-up.
    Конвертация: Revit_X = ARKit_X, Revit_Y = ARKit_Z, Revit_Z = ARKit_Y

    Возвращает: (points, colors) где points - список XYZ, colors - список (r,g,b) или None
    """
    vertex_count, has_colors, is_binary, header_size = parse_ply_header(file_path)

    if vertex_count is None:
        return None, None

    if not is_binary:
        return parse_ply_ascii(file_path, with_colors, progress_callback)

    points = []
    colors = [] if with_colors and has_colors else None

    # Структура: 3 float (xyz) + 3 uchar (rgb) = 12 + 3 = 15 bytes per vertex
    vertex_size = 12 + (3 if has_colors else 0)

    with open(file_path, 'rb') as f:
        f.seek(header_size)

        update_interval = max(1, vertex_count // 100)

        for i in range(vertex_count):
            # Читаем координаты (float32 little endian)
            data = f.read(12)
            if len(data) < 12:
                break
            x, y, z = struct.unpack('<fff', data)

            # Конвертация из ARKit (Y-up) в Revit (Z-up)
            # ARKit: X-right, Y-up, Z-backward
            # Revit: X-right, Y-forward, Z-up
            revit_x = x
            revit_y = -z  # ARKit Z backward -> Revit Y forward (инвертируем)
            revit_z = y   # ARKit Y up -> Revit Z up

            # Координаты в метрах -> футы
            points.append(XYZ(m_to_feet(revit_x), m_to_feet(revit_y), m_to_feet(revit_z)))

            # Читаем цвета если есть
            if has_colors:
                color_data = f.read(3)
                if len(color_data) == 3:
                    r, g, b = struct.unpack('<BBB', color_data)
                    if colors is not None:
                        colors.append((r, g, b))

            # Обновляем прогресс
            if progress_callback and i % update_interval == 0:
                progress_callback(int(100.0 * i / vertex_count))

    if progress_callback:
        progress_callback(100)

    return points, colors


def parse_ply_ascii(file_path, with_colors=True, progress_callback=None):
    """Парсит ASCII PLY файл."""
    vertex_count, has_colors, is_binary, header_size = parse_ply_header(file_path)

    if vertex_count is None:
        return None, None

    points = []
    colors = [] if with_colors and has_colors else None

    with codecs.open(file_path, 'r', 'utf-8') as f:
        # Пропускаем заголовок
        for line in f:
            if line.strip() == 'end_header':
                break

        update_interval = max(1, vertex_count // 100)
        i = 0

        for line in f:
            if i >= vertex_count:
                break

            parts = line.strip().split()
            if len(parts) >= 3:
                x, y, z = float(parts[0]), float(parts[1]), float(parts[2])

                # Конвертация из ARKit в Revit
                revit_x = x
                revit_y = -z
                revit_z = y

                points.append(XYZ(m_to_feet(revit_x), m_to_feet(revit_y), m_to_feet(revit_z)))

                if has_colors and len(parts) >= 6 and colors is not None:
                    r, g, b = int(parts[3]), int(parts[4]), int(parts[5])
                    colors.append((r, g, b))

            i += 1
            if progress_callback and i % update_interval == 0:
                progress_callback(int(100.0 * i / vertex_count))

    if progress_callback:
        progress_callback(100)

    return points, colors


def get_points_bounds(points):
    """Получить границы точек (min_z, max_z, center_z)."""
    if not points:
        return 0, 0, 0

    min_z = min(p.Z for p in points)
    max_z = max(p.Z for p in points)
    center_z = (min_z + max_z) / 2.0

    return min_z, max_z, center_z


def create_point_cloud(doc, points, name, cross_size_mm=10, max_points=30000):
    """
    Создаёт представление точечного облака через линии.
    Каждая точка представлена как маленький крест из линий.

    Args:
        doc: Revit document
        points: список XYZ точек
        name: имя для DirectShape
        cross_size_mm: размер креста в мм
        max_points: максимальное количество точек для отображения
    """
    # Создаём DirectShape с линиями
    ds = DirectShape.CreateElement(doc, ElementId(BuiltInCategory.OST_GenericModel))

    curves = []
    cross_size = mm_to_feet(cross_size_mm)

    # Ограничиваем количество точек
    actual_max = min(len(points), max_points)
    step = max(1, len(points) // actual_max)

    for i in range(0, len(points), step):
        p = points[i]
        try:
            # Крест в трёх плоскостях
            curves.append(Line.CreateBound(
                XYZ(p.X - cross_size, p.Y, p.Z),
                XYZ(p.X + cross_size, p.Y, p.Z)
            ))
            curves.append(Line.CreateBound(
                XYZ(p.X, p.Y - cross_size, p.Z),
                XYZ(p.X, p.Y + cross_size, p.Z)
            ))
            curves.append(Line.CreateBound(
                XYZ(p.X, p.Y, p.Z - cross_size),
                XYZ(p.X, p.Y, p.Z + cross_size)
            ))
        except:
            continue

    if curves:
        ds.SetShape(curves)

    return ds, len(curves) // 3  # Возвращаем количество отображённых точек


def get_or_create_level(doc, elevation, name):
    """Получить существующий или создать новый уровень."""
    # Ищем существующий уровень на этой отметке (с допуском 10см)
    tolerance = mm_to_feet(100)

    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    for level in levels:
        if abs(level.Elevation - elevation) < tolerance:
            return level

    # Создаём новый уровень
    level = Level.Create(doc, elevation)
    level.Name = name
    return level


def create_floor_plan(doc, level, view_depth_mm=100):
    """
    Создать план этажа для уровня.

    Args:
        doc: Revit document
        level: уровень для плана
        view_depth_mm: глубина вида в мм (по умолчанию 100мм = 10см)
    """
    # Ищем тип вида плана
    view_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
    floor_plan_type = None

    for vt in view_types:
        if vt.ViewFamily.ToString() == "FloorPlan":
            floor_plan_type = vt
            break

    if floor_plan_type is None:
        return None

    # Создаём план
    plan = ViewPlan.Create(doc, floor_plan_type.Id, level.Id)

    # Настраиваем View Range для ограничения глубины
    # Это скрывает всё что ниже view_depth от секущей плоскости
    view_range = plan.GetViewRange()

    # Устанавливаем все плоскости относительно уровня
    depth_feet = mm_to_feet(view_depth_mm)

    # Cut plane на уровне (offset = 0)
    view_range.SetOffset(PlanViewPlane.CutPlane, 0)
    # Bottom чуть ниже cut plane
    view_range.SetOffset(PlanViewPlane.BottomClipPlane, -depth_feet)
    # View Depth = Bottom (минимальная глубина)
    view_range.SetOffset(PlanViewPlane.ViewDepthPlane, -depth_feet)
    # Top выше уровня
    view_range.SetOffset(PlanViewPlane.TopClipPlane, depth_feet * 2)

    plan.SetViewRange(view_range)

    return plan


class PLYLoaderForm(Form):
    """Диалог загрузки PLY файла."""

    def __init__(self):
        self.result = None
        self.ply_path = None
        self.setup_form()

    def setup_form(self):
        self.Text = "SLAM PLY iOS - Загрузка сканирования"
        self.Width = 450
        self.Height = 400
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # Выбор файла
        lbl_file = Label()
        lbl_file.Text = "PLY файл:"
        lbl_file.Location = Point(15, y)
        lbl_file.Size = Size(70, 20)
        self.Controls.Add(lbl_file)

        self.txt_file = Label()
        self.txt_file.Text = "Файл не выбран"
        self.txt_file.Location = Point(90, y)
        self.txt_file.Size = Size(240, 20)
        self.txt_file.ForeColor = Color.Gray
        self.Controls.Add(self.txt_file)

        btn_browse = Button()
        btn_browse.Text = "Обзор..."
        btn_browse.Location = Point(340, y - 3)
        btn_browse.Size = Size(80, 25)
        btn_browse.Click += self.on_browse
        self.Controls.Add(btn_browse)

        y += 40

        # Информация о файле
        self.lbl_info = Label()
        self.lbl_info.Text = ""
        self.lbl_info.Location = Point(15, y)
        self.lbl_info.Size = Size(400, 40)
        self.lbl_info.ForeColor = Color.DarkBlue
        self.Controls.Add(self.lbl_info)

        y += 50

        # Группа настроек отображения
        grp_display = GroupBox()
        grp_display.Text = "Настройки отображения"
        grp_display.Location = Point(15, y)
        grp_display.Size = Size(405, 75)

        # Размер точки
        lbl_size = Label()
        lbl_size.Text = "Размер точки (мм):"
        lbl_size.Location = Point(15, 22)
        lbl_size.Size = Size(120, 20)
        grp_display.Controls.Add(lbl_size)

        self.txt_point_size = TextBox()
        self.txt_point_size.Text = "10"
        self.txt_point_size.Location = Point(140, 20)
        self.txt_point_size.Size = Size(50, 20)
        grp_display.Controls.Add(self.txt_point_size)

        # Макс. количество точек
        lbl_max = Label()
        lbl_max.Text = "Макс. точек:"
        lbl_max.Location = Point(210, 22)
        lbl_max.Size = Size(80, 20)
        grp_display.Controls.Add(lbl_max)

        self.txt_max_points = TextBox()
        self.txt_max_points.Text = "30000"
        self.txt_max_points.Location = Point(295, 20)
        self.txt_max_points.Size = Size(70, 20)
        grp_display.Controls.Add(self.txt_max_points)

        # Подсказка
        lbl_hint = Label()
        lbl_hint.Text = "Точки отображаются как 3D-кресты из линий"
        lbl_hint.Location = Point(15, 48)
        lbl_hint.Size = Size(380, 20)
        lbl_hint.ForeColor = Color.Gray
        grp_display.Controls.Add(lbl_hint)

        self.Controls.Add(grp_display)

        y += 85

        # Опция создания планов
        self.chk_create_plans = CheckBox()
        self.chk_create_plans.Text = "Создать 3 плана (низ, середина, верх)"
        self.chk_create_plans.Location = Point(15, y)
        self.chk_create_plans.Size = Size(300, 25)
        self.chk_create_plans.Checked = True
        self.Controls.Add(self.chk_create_plans)

        y += 35

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

        y += 35

        # Кнопки
        btn_ok = Button()
        btn_ok.Text = "Загрузить"
        btn_ok.Location = Point(240, y)
        btn_ok.Size = Size(85, 30)
        btn_ok.Click += self.on_ok
        self.Controls.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(335, y)
        btn_cancel.Size = Size(85, 30)
        btn_cancel.Click += self.on_cancel
        self.Controls.Add(btn_cancel)

        self.AcceptButton = btn_ok
        self.CancelButton = btn_cancel

    def on_browse(self, sender, args):
        dialog = OpenFileDialog()
        dialog.Filter = "PLY files (*.ply)|*.ply|All files (*.*)|*.*"
        dialog.Title = "Выберите PLY файл сканирования"

        if dialog.ShowDialog() == DialogResult.OK:
            self.ply_path = dialog.FileName
            self.txt_file.Text = os.path.basename(self.ply_path)
            self.txt_file.ForeColor = Color.Black

            # Читаем информацию о файле
            vertex_count, has_colors, is_binary, _ = parse_ply_header(self.ply_path)
            if vertex_count:
                info = "Точек: {:,}".format(vertex_count).replace(',', ' ')
                info += " | Цвета: {}".format("Да" if has_colors else "Нет")
                info += " | Формат: {}".format("Binary" if is_binary else "ASCII")
                self.lbl_info.Text = info

    def update_progress(self, value):
        self.progress.Value = value
        System.Windows.Forms.Application.DoEvents()

    def on_ok(self, sender, args):
        if not self.ply_path or not os.path.exists(self.ply_path):
            show_warning("Внимание", "Выберите PLY файл")
            return

        # Валидация числовых полей
        try:
            point_size = int(self.txt_point_size.Text)
            max_points = int(self.txt_max_points.Text)
            if point_size < 1 or point_size > 100:
                show_warning("Внимание", "Размер точки должен быть от 1 до 100 мм")
                return
            if max_points < 1000 or max_points > 100000:
                show_warning("Внимание", "Количество точек должно быть от 1000 до 100000")
                return
        except ValueError:
            show_warning("Внимание", "Введите корректные числа")
            return

        self.result = {
            "path": self.ply_path,
            "point_size_mm": point_size,
            "max_points": max_points,
            "create_plans": self.chk_create_plans.Checked
        }
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


def main():
    """Основная функция."""
    form = PLYLoaderForm()
    if form.ShowDialog() != DialogResult.OK:
        return

    opts = form.result
    ply_path = opts["path"]
    point_size_mm = opts["point_size_mm"]
    max_points = opts["max_points"]
    create_plans = opts["create_plans"]

    # Парсим PLY файл
    output.print_md("## Загрузка PLY файла...")
    output.print_md("Файл: {}".format(os.path.basename(ply_path)))

    points, colors = parse_ply_binary(ply_path, with_colors=False)

    if not points:
        show_error("Ошибка", "Не удалось прочитать PLY файл")
        return

    output.print_md("Загружено точек: **{:,}**".format(len(points)).replace(',', ' '))

    # Получаем границы
    min_z, max_z, center_z = get_points_bounds(points)
    output.print_md("Диапазон высот: {:.2f}м - {:.2f}м".format(min_z / 3.28084, max_z / 3.28084))

    # Создаём геометрию в Revit
    with Transaction(doc, "Загрузка SLAM PLY") as t:
        t.Start()

        # Создаём точечное облако
        scan_name = os.path.splitext(os.path.basename(ply_path))[0]

        output.print_md("Создание точечного облака...")
        output.print_md("- Размер точки: {} мм".format(point_size_mm))
        output.print_md("- Макс. точек: {:,}".format(max_points).replace(',', ' '))

        ds, displayed_points = create_point_cloud(doc, points, scan_name, point_size_mm, max_points)
        output.print_md("Отображено точек: **{:,}**".format(displayed_points).replace(',', ' '))

        # Создаём планы если нужно
        if create_plans:
            output.print_md("## Создание планов...")

            # Три отметки: низ, середина, верх
            elevations = [
                (min_z, "SLAM_Низ_{:.1f}м".format(min_z / 3.28084)),
                (center_z, "SLAM_Середина_{:.1f}м".format(center_z / 3.28084)),
                (max_z - mm_to_feet(100), "SLAM_Верх_{:.1f}м".format((max_z - mm_to_feet(100)) / 3.28084))
            ]

            plan_errors = []
            for elev, name in elevations:
                try:
                    level = get_or_create_level(doc, elev, name)
                    plan = create_floor_plan(doc, level)
                    if plan:
                        plan.Name = name
                        output.print_md("- Создан план: {}".format(name))
                except Exception as e:
                    plan_errors.append("{}: {}".format(name, str(e)))
                    continue

            if plan_errors:
                show_warning("Предупреждение", "Некоторые планы не созданы", details="\n".join(plan_errors))

        t.Commit()

    show_success(
        "Загрузка завершена",
        "Отображено {:,} из {:,} точек".format(displayed_points, len(points)).replace(',', ' '),
        details="Файл: {}\nРазмер точки: {} мм\nПланы: {}".format(
            os.path.basename(ply_path),
            point_size_mm,
            "Созданы" if create_plans else "Не создавались"
        )
    )


if __name__ == "__main__":
    main()
