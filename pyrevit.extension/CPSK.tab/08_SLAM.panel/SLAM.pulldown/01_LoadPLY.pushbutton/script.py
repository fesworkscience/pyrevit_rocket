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
    FormStartPosition, FormBorderStyle, DockStyle, AnchorStyles,
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
    Transaction, XYZ, Line, Plane,
    DirectShape, DirectShapeLibrary, DirectShapeType,
    ElementId, BuiltInCategory,
    TessellatedShapeBuilder, TessellatedFace, ShapeBuilderTarget, ShapeBuilderFallback,
    ViewPlan, Level, FilteredElementCollector,
    BoundingBoxXYZ, ViewFamilyType
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


def create_directshape_from_points(doc, points, colors, name, use_colors=True):
    """
    Создаёт DirectShape из точек.
    Для визуализации создаём маленькие тетраэдры в каждой точке.
    """
    # Размер точки (в футах) ~ 5мм
    point_size = mm_to_feet(5)

    # Создаём DirectShape
    ds_type = DirectShape.CreateElement(doc, ElementId(BuiltInCategory.OST_GenericModel))

    builder = TessellatedShapeBuilder()
    builder.OpenConnectedFaceSet(False)

    # Для каждой точки создаём маленький тетраэдр
    # Ограничиваем количество для производительности
    max_points = min(len(points), 50000)
    step = max(1, len(points) // max_points)

    for i in range(0, len(points), step):
        p = points[i]

        # Вершины тетраэдра
        v0 = p
        v1 = XYZ(p.X + point_size, p.Y, p.Z)
        v2 = XYZ(p.X, p.Y + point_size, p.Z)
        v3 = XYZ(p.X, p.Y, p.Z + point_size)

        # 4 грани тетраэдра
        try:
            builder.AddFace(TessellatedFace([v0, v2, v1], ElementId.InvalidElementId))
            builder.AddFace(TessellatedFace([v0, v1, v3], ElementId.InvalidElementId))
            builder.AddFace(TessellatedFace([v0, v3, v2], ElementId.InvalidElementId))
            builder.AddFace(TessellatedFace([v1, v2, v3], ElementId.InvalidElementId))
        except:
            continue

    builder.CloseConnectedFaceSet()
    builder.Target = ShapeBuilderTarget.Solid
    builder.Fallback = ShapeBuilderFallback.Mesh

    result = builder.Build()

    if result.Outcome == result.Outcome.Success:
        ds_type.SetShape(list(result.GetGeometricalObjects()))

    return ds_type


def create_point_cloud_simple(doc, points, name):
    """
    Создаёт упрощённое представление точечного облака через линии.
    Каждая точка представлена как маленький крест из линий.
    """
    # Создаём DirectShape с линиями
    ds = DirectShape.CreateElement(doc, ElementId(BuiltInCategory.OST_GenericModel))

    curves = []
    cross_size = mm_to_feet(10)  # 10мм крест

    # Ограничиваем количество точек
    max_points = min(len(points), 20000)
    step = max(1, len(points) // max_points)

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

    return ds


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


def create_floor_plan(doc, level):
    """Создать план этажа для уровня."""
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

        # Группа опций цвета
        grp_color = GroupBox()
        grp_color.Text = "Цвета"
        grp_color.Location = Point(15, y)
        grp_color.Size = Size(200, 70)

        self.rb_with_colors = RadioButton()
        self.rb_with_colors.Text = "С цветами"
        self.rb_with_colors.Location = Point(15, 20)
        self.rb_with_colors.Checked = True
        grp_color.Controls.Add(self.rb_with_colors)

        self.rb_no_colors = RadioButton()
        self.rb_no_colors.Text = "Без цветов"
        self.rb_no_colors.Location = Point(15, 42)
        grp_color.Controls.Add(self.rb_no_colors)

        self.Controls.Add(grp_color)

        # Группа опций представления
        grp_type = GroupBox()
        grp_type.Text = "Тип представления"
        grp_type.Location = Point(225, y)
        grp_type.Size = Size(200, 70)

        self.rb_directshape = RadioButton()
        self.rb_directshape.Text = "DirectShape (меши)"
        self.rb_directshape.Location = Point(15, 20)
        self.rb_directshape.Checked = True
        grp_type.Controls.Add(self.rb_directshape)

        self.rb_points = RadioButton()
        self.rb_points.Text = "Точки (линии)"
        self.rb_points.Location = Point(15, 42)
        grp_type.Controls.Add(self.rb_points)

        self.Controls.Add(grp_type)

        y += 80

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

                # Включаем/выключаем опцию цветов
                if not has_colors:
                    self.rb_no_colors.Checked = True
                    self.rb_with_colors.Enabled = False
                else:
                    self.rb_with_colors.Enabled = True

    def update_progress(self, value):
        self.progress.Value = value
        System.Windows.Forms.Application.DoEvents()

    def on_ok(self, sender, args):
        if not self.ply_path or not os.path.exists(self.ply_path):
            show_warning("Внимание", "Выберите PLY файл")
            return

        self.result = {
            "path": self.ply_path,
            "with_colors": self.rb_with_colors.Checked,
            "use_directshape": self.rb_directshape.Checked,
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
    with_colors = opts["with_colors"]
    use_directshape = opts["use_directshape"]
    create_plans = opts["create_plans"]

    # Парсим PLY файл
    output.print_md("## Загрузка PLY файла...")
    output.print_md("Файл: {}".format(os.path.basename(ply_path)))

    points, colors = parse_ply_binary(ply_path, with_colors)

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

        # Создаём DirectShape или точечное облако
        scan_name = os.path.splitext(os.path.basename(ply_path))[0]

        if use_directshape:
            output.print_md("Создание DirectShape...")
            ds = create_directshape_from_points(doc, points, colors, scan_name, with_colors)
            output.print_md("DirectShape создан")
        else:
            output.print_md("Создание точечного облака...")
            ds = create_point_cloud_simple(doc, points, scan_name)
            output.print_md("Точечное облако создано")

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
        "Загружено {:,} точек".format(len(points)).replace(',', ' '),
        details="Файл: {}\nТип: {}\nПланы: {}".format(
            os.path.basename(ply_path),
            "DirectShape" if use_directshape else "Точки",
            "Созданы" if create_plans else "Не создавались"
        )
    )


if __name__ == "__main__":
    main()
