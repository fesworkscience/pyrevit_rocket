# -*- coding: utf-8 -*-
"""
SLAM PLY iOS - Загрузка PLY файлов из iOS LiDAR сканера.
Поддерживает binary little endian формат с цветами.
Создаёт DirectShape по слоям высоты для управления видимостью на планах.
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
    Form, Label, Button, CheckBox, GroupBox,
    OpenFileDialog, ProgressBar, DialogResult,
    FormStartPosition, FormBorderStyle, TextBox
)
from System.Drawing import Point, Size, Color

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from cpsk_notify import show_error, show_warning, show_success
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


def feet_to_m(feet):
    """Конвертация футов в метры."""
    return feet / 3.28084


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


def parse_ply_binary(file_path, progress_callback=None):
    """
    Парсит binary PLY файл.
    ARKit использует Y-up, Revit использует Z-up.
    Конвертация: Revit_X = ARKit_X, Revit_Y = -ARKit_Z, Revit_Z = ARKit_Y

    Возвращает: список точек [(x, y, z), ...] в метрах
    """
    vertex_count, has_colors, is_binary, header_size = parse_ply_header(file_path)

    if vertex_count is None:
        return None

    if not is_binary:
        return parse_ply_ascii(file_path, progress_callback)

    points = []

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
            revit_x = x
            revit_y = -z
            revit_z = y

            points.append((revit_x, revit_y, revit_z))

            # Пропускаем цвета если есть
            if has_colors:
                f.read(3)

            # Обновляем прогресс
            if progress_callback and i % update_interval == 0:
                progress_callback(int(50.0 * i / vertex_count))

    if progress_callback:
        progress_callback(50)

    return points


def parse_ply_ascii(file_path, progress_callback=None):
    """Парсит ASCII PLY файл."""
    vertex_count, has_colors, is_binary, header_size = parse_ply_header(file_path)

    if vertex_count is None:
        return None

    points = []

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

                points.append((revit_x, revit_y, revit_z))

            i += 1
            if progress_callback and i % update_interval == 0:
                progress_callback(int(50.0 * i / vertex_count))

    if progress_callback:
        progress_callback(50)

    return points


def get_points_bounds(points):
    """Получить границы точек (min_z, max_z) в метрах."""
    if not points:
        return 0, 0

    min_z = min(p[2] for p in points)
    max_z = max(p[2] for p in points)

    return min_z, max_z


def split_points_by_layers(points, layer_height_m):
    """
    Разбить точки на слои по высоте.

    Args:
        points: список точек [(x, y, z), ...] в метрах
        layer_height_m: высота слоя в метрах

    Returns:
        dict {layer_index: [(x, y, z), ...]}
    """
    if not points:
        return {}

    min_z, max_z = get_points_bounds(points)

    layers = {}

    for p in points:
        # Определяем индекс слоя
        layer_idx = int((p[2] - min_z) / layer_height_m)

        if layer_idx not in layers:
            layers[layer_idx] = []
        layers[layer_idx].append(p)

    return layers, min_z


def create_layer_directshape(doc, points, name, cross_size_mm=10, max_points_per_layer=10000):
    """
    Создаёт DirectShape для одного слоя точек.

    Args:
        doc: Revit document
        points: список точек [(x, y, z), ...] в метрах
        name: имя для DirectShape
        cross_size_mm: размер креста в мм
        max_points_per_layer: максимум точек на слой

    Returns:
        (DirectShape, displayed_count)
    """
    ds = DirectShape.CreateElement(doc, ElementId(BuiltInCategory.OST_GenericModel))

    curves = []
    cross_size = mm_to_feet(cross_size_mm)

    # Ограничиваем количество точек
    step = max(1, len(points) // max_points_per_layer)
    displayed = 0

    for i in range(0, len(points), step):
        p = points[i]
        # Конвертируем в футы
        px = m_to_feet(p[0])
        py = m_to_feet(p[1])
        pz = m_to_feet(p[2])

        try:
            # Крест в трёх плоскостях
            curves.append(Line.CreateBound(
                XYZ(px - cross_size, py, pz),
                XYZ(px + cross_size, py, pz)
            ))
            curves.append(Line.CreateBound(
                XYZ(px, py - cross_size, pz),
                XYZ(px, py + cross_size, pz)
            ))
            curves.append(Line.CreateBound(
                XYZ(px, py, pz - cross_size),
                XYZ(px, py, pz + cross_size)
            ))
            displayed += 1
        except Exception:
            continue

    if curves:
        ds.SetShape(curves)

    # Устанавливаем имя через параметр Mark (не критично если не получится)
    mark_param = ds.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
    if mark_param:
        mark_param.Set(name)

    return ds, displayed


def get_or_create_level(doc, elevation, name):
    """Получить существующий или создать новый уровень."""
    tolerance = mm_to_feet(100)

    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    for level in levels:
        if abs(level.Elevation - elevation) < tolerance:
            return level

    level = Level.Create(doc, elevation)
    level.Name = name
    return level


def create_floor_plan(doc, level, view_depth_mm=100):
    """Создать план этажа для уровня."""
    view_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
    floor_plan_type = None

    for vt in view_types:
        if vt.ViewFamily.ToString() == "FloorPlan":
            floor_plan_type = vt
            break

    if floor_plan_type is None:
        return None

    plan = ViewPlan.Create(doc, floor_plan_type.Id, level.Id)

    view_range = plan.GetViewRange()
    depth_feet = mm_to_feet(view_depth_mm)

    view_range.SetOffset(PlanViewPlane.CutPlane, 0)
    view_range.SetOffset(PlanViewPlane.BottomClipPlane, -depth_feet)
    view_range.SetOffset(PlanViewPlane.ViewDepthPlane, -depth_feet)
    view_range.SetOffset(PlanViewPlane.TopClipPlane, depth_feet * 2)

    plan.SetViewRange(view_range)

    return plan


# Импорт BuiltInParameter для Mark
from Autodesk.Revit.DB import BuiltInParameter


class PLYLoaderForm(Form):
    """Диалог загрузки PLY файла."""

    def __init__(self):
        self.result = None
        self.ply_path = None
        self.setup_form()

    def setup_form(self):
        self.Text = "SLAM PLY iOS - Загрузка сканирования"
        self.Width = 450
        self.Height = 380
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

        # Группа настроек слоёв
        grp_layers = GroupBox()
        grp_layers.Text = "Настройки слоёв"
        grp_layers.Location = Point(15, y)
        grp_layers.Size = Size(405, 80)

        # Высота слоя
        lbl_layer = Label()
        lbl_layer.Text = "Высота слоя (мм):"
        lbl_layer.Location = Point(15, 25)
        lbl_layer.Size = Size(110, 20)
        grp_layers.Controls.Add(lbl_layer)

        self.txt_layer_height = TextBox()
        self.txt_layer_height.Text = "500"
        self.txt_layer_height.Location = Point(130, 23)
        self.txt_layer_height.Size = Size(60, 20)
        grp_layers.Controls.Add(self.txt_layer_height)

        # Размер точки
        lbl_size = Label()
        lbl_size.Text = "Размер точки (мм):"
        lbl_size.Location = Point(210, 25)
        lbl_size.Size = Size(115, 20)
        grp_layers.Controls.Add(lbl_size)

        self.txt_point_size = TextBox()
        self.txt_point_size.Text = "10"
        self.txt_point_size.Location = Point(330, 23)
        self.txt_point_size.Size = Size(50, 20)
        grp_layers.Controls.Add(self.txt_point_size)

        # Подсказка
        lbl_hint = Label()
        lbl_hint.Text = "Каждый слой - отдельный DirectShape для управления видимостью"
        lbl_hint.Location = Point(15, 52)
        lbl_hint.Size = Size(380, 20)
        lbl_hint.ForeColor = Color.Gray
        grp_layers.Controls.Add(lbl_hint)

        self.Controls.Add(grp_layers)

        y += 90

        # Опция создания планов
        self.chk_create_plans = CheckBox()
        self.chk_create_plans.Text = "Создать планы для каждого слоя"
        self.chk_create_plans.Location = Point(15, y)
        self.chk_create_plans.Size = Size(300, 25)
        self.chk_create_plans.Checked = False
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

    def set_status(self, text):
        self.lbl_status.Text = text
        System.Windows.Forms.Application.DoEvents()

    def on_ok(self, sender, args):
        if not self.ply_path or not os.path.exists(self.ply_path):
            show_warning("Внимание", "Выберите PLY файл")
            return

        # Валидация
        try:
            layer_height = int(self.txt_layer_height.Text)
            point_size = int(self.txt_point_size.Text)
            if layer_height < 100 or layer_height > 5000:
                show_warning("Внимание", "Высота слоя должна быть от 100 до 5000 мм")
                return
            if point_size < 1 or point_size > 100:
                show_warning("Внимание", "Размер точки должен быть от 1 до 100 мм")
                return
        except ValueError:
            show_warning("Внимание", "Введите корректные числа")
            return

        self.result = {
            "path": self.ply_path,
            "layer_height_mm": layer_height,
            "point_size_mm": point_size,
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
    layer_height_mm = opts["layer_height_mm"]
    point_size_mm = opts["point_size_mm"]
    create_plans = opts["create_plans"]

    layer_height_m = layer_height_mm / 1000.0

    # Собираем лог
    log = []

    # Парсим PLY
    log.append("=== Загрузка PLY файла ===")
    log.append("Файл: {}".format(os.path.basename(ply_path)))

    points = parse_ply_binary(ply_path)

    if not points:
        show_error("Ошибка", "Не удалось прочитать PLY файл")
        return

    log.append("Загружено точек: {:,}".format(len(points)).replace(',', ' '))

    min_z, max_z = get_points_bounds(points)
    log.append("Диапазон высот: {:.2f}м - {:.2f}м".format(min_z, max_z))
    log.append("")

    # Разбиваем на слои
    log.append("=== Разбиение на слои ===")
    log.append("Высота слоя: {} мм".format(layer_height_mm))

    layers, base_z = split_points_by_layers(points, layer_height_m)
    log.append("Создано слоёв: {}".format(len(layers)))
    log.append("")

    # Создаём геометрию
    scan_name = os.path.splitext(os.path.basename(ply_path))[0]

    with Transaction(doc, "Загрузка SLAM PLY") as t:
        t.Start()

        log.append("=== Создание DirectShape ===")
        total_displayed = 0

        for layer_idx in sorted(layers.keys()):
            layer_points = layers[layer_idx]

            # Высоты слоя в метрах
            z_from = base_z + layer_idx * layer_height_m
            z_to = z_from + layer_height_m

            layer_name = "SLAM_{:.1f}-{:.1f}m".format(z_from, z_to)

            ds, displayed = create_layer_directshape(
                doc, layer_points, layer_name, point_size_mm
            )
            total_displayed += displayed

            log.append("+ {} ({} точек)".format(layer_name, displayed))

        log.append("")
        log.append("Всего отображено: {:,} точек".format(total_displayed).replace(',', ' '))

        # Создаём планы если нужно
        if create_plans:
            log.append("")
            log.append("=== Планы ===")

            for layer_idx in sorted(layers.keys()):
                z_from = base_z + layer_idx * layer_height_m
                z_to = z_from + layer_height_m
                center_z = (z_from + z_to) / 2.0

                level_name = "SLAM_{:.1f}-{:.1f}m".format(z_from, z_to)

                try:
                    level = get_or_create_level(doc, m_to_feet(center_z), level_name)
                    plan = create_floor_plan(doc, level, layer_height_mm)
                    if plan:
                        plan.Name = level_name
                        log.append("+ {}".format(level_name))
                except Exception as e:
                    log.append("- {} (ошибка: {})".format(level_name, str(e)))
                    continue

        t.Commit()

    show_success(
        "Загрузка завершена",
        "Создано {} слоёв, {:,} точек".format(len(layers), total_displayed).replace(',', ' '),
        details="\n".join(log)
    )


if __name__ == "__main__":
    main()
