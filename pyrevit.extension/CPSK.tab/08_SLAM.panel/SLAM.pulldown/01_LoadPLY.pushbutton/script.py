# -*- coding: utf-8 -*-
"""
SLAM PLY iOS - Загрузка PLY файлов из iOS LiDAR сканера.

Возможности:
- Потоковое чтение больших файлов
- Фильтрация: Voxel Grid, Statistical Outlier, Radius Outlier
- Визуализация: раскраска по высоте, оригинальные цвета
- Разбиение на слои по высоте для управления видимостью
"""

__title__ = "SLAM PLY\niOS"
__author__ = "CPSK"

import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, CheckBox, GroupBox, ComboBox,
    OpenFileDialog, ProgressBar, DialogResult,
    FormStartPosition, FormBorderStyle, TextBox
)
from System.Drawing import Point, Size, Color

# Добавляем lib и текущую папку в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from cpsk_notify import show_error, show_warning, show_success
from cpsk_auth import require_auth

if not require_auth():
    sys.exit()

# Импортируем локальные модули
from ply_parser import parse_header, parse_ply_full, get_bounds
from ply_filters import voxel_grid_filter, statistical_outlier_filter, radius_outlier_filter
from ply_visualization import (
    apply_colors, COLOR_MODE_NONE, COLOR_MODE_HEIGHT, COLOR_MODE_ORIGINAL,
    has_colors, get_height_color, group_points_by_color
)

from pyrevit import revit, script
from Autodesk.Revit.DB import (
    Transaction, XYZ, Line,
    DirectShape, ElementId, BuiltInCategory, BuiltInParameter,
    ViewPlan, View3D, Level, FilteredElementCollector, ViewFamilyType,
    PlanViewPlane, OverrideGraphicSettings
)
from Autodesk.Revit.DB import Color as RevitColor

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


def mm_to_feet(mm):
    """Конвертация мм в футы."""
    return mm / 304.8


def m_to_feet(m):
    """Конвертация метров в футы."""
    return m * 3.28084


def split_points_by_layers(points, layer_height_m):
    """
    Разбить точки на слои по высоте.

    Args:
        points: список точек [(x, y, z, r, g, b), ...]
        layer_height_m: высота слоя в метрах

    Returns:
        (dict {layer_index: [points]}, min_z)
    """
    if not points:
        return {}, 0

    min_z = min(p[2] for p in points)

    layers = {}

    for p in points:
        layer_idx = int((p[2] - min_z) / layer_height_m)
        if layer_idx not in layers:
            layers[layer_idx] = []
        layers[layer_idx].append(p)

    return layers, min_z


def create_layer_directshape(doc, points, name, cross_size_mm=10, max_points_per_layer=10000):
    """
    Создаёт DirectShape для одного слоя точек.
    """
    ds = DirectShape.CreateElement(doc, ElementId(BuiltInCategory.OST_GenericModel))

    curves = []
    cross_size = mm_to_feet(cross_size_mm)

    step = max(1, len(points) // max_points_per_layer)
    displayed = 0

    for i in range(0, len(points), step):
        p = points[i]
        px = m_to_feet(p[0])
        py = m_to_feet(p[1])
        pz = m_to_feet(p[2])

        try:
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


def apply_layer_color_override(view, element_id, r, g, b):
    """
    Применить цвет к элементу через Override Graphics на виде.

    Args:
        view: вид Revit
        element_id: ID элемента
        r, g, b: компоненты цвета 0-255
    """
    ogs = OverrideGraphicSettings()
    color = RevitColor(int(r), int(g), int(b))

    # Устанавливаем цвет линий проекции
    ogs.SetProjectionLineColor(color)

    # Применяем к элементу на виде
    view.SetElementOverrides(element_id, ogs)


class PLYLoaderForm(Form):
    """Диалог загрузки PLY файла."""

    def __init__(self):
        self.result = None
        self.ply_path = None
        self.ply_header = None
        self.setup_form()

    def setup_form(self):
        self.Text = "SLAM PLY iOS - Загрузка сканирования"
        self.Width = 480
        self.Height = 580
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # === Выбор файла ===
        lbl_file = Label()
        lbl_file.Text = "PLY файл:"
        lbl_file.Location = Point(15, y)
        lbl_file.Size = Size(70, 20)
        self.Controls.Add(lbl_file)

        self.txt_file = Label()
        self.txt_file.Text = "Файл не выбран"
        self.txt_file.Location = Point(90, y)
        self.txt_file.Size = Size(270, 20)
        self.txt_file.ForeColor = Color.Gray
        self.Controls.Add(self.txt_file)

        btn_browse = Button()
        btn_browse.Text = "Обзор..."
        btn_browse.Location = Point(370, y - 3)
        btn_browse.Size = Size(80, 25)
        btn_browse.Click += self.on_browse
        self.Controls.Add(btn_browse)

        y += 30

        self.lbl_info = Label()
        self.lbl_info.Text = ""
        self.lbl_info.Location = Point(15, y)
        self.lbl_info.Size = Size(440, 20)
        self.lbl_info.ForeColor = Color.DarkBlue
        self.Controls.Add(self.lbl_info)

        y += 30

        # === Фильтрация ===
        grp_filters = GroupBox()
        grp_filters.Text = "Фильтрация точек"
        grp_filters.Location = Point(15, y)
        grp_filters.Size = Size(435, 130)

        # Voxel Grid
        self.chk_voxel = CheckBox()
        self.chk_voxel.Text = "Voxel Grid (равномерное прореживание)"
        self.chk_voxel.Location = Point(15, 22)
        self.chk_voxel.Size = Size(280, 20)
        self.chk_voxel.Checked = True
        grp_filters.Controls.Add(self.chk_voxel)

        lbl_voxel = Label()
        lbl_voxel.Text = "Размер (мм):"
        lbl_voxel.Location = Point(300, 22)
        lbl_voxel.Size = Size(75, 20)
        grp_filters.Controls.Add(lbl_voxel)

        self.txt_voxel_size = TextBox()
        self.txt_voxel_size.Text = "50"
        self.txt_voxel_size.Location = Point(380, 20)
        self.txt_voxel_size.Size = Size(40, 20)
        grp_filters.Controls.Add(self.txt_voxel_size)

        # Statistical Outlier
        self.chk_statistical = CheckBox()
        self.chk_statistical.Text = "Statistical Outlier (удаление шума)"
        self.chk_statistical.Location = Point(15, 50)
        self.chk_statistical.Size = Size(280, 20)
        self.chk_statistical.Checked = False
        grp_filters.Controls.Add(self.chk_statistical)

        lbl_stat_k = Label()
        lbl_stat_k.Text = "K соседей:"
        lbl_stat_k.Location = Point(300, 50)
        lbl_stat_k.Size = Size(65, 20)
        grp_filters.Controls.Add(lbl_stat_k)

        self.txt_stat_k = TextBox()
        self.txt_stat_k.Text = "20"
        self.txt_stat_k.Location = Point(370, 48)
        self.txt_stat_k.Size = Size(35, 20)
        grp_filters.Controls.Add(self.txt_stat_k)

        # Radius Outlier
        self.chk_radius = CheckBox()
        self.chk_radius.Text = "Radius Outlier (удаление изолированных)"
        self.chk_radius.Location = Point(15, 78)
        self.chk_radius.Size = Size(280, 20)
        self.chk_radius.Checked = False
        grp_filters.Controls.Add(self.chk_radius)

        lbl_radius = Label()
        lbl_radius.Text = "Радиус (мм):"
        lbl_radius.Location = Point(300, 78)
        lbl_radius.Size = Size(70, 20)
        grp_filters.Controls.Add(lbl_radius)

        self.txt_radius = TextBox()
        self.txt_radius.Text = "100"
        self.txt_radius.Location = Point(375, 76)
        self.txt_radius.Size = Size(40, 20)
        grp_filters.Controls.Add(self.txt_radius)

        # Подсказка
        lbl_filter_hint = Label()
        lbl_filter_hint.Text = "Фильтры применяются последовательно"
        lbl_filter_hint.Location = Point(15, 105)
        lbl_filter_hint.Size = Size(400, 18)
        lbl_filter_hint.ForeColor = Color.Gray
        grp_filters.Controls.Add(lbl_filter_hint)

        self.Controls.Add(grp_filters)

        y += 140

        # === Визуализация ===
        grp_visual = GroupBox()
        grp_visual.Text = "Визуализация"
        grp_visual.Location = Point(15, y)
        grp_visual.Size = Size(435, 80)

        lbl_color = Label()
        lbl_color.Text = "Раскраска:"
        lbl_color.Location = Point(15, 22)
        lbl_color.Size = Size(70, 20)
        grp_visual.Controls.Add(lbl_color)

        self.cmb_color_mode = ComboBox()
        self.cmb_color_mode.Location = Point(90, 20)
        self.cmb_color_mode.Size = Size(200, 25)
        self.cmb_color_mode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self.cmb_color_mode.Items.Add("Без цвета (серый)")
        self.cmb_color_mode.Items.Add("По высоте (градиент)")
        self.cmb_color_mode.Items.Add("Оригинальные цвета из PLY")
        self.cmb_color_mode.SelectedIndex = 1  # По высоте по умолчанию
        grp_visual.Controls.Add(self.cmb_color_mode)

        # Квантизация цветов
        lbl_quantize = Label()
        lbl_quantize.Text = "Квантизация:"
        lbl_quantize.Location = Point(15, 50)
        lbl_quantize.Size = Size(80, 20)
        grp_visual.Controls.Add(lbl_quantize)

        self.cmb_quantize = ComboBox()
        self.cmb_quantize.Location = Point(100, 48)
        self.cmb_quantize.Size = Size(150, 25)
        self.cmb_quantize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self.cmb_quantize.Items.Add("4 (64 цвета)")
        self.cmb_quantize.Items.Add("8 (512 цветов)")
        self.cmb_quantize.Items.Add("16 (4096 цветов)")
        self.cmb_quantize.Items.Add("32 (32768 цветов)")
        self.cmb_quantize.SelectedIndex = 1  # 8 по умолчанию
        grp_visual.Controls.Add(self.cmb_quantize)

        lbl_quantize_hint = Label()
        lbl_quantize_hint.Text = "(для оригинальных цветов)"
        lbl_quantize_hint.Location = Point(260, 50)
        lbl_quantize_hint.Size = Size(160, 20)
        lbl_quantize_hint.ForeColor = Color.Gray
        grp_visual.Controls.Add(lbl_quantize_hint)

        self.Controls.Add(grp_visual)

        y += 90

        # === Настройки слоёв ===
        grp_layers = GroupBox()
        grp_layers.Text = "Настройки слоёв"
        grp_layers.Location = Point(15, y)
        grp_layers.Size = Size(435, 80)

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

        lbl_hint = Label()
        lbl_hint.Text = "Каждый слой - отдельный DirectShape для управления видимостью"
        lbl_hint.Location = Point(15, 52)
        lbl_hint.Size = Size(410, 20)
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
        self.progress.Size = Size(435, 20)
        self.progress.Minimum = 0
        self.progress.Maximum = 100
        self.Controls.Add(self.progress)

        y += 30

        self.lbl_status = Label()
        self.lbl_status.Text = ""
        self.lbl_status.Location = Point(15, y)
        self.lbl_status.Size = Size(435, 20)
        self.Controls.Add(self.lbl_status)

        y += 35

        # Кнопки
        btn_ok = Button()
        btn_ok.Text = "Загрузить"
        btn_ok.Location = Point(270, y)
        btn_ok.Size = Size(85, 30)
        btn_ok.Click += self.on_ok
        self.Controls.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(365, y)
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

            self.ply_header = parse_header(self.ply_path)
            if self.ply_header:
                info = "Точек: {:,}".format(self.ply_header.vertex_count).replace(',', ' ')
                info += " | Цвета: {}".format("Да" if self.ply_header.has_colors else "Нет")
                info += " | {}".format(self.ply_header.format)
                self.lbl_info.Text = info

                # Если нет цветов, меняем режим раскраски
                if not self.ply_header.has_colors:
                    self.cmb_color_mode.SelectedIndex = 1  # По высоте

    def update_progress(self, value):
        self.progress.Value = min(100, max(0, value))
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
            voxel_size = int(self.txt_voxel_size.Text)
            stat_k = int(self.txt_stat_k.Text)
            radius = int(self.txt_radius.Text)

            if layer_height < 100 or layer_height > 5000:
                show_warning("Внимание", "Высота слоя должна быть от 100 до 5000 мм")
                return
            if point_size < 1 or point_size > 100:
                show_warning("Внимание", "Размер точки должен быть от 1 до 100 мм")
                return
            if voxel_size < 10 or voxel_size > 500:
                show_warning("Внимание", "Размер воксела должен быть от 10 до 500 мм")
                return
        except ValueError:
            show_warning("Внимание", "Введите корректные числа")
            return

        # Режим раскраски
        color_modes = [COLOR_MODE_NONE, COLOR_MODE_HEIGHT, COLOR_MODE_ORIGINAL]
        color_mode = color_modes[self.cmb_color_mode.SelectedIndex]

        # Уровень квантизации
        quantize_levels = [4, 8, 16, 32]
        quantize_level = quantize_levels[self.cmb_quantize.SelectedIndex]

        self.result = {
            "path": self.ply_path,
            "layer_height_mm": layer_height,
            "point_size_mm": point_size,
            "create_plans": self.chk_create_plans.Checked,
            # Фильтры
            "use_voxel": self.chk_voxel.Checked,
            "voxel_size_mm": voxel_size,
            "use_statistical": self.chk_statistical.Checked,
            "stat_k": stat_k,
            "use_radius": self.chk_radius.Checked,
            "radius_mm": radius,
            # Визуализация
            "color_mode": color_mode,
            "quantize_level": quantize_level
        }
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


def main():
    """Основная функция."""
    # Проверяем активный вид ДО показа формы
    active_view = uidoc.ActiveView
    # ViewPlan и View3D поддерживают Override Graphics
    if not isinstance(active_view, (ViewPlan, View3D)):
        show_error(
            "Неподдерживаемый вид",
            "Текущий вид не поддерживает раскраску элементов.",
            details="Переключитесь на 3D вид или план этажа перед загрузкой.\nТекущий тип вида: {}".format(type(active_view).__name__)
        )
        return

    form = PLYLoaderForm()
    if form.ShowDialog() != DialogResult.OK:
        return

    opts = form.result
    ply_path = opts["path"]
    layer_height_mm = opts["layer_height_mm"]
    point_size_mm = opts["point_size_mm"]
    create_plans = opts["create_plans"]
    color_mode = opts["color_mode"]
    quantize_level = opts["quantize_level"]

    layer_height_m = layer_height_mm / 1000.0

    log = []
    log.append("=== Загрузка PLY файла ===")
    log.append("Файл: {}".format(os.path.basename(ply_path)))

    # Парсим PLY
    points = parse_ply_full(ply_path)

    if not points:
        show_error("Ошибка", "Не удалось прочитать PLY файл")
        return

    log.append("Загружено точек: {:,}".format(len(points)).replace(',', ' '))

    # === Применяем фильтры ===
    log.append("")
    log.append("=== Фильтрация ===")

    # Voxel Grid
    if opts["use_voxel"]:
        voxel_size_m = opts["voxel_size_mm"] / 1000.0
        before = len(points)
        points = voxel_grid_filter(points, voxel_size_m)
        log.append("Voxel Grid ({} мм): {:,} -> {:,}".format(
            opts["voxel_size_mm"], before, len(points)
        ).replace(',', ' '))

    # Statistical Outlier
    if opts["use_statistical"]:
        before = len(points)
        points = statistical_outlier_filter(points, k_neighbors=opts["stat_k"])
        log.append("Statistical Outlier (K={}): {:,} -> {:,}".format(
            opts["stat_k"], before, len(points)
        ).replace(',', ' '))

    # Radius Outlier
    if opts["use_radius"]:
        radius_m = opts["radius_mm"] / 1000.0
        before = len(points)
        points = radius_outlier_filter(points, radius_m=radius_m)
        log.append("Radius Outlier ({} мм): {:,} -> {:,}".format(
            opts["radius_mm"], before, len(points)
        ).replace(',', ' '))

    if not points:
        show_error("Ошибка", "После фильтрации не осталось точек")
        return

    # === Применяем цвета ===
    log.append("")
    log.append("=== Визуализация ===")

    bounds = get_bounds(points)
    min_z, max_z = bounds[4], bounds[5]

    points = apply_colors(points, color_mode, min_z, max_z)

    color_mode_names = {
        COLOR_MODE_NONE: "Без цвета",
        COLOR_MODE_HEIGHT: "По высоте",
        COLOR_MODE_ORIGINAL: "Оригинальные"
    }
    log.append("Раскраска: {}".format(color_mode_names.get(color_mode, "Неизвестно")))
    log.append("Диапазон высот: {:.2f}м - {:.2f}м".format(min_z, max_z))

    # === Разбиваем на слои ===
    log.append("")
    log.append("=== Разбиение на слои ===")
    log.append("Высота слоя: {} мм".format(layer_height_mm))

    layers, base_z = split_points_by_layers(points, layer_height_m)
    log.append("Создано слоёв: {}".format(len(layers)))
    log.append("")

    # === Создаём геометрию ===
    scan_name = os.path.splitext(os.path.basename(ply_path))[0]

    # Получаем активный вид для применения цветов
    active_view = uidoc.ActiveView

    # Храним созданные DirectShape и планы для раскраски
    created_elements = []  # [(ds, r, g, b), ...]
    created_plans = []  # [plan, ...]

    with Transaction(doc, "Загрузка SLAM PLY") as t:
        t.Start()

        log.append("=== Создание DirectShape ===")
        total_displayed = 0

        # Режим оригинальных цветов - группируем по цветам
        if color_mode == COLOR_MODE_ORIGINAL:
            log.append("Режим: оригинальные цвета (квантизация {})".format(quantize_level))

            for layer_idx in sorted(layers.keys()):
                layer_points = layers[layer_idx]

                z_from = base_z + layer_idx * layer_height_m
                z_to = z_from + layer_height_m

                # Группируем точки слоя по цветам
                color_groups = group_points_by_color(layer_points, levels=quantize_level)
                log.append("Слой {:.1f}-{:.1f}m: {} цветовых групп".format(
                    z_from, z_to, len(color_groups)))

                for color_key, group_points in color_groups.items():
                    r, g, b = color_key
                    group_name = "SLAM_{:.1f}m_RGB({},{},{})".format(z_from, r, g, b)

                    ds, displayed = create_layer_directshape(
                        doc, group_points, group_name, point_size_mm
                    )
                    total_displayed += displayed

                    # Сохраняем для раскраски планов
                    created_elements.append((ds, r, g, b))

                    # Применяем цвет на активном виде
                    apply_layer_color_override(active_view, ds.Id, r, g, b)

        else:
            # Режим по высоте или без цвета - один DirectShape на слой
            for layer_idx in sorted(layers.keys()):
                layer_points = layers[layer_idx]

                z_from = base_z + layer_idx * layer_height_m
                z_to = z_from + layer_height_m

                layer_name = "SLAM_{:.1f}-{:.1f}m".format(z_from, z_to)

                ds, displayed = create_layer_directshape(
                    doc, layer_points, layer_name, point_size_mm
                )
                total_displayed += displayed

                # Определяем цвет слоя
                if color_mode == COLOR_MODE_HEIGHT:
                    layer_center_z = (z_from + z_to) / 2.0
                    r, g, b = get_height_color(layer_center_z, min_z, max_z)
                else:
                    r, g, b = 128, 128, 128  # Серый для режима без цвета

                # Сохраняем для раскраски планов
                created_elements.append((ds, r, g, b))

                # Применяем цвет на активном виде
                if color_mode != COLOR_MODE_NONE:
                    apply_layer_color_override(active_view, ds.Id, r, g, b)

                log.append("+ {} ({} точек)".format(layer_name, displayed))

        log.append("")
        log.append("Всего отображено: {:,} точек".format(total_displayed).replace(',', ' '))

        # Создаём планы
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
                        created_plans.append(plan)
                        log.append("+ {}".format(level_name))
                except Exception as e:
                    log.append("- {} (ошибка: {})".format(level_name, str(e)))
                    continue

            # Раскрашиваем элементы на созданных планах
            if created_plans and color_mode != COLOR_MODE_NONE:
                log.append("")
                log.append("=== Раскраска планов ===")
                for plan in created_plans:
                    for ds, r, g, b in created_elements:
                        apply_layer_color_override(plan, ds.Id, r, g, b)
                    log.append("+ {}".format(plan.Name))

        t.Commit()

    show_success(
        "Загрузка завершена",
        "Создано {} слоёв, {:,} точек".format(len(layers), total_displayed).replace(',', ' '),
        details="\n".join(log)
    )


if __name__ == "__main__":
    main()
