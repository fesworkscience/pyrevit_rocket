# -*- coding: utf-8 -*-
"""
Обрезать арматуру - обрезка стержней вокруг отверстий в плите.

Workflow:
1. Выбрать плиту
2. Выбрать отверстия
3. Настроить параметры
4. Обрезать арматуру
"""

__title__ = "Обрезать\nарматуру"
__author__ = "CPSK"

# 1. Стандартные импорты
import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

# 2. System и WinForms импорты
import System
from System.Windows.Forms import (
    Form, Label, Button, Panel, CheckBox, NumericUpDown,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, AnchorStyles, GroupBox, TextBox,
    CheckedListBox, ScrollBars
)
from System.Drawing import Point, Size, Font, FontStyle

# 3. Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# 4. Проверка авторизации
from cpsk_auth import require_auth
from cpsk_notify import show_error, show_success, show_warning, show_info
from cpsk_config import get_setting, set_setting

if not require_auth():
    sys.exit()

# 5. pyrevit и Revit API
from pyrevit import revit

from Autodesk.Revit.DB import (
    Transaction, BuiltInCategory, FilteredElementCollector,
    ElementId, StorageType, BuiltInParameter
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

# 6. Импорт модулей для работы с арматурой
from cpsk_rebar_utils import (
    get_rebars_in_host, get_openings_in_floor, find_rebars_intersecting_opening,
    get_opening_solid, split_line_by_solid, split_line_by_rect_2d,
    create_rebar_from_curve, get_rebar_centerline, mm_to_feet,
    get_area_reinforcement_in_host, convert_area_reinforcement_to_rebars,
    get_all_openings_in_floor, find_rebars_intersecting_opening_2d,
    check_rebar_intersects_opening_2d, get_opening_2d_bounds, create_solid_from_curves
)
from cpsk_shared_params import (
    ensure_rebar_cut_param, add_opening_to_rebar, get_rebar_cut_data,
    get_shared_param_info, ensure_rebar_cut_param_with_info, REBAR_CUT_DATA_PARAM
)

# 7. Настройки
doc = revit.doc
uidoc = revit.uidoc
app = doc.Application


# === ФОРМА ===

class CutRebarForm(Form):
    """Форма для обрезки арматуры."""

    def __init__(self):
        self.selected_floor = None
        self.selected_floor_id = None
        self.selected_openings = []
        self.selected_opening_ids = []
        self.action = None  # "select_floor", "select_openings", "cut"

        self.setup_form()
        self.load_settings()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Обрезать арматуру"
        self.Width = 400
        self.Height = 450
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 10

        # Инструкция
        lbl_info = Label()
        lbl_info.Text = "1. Выберите плиту\n2. Выберите отверстия\n3. Нажмите 'Обрезать'"
        lbl_info.Location = Point(10, y)
        lbl_info.Size = Size(370, 45)
        self.Controls.Add(lbl_info)
        y += 55

        # --- Группа выбора ---
        grp_select = GroupBox()
        grp_select.Text = "Выбор элементов"
        grp_select.Location = Point(10, y)
        grp_select.Size = Size(370, 100)
        self.Controls.Add(grp_select)

        # Плита
        lbl_floor = Label()
        lbl_floor.Text = "Плита:"
        lbl_floor.Location = Point(10, 25)
        lbl_floor.Size = Size(50, 20)
        grp_select.Controls.Add(lbl_floor)

        self.txt_floor = TextBox()
        self.txt_floor.Location = Point(65, 22)
        self.txt_floor.Size = Size(200, 20)
        self.txt_floor.ReadOnly = True
        self.txt_floor.Text = "[не выбрана]"
        grp_select.Controls.Add(self.txt_floor)

        self.btn_select_floor = Button()
        self.btn_select_floor.Text = "Выбрать"
        self.btn_select_floor.Location = Point(275, 20)
        self.btn_select_floor.Size = Size(80, 25)
        self.btn_select_floor.Click += self.on_select_floor
        grp_select.Controls.Add(self.btn_select_floor)

        # Отверстия
        lbl_openings = Label()
        lbl_openings.Text = "Отверстий:"
        lbl_openings.Location = Point(10, 60)
        lbl_openings.Size = Size(70, 20)
        grp_select.Controls.Add(lbl_openings)

        self.txt_openings = TextBox()
        self.txt_openings.Location = Point(85, 57)
        self.txt_openings.Size = Size(180, 20)
        self.txt_openings.ReadOnly = True
        self.txt_openings.Text = "0"
        grp_select.Controls.Add(self.txt_openings)

        self.btn_select_openings = Button()
        self.btn_select_openings.Text = "Выбрать"
        self.btn_select_openings.Location = Point(275, 55)
        self.btn_select_openings.Size = Size(80, 25)
        self.btn_select_openings.Click += self.on_select_openings
        grp_select.Controls.Add(self.btn_select_openings)

        y += 110

        # --- Группа настроек ---
        grp_settings = GroupBox()
        grp_settings.Text = "Настройки"
        grp_settings.Location = Point(10, y)
        grp_settings.Size = Size(370, 130)
        self.Controls.Add(grp_settings)

        # Отступ от края
        lbl_offset = Label()
        lbl_offset.Text = "Отступ от края отверстия (мм):"
        lbl_offset.Location = Point(10, 25)
        lbl_offset.Size = Size(200, 20)
        grp_settings.Controls.Add(lbl_offset)

        self.num_offset = NumericUpDown()
        self.num_offset.Location = Point(220, 22)
        self.num_offset.Size = Size(80, 20)
        self.num_offset.Minimum = 0
        self.num_offset.Maximum = 500
        self.num_offset.Value = 50
        grp_settings.Controls.Add(self.num_offset)

        # Минимальная длина
        lbl_min_len = Label()
        lbl_min_len.Text = "Мин. длина стержня (мм):"
        lbl_min_len.Location = Point(10, 55)
        lbl_min_len.Size = Size(200, 20)
        grp_settings.Controls.Add(lbl_min_len)

        self.num_min_len = NumericUpDown()
        self.num_min_len.Location = Point(220, 52)
        self.num_min_len.Size = Size(80, 20)
        self.num_min_len.Minimum = 0
        self.num_min_len.Maximum = 1000
        self.num_min_len.Value = 100
        grp_settings.Controls.Add(self.num_min_len)

        # Чекбоксы
        self.chk_copy_tags = CheckBox()
        self.chk_copy_tags.Text = "Переносить выноски"
        self.chk_copy_tags.Location = Point(10, 85)
        self.chk_copy_tags.Size = Size(170, 20)
        self.chk_copy_tags.Checked = True
        grp_settings.Controls.Add(self.chk_copy_tags)

        self.chk_copy_params = CheckBox()
        self.chk_copy_params.Text = "Копировать параметры"
        self.chk_copy_params.Location = Point(190, 85)
        self.chk_copy_params.Size = Size(170, 20)
        self.chk_copy_params.Checked = True
        grp_settings.Controls.Add(self.chk_copy_params)

        y += 140

        # --- Preview ---
        self.lbl_preview = Label()
        self.lbl_preview.Text = "Стержней к обрезке: 0"
        self.lbl_preview.Location = Point(10, y)
        self.lbl_preview.Size = Size(370, 20)
        self.lbl_preview.Font = Font(self.lbl_preview.Font, FontStyle.Bold)
        self.Controls.Add(self.lbl_preview)
        y += 30

        # --- Лог ---
        self.lbl_log = Label()
        self.lbl_log.Text = ""
        self.lbl_log.Location = Point(10, y)
        self.lbl_log.Size = Size(370, 40)
        self.Controls.Add(self.lbl_log)
        y += 50

        # --- Кнопки ---
        self.btn_cut = Button()
        self.btn_cut.Text = "Обрезать"
        self.btn_cut.Location = Point(100, y)
        self.btn_cut.Size = Size(90, 30)
        self.btn_cut.Click += self.on_cut
        self.Controls.Add(self.btn_cut)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(200, y)
        btn_close.Size = Size(90, 30)
        btn_close.Click += self.on_close
        self.Controls.Add(btn_close)

    def load_settings(self):
        """Загрузить последние настройки."""
        try:
            offset = get_setting("smart_openings.offset_mm", 50)
            min_len = get_setting("smart_openings.min_length_mm", 100)

            self.num_offset.Value = int(offset)
            self.num_min_len.Value = int(min_len)
        except Exception:
            return

    def save_settings(self):
        """Сохранить настройки."""
        try:
            set_setting("smart_openings.offset_mm", int(self.num_offset.Value))
            set_setting("smart_openings.min_length_mm", int(self.num_min_len.Value))
        except Exception:
            return

    def on_select_floor(self, sender, args):
        """Обработчик выбора плиты."""
        self.action = "select_floor"
        self.DialogResult = DialogResult.Retry
        self.Close()

    def on_select_openings(self, sender, args):
        """Обработчик выбора отверстий."""
        if self.selected_floor is None:
            show_warning("Внимание", "Сначала выберите плиту")
            return
        self.action = "select_openings"
        self.DialogResult = DialogResult.Retry
        self.Close()

    def on_cut(self, sender, args):
        """Обработчик кнопки обрезки."""
        if self.selected_floor is None:
            show_warning("Внимание", "Сначала выберите плиту")
            return
        if not self.selected_openings:
            show_warning("Внимание", "Выберите хотя бы одно отверстие")
            return

        self.save_settings()
        self.action = "cut"
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_close(self, sender, args):
        """Обработчик закрытия."""
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def update_floor_display(self):
        """Обновить отображение выбранной плиты."""
        if self.selected_floor is not None:
            try:
                type_id = self.selected_floor.GetTypeId()
                floor_type = doc.GetElement(type_id)
                name = floor_type.get_Parameter(
                    BuiltInParameter.SYMBOL_NAME_PARAM
                ).AsString() if floor_type else "Плита"
                self.txt_floor.Text = name
            except Exception:
                self.txt_floor.Text = "ID: {}".format(self.selected_floor.Id.IntegerValue)
                return
        else:
            self.txt_floor.Text = "[не выбрана]"

    def update_openings_display(self):
        """Обновить отображение выбранных отверстий."""
        count = len(self.selected_openings)
        self.txt_openings.Text = str(count)
        self.update_preview()

    def update_preview(self):
        """Обновить preview количества стержней."""
        if self.selected_floor is None or not self.selected_openings:
            self.lbl_preview.Text = "Стержней к обрезке: 0"
            return

        try:
            total = 0
            for opening in self.selected_openings:
                rebars = find_rebars_intersecting_opening(doc, opening, self.selected_floor)
                total += len(rebars)
            self.lbl_preview.Text = "Стержней к обрезке: {}".format(total)
        except Exception:
            self.lbl_preview.Text = "Ошибка подсчёта"
            return

    def restore_selection(self, floor_id, opening_ids):
        """Восстановить выбор после переоткрытия формы."""
        if floor_id is not None:
            try:
                self.selected_floor = doc.GetElement(floor_id)
                self.selected_floor_id = floor_id
                self.update_floor_display()
            except Exception:
                return

        if opening_ids:
            self.selected_openings = []
            self.selected_opening_ids = []
            for oid in opening_ids:
                try:
                    opening = doc.GetElement(oid)
                    if opening is not None:
                        self.selected_openings.append(opening)
                        self.selected_opening_ids.append(oid)
                except Exception:
                    continue
            self.update_openings_display()


# === ДИАЛОГ ВЫБОРА ОТВЕРСТИЙ ===

class OpeningSelectionForm(Form):
    """Форма для выбора отверстий из списка."""

    def __init__(self, openings_data):
        """
        Args:
            openings_data: list of dict from get_all_openings_in_floor
        """
        self.openings_data = openings_data
        self.selected_indices = []
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Выбор отверстий"
        self.Width = 450
        self.Height = 400
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 10

        # Инструкция
        lbl_info = Label()
        lbl_info.Text = "Выберите отверстия для обрезки арматуры:"
        lbl_info.Location = Point(10, y)
        lbl_info.Size = Size(420, 20)
        self.Controls.Add(lbl_info)
        y += 25

        # Список отверстий с чекбоксами
        self.checklist = CheckedListBox()
        self.checklist.Location = Point(10, y)
        self.checklist.Size = Size(420, 260)
        self.checklist.CheckOnClick = True

        for op_data in self.openings_data:
            self.checklist.Items.Add(op_data['name'], False)

        self.Controls.Add(self.checklist)
        y += 270

        # Кнопки Выбрать все / Снять все
        btn_select_all = Button()
        btn_select_all.Text = "Выбрать все"
        btn_select_all.Location = Point(10, y)
        btn_select_all.Size = Size(100, 25)
        btn_select_all.Click += self.on_select_all
        self.Controls.Add(btn_select_all)

        btn_deselect_all = Button()
        btn_deselect_all.Text = "Снять все"
        btn_deselect_all.Location = Point(120, y)
        btn_deselect_all.Size = Size(100, 25)
        btn_deselect_all.Click += self.on_deselect_all
        self.Controls.Add(btn_deselect_all)

        y += 35

        # Кнопки OK / Cancel
        btn_ok = Button()
        btn_ok.Text = "OK"
        btn_ok.Location = Point(140, y)
        btn_ok.Size = Size(80, 30)
        btn_ok.Click += self.on_ok
        self.Controls.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(230, y)
        btn_cancel.Size = Size(80, 30)
        btn_cancel.Click += self.on_cancel
        self.Controls.Add(btn_cancel)

    def on_select_all(self, sender, args):
        for i in range(self.checklist.Items.Count):
            self.checklist.SetItemChecked(i, True)

    def on_deselect_all(self, sender, args):
        for i in range(self.checklist.Items.Count):
            self.checklist.SetItemChecked(i, False)

    def on_ok(self, sender, args):
        self.selected_indices = []
        for i in range(self.checklist.Items.Count):
            if self.checklist.GetItemChecked(i):
                self.selected_indices.append(i)

        if not self.selected_indices:
            show_warning("Внимание", "Выберите хотя бы одно отверстие")
            return

        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


# === ГЛАВНАЯ ЛОГИКА ===

def cut_rebars_around_openings_data(floor, openings_data, offset_mm, min_length_mm, copy_tags, copy_params):
    """
    Обрезать арматуру вокруг отверстий (работает с dict данными).

    Args:
        floor: Floor element
        openings_data: list of dict from get_all_openings_in_floor
        offset_mm: Отступ от края в мм
        min_length_mm: Минимальная длина стержня в мм
        copy_tags: Переносить выноски
        copy_params: Копировать параметры

    Returns:
        tuple (created_count, deleted_count, errors)
    """
    import codecs
    import os

    # Путь к лог-файлу
    log_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))), "log_cut_rebar.log")

    def log(msg):
        """Записать в лог - игнорируем ошибки записи."""
        try:
            with codecs.open(log_path, 'a', 'utf-8') as f:
                f.write(msg + "\n")
        except IOError:
            return
        except OSError:
            return

    # Начало логирования
    log("\n" + "=" * 60)
    log("=== CUT REBARS START ===")
    log("Floor ID: {}".format(floor.Id.IntegerValue))
    log("Openings count: {}".format(len(openings_data)))
    log("Offset: {} mm, Min length: {} mm".format(offset_mm, min_length_mm))
    log("Copy params: {}, Copy tags: {}".format(copy_params, copy_tags))

    # === СТАТИСТИКА ДО ОБРЕЗКИ ===
    log("\n=== REBARS BEFORE CUTTING ===")
    rebars_before = get_rebars_in_host(doc, floor)
    log("Total rebar count BEFORE: {}".format(len(rebars_before)))

    total_length_before = 0.0
    for rb in rebars_before:
        curve = get_rebar_centerline(rb)
        if curve:
            length_mm = curve.Length * 304.8
            total_length_before += length_mm
            log("  Rebar ID={}: {:.0f} mm".format(rb.Id.IntegerValue, length_mm))
        else:
            log("  Rebar ID={}: NO CENTERLINE".format(rb.Id.IntegerValue))

    log("Total length BEFORE: {:.0f} mm".format(total_length_before))
    log("=" * 40)

    created_count = 0
    deleted_count = 0
    errors = []

    offset_feet = mm_to_feet(offset_mm)
    min_length_feet = mm_to_feet(min_length_mm)

    with Transaction(doc, "Обрезать арматуру") as t:
        t.Start()

        # Убедиться что параметр существует
        ensure_rebar_cut_param(doc, app)

        # Регенерация чтобы параметр стал доступен
        doc.Regenerate()

        for op_idx, op_data in enumerate(openings_data):
            opening_id = op_data['id']
            curves = op_data.get('curves', [])

            log("\n--- Opening {}: {} ---".format(op_idx + 1, op_data['name']))
            log("  Type: {}, ID: {}".format(op_data['type'], opening_id))
            log("  Curves count: {}".format(len(curves)))

            # Для семейств показать какой bbox использовался
            if op_data['type'] == 'family':
                bbox_source = op_data.get('bbox_source', 'unknown')
                elem_size = op_data.get('elem_size', (0, 0))
                curves_size = op_data.get('curves_size', (0, 0))
                log("  FAMILY bbox_source: {}".format(bbox_source))
                log("  FAMILY element size: {} x {} mm".format(int(elem_size[0]), int(elem_size[1])))
                log("  FAMILY curves (opening) size: {} x {} mm".format(int(curves_size[0]), int(curves_size[1])))

            if not curves:
                errors.append("Нет контура для: {}".format(op_data['name']))
                log("  ERROR: No curves!")
                continue

            # Получить 2D границы отверстия
            bounds = get_opening_2d_bounds(curves)
            if bounds[0] == float('inf'):
                errors.append("Не удалось получить границы для: {}".format(op_data['name']))
                log("  ERROR: Could not get bounds!")
                continue

            rect_min_x, rect_min_y, rect_max_x, rect_max_y = bounds
            log("  Bounds: X({:.2f} - {:.2f}) Y({:.2f} - {:.2f})".format(
                rect_min_x, rect_max_x, rect_min_y, rect_max_y))

            # Найти пересекающиеся стержни через 2D проверку
            rebars = find_rebars_intersecting_opening_2d(doc, curves, floor, debug_log=True)
            log("  Intersecting rebars: {}".format(len(rebars)))

            if not rebars:
                log("  No rebars to cut for this opening")
                continue

            for rebar in rebars:
                rebar_id = rebar.Id.IntegerValue
                log("\n  Processing Rebar ID={}".format(rebar_id))

                try:
                    # Получить centerline
                    curve = get_rebar_centerline(rebar)
                    if curve is None:
                        errors.append("Нет centerline для стержня {}".format(rebar_id))
                        log("    ERROR: No centerline!")
                        continue

                    p0 = curve.GetEndPoint(0)
                    p1 = curve.GetEndPoint(1)
                    log("    Centerline: ({:.2f},{:.2f},{:.2f}) - ({:.2f},{:.2f},{:.2f})".format(
                        p0.X, p0.Y, p0.Z, p1.X, p1.Y, p1.Z))
                    log("    Length: {:.2f} ft ({:.0f} mm)".format(curve.Length, curve.Length * 304.8))

                    # Разделить линию через 2D метод (основной)
                    parts = split_line_by_rect_2d(
                        curve, rect_min_x, rect_min_y, rect_max_x, rect_max_y, offset_feet
                    )
                    log("    split_line_by_rect_2d result: {}".format(
                        len(parts) if parts else "None"))

                    # Fallback: попробовать solid метод если 2D не сработал
                    if parts is None or len(parts) == 0:
                        log("    Trying solid fallback...")
                        solid = create_solid_from_curves(curves, floor)
                        if solid is not None:
                            parts = split_line_by_solid(curve, solid, offset_feet)
                            log("    split_line_by_solid result: {}".format(
                                len(parts) if parts else "None"))
                        else:
                            log("    Could not create solid from curves")

                    if parts is None or len(parts) == 0:
                        errors.append("Не удалось разделить стержень {}".format(rebar_id))
                        log("    ERROR: Could not split rebar!")
                        continue

                    # Логируем части
                    for part_name, part_curve in parts:
                        pp0 = part_curve.GetEndPoint(0)
                        pp1 = part_curve.GetEndPoint(1)
                        log("    Part '{}': ({:.2f},{:.2f}) - ({:.2f},{:.2f}), len={:.0f}mm".format(
                            part_name, pp0.X, pp0.Y, pp1.X, pp1.Y, part_curve.Length * 304.8))

                    # Получить существующие данные о разрезах
                    existing_guids = get_rebar_cut_data(rebar)
                    log("    Existing cut data: {}".format(existing_guids))

                    # Создать новые стержни
                    new_rebars = []
                    for part_name, part_curve in parts:
                        # Проверить минимальную длину
                        if part_curve.Length < min_length_feet:
                            log("    Skipping part '{}' - too short ({:.0f}mm < {:.0f}mm)".format(
                                part_name, part_curve.Length * 304.8, min_length_mm))
                            continue

                        # Логируем кривую которую передаём
                        pc0 = part_curve.GetEndPoint(0)
                        pc1 = part_curve.GetEndPoint(1)
                        log("    INPUT curve for '{}': ({:.4f},{:.4f})-({:.4f},{:.4f}) len={:.0f}mm".format(
                            part_name, pc0.X, pc0.Y, pc1.X, pc1.Y, part_curve.Length * 304.8))

                        new_rebar = create_rebar_from_curve(doc, rebar, part_curve, floor)
                        if new_rebar is not None:
                            new_rebars.append(new_rebar)
                            created_count += 1
                            log("    Created new rebar ID={} from part '{}'".format(
                                new_rebar.Id.IntegerValue, part_name))

                            # ВАЖНО: Регенерация СРАЗУ после создания каждого стержня!
                            doc.Regenerate()

                            # VERIFICATION: проверяем РЕАЛЬНУЮ геометрию созданного стержня
                            actual_curve = get_rebar_centerline(new_rebar)
                            if actual_curve:
                                actual_len_mm = actual_curve.Length * 304.8
                                input_len_mm = part_curve.Length * 304.8
                                diff_mm = abs(actual_len_mm - input_len_mm)
                                log("      VERIFY: INPUT={:.0f}mm, ACTUAL={:.0f}mm, DIFF={:.0f}mm".format(
                                    input_len_mm, actual_len_mm, diff_mm))
                                if diff_mm > 10:
                                    log("      !!! GEOMETRY MISMATCH !!! Revit ignoring input curve!")
                            else:
                                log("      VERIFY: Could not get actual centerline after Regenerate!")
                        else:
                            log("    FAILED to create rebar from part '{}'".format(part_name))

                    # Регенерация чтобы параметры стали доступны на новых стержнях
                    if new_rebars:
                        doc.Regenerate()

                    # СНАЧАЛА копировать параметры (если включено)
                    # НО: НЕ копируем если исходный стержень уже был обрезан ранее!
                    # Потому что копирование параметров ломает геометрию новых стержней
                    if copy_params and not existing_guids:
                        for new_rebar in new_rebars:
                            copy_instance_params(rebar, new_rebar, log)
                        log("    Copied params to {} new rebars".format(len(new_rebars)))

                        # VERIFY AFTER COPY: проверяем что геометрия не изменилась
                        doc.Regenerate()
                        for new_rebar in new_rebars:
                            after_curve = get_rebar_centerline(new_rebar)
                            if after_curve:
                                after_len = after_curve.Length * 304.8
                                log("      AFTER COPY VERIFY ID={}: {:.0f}mm".format(
                                    new_rebar.Id.IntegerValue, after_len))
                    elif copy_params and existing_guids:
                        log("    SKIP copy params - source already cut (existing_guids={}), would break geometry".format(
                            len(existing_guids)))

                    # Записать данные о разрезе (ПОСЛЕ копирования параметров!)
                    opening_guid = None
                    if op_data['type'] == 'element' and op_data.get('element') is not None:
                        opening_guid = op_data['element'].UniqueId
                    elif op_data['type'] == 'family' and op_data.get('element') is not None:
                        opening_guid = op_data['element'].UniqueId
                    elif op_data['type'] == 'sketch':
                        # Для sketch используем ID как строку
                        opening_guid = "sketch_{}".format(op_data['id'])

                    log("    Opening GUID: {}".format(opening_guid))

                    if opening_guid and new_rebars:
                        for new_rebar in new_rebars:
                            # Добавляем existing_guids + новый opening_guid
                            guids_to_write = list(existing_guids)
                            if opening_guid not in guids_to_write:
                                guids_to_write.append(opening_guid)
                            from cpsk_shared_params import set_rebar_cut_data
                            result = set_rebar_cut_data(new_rebar, guids_to_write)
                            if not result:
                                errors.append("Не удалось записать параметр для стержня {}".format(
                                    new_rebar.Id.IntegerValue))
                                log("    FAILED to write cut data to rebar {}".format(
                                    new_rebar.Id.IntegerValue))
                            else:
                                log("    Written cut data to rebar {}: {}".format(
                                    new_rebar.Id.IntegerValue, guids_to_write))

                    # Удалить исходный стержень
                    if new_rebars:
                        doc.Delete(rebar.Id)
                        deleted_count += 1
                        log("    Deleted original rebar ID={}".format(rebar_id))

                except Exception as e:
                    errors.append("Ошибка стержня {}: {}".format(
                        rebar.Id.IntegerValue, str(e)
                    ))
                    log("    EXCEPTION: {}".format(str(e)))
                    continue

            # ВАЖНО: Регенерация после каждого отверстия!
            # Это нужно чтобы новые стержни получили правильную геометрию
            # перед тем как их будет обрабатывать следующее отверстие
            doc.Regenerate()
            log("  [Regenerate after opening {}]".format(op_idx + 1))

        t.Commit()

    # === СТАТИСТИКА ПОСЛЕ ОБРЕЗКИ ===
    log("\n=== REBARS AFTER CUTTING ===")
    rebars_after = get_rebars_in_host(doc, floor)
    log("Total rebar count AFTER: {}".format(len(rebars_after)))

    total_length_after = 0.0
    for rb in rebars_after:
        curve = get_rebar_centerline(rb)
        if curve:
            length_mm = curve.Length * 304.8
            total_length_after += length_mm
            log("  Rebar ID={}: {:.0f} mm".format(rb.Id.IntegerValue, length_mm))
        else:
            log("  Rebar ID={}: NO CENTERLINE".format(rb.Id.IntegerValue))

    log("Total length AFTER: {:.0f} mm".format(total_length_after))
    log("=" * 40)

    # === СРАВНЕНИЕ ===
    log("\n=== COMPARISON ===")
    log("Rebars BEFORE: {}, AFTER: {}".format(len(rebars_before), len(rebars_after)))
    log("Total length BEFORE: {:.0f} mm".format(total_length_before))
    log("Total length AFTER: {:.0f} mm".format(total_length_after))
    length_diff = total_length_after - total_length_before
    log("Length DIFFERENCE: {:.0f} mm ({})".format(
        length_diff,
        "INCREASED!" if length_diff > 0 else ("DECREASED" if length_diff < 0 else "SAME")
    ))
    log("=" * 40)

    log("\n=== CUT REBARS FINISHED ===")
    log("Created: {}, Deleted: {}, Errors: {}".format(created_count, deleted_count, len(errors)))
    if errors:
        log("Errors:")
        for err in errors:
            log("  - {}".format(err))
    log("=" * 60)

    return created_count, deleted_count, errors


def copy_instance_params(source_rebar, target_rebar, log_func=None):
    """Копировать instance параметры (кроме геометрических и CPSK_RebarCutData)."""
    # Имена групп параметров которые НЕ нужно копировать (геометрия)
    # Проверяем по имени группы т.к. работает во всех версиях Revit
    SKIP_GROUP_NAMES = {
        'PG_GEOMETRY',           # Размеры / Dimensions
        'PG_CONSTRAINTS',        # Зависимости / Constraints
        'PG_REBAR_ARRAY',        # Набор арматурных стержней
        'PG_REBAR_SYSTEM_LAYERS', # Слои системы арматуры
    }

    copied_params = []
    skipped_params = []

    try:
        for param in source_rebar.Parameters:
            if param.IsReadOnly:
                continue
            if not param.HasValue:
                continue

            param_name = param.Definition.Name

            # Пропускаем наш параметр - он устанавливается отдельно
            if param_name == REBAR_CUT_DATA_PARAM:
                skipped_params.append("{} (our param)".format(param_name))
                continue

            # Пропускаем параметры из геометрических групп
            # Проверяем через GetGroupTypeId (Revit 2022+) или ParameterGroup (старые версии)
            skip_param = False
            definition = param.Definition
            group_info = ""

            # Попробуем новый API (Revit 2022+)
            if hasattr(definition, 'GetGroupTypeId'):
                group_type_id = definition.GetGroupTypeId()
                if group_type_id is not None:
                    group_name = str(group_type_id.TypeId) if hasattr(group_type_id, 'TypeId') else str(group_type_id)
                    group_info = group_name
                    for skip_name in SKIP_GROUP_NAMES:
                        if skip_name.lower() in group_name.lower():
                            skip_param = True
                            break

            # Fallback: старый API
            if not skip_param and hasattr(definition, 'ParameterGroup'):
                group_str = str(definition.ParameterGroup)
                group_info = group_str
                for skip_name in SKIP_GROUP_NAMES:
                    if skip_name in group_str:
                        skip_param = True
                        break

            if skip_param:
                skipped_params.append("{} ({})".format(param_name, group_info))
                continue

            target_param = target_rebar.LookupParameter(param_name)
            if target_param is None or target_param.IsReadOnly:
                continue

            copied_params.append("{} ({})".format(param_name, group_info))

            try:
                if param.StorageType == StorageType.String:
                    target_param.Set(param.AsString() or "")
                elif param.StorageType == StorageType.Double:
                    target_param.Set(param.AsDouble())
                elif param.StorageType == StorageType.Integer:
                    target_param.Set(param.AsInteger())
                elif param.StorageType == StorageType.ElementId:
                    target_param.Set(param.AsElementId())
            except Exception:
                continue

        # Логируем что было скопировано и пропущено
        if log_func:
            if skipped_params:
                log_func("      SKIPPED params: {}".format(", ".join(skipped_params[:5])))
                if len(skipped_params) > 5:
                    log_func("        ... and {} more".format(len(skipped_params) - 5))
            if copied_params:
                log_func("      COPIED params: {}".format(", ".join(copied_params[:5])))
                if len(copied_params) > 5:
                    log_func("        ... and {} more".format(len(copied_params) - 5))
    except Exception:
        return


def check_area_reinforcement(floor):
    """
    Проверить наличие Area Reinforcement в плите.

    Returns:
        list of AreaReinforcement или пустой список
    """
    return get_area_reinforcement_in_host(doc, floor)


def handle_area_reinforcement(floor, area_reinforcements):
    """
    Обработать Area Reinforcement - предложить конвертацию.

    Returns:
        bool: True если продолжать, False если отмена
    """
    from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon

    count = len(area_reinforcements)
    total_rebars = 0
    for ar in area_reinforcements:
        try:
            total_rebars += ar.GetRebarInSystemIds().Count
        except Exception:
            continue

    msg = "Обнаружено Area Reinforcement!\n\n"
    msg += "Зон: {}\n".format(count)
    msg += "Стержней в зонах: {}\n\n".format(total_rebars)
    msg += "Для обрезки арматуры необходимо конвертировать\n"
    msg += "Area Reinforcement в отдельные стержни.\n\n"
    msg += "Конвертировать?"

    result = MessageBox.Show(
        msg,
        "Area Reinforcement",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question
    )

    if result == DialogResult.Yes:
        # Конвертировать
        with Transaction(doc, "Конвертация Area Reinforcement") as t:
            t.Start()
            created = 0
            for ar in area_reinforcements:
                rebars = convert_area_reinforcement_to_rebars(doc, ar, floor)
                created += len(rebars)
            t.Commit()

        show_success("Конвертация", "Создано {} стержней".format(created))
        return True
    else:
        show_warning(
            "Конвертация обязательна",
            "Без конвертации Area Reinforcement в отдельные стержни обрезка арматуры невозможна."
        )
        return False


# === MAIN ===

def ensure_shared_parameter():
    """
    Проверить и создать параметр ФОП если нужно.

    Returns:
        bool: True если параметр готов к работе
    """
    # Проверить текущее состояние
    param_info = get_shared_param_info(doc, app)

    if param_info['is_bound']:
        # Параметр уже существует
        show_info(
            "Параметр ФОП",
            "Параметр {} готов к работе".format(REBAR_CUT_DATA_PARAM),
            details="Статус: Привязан к категории Structural Rebar\n"
                    "Файл ФОП: {}".format(param_info['file_path'])
        )
        return True

    # Нужно создать параметр
    show_info(
        "Создание параметра ФОП",
        "Параметр {} будет создан и привязан к арматуре".format(REBAR_CUT_DATA_PARAM),
        details="Файл ФОП: {}\n\n"
                "Параметр будет виден в свойствах арматуры.".format(param_info['file_path'])
    )

    # Создать в транзакции
    with Transaction(doc, "Создать параметр ФОП") as t:
        t.Start()
        success, message, was_created = ensure_rebar_cut_param_with_info(doc, app)
        t.Commit()

    if success:
        if was_created:
            show_success("Параметр ФОП создан", message)
        return True
    else:
        show_error("Ошибка", "Не удалось создать параметр ФОП", details=message)
        return False


def main():
    """Главная функция - упрощённый поток."""

    # ШАГ 0: Проверить/создать параметр ФОП
    if not ensure_shared_parameter():
        return

    # ШАГ 1: Выбор плиты
    try:
        show_info("Обрезка арматуры", "Выберите плиту/перекрытие для обрезки арматуры")

        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            "Выберите плиту"
        )
        floor = doc.GetElement(ref.ElementId)

        # Проверка категории
        if floor.Category is None:
            show_error("Ошибка", "Выбранный элемент не имеет категории")
            return

        cat_id = floor.Category.Id.IntegerValue
        valid_categories = [
            int(BuiltInCategory.OST_Floors),
            int(BuiltInCategory.OST_StructuralFoundation),
            int(BuiltInCategory.OST_Ceilings)
        ]

        if cat_id not in valid_categories:
            show_error("Ошибка", "Выбранный элемент не является плитой",
                       details="Категория: {} (ID: {})\n\nДопустимые: Перекрытия, Фундаменты, Потолки".format(
                           floor.Category.Name, cat_id))
            return

    except OperationCanceledException:
        return
    except Exception as e:
        show_error("Ошибка", "Не удалось выбрать плиту", details=str(e))
        return

    # Проверить Area Reinforcement
    area_reinfs = check_area_reinforcement(floor)
    if area_reinfs:
        if not handle_area_reinforcement(floor, area_reinfs):
            return

    # ШАГ 2: Найти все отверстия в плите
    all_openings = get_all_openings_in_floor(doc, floor)

    # DEBUG: показать полную информацию
    from cpsk_rebar_utils import (
        get_rebars_in_host, get_openings_in_floor, get_floor_sketch_openings,
        get_nested_family_openings
    )

    # Собрать статистику
    all_rebars = get_rebars_in_host(doc, floor)
    element_openings = get_openings_in_floor(doc, floor)
    sketch_openings = get_floor_sketch_openings(floor)
    family_openings = get_nested_family_openings(doc, floor)

    debug_info = "=== ПОЛНЫЙ DEBUG ЛОГ ===\n\n"
    debug_info += "ПЛИТА: ID={}\n".format(floor.Id.IntegerValue)
    floor_bbox = floor.get_BoundingBox(None)
    if floor_bbox:
        debug_info += "  BBox: X({:.2f}-{:.2f}) Y({:.2f}-{:.2f}) Z({:.2f}-{:.2f})\n\n".format(
            floor_bbox.Min.X, floor_bbox.Max.X,
            floor_bbox.Min.Y, floor_bbox.Max.Y,
            floor_bbox.Min.Z, floor_bbox.Max.Z
        )

    debug_info += "СТЕРЖНИ В ПЛИТЕ: {}\n".format(len(all_rebars))
    # Показать XY диапазон всех стержней
    if all_rebars:
        from cpsk_rebar_utils import get_rebar_centerline
        all_x = []
        all_y = []
        method_stats = {}  # Статистика по методам
        for rb in all_rebars[:5]:  # Первые 5 стержней для детального вывода
            debug_info += "  Стержень {}: тип={}\n".format(rb.Id.IntegerValue, type(rb).__name__)
            curve, method = get_rebar_centerline(rb, return_method=True)
            if curve:
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
                all_x.extend([p0.X, p1.X])
                all_y.extend([p0.Y, p1.Y])
                debug_info += "    XY: ({:.2f},{:.2f})-({:.2f},{:.2f}) [{}]\n".format(
                    p0.X, p0.Y, p1.X, p1.Y, method)
            else:
                debug_info += "    НЕТ ГЕОМЕТРИИ! [{}]\n".format(method)

        # Собрать статистику по всем стержням
        for rb in all_rebars:
            curve, method = get_rebar_centerline(rb, return_method=True)
            method_stats[method] = method_stats.get(method, 0) + 1
            if curve:
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
                all_x.extend([p0.X, p1.X])
                all_y.extend([p0.Y, p1.Y])

        if all_x:
            debug_info += "  Общий диапазон X: {:.2f} - {:.2f}\n".format(min(all_x), max(all_x))
            debug_info += "  Общий диапазон Y: {:.2f} - {:.2f}\n".format(min(all_y), max(all_y))

        # Показать статистику методов
        debug_info += "  Методы определения centerline:\n"
        for method, count in sorted(method_stats.items(), key=lambda x: -x[1]):
            debug_info += "    {}: {} стержней\n".format(method, count)
    debug_info += "\n"

    debug_info += "ОТВЕРСТИЯ (Opening элементы): {}\n".format(len(element_openings))
    debug_info += "ОТВЕРСТИЯ (Sketch контуры): {}\n".format(len(sketch_openings))
    debug_info += "ОТВЕРСТИЯ (Вложенные семейства): {}\n".format(len(family_openings))
    debug_info += "ВСЕГО ОТВЕРСТИЙ: {}\n\n".format(len(all_openings))

    # DEBUG: показать ВСЕ семейства пересекающие плиту
    from Autodesk.Revit.DB import FilteredElementCollector, FamilyInstance
    debug_info += "=== ВСЕ СЕМЕЙСТВА В ЗОНЕ ПЛИТЫ ===\n"
    try:
        all_families = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
        families_in_floor = []
        for fam in all_families:
            fam_bbox = fam.get_BoundingBox(None)
            if fam_bbox and floor_bbox:
                # Проверить пересечение XY
                if not (fam_bbox.Max.X < floor_bbox.Min.X or fam_bbox.Min.X > floor_bbox.Max.X or
                        fam_bbox.Max.Y < floor_bbox.Min.Y or fam_bbox.Min.Y > floor_bbox.Max.Y):
                    families_in_floor.append(fam)

        debug_info += "Найдено семейств в зоне плиты: {}\n".format(len(families_in_floor))
        for fam in families_in_floor[:10]:  # Первые 10
            fam_type = doc.GetElement(fam.GetTypeId())
            fam_name = fam_type.FamilyName if fam_type else "?"
            cat_name = fam.Category.Name if fam.Category else "?"
            debug_info += "  {} | {} | ID:{}\n".format(cat_name, fam_name, fam.Id.IntegerValue)
    except Exception as e:
        debug_info += "Ошибка поиска семейств: {}\n".format(str(e))
        show_warning("Debug", "Ошибка в debug: {}".format(str(e)))
    debug_info += "\n"

    debug_info += "=== ДЕТАЛИ ПО ОТВЕРСТИЯМ ===\n\n"
    total_intersecting = 0

    for i, op in enumerate(all_openings):
        debug_info += "{}. {}\n".format(i + 1, op['name'])
        debug_info += "   Тип: {}\n".format(op['type'])

        curves = op.get('curves', [])
        debug_info += "   Кривых в контуре: {}\n".format(len(curves))

        if curves:
            # Показать bounds контура
            bounds = get_opening_2d_bounds(curves)
            debug_info += "   Bounds XY: ({:.2f},{:.2f}) - ({:.2f},{:.2f})\n".format(
                bounds[0], bounds[1], bounds[2], bounds[3]
            )

            # Проверить пересечения (используется алгоритм Cohen-Sutherland)
            # Включаем debug_log для первого отверстия
            use_debug = (i == 0)
            intersecting = find_rebars_intersecting_opening_2d(doc, curves, floor, debug_log=use_debug)
            debug_info += "   ПЕРЕСЕКАЮЩИХ СТЕРЖНЕЙ: {} (метод: Cohen-Sutherland 2D)\n".format(len(intersecting))
            total_intersecting += len(intersecting)

            # Показать ID первых 3 пересекающих стержней с методами centerline
            if intersecting:
                rebar_details = []
                for r in intersecting[:3]:
                    curve, method = get_rebar_centerline(r, return_method=True)
                    rebar_details.append("{} [{}]".format(r.Id.IntegerValue, method))
                debug_info += "   Примеры: {}\n".format(", ".join(rebar_details))
        else:
            debug_info += "   ОШИБКА: Нет кривых!\n"

        debug_info += "\n"

    debug_info += "=== ИТОГО ===\n"
    debug_info += "Всего пересекающих стержней: {}\n".format(total_intersecting)
    debug_info += "Алгоритм пересечения: Cohen-Sutherland line-rect 2D\n"

    show_info("DEBUG: Полный лог", "Диагностика", details=debug_info)

    if not all_openings:
        show_warning("Нет отверстий",
                     "В выбранной плите не найдено отверстий.\n\n"
                     "Поддерживаются:\n"
                     "- Элементы Opening\n"
                     "- Отверстия в контуре плиты (Edit Boundary)")
        return

    # ШАГ 3: Показать диалог выбора отверстий
    sel_form = OpeningSelectionForm(all_openings)
    if sel_form.ShowDialog() != DialogResult.OK:
        return

    # Получить выбранные отверстия
    selected_openings = [all_openings[i] for i in sel_form.selected_indices]

    if not selected_openings:
        show_warning("Внимание", "Не выбрано ни одного отверстия")
        return

    # ШАГ 4: Показать форму настроек и выполнить обрезку
    form = CutRebarForm()
    form.selected_floor = floor
    form.selected_floor_id = floor.Id
    form.update_floor_display()

    # Установить selected_openings чтобы прошла валидация в on_cut
    form.selected_openings = selected_openings  # list of dict
    form.selected_opening_ids = [op['id'] for op in selected_openings]

    # Показать количество отверстий
    form.txt_openings.Text = str(len(selected_openings))

    # Скрыть кнопки выбора (уже выбрано)
    form.btn_select_floor.Enabled = False
    form.btn_select_openings.Enabled = False

    # Показать preview (количество стержней считается при обрезке)
    form.lbl_preview.Text = "Отверстий выбрано: {}".format(len(selected_openings))

    result = form.ShowDialog()

    if result != DialogResult.OK:
        return

    # ШАГ 5: Выполнить обрезку
    offset_mm = int(form.num_offset.Value)
    min_length_mm = int(form.num_min_len.Value)
    copy_tags = form.chk_copy_tags.Checked
    copy_params = form.chk_copy_params.Checked

    try:
        created, deleted, errors = cut_rebars_around_openings_data(
            floor, selected_openings, offset_mm, min_length_mm, copy_tags, copy_params
        )

        msg = "Создано стержней: {}\nУдалено стержней: {}".format(created, deleted)
        if errors:
            msg += "\n\nОшибки:\n" + "\n".join(errors[:5])
            if len(errors) > 5:
                msg += "\n... и ещё {}".format(len(errors) - 5)

        show_success("Готово", msg)

    except Exception as e:
        show_error("Ошибка", "Не удалось обрезать арматуру", details=str(e))


if __name__ == "__main__":
    main()
