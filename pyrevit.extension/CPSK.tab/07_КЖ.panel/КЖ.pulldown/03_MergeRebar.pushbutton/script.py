# -*- coding: utf-8 -*-
"""
Склеить арматуру - склейка стержней и удаление отверстия.

Workflow:
1. Выбрать плиту
2. Выбрать отверстия из списка (включая вложенные семейства)
3. Стержни склеиваются автоматически
4. Отверстия удаляются (для sketch-контуров - предупреждение)
"""

__title__ = "Склеить\nарматуру"
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
    Form, Label, Button, Panel, TextBox, GroupBox, CheckBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, AnchorStyles, ScrollBars, CheckedListBox
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
    can_merge_rebars, merge_rebars, get_rebar_endpoints,
    get_rebar_centerline, mm_to_feet, feet_to_mm,
    get_all_openings_in_floor, find_rebars_intersecting_opening_2d,
    get_rebars_in_host
)
from cpsk_shared_params import (
    find_rebars_by_opening_guid, remove_opening_from_rebar,
    get_rebar_cut_data, set_rebar_cut_data, ensure_rebar_cut_param,
    REBAR_CUT_DATA_PARAM, get_shared_param_info, ensure_rebar_cut_param_with_info
)

# 7. Настройки
doc = revit.doc
uidoc = revit.uidoc
app = doc.Application


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
        self.Text = "Выбор отверстий для склейки"
        self.Width = 500
        self.Height = 450
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 10

        # Инструкция
        lbl_info = Label()
        lbl_info.Text = "Выберите отверстия для склейки арматуры.\nСтержни будут объединены, отверстия удалены."
        lbl_info.Location = Point(10, y)
        lbl_info.Size = Size(470, 35)
        self.Controls.Add(lbl_info)
        y += 40

        # Предупреждение о sketch-контурах
        has_sketch = any(op['type'] == 'sketch' for op in self.openings_data)
        if has_sketch:
            lbl_warn = Label()
            lbl_warn.Text = "! Sketch-контуры нужно удалить вручную (Edit Boundary)"
            lbl_warn.Location = Point(10, y)
            lbl_warn.Size = Size(470, 20)
            lbl_warn.Font = Font(lbl_warn.Font, FontStyle.Italic)
            self.Controls.Add(lbl_warn)
            y += 25

        # Список отверстий с чекбоксами
        self.checklist = CheckedListBox()
        self.checklist.Location = Point(10, y)
        self.checklist.Size = Size(470, 260)
        self.checklist.CheckOnClick = True

        for op_data in self.openings_data:
            # Добавить информацию о типе
            type_suffix = ""
            if op_data['type'] == 'sketch':
                type_suffix = " [SKETCH - удалить вручную]"
            elif op_data['type'] == 'family':
                type_suffix = " [Семейство]"
            elif op_data['type'] == 'element':
                type_suffix = " [Opening]"

            # Показать количество связанных стержней
            rebar_count = 0
            if op_data['type'] == 'element' and op_data.get('element'):
                rebars = find_rebars_by_opening_guid(doc, op_data['element'].UniqueId)
                rebar_count = len(rebars)

            name = "{}{}".format(op_data['name'], type_suffix)
            if rebar_count > 0:
                name += " - {} стержней".format(rebar_count)

            self.checklist.Items.Add(name, False)

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
        btn_ok.Text = "Склеить"
        btn_ok.Location = Point(160, y)
        btn_ok.Size = Size(80, 30)
        btn_ok.Click += self.on_ok
        self.Controls.Add(btn_ok)

        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(250, y)
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

def try_delete_element(doc, element):
    """Попытаться удалить элемент. Возвращает None при успехе, текст ошибки при неудаче."""
    try:
        doc.Delete(element.Id)
        return None
    except Exception as e:
        return "Не удалось удалить: {}".format(str(e))


def find_rebar_pairs_for_merging(rebars, tolerance_feet):
    """
    Найти пары стержней которые можно склеить.

    Args:
        rebars: list of Rebar elements
        tolerance_feet: допуск расстояния между концами

    Returns:
        list of (rebar1, rebar2, merge_info) tuples
    """
    pairs = []
    used_ids = set()

    for i, rebar1 in enumerate(rebars):
        if rebar1.Id.IntegerValue in used_ids:
            continue

        for j, rebar2 in enumerate(rebars):
            if i >= j:
                continue
            if rebar2.Id.IntegerValue in used_ids:
                continue

            can_merge, info = can_merge_rebars(rebar1, rebar2, tolerance_feet)

            if can_merge:
                pairs.append((rebar1, rebar2, info))
                used_ids.add(rebar1.Id.IntegerValue)
                used_ids.add(rebar2.Id.IntegerValue)
                break

    return pairs


def merge_rebars_for_opening_data(op_data, floor):
    """
    Склеить стержни для данного отверстия.

    Args:
        op_data: dict from get_all_openings_in_floor
        floor: Floor element (host)

    Returns:
        tuple (merged_count, warning_count, deleted_opening, error_message)
    """
    merged_count = 0
    warning_count = 0
    deleted_opening = False
    errors = []

    opening_type = op_data['type']
    opening_id = op_data['id']
    curves = op_data.get('curves', [])

    # Для element-based openings используем GUID для поиска связанных стержней
    if opening_type == 'element' and op_data.get('element'):
        opening = op_data['element']
        opening_guid = opening.UniqueId

        # Найти стержни по GUID (которые были обрезаны этим отверстием)
        rebars = find_rebars_by_opening_guid(doc, opening_guid)

        if not rebars:
            # Попробовать найти по геометрии
            if curves:
                rebars = find_rebars_intersecting_opening_2d(doc, curves, floor)

        if not rebars:
            return 0, 0, False, "Нет стержней для склейки"

        tolerance_feet = mm_to_feet(100)  # 100mm tolerance для склейки
        pairs = find_rebar_pairs_for_merging(rebars, tolerance_feet)

        if pairs:
            for rebar1, rebar2, info in pairs:
                new_rebar = merge_rebars(doc, rebar1, rebar2, floor)

                if new_rebar is not None:
                    # Объединить списки cut data
                    guids1 = get_rebar_cut_data(rebar1)
                    guids2 = get_rebar_cut_data(rebar2)
                    merged_guids = list(set(guids1 + guids2))

                    # Удалить текущее отверстие из списка
                    if opening_guid in merged_guids:
                        merged_guids.remove(opening_guid)

                    # Записать в новый стержень
                    if merged_guids:
                        set_rebar_cut_data(new_rebar, merged_guids)

                    # Копировать параметры
                    line1 = get_rebar_centerline(rebar1)
                    line2 = get_rebar_centerline(rebar2)
                    if line1 and line2:
                        source = rebar1 if line1.Length >= line2.Length else rebar2
                        copy_instance_params(source, new_rebar)

                    # Удалить старые стержни
                    doc.Delete(rebar1.Id)
                    doc.Delete(rebar2.Id)

                    merged_count += 1

        # Стержни без пары - просто убрать отверстие из списка
        processed_ids = set()
        for r1, r2, _ in pairs:
            processed_ids.add(r1.Id.IntegerValue)
            processed_ids.add(r2.Id.IntegerValue)

        for rebar in rebars:
            if rebar.Id.IntegerValue not in processed_ids:
                remove_opening_from_rebar(rebar, opening_guid)
                warning_count += 1

        # Удалить отверстие
        delete_error = try_delete_element(doc, opening)
        if delete_error:
            errors.append(delete_error)
        else:
            deleted_opening = True

    elif opening_type == 'family' and op_data.get('element'):
        # Вложенное семейство - найти пересекающиеся стержни и попробовать склеить
        family_elem = op_data['element']

        if curves:
            rebars = find_rebars_intersecting_opening_2d(doc, curves, floor)

            if rebars:
                tolerance_feet = mm_to_feet(100)
                pairs = find_rebar_pairs_for_merging(rebars, tolerance_feet)

                for rebar1, rebar2, info in pairs:
                    new_rebar = merge_rebars(doc, rebar1, rebar2, floor)

                    if new_rebar is not None:
                        # Копировать параметры
                        line1 = get_rebar_centerline(rebar1)
                        line2 = get_rebar_centerline(rebar2)
                        if line1 and line2:
                            source = rebar1 if line1.Length >= line2.Length else rebar2
                            copy_instance_params(source, new_rebar)

                        doc.Delete(rebar1.Id)
                        doc.Delete(rebar2.Id)
                        merged_count += 1

        # Попытаться удалить семейство
        delete_error = try_delete_element(doc, family_elem)
        if delete_error:
            errors.append(delete_error)
        else:
            deleted_opening = True

    elif opening_type == 'sketch':
        # Sketch-based отверстие - нельзя удалить программно
        # Можно только склеить стержни
        if curves:
            rebars = find_rebars_intersecting_opening_2d(doc, curves, floor)

            if rebars:
                tolerance_feet = mm_to_feet(100)
                pairs = find_rebar_pairs_for_merging(rebars, tolerance_feet)

                for rebar1, rebar2, info in pairs:
                    new_rebar = merge_rebars(doc, rebar1, rebar2, floor)

                    if new_rebar is not None:
                        line1 = get_rebar_centerline(rebar1)
                        line2 = get_rebar_centerline(rebar2)
                        if line1 and line2:
                            source = rebar1 if line1.Length >= line2.Length else rebar2
                            copy_instance_params(source, new_rebar)

                        doc.Delete(rebar1.Id)
                        doc.Delete(rebar2.Id)
                        merged_count += 1

        errors.append("Sketch-контур нужно удалить вручную (Edit Boundary)")

    error_msg = "\n".join(errors) if errors else None
    return merged_count, warning_count, deleted_opening, error_msg


def copy_instance_params(source_rebar, target_rebar):
    """Копировать instance параметры."""
    try:
        for param in source_rebar.Parameters:
            if param.IsReadOnly:
                continue
            if not param.HasValue:
                continue
            # Пропустить системный параметр CPSK_RebarCutData
            if param.Definition.Name == REBAR_CUT_DATA_PARAM:
                continue

            target_param = target_rebar.LookupParameter(param.Definition.Name)
            if target_param is None or target_param.IsReadOnly:
                continue

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
    except Exception:
        return


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
        show_info("Склейка арматуры", "Выберите плиту/перекрытие для склейки арматуры")

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

    # ШАГ 2: Найти все отверстия в плите
    all_openings = get_all_openings_in_floor(doc, floor)

    if not all_openings:
        show_warning("Нет отверстий",
                     "В выбранной плите не найдено отверстий.\n\n"
                     "Поддерживаются:\n"
                     "- Элементы Opening\n"
                     "- Отверстия в контуре плиты (Sketch)\n"
                     "- Вложенные семейства с отверстиями")
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

    # ШАГ 4: Выполнить склейку
    total_merged = 0
    total_warnings = 0
    total_deleted = 0
    all_errors = []
    sketch_warnings = []

    with Transaction(doc, "Склеить арматуру") as t:
        t.Start()

        # Убедиться что параметр существует
        ensure_rebar_cut_param(doc, app)

        for op_data in selected_openings:
            try:
                merged, warnings, deleted, error = merge_rebars_for_opening_data(op_data, floor)

                total_merged += merged
                total_warnings += warnings
                if deleted:
                    total_deleted += 1

                if error:
                    if "Sketch-контур" in error:
                        sketch_warnings.append(op_data['name'])
                    else:
                        all_errors.append("{}: {}".format(op_data['name'], error))

            except Exception as e:
                error_msg = "{}: {}".format(op_data['name'], str(e))
                all_errors.append(error_msg)
                continue  # Продолжаем обработку остальных отверстий

        t.Commit()

    # ШАГ 5: Показать результат
    msg = "Склеено пар стержней: {}\n".format(total_merged)
    msg += "Удалено отверстий: {}\n".format(total_deleted)

    if total_warnings > 0:
        msg += "Стержней без пары: {}\n".format(total_warnings)

    if sketch_warnings:
        msg += "\nSketch-контуры (удалить вручную через Edit Boundary):\n"
        for name in sketch_warnings:
            msg += "  - {}\n".format(name)

    if all_errors:
        msg += "\nОшибки:\n"
        for err in all_errors[:5]:
            msg += "  - {}\n".format(err)
        if len(all_errors) > 5:
            msg += "  ... и ещё {}\n".format(len(all_errors) - 5)

    if total_merged > 0 or total_deleted > 0:
        show_success("Готово", msg)
    else:
        show_warning("Результат", msg)


if __name__ == "__main__":
    main()
