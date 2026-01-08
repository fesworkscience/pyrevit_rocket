# -*- coding: utf-8 -*-
"""
Заполнение ведомости эскизов.
Размещает аннотации семейств (CPSK_VD_F) на листах со спецификациями CPSK_VD_.
"""

__title__ = "Заполнить\nэскизы"
__author__ = "CPSK"

import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, CheckedListBox, Panel,
    FormStartPosition, FormBorderStyle, DialogResult
)
from System.Drawing import Point, Size, Font

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Проверка авторизации
from cpsk_auth import require_auth
if not require_auth():
    sys.exit()

from cpsk_notify import show_error, show_warning, show_success, show_info

from pyrevit import revit

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, ViewSheet, Transaction,
    ScheduleSheetInstance, Family, FamilySymbol, FamilyInstance,
    SectionType, ScheduleFilter, ScheduleFilterType, ScheduleFieldType,
    XYZ, BuiltInCategory, ElementId, StorageType
)

doc = revit.doc


class SpecificationItem(object):
    """Элемент списка спецификаций."""
    def __init__(self, specification, display_text):
        self.specification = specification
        self.display_text = display_text

    def __str__(self):
        return self.display_text


class SpecificationSelectionDialog(Form):
    """Диалог выбора спецификаций."""

    def __init__(self, specifications):
        self.all_specs = specifications
        self.selected_specs = []
        self.setup_form()
        self.load_specs()

    def setup_form(self):
        self.Text = "Выбор спецификаций для заполнения эскизов"
        self.Width = 500
        self.Height = 400
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        title_label = Label()
        title_label.Text = "Выберите спецификации CPSK_VD_ для размещения эскизов:"
        title_label.Location = Point(12, 9)
        title_label.Size = Size(460, 40)
        title_label.Font = Font("Segoe UI", 9)

        self.spec_list = CheckedListBox()
        self.spec_list.Location = Point(12, 55)
        self.spec_list.Size = Size(460, 260)
        self.spec_list.CheckOnClick = True

        select_all_btn = Button()
        select_all_btn.Text = "Выбрать все"
        select_all_btn.Location = Point(12, 325)
        select_all_btn.Size = Size(100, 23)
        select_all_btn.Click += self.on_select_all

        select_none_btn = Button()
        select_none_btn.Text = "Снять все"
        select_none_btn.Location = Point(118, 325)
        select_none_btn.Size = Size(100, 23)
        select_none_btn.Click += self.on_select_none

        ok_btn = Button()
        ok_btn.Text = "OK"
        ok_btn.Location = Point(316, 325)
        ok_btn.Size = Size(75, 23)
        ok_btn.Click += self.on_ok

        cancel_btn = Button()
        cancel_btn.Text = "Отмена"
        cancel_btn.Location = Point(397, 325)
        cancel_btn.Size = Size(75, 23)
        cancel_btn.DialogResult = DialogResult.Cancel

        self.Controls.Add(title_label)
        self.Controls.Add(self.spec_list)
        self.Controls.Add(select_all_btn)
        self.Controls.Add(select_none_btn)
        self.Controls.Add(ok_btn)
        self.Controls.Add(cancel_btn)

    def load_specs(self):
        self.spec_list.Items.Clear()
        for spec in self.all_specs:
            display_text = "{} (ID: {})".format(spec.Name, spec.Id.IntegerValue)
            item = SpecificationItem(spec, display_text)
            self.spec_list.Items.Add(item)
        for i in range(self.spec_list.Items.Count):
            self.spec_list.SetItemChecked(i, True)

    def on_select_all(self, sender, args):
        for i in range(self.spec_list.Items.Count):
            self.spec_list.SetItemChecked(i, True)

    def on_select_none(self, sender, args):
        for i in range(self.spec_list.Items.Count):
            self.spec_list.SetItemChecked(i, False)

    def on_ok(self, sender, args):
        self.selected_specs = []
        for i in range(self.spec_list.Items.Count):
            if self.spec_list.GetItemChecked(i):
                item = self.spec_list.Items[i]
                self.selected_specs.append(item.specification)

        if not self.selected_specs:
            show_warning("Предупреждение", "Выберите хотя бы одну спецификацию")
            return

        self.DialogResult = DialogResult.OK
        self.Close()

    def get_selected(self):
        return self.selected_specs


def adjust_column_widths(schedule):
    """Настройка ширины столбцов."""
    try:
        table_data = schedule.GetTableData()
        body_section = table_data.GetSectionData(SectionType.Body)
        num_columns = body_section.NumberOfColumns

        width_20mm = 20.0 / 304.8
        width_70mm = 70.0 / 304.8

        if num_columns >= 2:
            body_section.SetColumnWidth(0, width_20mm / 2)
            body_section.SetColumnWidth(1, width_20mm / 2)

        if num_columns >= 5:
            body_section.SetColumnWidth(3, width_70mm / 2)
            body_section.SetColumnWidth(4, width_70mm / 2)

        header_section = table_data.GetSectionData(SectionType.Header)
        if header_section.NumberOfColumns >= 5:
            header_section.SetColumnWidth(0, width_20mm / 2)
            header_section.SetColumnWidth(1, width_20mm / 2)
            header_section.SetColumnWidth(3, width_70mm / 2)
            header_section.SetColumnWidth(4, width_70mm / 2)
    except Exception:
        pass


def hide_all_rows(schedule_def):
    """Добавить фильтр скрытия."""
    field_id = None
    field_name = ""

    for i in range(schedule_def.GetFieldCount()):
        field = schedule_def.GetField(i)
        name = field.GetName()
        if name == "Расч_Метка основы":
            field_id = field.FieldId
            field_name = name
            break

    if field_id is None:
        for i in range(schedule_def.GetFieldCount()):
            field = schedule_def.GetField(i)
            if field.FieldType in [ScheduleFieldType.Instance, ScheduleFieldType.ElementType]:
                field_id = field.FieldId
                field_name = field.GetName()
                break

    if field_id is None:
        return None

    # Проверить, есть ли уже фильтр скрытия
    filters = schedule_def.GetFilters()
    for f in filters:
        if f.FilterType == ScheduleFilterType.Equal:
            try:
                if f.GetStringValue() == "##HIDE_ALL_ROWS##":
                    return field_name
            except Exception:
                pass

    # Добавить фильтр
    if len(filters) < 8:
        hide_filter = ScheduleFilter(field_id, ScheduleFilterType.Equal, "##HIDE_ALL_ROWS##")
        schedule_def.AddFilter(hide_filter)

    return field_name


def get_elements_from_schedule(schedule):
    """Получить элементы из спецификации."""
    try:
        collector = FilteredElementCollector(doc, schedule.Id)
        return list(collector.ToElements())
    except Exception:
        return []


def find_annotation_families():
    """Найти семейства аннотаций CPSK_VD_F."""
    families = {}
    collector = FilteredElementCollector(doc).OfClass(Family)
    for family in collector:
        if family.Name.startswith("CPSK_VD_F"):
            families[family.Name] = family
    return families


def get_schedule_on_sheets(schedule):
    """Получить листы, на которых размещена спецификация."""
    result = []
    collector = FilteredElementCollector(doc).OfClass(ScheduleSheetInstance)
    for instance in collector:
        if instance.ScheduleId == schedule.Id:
            sheet = doc.GetElement(instance.OwnerViewId)
            if sheet and isinstance(sheet, ViewSheet):
                result.append((sheet, instance))
    return result


def get_element_shape(element):
    """Получить форму элемента."""
    try:
        param = element.LookupParameter("Форма")
        if param and param.HasValue:
            if param.StorageType == StorageType.String:
                return param.AsString() or ""
            return param.AsValueString() or ""
    except Exception:
        pass
    return ""


def get_element_dimensions(element):
    """Получить размеры элемента."""
    try:
        param_a = element.LookupParameter("ADSK_A") or element.LookupParameter("A")
        param_b = element.LookupParameter("ADSK_B") or element.LookupParameter("B")
        param_c = element.LookupParameter("ADSK_C") or element.LookupParameter("C")

        def get_value(param):
            if param and param.HasValue:
                if param.StorageType == StorageType.Integer:
                    return str(param.AsInteger())
                elif param.StorageType == StorageType.Double:
                    val = param.AsValueString()
                    return val if val else str(param.AsDouble())
                return param.AsString() or "?"
            return "?"

        return "{}x{}x{}".format(get_value(param_a), get_value(param_b), get_value(param_c))
    except Exception:
        return "?x?x?"


def get_element_composite_key(element):
    """Получить составной ключ элемента."""
    parts = []
    shape = get_element_shape(element)
    if shape:
        parts.append("Shape:{}".format(shape))

    dims = get_element_dimensions(element)
    if dims != "?x?x?":
        parts.append("Dims:{}".format(dims))

    # Марка
    mark_param = element.LookupParameter("Марка") or element.LookupParameter("ADSK_Марка")
    if mark_param and mark_param.HasValue:
        mark = mark_param.AsString() if mark_param.StorageType == StorageType.String else mark_param.AsValueString()
        if mark:
            parts.append("Mark:{}".format(mark))

    return "|".join(parts)


def find_matching_family_type(families, shape_name):
    """Найти подходящий типоразмер семейства."""
    for family in families.values():
        symbol_ids = family.GetFamilySymbolIds()
        for symbol_id in symbol_ids:
            symbol = doc.GetElement(symbol_id)
            if symbol:
                sym_name = symbol.Name
                if shape_name in sym_name or sym_name in shape_name or sym_name.lower() == shape_name.lower():
                    return symbol
    return None


def transfer_parameters(family_instance, source_element, schedule):
    """Передать параметры от элемента к семейству."""
    try:
        # Установить ID спецификации
        spec_param = family_instance.LookupParameter("CPSK_VD_ID_Спецификации")
        if spec_param and not spec_param.IsReadOnly:
            if spec_param.StorageType == StorageType.Integer:
                spec_param.Set(schedule.Id.IntegerValue)
            elif spec_param.StorageType == StorageType.String:
                spec_param.Set(str(schedule.Id.IntegerValue))

        # Передать параметры от источника
        source_params = {}
        for param in source_element.Parameters:
            if param.HasValue:
                name = param.Definition.Name
                if param.StorageType == StorageType.String:
                    source_params[name] = param.AsString()
                elif param.StorageType == StorageType.Integer:
                    source_params[name] = param.AsInteger()
                elif param.StorageType == StorageType.Double:
                    source_params[name] = param.AsDouble()

        # Установить соответствующие параметры
        for target_param in family_instance.Parameters:
            if target_param.IsReadOnly:
                continue
            target_name = target_param.Definition.Name
            if target_name in source_params:
                try:
                    value = source_params[target_name]
                    if target_param.StorageType == StorageType.String:
                        target_param.Set(str(value) if value else "")
                    elif target_param.StorageType == StorageType.Integer and isinstance(value, int):
                        target_param.Set(value)
                    elif target_param.StorageType == StorageType.Double and isinstance(value, float):
                        target_param.Set(value)
                except Exception:
                    pass
    except Exception:
        pass


def place_family_instance(sheet, schedule_instance, family_symbol, row_index, elements, schedule):
    """Разместить экземпляр семейства на листе."""
    try:
        if not family_symbol.IsActive:
            family_symbol.Activate()
            doc.Regenerate()

        # Получить границы спецификации
        bounds = schedule_instance.get_BoundingBox(sheet)
        if not bounds:
            return None

        # Получить данные таблицы
        current_schedule = doc.GetElement(schedule_instance.ScheduleId)
        table_data = current_schedule.GetTableData()
        header_section = table_data.GetSectionData(SectionType.Header)

        header_height = 0
        for i in range(header_section.NumberOfRows):
            header_height += header_section.GetRowHeight(i)

        # Рассчитать точку размещения
        viewport_width = bounds.Max.X - bounds.Min.X
        viewport_width_mm = viewport_width * 304.8
        centering_offset_mm = (viewport_width_mm - 90.0) / 2.0
        centering_offset = centering_offset_mm / 304.8

        additional_offset = 8.493 / 304.8
        base_offset_y = -(header_height / 2.0) - additional_offset
        row_height = 24.0 / 304.8

        placement_point = XYZ(
            bounds.Min.X + centering_offset,
            bounds.Max.Y - header_height + base_offset_y - (row_index * row_height),
            0
        )

        # Создать экземпляр семейства
        instance = doc.Create.NewFamilyInstance(placement_point, family_symbol, sheet)

        if instance and elements:
            transfer_parameters(instance, elements[0], schedule)

        return instance

    except Exception as ex:
        show_warning("Ошибка размещения", "Не удалось разместить семейство: {}".format(str(ex)))
        return None


def process_schedule(schedule):
    """Обработать спецификацию."""
    schedule_def = schedule.Definition

    # Настроить ширину столбцов
    adjust_column_widths(schedule)

    # Получить элементы
    elements = get_elements_from_schedule(schedule)
    if not elements:
        return 0

    # Скрыть строки
    hide_all_rows(schedule_def)

    # Найти семейства аннотаций
    annotation_families = find_annotation_families()
    if not annotation_families:
        show_warning("Предупреждение", "Не найдено семейств аннотаций с префиксом CPSK_VD_F")
        return 0

    # Получить листы со спецификацией
    schedule_views = get_schedule_on_sheets(schedule)
    if not schedule_views:
        show_warning("Предупреждение", "Спецификация '{}' не размещена ни на одном листе".format(schedule.Name))
        return 0

    # Группировать элементы по составному ключу
    elements_by_key = {}
    for elem in elements:
        key = get_element_composite_key(elem)
        if key:
            if key not in elements_by_key:
                elements_by_key[key] = []
            elements_by_key[key].append(elem)

    # Разместить семейства
    placed_count = 0
    row_index = 0

    for key in sorted(elements_by_key.keys()):
        group_elements = elements_by_key[key]
        shape = get_element_shape(group_elements[0])

        family_type = find_matching_family_type(annotation_families, shape)
        if family_type:
            for sheet, schedule_instance in schedule_views:
                instance = place_family_instance(
                    sheet, schedule_instance, family_type,
                    row_index, group_elements, schedule
                )
                if instance:
                    placed_count += 1
            row_index += 1

    return placed_count


def main():
    """Основная функция."""
    # Найти спецификации CPSK_VD_
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    cpsk_schedules = [s for s in collector if s.Name.startswith("CPSK_VD_")]
    cpsk_schedules.sort(key=lambda x: x.Name)

    if not cpsk_schedules:
        show_info("Информация", "В проекте не найдено спецификаций с префиксом CPSK_VD_")
        return

    # Показать диалог выбора
    dialog = SpecificationSelectionDialog(cpsk_schedules)
    result = dialog.ShowDialog()

    if result != DialogResult.OK:
        return

    selected = dialog.get_selected()
    if not selected:
        return

    # Обработать спецификации
    with revit.Transaction("Заполнить ведомость эскизов"):
        total_placed = 0
        processed_count = 0
        errors = []

        for schedule in selected:
            try:
                placed = process_schedule(schedule)
                total_placed += placed
                if placed > 0:
                    processed_count += 1
            except Exception as ex:
                errors.append("{}: {}".format(schedule.Name, str(ex)))

        # Показать результат
        if processed_count > 0:
            message = "Обработано {} спецификаций, размещено {} аннотаций".format(processed_count, total_placed)
            if errors:
                details = "Ошибки:\n" + "\n".join(errors)
                show_warning("Частично выполнено", message, details=details)
            else:
                show_success("Успех", message)
        else:
            if errors:
                show_error("Ошибка", "Не удалось обработать спецификации",
                          details="\n".join(errors))
            else:
                show_info("Информация", "Нет элементов для обработки или не найдены подходящие семейства аннотаций")


if __name__ == "__main__":
    main()
