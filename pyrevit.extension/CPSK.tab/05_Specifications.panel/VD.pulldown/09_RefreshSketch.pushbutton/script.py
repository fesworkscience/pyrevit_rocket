#! python3
# -*- coding: utf-8 -*-
"""
Обновление эскизов ВД.
Удаляет существующие аннотации и пересоздает их заново.
"""

__title__ = "Обновить\nэскизы"
__author__ = "CPSK"

import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, CheckedListBox,
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
    FilteredElementCollector, ViewSchedule, FamilyInstance, ViewSheet,
    BuiltInCategory, Transaction, ScheduleSheetInstance, Family,
    FamilySymbol, SectionType, ScheduleFilter, ScheduleFilterType,
    ScheduleFieldType, XYZ, StorageType
)

doc = revit.doc


class SpecificationItem(object):
    """Элемент списка спецификаций."""
    def __init__(self, specification, annotation_count):
        self.specification = specification
        self.annotation_count = annotation_count

    def __str__(self):
        return "{} ({} аннотаций)".format(self.specification.Name, self.annotation_count)


class SpecificationSelectionDialog(Form):
    """Диалог выбора спецификаций для обновления."""

    def __init__(self, specs_with_annotations):
        self.specs_data = specs_with_annotations
        self.selected_specs = []
        self.setup_form()
        self.load_specs()

    def setup_form(self):
        self.Text = "Обновление эскизов ВД"
        self.Width = 550
        self.Height = 450
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        title_label = Label()
        title_label.Text = "Выберите спецификации CPSK_VD_ для обновления эскизов:\n(Существующие аннотации будут удалены и созданы заново)"
        title_label.Location = Point(12, 9)
        title_label.Size = Size(510, 40)
        title_label.Font = Font("Segoe UI", 9)

        self.spec_list = CheckedListBox()
        self.spec_list.Location = Point(12, 55)
        self.spec_list.Size = Size(510, 300)
        self.spec_list.CheckOnClick = True

        select_all_btn = Button()
        select_all_btn.Text = "Выбрать все"
        select_all_btn.Location = Point(12, 365)
        select_all_btn.Size = Size(100, 23)
        select_all_btn.Click += self.on_select_all

        select_none_btn = Button()
        select_none_btn.Text = "Снять все"
        select_none_btn.Location = Point(118, 365)
        select_none_btn.Size = Size(100, 23)
        select_none_btn.Click += self.on_select_none

        ok_btn = Button()
        ok_btn.Text = "Обновить"
        ok_btn.Location = Point(366, 365)
        ok_btn.Size = Size(75, 23)
        ok_btn.Click += self.on_ok

        cancel_btn = Button()
        cancel_btn.Text = "Отмена"
        cancel_btn.Location = Point(447, 365)
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
        for spec, count in self.specs_data:
            item = SpecificationItem(spec, count)
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


def find_annotations_for_schedule(schedule):
    """Найти аннотации, связанные со спецификацией."""
    annotations = []

    # Найти листы со спецификацией
    collector = FilteredElementCollector(doc).OfClass(ScheduleSheetInstance)
    for instance in collector:
        if instance.ScheduleId == schedule.Id:
            sheet = doc.GetElement(instance.OwnerViewId)
            if sheet and isinstance(sheet, ViewSheet):
                # Найти аннотации на листе
                annot_collector = FilteredElementCollector(doc, sheet.Id)
                annot_collector.OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_GenericAnnotation)
                for annot in annot_collector:
                    try:
                        family_name = annot.Symbol.Family.Name
                        if family_name.startswith("CPSK_VD_F"):
                            # Проверить параметр привязки к спецификации
                            param = annot.LookupParameter("CPSK_VD_ID_Спецификации")
                            if param and param.HasValue:
                                if param.StorageType == StorageType.Integer:
                                    if param.AsInteger() == schedule.Id.IntegerValue:
                                        annotations.append(annot)
                                elif param.StorageType == StorageType.String:
                                    if param.AsString() == str(schedule.Id.IntegerValue):
                                        annotations.append(annot)
                            else:
                                # Если параметра нет, добавляем все аннотации CPSK_VD_F
                                annotations.append(annot)
                    except Exception:
                        pass

    return annotations


def get_schedule_on_sheets(schedule):
    """Получить листы со спецификацией."""
    result = []
    collector = FilteredElementCollector(doc).OfClass(ScheduleSheetInstance)
    for instance in collector:
        if instance.ScheduleId == schedule.Id:
            sheet = doc.GetElement(instance.OwnerViewId)
            if sheet and isinstance(sheet, ViewSheet):
                result.append((sheet, instance))
    return result


def find_annotation_families():
    """Найти семейства аннотаций CPSK_VD_F."""
    families = {}
    collector = FilteredElementCollector(doc).OfClass(Family)
    for family in collector:
        if family.Name.startswith("CPSK_VD_F"):
            families[family.Name] = family
    return families


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


def find_matching_family_type(families, shape_name):
    """Найти подходящий типоразмер."""
    for family in families.values():
        symbol_ids = family.GetFamilySymbolIds()
        for symbol_id in symbol_ids:
            symbol = doc.GetElement(symbol_id)
            if symbol:
                sym_name = symbol.Name
                if shape_name in sym_name or sym_name in shape_name:
                    return symbol
    return None


def process_schedule(schedule, annotation_families):
    """Обработать спецификацию: удалить старые и создать новые аннотации."""
    deleted_count = 0
    created_count = 0

    # Удалить существующие аннотации
    existing_annotations = find_annotations_for_schedule(schedule)
    for annot in existing_annotations:
        try:
            doc.Delete(annot.Id)
            deleted_count += 1
        except Exception:
            pass

    # Получить элементы спецификации
    try:
        elem_collector = FilteredElementCollector(doc, schedule.Id)
        elements = list(elem_collector.ToElements())
    except Exception:
        elements = []

    if not elements:
        return deleted_count, created_count

    # Получить листы со спецификацией
    schedule_views = get_schedule_on_sheets(schedule)
    if not schedule_views:
        return deleted_count, created_count

    # Группировать элементы по форме
    elements_by_shape = {}
    for elem in elements:
        shape = get_element_shape(elem)
        if shape:
            if shape not in elements_by_shape:
                elements_by_shape[shape] = []
            elements_by_shape[shape].append(elem)

    # Разместить новые аннотации
    row_index = 0
    for shape_name in sorted(elements_by_shape.keys()):
        group_elements = elements_by_shape[shape_name]
        family_type = find_matching_family_type(annotation_families, shape_name)

        if family_type:
            for sheet, schedule_instance in schedule_views:
                try:
                    if not family_type.IsActive:
                        family_type.Activate()
                        doc.Regenerate()

                    bounds = schedule_instance.get_BoundingBox(sheet)
                    if bounds:
                        # Рассчитать позицию
                        table_data = schedule.GetTableData()
                        header_section = table_data.GetSectionData(SectionType.Header)
                        header_height = 0
                        for i in range(header_section.NumberOfRows):
                            header_height += header_section.GetRowHeight(i)

                        row_height = 24.0 / 304.8
                        placement_point = XYZ(
                            bounds.Min.X + 0.1,
                            bounds.Max.Y - header_height - (row_index * row_height),
                            0
                        )

                        instance = doc.Create.NewFamilyInstance(placement_point, family_type, sheet)
                        if instance:
                            # Установить параметр привязки
                            param = instance.LookupParameter("CPSK_VD_ID_Спецификации")
                            if param and not param.IsReadOnly:
                                if param.StorageType == StorageType.Integer:
                                    param.Set(schedule.Id.IntegerValue)
                                elif param.StorageType == StorageType.String:
                                    param.Set(str(schedule.Id.IntegerValue))
                            created_count += 1
                except Exception:
                    pass
            row_index += 1

    return deleted_count, created_count


def main():
    """Основная функция."""
    # Найти спецификации CPSK_VD_
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    cpsk_schedules = [s for s in collector if s.Name.startswith("CPSK_VD_")]
    cpsk_schedules.sort(key=lambda x: x.Name)

    if not cpsk_schedules:
        show_info("Информация", "В проекте не найдено спецификаций с префиксом CPSK_VD_")
        return

    # Подсчитать аннотации для каждой спецификации
    specs_with_counts = []
    for spec in cpsk_schedules:
        annotations = find_annotations_for_schedule(spec)
        specs_with_counts.append((spec, len(annotations)))

    # Показать диалог выбора
    dialog = SpecificationSelectionDialog(specs_with_counts)
    result = dialog.ShowDialog()

    if result != DialogResult.OK:
        return

    selected_specs = dialog.selected_specs
    if not selected_specs:
        return

    # Найти семейства аннотаций
    annotation_families = find_annotation_families()
    if not annotation_families:
        show_warning("Предупреждение", "Не найдено семейств аннотаций CPSK_VD_F")
        return

    # Обработать спецификации
    with revit.Transaction("Обновить эскизы ВД"):
        total_deleted = 0
        total_created = 0

        for spec in selected_specs:
            deleted, created = process_schedule(spec, annotation_families)
            total_deleted += deleted
            total_created += created

        # Показать результат
        message = "Обновление завершено:\n"
        message += "- Удалено аннотаций: {}\n".format(total_deleted)
        message += "- Создано аннотаций: {}".format(total_created)

        if total_created > 0 or total_deleted > 0:
            show_success("Успех", message)
        else:
            show_info("Информация", "Нет изменений")


if __name__ == "__main__":
    main()
