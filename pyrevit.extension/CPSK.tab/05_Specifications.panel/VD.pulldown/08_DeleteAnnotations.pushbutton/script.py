#! python3
# -*- coding: utf-8 -*-
"""
Удаление аннотаций ВД.
Удаляет все семейства аннотаций CPSK_VD_F, размещенные на листах.
"""

__title__ = "Удалить\nаннотации"
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
    FilteredElementCollector, FamilyInstance, ViewSheet,
    BuiltInCategory, Transaction, ElementId
)

doc = revit.doc


class SheetItem(object):
    """Элемент списка листов."""
    def __init__(self, sheet, annotation_count):
        self.sheet = sheet
        self.annotation_count = annotation_count

    def __str__(self):
        return "{} - {} ({} аннотаций)".format(
            self.sheet.SheetNumber, self.sheet.Name, self.annotation_count)


class SheetSelectionDialog(Form):
    """Диалог выбора листов для удаления аннотаций."""

    def __init__(self, sheets_with_annotations):
        self.sheets_data = sheets_with_annotations
        self.selected_sheets = []
        self.setup_form()
        self.load_sheets()

    def setup_form(self):
        self.Text = "Удаление аннотаций CPSK_VD_F"
        self.Width = 550
        self.Height = 450
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        title_label = Label()
        title_label.Text = "Выберите листы для удаления аннотаций CPSK_VD_F:"
        title_label.Location = Point(12, 9)
        title_label.Size = Size(510, 40)
        title_label.Font = Font("Segoe UI", 9)

        self.sheet_list = CheckedListBox()
        self.sheet_list.Location = Point(12, 55)
        self.sheet_list.Size = Size(510, 300)
        self.sheet_list.CheckOnClick = True

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
        ok_btn.Text = "Удалить"
        ok_btn.Location = Point(366, 365)
        ok_btn.Size = Size(75, 23)
        ok_btn.Click += self.on_ok

        cancel_btn = Button()
        cancel_btn.Text = "Отмена"
        cancel_btn.Location = Point(447, 365)
        cancel_btn.Size = Size(75, 23)
        cancel_btn.DialogResult = DialogResult.Cancel

        self.Controls.Add(title_label)
        self.Controls.Add(self.sheet_list)
        self.Controls.Add(select_all_btn)
        self.Controls.Add(select_none_btn)
        self.Controls.Add(ok_btn)
        self.Controls.Add(cancel_btn)

    def load_sheets(self):
        self.sheet_list.Items.Clear()
        for sheet, annotations in self.sheets_data:
            item = SheetItem(sheet, len(annotations))
            self.sheet_list.Items.Add(item)
        # Выбрать все по умолчанию
        for i in range(self.sheet_list.Items.Count):
            self.sheet_list.SetItemChecked(i, True)

    def on_select_all(self, sender, args):
        for i in range(self.sheet_list.Items.Count):
            self.sheet_list.SetItemChecked(i, True)

    def on_select_none(self, sender, args):
        for i in range(self.sheet_list.Items.Count):
            self.sheet_list.SetItemChecked(i, False)

    def on_ok(self, sender, args):
        self.selected_sheets = []
        for i in range(self.sheet_list.Items.Count):
            if self.sheet_list.GetItemChecked(i):
                item = self.sheet_list.Items[i]
                self.selected_sheets.append(item.sheet)

        if not self.selected_sheets:
            show_warning("Предупреждение", "Выберите хотя бы один лист")
            return

        # Подсчитать аннотации
        total_count = 0
        for sheet in self.selected_sheets:
            for s, annotations in self.sheets_data:
                if s.Id == sheet.Id:
                    total_count += len(annotations)
                    break

        # Сохранить количество для отображения
        self.total_to_delete = total_count
        self.DialogResult = DialogResult.OK
        self.Close()


def find_cpsk_vd_annotations():
    """Найти все аннотации CPSK_VD_F на листах."""
    sheets_with_annotations = []

    # Получить все листы
    sheet_collector = FilteredElementCollector(doc).OfClass(ViewSheet)

    for sheet in sheet_collector:
        try:
            # Найти аннотации на листе
            annotation_collector = FilteredElementCollector(doc, sheet.Id)
            annotations = annotation_collector.OfClass(FamilyInstance).OfCategory(
                BuiltInCategory.OST_GenericAnnotation
            )

            cpsk_annotations = []
            for annot in annotations:
                try:
                    family_name = annot.Symbol.Family.Name
                    if family_name.startswith("CPSK_VD_F"):
                        cpsk_annotations.append(annot)
                except Exception:
                    pass

            if cpsk_annotations:
                sheets_with_annotations.append((sheet, cpsk_annotations))
        except Exception:
            pass

    return sheets_with_annotations


def delete_annotations_on_sheets(selected_sheets, sheets_data):
    """Удалить аннотации на выбранных листах."""
    deleted_count = 0
    errors = []

    # Собрать ID аннотаций для удаления
    ids_to_delete = []
    for sheet in selected_sheets:
        for s, annotations in sheets_data:
            if s.Id == sheet.Id:
                for annot in annotations:
                    ids_to_delete.append(annot.Id)
                break

    # Удалить аннотации
    for elem_id in ids_to_delete:
        try:
            doc.Delete(elem_id)
            deleted_count += 1
        except Exception as ex:
            errors.append("ID {}: {}".format(elem_id.IntegerValue, str(ex)))

    return deleted_count, errors


def main():
    """Основная функция."""
    # Найти все аннотации CPSK_VD_F
    sheets_with_annotations = find_cpsk_vd_annotations()

    if not sheets_with_annotations:
        show_info("Информация", "Не найдено аннотаций CPSK_VD_F на листах проекта")
        return

    # Подсчитать общее количество
    total_annotations = sum(len(annotations) for _, annotations in sheets_with_annotations)

    # Показать диалог выбора
    dialog = SheetSelectionDialog(sheets_with_annotations)
    result = dialog.ShowDialog()

    if result != DialogResult.OK:
        return

    selected_sheets = dialog.selected_sheets
    if not selected_sheets:
        return

    # Удалить аннотации
    with revit.Transaction("Удалить аннотации CPSK_VD_F"):
        deleted_count, errors = delete_annotations_on_sheets(selected_sheets, sheets_with_annotations)

        # Показать результат
        if deleted_count > 0:
            message = "Удалено {} аннотаций с {} листов".format(deleted_count, len(selected_sheets))
            if errors:
                details = "Ошибки:\n" + "\n".join(errors[:10])
                if len(errors) > 10:
                    details += "\n... и еще {} ошибок".format(len(errors) - 10)
                show_warning("Частично выполнено", message, details=details)
            else:
                show_success("Успех", message)
        else:
            if errors:
                show_error("Ошибка", "Не удалось удалить аннотации",
                          details="\n".join(errors[:10]))
            else:
                show_info("Информация", "Нет аннотаций для удаления")


if __name__ == "__main__":
    main()
