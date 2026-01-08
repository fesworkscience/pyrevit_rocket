#! python3
# -*- coding: utf-8 -*-
"""
Размещение аннотации CPSK_VD_F на листе.
Позволяет выбрать типоразмер и координаты для размещения.
"""

__title__ = "Разместить\nаннотацию"
__author__ = "CPSK"

import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, ComboBox, TextBox, GroupBox,
    FormStartPosition, FormBorderStyle, DialogResult,
    ComboBoxStyle, NumericUpDown
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
    FilteredElementCollector, Family, FamilySymbol, ViewSheet,
    BuiltInCategory, Transaction, XYZ
)

doc = revit.doc
uidoc = revit.uidoc


class FamilyTypeItem(object):
    """Элемент списка типоразмеров."""
    def __init__(self, family_symbol):
        self.family_symbol = family_symbol
        self.family_name = family_symbol.Family.Name
        self.type_name = family_symbol.Name

    def __str__(self):
        return "{} : {}".format(self.family_name, self.type_name)


class PlaceAnnotationDialog(Form):
    """Диалог размещения аннотации."""

    def __init__(self, family_types):
        self.family_types = family_types
        self.selected_type = None
        self.x_coord = 0.0
        self.y_coord = 0.0
        self.setup_form()
        self.load_types()

    def setup_form(self):
        self.Text = "Разместить аннотацию CPSK_VD_F"
        self.Width = 450
        self.Height = 280
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        # Выбор типоразмера
        type_label = Label()
        type_label.Text = "Типоразмер семейства:"
        type_label.Location = Point(15, 20)
        type_label.Size = Size(200, 20)

        self.type_combo = ComboBox()
        self.type_combo.Location = Point(15, 45)
        self.type_combo.Size = Size(405, 25)
        self.type_combo.DropDownStyle = ComboBoxStyle.DropDownList

        # Координаты
        coord_group = GroupBox()
        coord_group.Text = "Координаты размещения (мм)"
        coord_group.Location = Point(15, 85)
        coord_group.Size = Size(405, 100)

        x_label = Label()
        x_label.Text = "X:"
        x_label.Location = Point(15, 30)
        x_label.Size = Size(30, 20)

        self.x_input = NumericUpDown()
        self.x_input.Location = Point(50, 28)
        self.x_input.Size = Size(150, 25)
        self.x_input.Minimum = System.Decimal(-100000)
        self.x_input.Maximum = System.Decimal(100000)
        self.x_input.DecimalPlaces = 1
        self.x_input.Value = System.Decimal(0)

        y_label = Label()
        y_label.Text = "Y:"
        y_label.Location = Point(220, 30)
        y_label.Size = Size(30, 20)

        self.y_input = NumericUpDown()
        self.y_input.Location = Point(255, 28)
        self.y_input.Size = Size(130, 25)
        self.y_input.Minimum = System.Decimal(-100000)
        self.y_input.Maximum = System.Decimal(100000)
        self.y_input.DecimalPlaces = 1
        self.y_input.Value = System.Decimal(0)

        hint_label = Label()
        hint_label.Text = "Координаты относительно начала листа"
        hint_label.Location = Point(15, 65)
        hint_label.Size = Size(370, 20)

        coord_group.Controls.Add(x_label)
        coord_group.Controls.Add(self.x_input)
        coord_group.Controls.Add(y_label)
        coord_group.Controls.Add(self.y_input)
        coord_group.Controls.Add(hint_label)

        # Кнопки
        ok_btn = Button()
        ok_btn.Text = "Разместить"
        ok_btn.Location = Point(250, 200)
        ok_btn.Size = Size(85, 30)
        ok_btn.Click += self.on_ok

        cancel_btn = Button()
        cancel_btn.Text = "Отмена"
        cancel_btn.Location = Point(345, 200)
        cancel_btn.Size = Size(75, 30)
        cancel_btn.DialogResult = DialogResult.Cancel

        self.Controls.Add(type_label)
        self.Controls.Add(self.type_combo)
        self.Controls.Add(coord_group)
        self.Controls.Add(ok_btn)
        self.Controls.Add(cancel_btn)

    def load_types(self):
        self.type_combo.Items.Clear()
        for ft in self.family_types:
            item = FamilyTypeItem(ft)
            self.type_combo.Items.Add(item)
        if self.type_combo.Items.Count > 0:
            self.type_combo.SelectedIndex = 0

    def on_ok(self, sender, args):
        if not self.type_combo.SelectedItem:
            show_warning("Предупреждение", "Выберите типоразмер семейства")
            return

        self.selected_type = self.type_combo.SelectedItem.family_symbol
        self.x_coord = float(self.x_input.Value)
        self.y_coord = float(self.y_input.Value)

        self.DialogResult = DialogResult.OK
        self.Close()


def get_cpsk_annotation_types():
    """Получить все типоразмеры семейств CPSK_VD_F."""
    types = []
    collector = FilteredElementCollector(doc).OfClass(Family)

    for family in collector:
        if family.Name.startswith("CPSK_VD_F"):
            try:
                cat_id = family.FamilyCategory.Id.IntegerValue
                if cat_id == int(BuiltInCategory.OST_GenericAnnotation):
                    for type_id in family.GetFamilySymbolIds():
                        symbol = doc.GetElement(type_id)
                        if symbol:
                            types.append(symbol)
            except Exception:
                pass

    return sorted(types, key=lambda x: "{}_{}".format(x.Family.Name, x.Name))


def main():
    """Основная функция."""
    # Проверить, что активный вид - лист
    active_view = doc.ActiveView
    if not isinstance(active_view, ViewSheet):
        show_error("Ошибка", "Для размещения аннотации необходимо открыть лист.\nПожалуйста, откройте лист и повторите.")
        return

    # Получить типоразмеры семейств
    family_types = get_cpsk_annotation_types()

    if not family_types:
        show_warning("Предупреждение", "В проекте не найдено семейств аннотаций с префиксом CPSK_VD_F")
        return

    # Показать диалог
    dialog = PlaceAnnotationDialog(family_types)
    result = dialog.ShowDialog()

    if result != DialogResult.OK:
        return

    family_symbol = dialog.selected_type
    x_mm = dialog.x_coord
    y_mm = dialog.y_coord

    # Конвертировать мм в футы
    x_feet = x_mm / 304.8
    y_feet = y_mm / 304.8

    # Разместить семейство
    with revit.Transaction("Разместить аннотацию"):
        try:
            # Активировать типоразмер если нужно
            if not family_symbol.IsActive:
                family_symbol.Activate()
                doc.Regenerate()

            # Создать экземпляр
            placement_point = XYZ(x_feet, y_feet, 0)
            instance = doc.Create.NewFamilyInstance(placement_point, family_symbol, active_view)

            if instance:
                show_success(
                    "Успех",
                    "Аннотация '{}' размещена на листе".format(family_symbol.Name),
                    details="Лист: {} - {}\nКоординаты: X={} мм, Y={} мм".format(
                        active_view.SheetNumber, active_view.Name, x_mm, y_mm
                    )
                )
            else:
                show_error("Ошибка", "Не удалось создать экземпляр семейства")

        except Exception as ex:
            show_error("Ошибка", "Ошибка при размещении аннотации: {}".format(str(ex)))


if __name__ == "__main__":
    main()
