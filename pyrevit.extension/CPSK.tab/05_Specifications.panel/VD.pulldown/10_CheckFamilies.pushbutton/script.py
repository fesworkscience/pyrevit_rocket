#! python3
# -*- coding: utf-8 -*-
"""
Подробная проверка семейств форм арматуры.
Показывает загруженные формы, используемые формы, семейства аннотаций
и соответствие между ними.
"""

__title__ = "Проверить\nсемейства"
__author__ = "CPSK"

import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, Panel, TreeView, TreeNode,
    FormStartPosition, FormBorderStyle, DialogResult,
    DockStyle, RichTextBox, TabControl, TabPage, SaveFileDialog
)
from System.Drawing import Point, Size, Font

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

import codecs

# Проверка авторизации
from cpsk_auth import require_auth
if not require_auth():
    sys.exit()

from cpsk_notify import show_error, show_success

from pyrevit import revit

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, Family, FamilySymbol,
    BuiltInCategory, StorageType
)
from Autodesk.Revit.DB.Structure import RebarShape

doc = revit.doc


def get_loaded_rebar_shapes():
    """Получить все загруженные формы арматурных стержней."""
    shapes = []
    collector = FilteredElementCollector(doc).OfClass(RebarShape)
    for shape in collector:
        shapes.append(shape.Name)
    return sorted(shapes)


def get_used_rebar_shapes():
    """Получить все используемые формы арматуры в проекте."""
    shapes = {}  # shape_name -> count
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar)
    collector.WhereElementIsNotElementType()

    for elem in collector:
        param = elem.LookupParameter("Форма")
        if param and param.HasValue:
            if param.StorageType == StorageType.String:
                shape = param.AsString()
            else:
                shape = param.AsValueString()
            if shape:
                if shape not in shapes:
                    shapes[shape] = 0
                shapes[shape] += 1

    return shapes


def get_shapes_in_cpsk_schedules():
    """Получить формы арматуры в спецификациях CPSK_VD_."""
    schedule_shapes = {}  # schedule_name -> set of shapes

    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    cpsk_schedules = [s for s in collector if s.Name.startswith("CPSK_VD_")]

    for schedule in cpsk_schedules:
        shapes = set()
        try:
            elem_collector = FilteredElementCollector(doc, schedule.Id)
            elem_collector.OfCategory(BuiltInCategory.OST_Rebar)
            for elem in elem_collector:
                param = elem.LookupParameter("Форма")
                if param and param.HasValue:
                    if param.StorageType == StorageType.String:
                        shape = param.AsString()
                    else:
                        shape = param.AsValueString()
                    if shape:
                        shapes.add(shape)
        except Exception:
            pass
        schedule_shapes[schedule.Name] = shapes

    return schedule_shapes


def get_cpsk_annotation_families():
    """Получить семейства аннотаций CPSK_VD_F."""
    families_info = {}  # family_name -> list of type_names

    collector = FilteredElementCollector(doc).OfClass(Family)
    for family in collector:
        if family.Name.startswith("CPSK_VD_F"):
            try:
                cat_id = family.FamilyCategory.Id.IntegerValue
                if cat_id == int(BuiltInCategory.OST_GenericAnnotation):
                    types = []
                    for type_id in family.GetFamilySymbolIds():
                        symbol = doc.GetElement(type_id)
                        if symbol:
                            types.append(symbol.Name)
                    families_info[family.Name] = sorted(types)
            except Exception:
                pass

    return families_info


def check_shape_correspondence(used_shapes, annotation_families):
    """Проверить соответствие форм и семейств."""
    # Собрать все типоразмеры
    all_types = set()
    for types in annotation_families.values():
        for t in types:
            all_types.add(t)

    found = []
    missing = []

    for shape in used_shapes.keys():
        shape_found = False
        for type_name in all_types:
            if shape in type_name or type_name in shape or shape.lower() == type_name.lower():
                shape_found = True
                break
        if shape_found:
            found.append(shape)
        else:
            missing.append(shape)

    return found, missing


class FamilyCheckDialog(Form):
    """Диалог с результатами проверки семейств."""

    def __init__(self, report_data):
        self.report_data = report_data
        self.setup_form()

    def setup_form(self):
        self.Text = "Проверка семейств форм арматуры"
        self.Width = 800
        self.Height = 600
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        # Вкладки
        tabs = TabControl()
        tabs.Location = Point(10, 10)
        tabs.Size = Size(765, 500)

        # Вкладка: Загруженные формы
        loaded_tab = TabPage()
        loaded_tab.Text = "Загруженные формы"
        loaded_text = self.format_loaded_shapes()
        loaded_box = RichTextBox()
        loaded_box.Dock = DockStyle.Fill
        loaded_box.ReadOnly = True
        loaded_box.Text = loaded_text
        loaded_box.Font = Font("Consolas", 9)
        loaded_tab.Controls.Add(loaded_box)
        tabs.TabPages.Add(loaded_tab)

        # Вкладка: Используемые формы
        used_tab = TabPage()
        used_tab.Text = "Используемые формы"
        used_text = self.format_used_shapes()
        used_box = RichTextBox()
        used_box.Dock = DockStyle.Fill
        used_box.ReadOnly = True
        used_box.Text = used_text
        used_box.Font = Font("Consolas", 9)
        used_tab.Controls.Add(used_box)
        tabs.TabPages.Add(used_tab)

        # Вкладка: Спецификации CPSK_VD_
        schedules_tab = TabPage()
        schedules_tab.Text = "Спецификации CPSK_VD_"
        schedules_text = self.format_schedule_shapes()
        schedules_box = RichTextBox()
        schedules_box.Dock = DockStyle.Fill
        schedules_box.ReadOnly = True
        schedules_box.Text = schedules_text
        schedules_box.Font = Font("Consolas", 9)
        schedules_tab.Controls.Add(schedules_box)
        tabs.TabPages.Add(schedules_tab)

        # Вкладка: Семейства аннотаций
        families_tab = TabPage()
        families_tab.Text = "Семейства CPSK_VD_F"
        families_text = self.format_annotation_families()
        families_box = RichTextBox()
        families_box.Dock = DockStyle.Fill
        families_box.ReadOnly = True
        families_box.Text = families_text
        families_box.Font = Font("Consolas", 9)
        families_tab.Controls.Add(families_box)
        tabs.TabPages.Add(families_tab)

        # Вкладка: Соответствие
        check_tab = TabPage()
        check_tab.Text = "Проверка соответствия"
        check_text = self.format_correspondence()
        check_box = RichTextBox()
        check_box.Dock = DockStyle.Fill
        check_box.ReadOnly = True
        check_box.Text = check_text
        check_box.Font = Font("Consolas", 9)
        check_tab.Controls.Add(check_box)
        tabs.TabPages.Add(check_tab)

        # Кнопки
        save_btn = Button()
        save_btn.Text = "Сохранить отчет"
        save_btn.Location = Point(580, 520)
        save_btn.Size = Size(100, 30)
        save_btn.Click += self.on_save

        close_btn = Button()
        close_btn.Text = "Закрыть"
        close_btn.Location = Point(690, 520)
        close_btn.Size = Size(85, 30)
        close_btn.DialogResult = DialogResult.OK

        self.Controls.Add(tabs)
        self.Controls.Add(save_btn)
        self.Controls.Add(close_btn)

    def format_loaded_shapes(self):
        lines = ["=== ЗАГРУЖЕННЫЕ ФОРМЫ АРМАТУРНЫХ СТЕРЖНЕЙ ===\n"]
        shapes = self.report_data["loaded_shapes"]
        lines.append("Всего загружено: {}\n".format(len(shapes)))
        for shape in shapes:
            lines.append("  - {}".format(shape))
        return "\n".join(lines)

    def format_used_shapes(self):
        lines = ["=== ИСПОЛЬЗОВАННЫЕ ФОРМЫ АРМАТУРЫ В ПРОЕКТЕ ===\n"]
        shapes = self.report_data["used_shapes"]
        lines.append("Всего используется: {} форм\n".format(len(shapes)))
        for shape, count in sorted(shapes.items()):
            lines.append("  - {} ({} элементов)".format(shape, count))
        return "\n".join(lines)

    def format_schedule_shapes(self):
        lines = ["=== ФОРМЫ В СПЕЦИФИКАЦИЯХ CPSK_VD_ ===\n"]
        schedules = self.report_data["schedule_shapes"]
        if not schedules:
            lines.append("Спецификации CPSK_VD_ не найдены")
        else:
            for schedule_name, shapes in sorted(schedules.items()):
                lines.append("{}:".format(schedule_name))
                if shapes:
                    for shape in sorted(shapes):
                        lines.append("    - {}".format(shape))
                else:
                    lines.append("    (нет форм)")
                lines.append("")
        return "\n".join(lines)

    def format_annotation_families(self):
        lines = ["=== СЕМЕЙСТВА АННОТАЦИЙ CPSK_VD_F ===\n"]
        families = self.report_data["annotation_families"]
        if not families:
            lines.append("Семейства CPSK_VD_F не найдены")
        else:
            for family_name, types in sorted(families.items()):
                lines.append("{}:".format(family_name))
                lines.append("  Типоразмеры ({})".format(len(types)))
                for t in types:
                    lines.append("    - {}".format(t))
                lines.append("")
        return "\n".join(lines)

    def format_correspondence(self):
        lines = ["=== ПРОВЕРКА СООТВЕТСТВИЯ ===\n"]
        found = self.report_data["found_shapes"]
        missing = self.report_data["missing_shapes"]

        if found:
            lines.append("[OK] Формы с соответствующими семействами ({}):\n".format(len(found)))
            for shape in sorted(found):
                lines.append("  - {}".format(shape))
            lines.append("")

        if missing:
            lines.append("[!] Формы БЕЗ соответствующих семейств ({}):\n".format(len(missing)))
            used_shapes = self.report_data["used_shapes"]
            for shape in sorted(missing):
                count = used_shapes.get(shape, 0)
                lines.append("  - {} ({} элементов)".format(shape, count))
        else:
            lines.append("\n[OK] Все используемые формы имеют соответствующие семейства!")

        return "\n".join(lines)

    def on_save(self, sender, args):
        dialog = SaveFileDialog()
        dialog.Filter = "Text files (*.txt)|*.txt"
        dialog.FileName = "family_check_report.txt"

        if dialog.ShowDialog() == DialogResult.OK:
            try:
                full_report = []
                full_report.append(self.format_loaded_shapes())
                full_report.append("\n\n")
                full_report.append(self.format_used_shapes())
                full_report.append("\n\n")
                full_report.append(self.format_schedule_shapes())
                full_report.append("\n\n")
                full_report.append(self.format_annotation_families())
                full_report.append("\n\n")
                full_report.append(self.format_correspondence())

                with codecs.open(dialog.FileName, 'w', 'utf-8') as f:
                    f.write("".join(full_report))

                show_success("Сохранено", "Отчет сохранен: {}".format(dialog.FileName))
            except Exception as ex:
                show_error("Ошибка", "Не удалось сохранить отчет: {}".format(str(ex)))


def main():
    """Основная функция."""
    # Собрать данные
    loaded_shapes = get_loaded_rebar_shapes()
    used_shapes = get_used_rebar_shapes()
    schedule_shapes = get_shapes_in_cpsk_schedules()
    annotation_families = get_cpsk_annotation_families()
    found_shapes, missing_shapes = check_shape_correspondence(used_shapes, annotation_families)

    report_data = {
        "loaded_shapes": loaded_shapes,
        "used_shapes": used_shapes,
        "schedule_shapes": schedule_shapes,
        "annotation_families": annotation_families,
        "found_shapes": found_shapes,
        "missing_shapes": missing_shapes
    }

    # Показать диалог
    dialog = FamilyCheckDialog(report_data)
    dialog.ShowDialog()


if __name__ == "__main__":
    main()
