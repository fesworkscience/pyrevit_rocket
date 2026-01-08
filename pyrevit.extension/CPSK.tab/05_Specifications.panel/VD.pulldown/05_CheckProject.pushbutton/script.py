#! python3
# -*- coding: utf-8 -*-
"""
Проверка проекта перед созданием ВД.
Проверяет соответствие форм арматуры и семейств аннотаций,
наличие спецификаций CPSK_VD_, параметры семейств.
"""

__title__ = "Проверить\nпроект"
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
    DockStyle, ScrollBars, RichTextBox, TabControl, TabPage
)
from System.Drawing import Point, Size, Font, FontStyle, Color

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

from cpsk_notify import show_error, show_info

from pyrevit import revit

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, Family, FamilySymbol,
    BuiltInCategory, StorageType
)

doc = revit.doc


# Статусы проверки
STATUS_SUCCESS = "success"
STATUS_WARNING = "warning"
STATUS_ERROR = "error"


def get_used_rebar_shapes():
    """Получить все используемые формы арматуры."""
    shapes = set()
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar)
    for elem in collector:
        param = elem.LookupParameter("Форма")
        if param and param.HasValue:
            shape = param.AsString() if param.StorageType == StorageType.String else param.AsValueString()
            if shape:
                shapes.add(shape)
    return sorted(list(shapes))


def get_loaded_annotation_family_types():
    """Получить загруженные типоразмеры семейств аннотаций CPSK_VD_F."""
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
                            types.append(symbol.Name)
            except Exception:
                pass
    return sorted(types)


def check_family_correspondence():
    """Проверка соответствия форм арматуры и типоразмеров семейств."""
    used_shapes = get_used_rebar_shapes()
    loaded_types = get_loaded_annotation_family_types()

    missing_types = []
    for shape in used_shapes:
        found = False
        for type_name in loaded_types:
            if shape in type_name or type_name in shape:
                found = True
                break
        if not found:
            missing_types.append(shape)

    if not missing_types:
        return {
            "status": STATUS_SUCCESS,
            "message": "Все формы арматуры ({}) имеют соответствующие типоразмеры семейств".format(len(used_shapes)),
            "used_shapes": used_shapes,
            "loaded_types": loaded_types,
            "missing_types": []
        }
    else:
        return {
            "status": STATUS_WARNING,
            "message": "Отсутствуют типоразмеры для {} форм арматуры".format(len(missing_types)),
            "used_shapes": used_shapes,
            "loaded_types": loaded_types,
            "missing_types": missing_types
        }


def check_cpsk_specifications():
    """Проверка наличия спецификаций CPSK_VD_."""
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    cpsk_schedules = [s for s in collector if s.Name.startswith("CPSK_VD_")]

    if not cpsk_schedules:
        return {
            "status": STATUS_ERROR,
            "message": "Не найдено спецификаций с префиксом CPSK_VD_",
            "specifications": [],
            "count": 0
        }
    else:
        return {
            "status": STATUS_SUCCESS,
            "message": "Найдено {} спецификаций CPSK_VD_".format(len(cpsk_schedules)),
            "specifications": [s.Name for s in cpsk_schedules],
            "count": len(cpsk_schedules)
        }


def check_annotation_family_parameters():
    """Проверка параметров семейств аннотаций."""
    collector = FilteredElementCollector(doc).OfClass(Family)
    annotation_families = []
    families_with_param = []
    families_without_param = []

    for family in collector:
        if family.Name.startswith("CPSK_VD_F"):
            try:
                cat_id = family.FamilyCategory.Id.IntegerValue
                if cat_id == int(BuiltInCategory.OST_GenericAnnotation):
                    annotation_families.append(family.Name)
                    has_param = False

                    # Проверить типоразмеры
                    for type_id in family.GetFamilySymbolIds():
                        symbol = doc.GetElement(type_id)
                        if symbol:
                            param = symbol.LookupParameter("CPSK_VD_ID_Спецификации")
                            if param:
                                has_param = True
                                break

                    if has_param:
                        families_with_param.append(family.Name)
                    else:
                        families_without_param.append(family.Name)
            except Exception:
                pass

    if not annotation_families:
        return {
            "status": STATUS_ERROR,
            "message": "Не найдено семейств аннотаций CPSK_VD_F",
            "found_families": [],
            "with_parameter": [],
            "without_parameter": []
        }
    elif families_without_param:
        return {
            "status": STATUS_WARNING,
            "message": "{} семейств не имеют параметра CPSK_VD_ID_Спецификации".format(len(families_without_param)),
            "found_families": annotation_families,
            "with_parameter": families_with_param,
            "without_parameter": families_without_param
        }
    else:
        return {
            "status": STATUS_SUCCESS,
            "message": "Все семейства аннотаций ({}) имеют параметр CPSK_VD_ID_Спецификации".format(len(annotation_families)),
            "found_families": annotation_families,
            "with_parameter": families_with_param,
            "without_parameter": []
        }


class ProjectCheckDialog(Form):
    """Диалог с результатами проверки проекта."""

    def __init__(self, results):
        self.results = results
        self.setup_form()

    def setup_form(self):
        self.Text = "Результаты проверки проекта"
        self.Width = 700
        self.Height = 550
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        # Общий статус
        status_panel = Panel()
        status_panel.Location = Point(10, 10)
        status_panel.Size = Size(665, 60)

        status_label = Label()
        status_label.Location = Point(10, 10)
        status_label.Size = Size(645, 40)
        status_label.Font = Font("Segoe UI", 11, FontStyle.Bold)

        overall_status = self.get_overall_status()
        if overall_status == STATUS_SUCCESS:
            status_label.Text = "Проект готов к созданию ВД"
            status_label.ForeColor = Color.Green
        elif overall_status == STATUS_WARNING:
            status_label.Text = "Проект имеет предупреждения"
            status_label.ForeColor = Color.Orange
        else:
            status_label.Text = "Проект требует исправлений"
            status_label.ForeColor = Color.Red

        status_panel.Controls.Add(status_label)

        # Вкладки с результатами
        tabs = TabControl()
        tabs.Location = Point(10, 80)
        tabs.Size = Size(665, 380)

        # Вкладка: Семейства
        family_tab = TabPage()
        family_tab.Text = "Семейства"
        family_text = self.format_family_check()
        family_box = RichTextBox()
        family_box.Dock = DockStyle.Fill
        family_box.ReadOnly = True
        family_box.Text = family_text
        family_box.Font = Font("Consolas", 9)
        family_tab.Controls.Add(family_box)
        tabs.TabPages.Add(family_tab)

        # Вкладка: Спецификации
        spec_tab = TabPage()
        spec_tab.Text = "Спецификации"
        spec_text = self.format_spec_check()
        spec_box = RichTextBox()
        spec_box.Dock = DockStyle.Fill
        spec_box.ReadOnly = True
        spec_box.Text = spec_text
        spec_box.Font = Font("Consolas", 9)
        spec_tab.Controls.Add(spec_box)
        tabs.TabPages.Add(spec_tab)

        # Вкладка: Аннотации
        annot_tab = TabPage()
        annot_tab.Text = "Аннотации"
        annot_text = self.format_annotation_check()
        annot_box = RichTextBox()
        annot_box.Dock = DockStyle.Fill
        annot_box.ReadOnly = True
        annot_box.Text = annot_text
        annot_box.Font = Font("Consolas", 9)
        annot_tab.Controls.Add(annot_box)
        tabs.TabPages.Add(annot_tab)

        # Кнопка закрытия
        close_btn = Button()
        close_btn.Text = "Закрыть"
        close_btn.Location = Point(580, 470)
        close_btn.Size = Size(95, 30)
        close_btn.DialogResult = DialogResult.OK

        self.Controls.Add(status_panel)
        self.Controls.Add(tabs)
        self.Controls.Add(close_btn)

    def get_overall_status(self):
        statuses = [
            self.results["family_check"]["status"],
            self.results["spec_check"]["status"],
            self.results["annotation_check"]["status"]
        ]
        if STATUS_ERROR in statuses:
            return STATUS_ERROR
        if STATUS_WARNING in statuses:
            return STATUS_WARNING
        return STATUS_SUCCESS

    def get_status_icon(self, status):
        if status == STATUS_SUCCESS:
            return "[OK]"
        elif status == STATUS_WARNING:
            return "[!]"
        else:
            return "[X]"

    def format_family_check(self):
        r = self.results["family_check"]
        lines = []
        lines.append("{} {}".format(self.get_status_icon(r["status"]), r["message"]))
        lines.append("")
        lines.append("Используемые формы арматуры ({}):\n".format(len(r["used_shapes"])))
        for shape in r["used_shapes"]:
            lines.append("  - {}".format(shape))
        lines.append("")
        lines.append("Загруженные типоразмеры ({}):\n".format(len(r["loaded_types"])))
        for t in r["loaded_types"]:
            lines.append("  - {}".format(t))
        if r["missing_types"]:
            lines.append("")
            lines.append("Отсутствующие типоразмеры ({}):\n".format(len(r["missing_types"])))
            for t in r["missing_types"]:
                lines.append("  - {}".format(t))
        return "\n".join(lines)

    def format_spec_check(self):
        r = self.results["spec_check"]
        lines = []
        lines.append("{} {}".format(self.get_status_icon(r["status"]), r["message"]))
        lines.append("")
        if r["specifications"]:
            lines.append("Найденные спецификации:\n")
            for spec in r["specifications"]:
                lines.append("  - {}".format(spec))
        return "\n".join(lines)

    def format_annotation_check(self):
        r = self.results["annotation_check"]
        lines = []
        lines.append("{} {}".format(self.get_status_icon(r["status"]), r["message"]))
        lines.append("")
        if r["with_parameter"]:
            lines.append("С параметром CPSK_VD_ID_Спецификации ({}):\n".format(len(r["with_parameter"])))
            for f in r["with_parameter"]:
                lines.append("  - {}".format(f))
        if r["without_parameter"]:
            lines.append("")
            lines.append("Без параметра ({}):\n".format(len(r["without_parameter"])))
            for f in r["without_parameter"]:
                lines.append("  - {}".format(f))
        return "\n".join(lines)


def main():
    """Основная функция."""
    try:
        # Выполнить проверки
        results = {
            "family_check": check_family_correspondence(),
            "spec_check": check_cpsk_specifications(),
            "annotation_check": check_annotation_family_parameters()
        }

        # Показать диалог с результатами
        dialog = ProjectCheckDialog(results)
        dialog.ShowDialog()

    except Exception as ex:
        show_error("Ошибка", "Ошибка при проверке проекта: {}".format(str(ex)))


if __name__ == "__main__":
    main()
