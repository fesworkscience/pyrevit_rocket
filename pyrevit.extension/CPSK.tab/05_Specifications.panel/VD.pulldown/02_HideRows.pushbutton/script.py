# -*- coding: utf-8 -*-
"""
Скрытие строк в спецификациях CPSK_VD_.
Добавляет фильтр с невозможным значением для скрытия всех строк.
"""

__title__ = "Скрыть\nстроки"
__author__ = "CPSK"

import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, CheckedListBox, Panel,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, AnchorStyles
)
from System.Drawing import Point, Size, Font, FontStyle

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
    FilteredElementCollector, ViewSchedule, Transaction,
    ScheduleFilter, ScheduleFilterType, ScheduleFieldType
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
    """Диалог выбора спецификаций для скрытия строк."""

    def __init__(self, specifications):
        self.all_specs = specifications
        self.selected_specs = []
        self.setup_form()
        self.load_specs()

    def setup_form(self):
        self.Text = "Выбор спецификаций для скрытия строк"
        self.Width = 500
        self.Height = 400
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        # Заголовок
        title_label = Label()
        title_label.Text = "Выберите спецификации CPSK_VD_ для скрытия строк:"
        title_label.Location = Point(12, 9)
        title_label.Size = Size(460, 40)
        title_label.Font = Font("Segoe UI", 9)

        # Список с чекбоксами
        self.spec_list = CheckedListBox()
        self.spec_list.Location = Point(12, 55)
        self.spec_list.Size = Size(460, 260)
        self.spec_list.CheckOnClick = True

        # Кнопки выбора
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

        # OK/Cancel
        ok_btn = Button()
        ok_btn.Text = "OK"
        ok_btn.Location = Point(316, 325)
        ok_btn.Size = Size(75, 23)
        ok_btn.DialogResult = DialogResult.OK
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
        # Выбрать все по умолчанию
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


def hide_rows_in_schedule(schedule):
    """Скрыть все строки в спецификации путем добавления фильтра."""
    try:
        schedule_def = schedule.Definition

        # Найти подходящее поле для фильтрации
        field_id = None
        for i in range(schedule_def.GetFieldCount()):
            field = schedule_def.GetField(i)
            if field.FieldType == ScheduleFieldType.Instance or field.FieldType == ScheduleFieldType.ElementType:
                field_id = field.FieldId
                break

        if field_id is None:
            return False, "Не найдено подходящее поле для фильтра"

        # Добавить фильтр с невозможным значением
        hide_filter = ScheduleFilter(field_id, ScheduleFilterType.Equal, "##HIDE_ALL_ROWS##")
        schedule_def.AddFilter(hide_filter)
        return True, None

    except Exception as ex:
        return False, str(ex)


def main():
    """Основная функция."""
    # Найти все спецификации CPSK_VD_
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
        show_warning("Предупреждение", "Не выбрано ни одной спецификации")
        return

    # Скрыть строки
    with revit.Transaction("Скрыть строки в спецификациях"):
        hidden_count = 0
        errors = []

        for schedule in selected:
            success, error = hide_rows_in_schedule(schedule)
            if success:
                hidden_count += 1
            else:
                errors.append("{}: {}".format(schedule.Name, error or "Неизвестная ошибка"))

        # Показать результат
        if hidden_count > 0:
            message = "Строки скрыты в {} из {} выбранных спецификаций".format(hidden_count, len(selected))
            if errors:
                details = "Ошибки:\n" + "\n".join(errors)
                show_warning("Частично выполнено", message, details=details)
            else:
                show_success("Успех", message)
        else:
            show_error("Ошибка", "Не удалось скрыть строки ни в одной спецификации",
                      details="\n".join(errors) if errors else None)


if __name__ == "__main__":
    main()
