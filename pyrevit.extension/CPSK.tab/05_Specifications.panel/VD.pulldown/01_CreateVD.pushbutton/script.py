# -*- coding: utf-8 -*-
"""
Создание копии спецификации с настройками ВД.
Копирует выбранную спецификацию с префиксом CPSK_VD_,
настраивает ширину столбцов и добавляет фильтр скрытия строк.
"""

__title__ = "Создать\nВД"
__author__ = "CPSK"

import clr
import os
import sys
from datetime import datetime

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, ListBox, Panel,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, BorderStyle, AnchorStyles
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
    ViewDuplicateOption, SectionType, ScheduleFilter,
    ScheduleFilterType, ScheduleFieldType
)

doc = revit.doc


class ScheduleListItem(object):
    """Элемент списка спецификаций."""
    def __init__(self, schedule, display_text):
        self.schedule = schedule
        self.display_text = display_text

    def __str__(self):
        return self.display_text


class ScheduleSelectionDialog(Form):
    """Диалог выбора спецификации."""

    def __init__(self, schedules, title="Выберите спецификацию"):
        self.all_schedules = schedules or []
        self.filtered_schedules = list(self.all_schedules)
        self.selected_schedule = None
        self.setup_form(title)
        self.populate_list()

    def setup_form(self, title):
        self.Text = title
        self.Width = 600
        self.Height = 500
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        # Заголовок
        self.title_label = Label()
        self.title_label.Text = title
        self.title_label.Font = Font("Segoe UI", 12, FontStyle.Bold)
        self.title_label.Location = Point(20, 20)
        self.title_label.Size = Size(560, 25)

        # Поиск
        search_label = Label()
        search_label.Text = "Поиск:"
        search_label.Location = Point(20, 55)
        search_label.Size = Size(50, 20)

        self.search_box = TextBox()
        self.search_box.Location = Point(80, 53)
        self.search_box.Size = Size(400, 23)
        self.search_box.TextChanged += self.on_search_changed

        clear_btn = Button()
        clear_btn.Text = "X"
        clear_btn.Location = Point(485, 53)
        clear_btn.Size = Size(25, 23)
        clear_btn.Click += self.on_clear_search

        # Список спецификаций
        list_label = Label()
        list_label.Text = "Доступные спецификации:"
        list_label.Location = Point(20, 90)
        list_label.Size = Size(200, 20)

        self.schedule_list = ListBox()
        self.schedule_list.Location = Point(20, 115)
        self.schedule_list.Size = Size(560, 280)
        self.schedule_list.BorderStyle = BorderStyle.FixedSingle
        self.schedule_list.Font = Font("Segoe UI", 9)
        self.schedule_list.SelectedIndexChanged += self.on_selection_changed
        self.schedule_list.DoubleClick += self.on_double_click

        # Информация
        self.info_label = Label()
        self.info_label.Location = Point(20, 405)
        self.info_label.Size = Size(560, 40)
        self.info_label.Text = "Найдено спецификаций: {}. Выберите спецификацию для создания копии с настройками ВД.".format(len(self.all_schedules))

        # Кнопки
        self.ok_btn = Button()
        self.ok_btn.Text = "Создать копию"
        self.ok_btn.Location = Point(400, 450)
        self.ok_btn.Size = Size(100, 30)
        self.ok_btn.Enabled = False
        self.ok_btn.Click += self.on_ok

        cancel_btn = Button()
        cancel_btn.Text = "Отмена"
        cancel_btn.Location = Point(510, 450)
        cancel_btn.Size = Size(80, 30)
        cancel_btn.DialogResult = DialogResult.Cancel

        # Добавление контролов
        self.Controls.Add(self.title_label)
        self.Controls.Add(search_label)
        self.Controls.Add(self.search_box)
        self.Controls.Add(clear_btn)
        self.Controls.Add(list_label)
        self.Controls.Add(self.schedule_list)
        self.Controls.Add(self.info_label)
        self.Controls.Add(self.ok_btn)
        self.Controls.Add(cancel_btn)

        self.AcceptButton = self.ok_btn
        self.CancelButton = cancel_btn

    def populate_list(self):
        self.schedule_list.Items.Clear()
        for schedule in self.filtered_schedules:
            category_name = self.get_category_name(schedule)
            display_text = schedule.Name
            if category_name:
                display_text = "{} ({})".format(schedule.Name, category_name)
            self.schedule_list.Items.Add(ScheduleListItem(schedule, display_text))
        self.update_info()

    def get_category_name(self, schedule):
        try:
            definition = schedule.Definition
            if definition:
                from Autodesk.Revit.DB import Category, ElementId
                category_id = definition.CategoryId
                if category_id != ElementId.InvalidElementId:
                    category = Category.GetCategory(schedule.Document, category_id)
                    return category.Name if category else ""
        except Exception:
            pass
        return ""

    def on_search_changed(self, sender, args):
        search_text = self.search_box.Text.lower().strip()
        if not search_text:
            self.filtered_schedules = list(self.all_schedules)
        else:
            self.filtered_schedules = [s for s in self.all_schedules if search_text in s.Name.lower()]
        self.populate_list()

    def on_clear_search(self, sender, args):
        self.search_box.Clear()
        self.search_box.Focus()

    def on_selection_changed(self, sender, args):
        if self.schedule_list.SelectedItem:
            self.selected_schedule = self.schedule_list.SelectedItem.schedule
            self.ok_btn.Enabled = True
        else:
            self.selected_schedule = None
            self.ok_btn.Enabled = False

    def on_double_click(self, sender, args):
        if self.schedule_list.SelectedItem:
            self.on_ok(sender, args)

    def on_ok(self, sender, args):
        if self.selected_schedule:
            self.DialogResult = DialogResult.OK
            self.Close()

    def update_info(self):
        self.info_label.Text = "Найдено спецификаций: {} из {}.".format(
            len(self.filtered_schedules), len(self.all_schedules))


def generate_unique_name(base_name):
    """Генерация уникального имени спецификации."""
    existing_names = set()
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    for schedule in collector:
        existing_names.add(schedule.Name)

    # Попробовать без суффикса
    if base_name not in existing_names:
        return base_name

    # Попробовать с номером
    for i in range(1, 101):
        candidate = "{}({})".format(base_name, i)
        if candidate not in existing_names:
            return candidate

    # Fallback с временной меткой
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return "{}_{}".format(base_name, timestamp)


def adjust_column_widths(schedule):
    """Настройка ширины столбцов: A+B=20мм, D+E=70мм."""
    try:
        table_data = schedule.GetTableData()
        body_section = table_data.GetSectionData(SectionType.Body)
        num_columns = body_section.NumberOfColumns

        # Конвертация мм в футы (Revit использует футы)
        width_20mm = 20.0 / 304.8  # 20мм в футах
        width_70mm = 70.0 / 304.8  # 70мм в футах

        # Столбцы A+B (0,1) = 20мм (по 10мм каждый)
        if num_columns >= 2:
            body_section.SetColumnWidth(0, width_20mm / 2)
            body_section.SetColumnWidth(1, width_20mm / 2)

        # Столбцы D+E (3,4) = 70мм (по 35мм каждый)
        if num_columns >= 5:
            body_section.SetColumnWidth(3, width_70mm / 2)
            body_section.SetColumnWidth(4, width_70mm / 2)

        # Также настроить заголовок
        header_section = table_data.GetSectionData(SectionType.Header)
        if header_section.NumberOfColumns >= 5:
            header_section.SetColumnWidth(0, width_20mm / 2)
            header_section.SetColumnWidth(1, width_20mm / 2)
            header_section.SetColumnWidth(3, width_70mm / 2)
            header_section.SetColumnWidth(4, width_70mm / 2)

        return True
    except Exception as ex:
        show_warning("Предупреждение", "Не удалось настроить ширину столбцов: {}".format(str(ex)))
        return False


def hide_all_rows(schedule):
    """Добавление фильтра для скрытия всех строк."""
    try:
        schedule_def = schedule.Definition

        # Найти поле "Расч_Метка основы"
        field_id = None
        for i in range(schedule_def.GetFieldCount()):
            field = schedule_def.GetField(i)
            if field.GetName() == "Расч_Метка основы":
                field_id = field.FieldId
                break

        if field_id:
            # Добавить фильтр с невозможным значением
            hide_filter = ScheduleFilter(field_id, ScheduleFilterType.Equal, "##HIDE_ALL_ROWS##")
            schedule_def.AddFilter(hide_filter)
            return True
        else:
            # Если поле не найдено, найти любое текстовое поле
            for i in range(schedule_def.GetFieldCount()):
                field = schedule_def.GetField(i)
                if field.FieldType == ScheduleFieldType.Instance or field.FieldType == ScheduleFieldType.ElementType:
                    field_id = field.FieldId
                    break

            if field_id:
                hide_filter = ScheduleFilter(field_id, ScheduleFilterType.Equal, "##HIDE_ALL_ROWS##")
                schedule_def.AddFilter(hide_filter)
                return True

        return False
    except Exception as ex:
        show_warning("Предупреждение", "Не удалось добавить фильтр скрытия: {}".format(str(ex)))
        return False


def main():
    """Основная функция."""
    # Получить все спецификации
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    all_schedules = sorted([s for s in collector], key=lambda x: x.Name)

    if not all_schedules:
        show_warning("Нет спецификаций", "В проекте не найдено спецификаций для копирования")
        return

    # Показать диалог выбора
    dialog = ScheduleSelectionDialog(all_schedules, "Выберите спецификацию для копирования")
    result = dialog.ShowDialog()

    if result != DialogResult.OK:
        return

    template_schedule = dialog.selected_schedule
    if not template_schedule:
        show_warning("Не выбрано", "Не выбрана спецификация для копирования")
        return

    # Создать копию спецификации
    with revit.Transaction("Создать спецификацию ВД"):
        try:
            # Дублировать спецификацию
            new_schedule_id = template_schedule.Duplicate(ViewDuplicateOption.Duplicate)
            new_schedule = doc.GetElement(new_schedule_id)

            if not new_schedule:
                show_error("Ошибка", "Не удалось создать копию спецификации")
                return

            # Сгенерировать имя
            base_name = "CPSK_VD_КЖ_Арматура_Ведомость деталей_часть"
            new_name = generate_unique_name(base_name)
            new_schedule.Name = new_name

            # Настроить спецификацию
            adjust_column_widths(new_schedule)
            hide_all_rows(new_schedule)

            # Показать результат
            details = "Исходная спецификация: '{}'\n".format(template_schedule.Name)
            details += "Настроены ширины столбцов (A+B=20мм, D+E=70мм)\n"
            details += "Добавлен фильтр скрытия строк"

            show_success(
                "Спецификация создана",
                "Создана новая спецификация ВД: '{}'".format(new_name),
                details=details
            )

        except Exception as ex:
            show_error("Ошибка", "Ошибка при создании спецификации: {}".format(str(ex)))


if __name__ == "__main__":
    main()
