#! python3
# -*- coding: utf-8 -*-
"""
Анализ почему элемент не попадает в спецификацию.
Собирает данные элемента и фильтры спецификации для анализа.
"""

__title__ = "Анализ\nэлемента"
__author__ = "CPSK"

import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, ListBox, Panel, TextBox,
    FormStartPosition, FormBorderStyle, DialogResult,
    DockStyle, RichTextBox, ScrollBars, GroupBox,
    ComboBox, ComboBoxStyle
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

from cpsk_notify import show_error, show_warning, show_info

from pyrevit import revit

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, Transaction,
    ScheduleFilterType, StorageType, ElementId, BuiltInCategory
)

doc = revit.doc
uidoc = revit.uidoc


class ScheduleItem(object):
    """Элемент списка спецификаций."""
    def __init__(self, schedule):
        self.schedule = schedule

    def __str__(self):
        return self.schedule.Name


class AnalysisResultDialog(Form):
    """Диалог с результатами анализа."""

    def __init__(self, element_info, schedule_info, analysis_result):
        self.element_info = element_info
        self.schedule_info = schedule_info
        self.analysis_result = analysis_result
        self.setup_form()

    def setup_form(self):
        self.Text = "Результат анализа элемента"
        self.Width = 700
        self.Height = 600
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        # Информация об элементе
        elem_group = GroupBox()
        elem_group.Text = "Информация об элементе"
        elem_group.Location = Point(10, 10)
        elem_group.Size = Size(665, 150)

        elem_text = RichTextBox()
        elem_text.Location = Point(10, 20)
        elem_text.Size = Size(645, 120)
        elem_text.ReadOnly = True
        elem_text.Text = self.element_info
        elem_text.Font = Font("Consolas", 9)
        elem_group.Controls.Add(elem_text)

        # Информация о спецификации
        spec_group = GroupBox()
        spec_group.Text = "Информация о спецификации"
        spec_group.Location = Point(10, 170)
        spec_group.Size = Size(665, 150)

        spec_text = RichTextBox()
        spec_text.Location = Point(10, 20)
        spec_text.Size = Size(645, 120)
        spec_text.ReadOnly = True
        spec_text.Text = self.schedule_info
        spec_text.Font = Font("Consolas", 9)
        spec_group.Controls.Add(spec_text)

        # Результат анализа
        result_group = GroupBox()
        result_group.Text = "Результат анализа"
        result_group.Location = Point(10, 330)
        result_group.Size = Size(665, 180)

        result_text = RichTextBox()
        result_text.Location = Point(10, 20)
        result_text.Size = Size(645, 150)
        result_text.ReadOnly = True
        result_text.Text = self.analysis_result
        result_text.Font = Font("Consolas", 9)
        result_group.Controls.Add(result_text)

        # Кнопка закрытия
        close_btn = Button()
        close_btn.Text = "Закрыть"
        close_btn.Location = Point(580, 520)
        close_btn.Size = Size(95, 30)
        close_btn.DialogResult = DialogResult.OK

        self.Controls.Add(elem_group)
        self.Controls.Add(spec_group)
        self.Controls.Add(result_group)
        self.Controls.Add(close_btn)


class ScheduleSelectionDialog(Form):
    """Диалог выбора спецификации для анализа."""

    def __init__(self, schedules):
        self.schedules = schedules
        self.selected_schedule = None
        self.setup_form()
        self.load_schedules()

    def setup_form(self):
        self.Text = "Выберите спецификацию для анализа"
        self.Width = 500
        self.Height = 350
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        label = Label()
        label.Text = "Выберите спецификацию, в которую элемент должен попадать:"
        label.Location = Point(12, 15)
        label.Size = Size(460, 20)

        self.schedule_combo = ComboBox()
        self.schedule_combo.Location = Point(12, 45)
        self.schedule_combo.Size = Size(460, 25)
        self.schedule_combo.DropDownStyle = ComboBoxStyle.DropDownList

        ok_btn = Button()
        ok_btn.Text = "Анализировать"
        ok_btn.Location = Point(300, 270)
        ok_btn.Size = Size(100, 30)
        ok_btn.Click += self.on_ok

        cancel_btn = Button()
        cancel_btn.Text = "Отмена"
        cancel_btn.Location = Point(410, 270)
        cancel_btn.Size = Size(70, 30)
        cancel_btn.DialogResult = DialogResult.Cancel

        self.Controls.Add(label)
        self.Controls.Add(self.schedule_combo)
        self.Controls.Add(ok_btn)
        self.Controls.Add(cancel_btn)

    def load_schedules(self):
        self.schedule_combo.Items.Clear()
        for schedule in self.schedules:
            self.schedule_combo.Items.Add(ScheduleItem(schedule))
        if self.schedule_combo.Items.Count > 0:
            self.schedule_combo.SelectedIndex = 0

    def on_ok(self, sender, args):
        if self.schedule_combo.SelectedItem:
            self.selected_schedule = self.schedule_combo.SelectedItem.schedule
            self.DialogResult = DialogResult.OK
            self.Close()
        else:
            show_warning("Предупреждение", "Выберите спецификацию для анализа")


def get_element_parameters(element):
    """Получить все параметры элемента."""
    params = {}
    for param in element.Parameters:
        try:
            name = param.Definition.Name
            if param.HasValue:
                if param.StorageType == StorageType.String:
                    params[name] = param.AsString() or ""
                elif param.StorageType == StorageType.Integer:
                    params[name] = str(param.AsInteger())
                elif param.StorageType == StorageType.Double:
                    value_str = param.AsValueString()
                    params[name] = value_str if value_str else str(param.AsDouble())
                elif param.StorageType == StorageType.ElementId:
                    elem_id = param.AsElementId()
                    if elem_id != ElementId.InvalidElementId:
                        linked_elem = doc.GetElement(elem_id)
                        if linked_elem:
                            params[name] = linked_elem.Name
                        else:
                            params[name] = str(elem_id.IntegerValue)
                    else:
                        params[name] = ""
            else:
                params[name] = "(пусто)"
        except Exception:
            pass
    return params


def get_schedule_filters(schedule):
    """Получить фильтры спецификации."""
    filters_info = []
    try:
        schedule_def = schedule.Definition
        filters = schedule_def.GetFilters()

        for i, f in enumerate(filters):
            try:
                field_id = f.FieldId
                field = schedule_def.GetField(field_id)
                field_name = field.GetName() if field else "Неизвестное поле"

                filter_type = f.FilterType
                filter_type_name = str(filter_type).replace("ScheduleFilterType.", "")

                # Получить значение фильтра
                filter_value = ""
                try:
                    if f.IsStringValue:
                        filter_value = f.GetStringValue()
                    elif f.IsDoubleValue:
                        filter_value = str(f.GetDoubleValue())
                    elif f.IsIntegerValue:
                        filter_value = str(f.GetIntegerValue())
                    elif f.IsElementIdValue:
                        filter_value = str(f.GetElementIdValue().IntegerValue)
                except Exception:
                    filter_value = "(не удалось получить)"

                filters_info.append({
                    "index": i,
                    "field": field_name,
                    "type": filter_type_name,
                    "value": filter_value
                })
            except Exception:
                pass
    except Exception:
        pass
    return filters_info


def get_schedule_fields(schedule):
    """Получить поля спецификации."""
    fields_info = []
    try:
        schedule_def = schedule.Definition
        for i in range(schedule_def.GetFieldCount()):
            field = schedule_def.GetField(i)
            fields_info.append({
                "name": field.GetName(),
                "hidden": field.IsHidden
            })
    except Exception:
        pass
    return fields_info


def check_element_in_schedule(element, schedule):
    """Проверить, попадает ли элемент в спецификацию."""
    try:
        collector = FilteredElementCollector(doc, schedule.Id)
        elements_in_schedule = list(collector.ToElementIds())
        return element.Id in elements_in_schedule
    except Exception:
        return False


def analyze_why_not_in_schedule(element, schedule, element_params):
    """Анализ почему элемент не попадает в спецификацию."""
    reasons = []

    # Проверить категорию
    try:
        schedule_def = schedule.Definition
        schedule_cat_id = schedule_def.CategoryId
        element_cat_id = element.Category.Id if element.Category else ElementId.InvalidElementId

        if schedule_cat_id != element_cat_id:
            schedule_cat = doc.GetElement(schedule_cat_id)
            element_cat = element.Category
            reasons.append("КАТЕГОРИЯ: Спецификация для категории '{}', элемент категории '{}'".format(
                schedule_cat.Name if schedule_cat else str(schedule_cat_id.IntegerValue),
                element_cat.Name if element_cat else "Неизвестно"
            ))
    except Exception:
        pass

    # Проверить фильтры
    filters = get_schedule_filters(schedule)
    for f in filters:
        field_name = f["field"]
        filter_type = f["type"]
        filter_value = f["value"]

        # Получить значение параметра элемента
        element_value = element_params.get(field_name, "(параметр отсутствует)")

        # Проверить соответствие фильтру
        match = False
        if filter_type == "Equal":
            match = (element_value == filter_value)
        elif filter_type == "NotEqual":
            match = (element_value != filter_value)
        elif filter_type == "Contains":
            match = (filter_value in element_value)
        elif filter_type == "NotContains":
            match = (filter_value not in element_value)
        elif filter_type == "BeginsWith":
            match = element_value.startswith(filter_value)
        elif filter_type == "EndsWith":
            match = element_value.endswith(filter_value)
        elif filter_type == "GreaterThan":
            try:
                match = float(element_value) > float(filter_value)
            except Exception:
                match = False
        elif filter_type == "LessThan":
            try:
                match = float(element_value) < float(filter_value)
            except Exception:
                match = False
        elif filter_type == "GreaterThanOrEqual":
            try:
                match = float(element_value) >= float(filter_value)
            except Exception:
                match = False
        elif filter_type == "LessThanOrEqual":
            try:
                match = float(element_value) <= float(filter_value)
            except Exception:
                match = False
        elif filter_type == "HasValue":
            match = (element_value and element_value != "(пусто)" and element_value != "(параметр отсутствует)")
        elif filter_type == "HasNoValue":
            match = (not element_value or element_value == "(пусто)" or element_value == "(параметр отсутствует)")

        if not match:
            reasons.append("ФИЛЬТР '{}': {} '{}' - значение элемента: '{}'".format(
                field_name, filter_type, filter_value, element_value
            ))

    if not reasons:
        reasons.append("Не найдено очевидных причин. Возможно:")
        reasons.append("  - Элемент находится в другом виде/фазе")
        reasons.append("  - Элемент скрыт в спецификации")
        reasons.append("  - Спецификация использует сложные фильтры")

    return reasons


def format_element_info(element, params):
    """Форматировать информацию об элементе."""
    lines = []
    lines.append("ID: {}".format(element.Id.IntegerValue))
    lines.append("Категория: {}".format(element.Category.Name if element.Category else "Неизвестно"))

    # Получить имя/тип
    try:
        type_id = element.GetTypeId()
        if type_id != ElementId.InvalidElementId:
            elem_type = doc.GetElement(type_id)
            if elem_type:
                lines.append("Тип: {}".format(elem_type.Name))
    except Exception:
        pass

    lines.append("")
    lines.append("Ключевые параметры:")

    # Показать важные параметры
    important_params = ["Марка", "Форма", "Комментарии", "ADSK_Марка", "ADSK_Позиция",
                       "Длина", "Диаметр стержня", "ADSK_A", "ADSK_B", "ADSK_C"]
    for param_name in important_params:
        if param_name in params:
            lines.append("  {}: {}".format(param_name, params[param_name]))

    return "\n".join(lines)


def format_schedule_info(schedule, filters, fields):
    """Форматировать информацию о спецификации."""
    lines = []
    lines.append("Имя: {}".format(schedule.Name))
    lines.append("ID: {}".format(schedule.Id.IntegerValue))

    try:
        cat_id = schedule.Definition.CategoryId
        cat = doc.GetElement(cat_id)
        lines.append("Категория: {}".format(cat.Name if cat else str(cat_id.IntegerValue)))
    except Exception:
        pass

    lines.append("")
    lines.append("Фильтры ({})".format(len(filters)))
    for f in filters:
        lines.append("  [{}] {} {} '{}'".format(f["index"], f["field"], f["type"], f["value"]))

    if not filters:
        lines.append("  (нет фильтров)")

    lines.append("")
    lines.append("Поля ({})".format(len(fields)))
    visible_fields = [f for f in fields if not f["hidden"]]
    for f in visible_fields[:10]:  # Показать первые 10
        lines.append("  - {}".format(f["name"]))
    if len(visible_fields) > 10:
        lines.append("  ... и еще {}".format(len(visible_fields) - 10))

    return "\n".join(lines)


def main():
    """Основная функция."""
    # Получить выбранный элемент
    selected_ids = uidoc.Selection.GetElementIds()

    if selected_ids.Count == 0:
        show_warning("Выберите элемент", "Пожалуйста, выберите элемент для анализа")
        return

    if selected_ids.Count > 1:
        show_info("Информация", "Выбрано несколько элементов. Будет проанализирован первый.")

    element = doc.GetElement(list(selected_ids)[0])
    if not element:
        show_error("Ошибка", "Не удалось получить выбранный элемент")
        return

    # Получить все спецификации
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    all_schedules = sorted([s for s in collector], key=lambda x: x.Name)

    if not all_schedules:
        show_warning("Нет спецификаций", "В проекте не найдено спецификаций")
        return

    # Показать диалог выбора спецификации
    dialog = ScheduleSelectionDialog(all_schedules)
    result = dialog.ShowDialog()

    if result != DialogResult.OK:
        return

    schedule = dialog.selected_schedule
    if not schedule:
        return

    # Собрать информацию
    element_params = get_element_parameters(element)
    schedule_filters = get_schedule_filters(schedule)
    schedule_fields = get_schedule_fields(schedule)

    # Проверить, попадает ли элемент в спецификацию
    is_in_schedule = check_element_in_schedule(element, schedule)

    # Анализ
    if is_in_schedule:
        analysis_result = "Элемент ПОПАДАЕТ в выбранную спецификацию.\n\nЕсли вы его не видите, возможные причины:\n"
        analysis_result += "- Строка скрыта фильтром ##HIDE_ALL_ROWS##\n"
        analysis_result += "- Элемент сгруппирован с другими элементами\n"
        analysis_result += "- Спецификация не обновлена"
    else:
        reasons = analyze_why_not_in_schedule(element, schedule, element_params)
        analysis_result = "Элемент НЕ ПОПАДАЕТ в спецификацию.\n\nВозможные причины:\n"
        for reason in reasons:
            analysis_result += "- {}\n".format(reason)

    # Форматировать информацию
    element_info = format_element_info(element, element_params)
    schedule_info = format_schedule_info(schedule, schedule_filters, schedule_fields)

    # Показать результат
    result_dialog = AnalysisResultDialog(element_info, schedule_info, analysis_result)
    result_dialog.ShowDialog()


if __name__ == "__main__":
    main()
