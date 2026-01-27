# -*- coding: utf-8 -*-
"""
Развернуть ВРС - показать скрытые столбцы с данными в Ведомостях расхода стали.
Показывает столбцы с диаметрами арматуры, где появились значения > 0.
"""

__title__ = "Развернуть\nВРС"
__author__ = "CPSK"
__doc__ = "Показывает скрытые столбцы с данными в Ведомостях расхода стали"

import clr
import os
import sys
import re

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, CheckedListBox, Panel,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, AnchorStyles, BorderStyle, FlatStyle,
    Application
)
from System.Drawing import (
    Point, Size, Font, FontStyle, Color,
    ContentAlignment
)

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from cpsk_auth import require_auth
if not require_auth():
    import sys
    sys.exit()

from cpsk_notify import show_error, show_warning, show_success, show_info

from pyrevit import revit

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSchedule, Transaction,
    ScheduleFieldType, SectionType
)

doc = revit.doc

# Паттерн для поиска столбцов с диаметрами (Ø + число)
DIAMETER_PATTERN = re.compile(r'^[OØ]\s*\d+(\.\d+)?\s*$', re.IGNORECASE)


class ExpandVRSDialog(Form):
    """Диалог выбора ведомостей расхода стали для разворачивания."""

    # Цветовая схема
    ACCENT_COLOR = Color.FromArgb(46, 139, 87)  # Зелёный акцент
    BG_COLOR = Color.FromArgb(250, 250, 250)
    HEADER_BG = Color.FromArgb(46, 139, 87)
    HEADER_FG = Color.White

    def __init__(self, schedules):
        self.all_schedules = schedules
        self.selected_schedules = []
        self._schedule_map = {}
        self._loading = True
        self.setup_form()
        self.load_schedules()
        self._loading = False

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Развернуть ВРС"
        self.Width = 550
        self.Height = 510
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.BackColor = self.BG_COLOR

        # === Заголовок ===
        header_panel = Panel()
        header_panel.Location = Point(0, 0)
        header_panel.Size = Size(550, 60)
        header_panel.BackColor = self.HEADER_BG

        title_label = Label()
        title_label.Text = "Развернуть ведомости расхода стали"
        title_label.Location = Point(20, 12)
        title_label.Size = Size(500, 24)
        title_label.Font = Font("Segoe UI", 14, FontStyle.Bold)
        title_label.ForeColor = self.HEADER_FG
        header_panel.Controls.Add(title_label)

        subtitle_label = Label()
        subtitle_label.Text = "Показать скрытые столбцы, в которых появились данные"
        subtitle_label.Location = Point(20, 36)
        subtitle_label.Size = Size(500, 18)
        subtitle_label.Font = Font("Segoe UI", 9)
        subtitle_label.ForeColor = Color.FromArgb(200, 230, 210)
        header_panel.Controls.Add(subtitle_label)

        self.Controls.Add(header_panel)

        # === Инструкция ===
        instruction_label = Label()
        instruction_label.Text = "Выберите ведомости для обработки:"
        instruction_label.Location = Point(20, 75)
        instruction_label.Size = Size(500, 20)
        instruction_label.Font = Font("Segoe UI", 10)
        instruction_label.ForeColor = Color.FromArgb(60, 60, 60)
        self.Controls.Add(instruction_label)

        # === Список спецификаций ===
        self.schedule_list = CheckedListBox()
        self.schedule_list.Location = Point(20, 100)
        self.schedule_list.Size = Size(495, 260)
        self.schedule_list.CheckOnClick = True
        self.schedule_list.Font = Font("Segoe UI", 9)
        self.schedule_list.BorderStyle = BorderStyle.FixedSingle
        self.Controls.Add(self.schedule_list)

        # === Панель кнопок выбора ===
        btn_select_all = Button()
        btn_select_all.Text = "Выбрать все"
        btn_select_all.Location = Point(20, 370)
        btn_select_all.Size = Size(100, 28)
        btn_select_all.FlatStyle = FlatStyle.Flat
        btn_select_all.FlatAppearance.BorderColor = self.ACCENT_COLOR
        btn_select_all.ForeColor = self.ACCENT_COLOR
        btn_select_all.Font = Font("Segoe UI", 9)
        btn_select_all.Click += self.on_select_all
        self.Controls.Add(btn_select_all)

        btn_select_none = Button()
        btn_select_none.Text = "Снять все"
        btn_select_none.Location = Point(130, 370)
        btn_select_none.Size = Size(100, 28)
        btn_select_none.FlatStyle = FlatStyle.Flat
        btn_select_none.FlatAppearance.BorderColor = Color.Gray
        btn_select_none.ForeColor = Color.Gray
        btn_select_none.Font = Font("Segoe UI", 9)
        btn_select_none.Click += self.on_select_none
        self.Controls.Add(btn_select_none)

        # === Счётчик ===
        self.count_label = Label()
        self.count_label.Location = Point(250, 374)
        self.count_label.Size = Size(150, 20)
        self.count_label.Font = Font("Segoe UI", 9)
        self.count_label.ForeColor = Color.Gray
        self.count_label.TextAlign = ContentAlignment.MiddleLeft
        self.Controls.Add(self.count_label)

        # === Кнопки OK/Отмена ===
        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(315, 410)
        btn_cancel.Size = Size(95, 32)
        btn_cancel.FlatStyle = FlatStyle.Flat
        btn_cancel.FlatAppearance.BorderColor = Color.Gray
        btn_cancel.Font = Font("Segoe UI", 10)
        btn_cancel.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_cancel)

        btn_ok = Button()
        btn_ok.Text = "Развернуть"
        btn_ok.Location = Point(420, 410)
        btn_ok.Size = Size(95, 32)
        btn_ok.FlatStyle = FlatStyle.Flat
        btn_ok.BackColor = self.ACCENT_COLOR
        btn_ok.ForeColor = Color.White
        btn_ok.FlatAppearance.BorderSize = 0
        btn_ok.Font = Font("Segoe UI", 10, FontStyle.Bold)
        btn_ok.Click += self.on_ok
        self.Controls.Add(btn_ok)

        self.AcceptButton = btn_ok
        self.CancelButton = btn_cancel

        self.schedule_list.ItemCheck += self.on_item_check

    def load_schedules(self):
        """Загрузка списка спецификаций."""
        self.schedule_list.Items.Clear()
        self._schedule_map.clear()
        for schedule in self.all_schedules:
            name = schedule.Name
            self._schedule_map[name] = schedule
            self.schedule_list.Items.Add(name)
            self.schedule_list.SetItemChecked(self.schedule_list.Items.Count - 1, True)
        self.update_count()

    def update_count(self):
        """Обновить счётчик выбранных."""
        checked = self.schedule_list.CheckedItems.Count
        total = self.schedule_list.Items.Count
        self.count_label.Text = "Выбрано: {} из {}".format(checked, total)

    def on_item_check(self, sender, args):
        """Обработчик изменения выбора элемента."""
        if self._loading:
            return
        if self.IsHandleCreated:
            self.BeginInvoke(System.Action(self.update_count))

    def on_select_all(self, sender, args):
        """Выбрать все."""
        for i in range(self.schedule_list.Items.Count):
            self.schedule_list.SetItemChecked(i, True)
        self.update_count()

    def on_select_none(self, sender, args):
        """Снять все."""
        for i in range(self.schedule_list.Items.Count):
            self.schedule_list.SetItemChecked(i, False)
        self.update_count()

    def on_ok(self, sender, args):
        """Обработка нажатия OK."""
        self.selected_schedules = []
        for i in range(self.schedule_list.Items.Count):
            if self.schedule_list.GetItemChecked(i):
                name = self.schedule_list.Items[i]
                if name in self._schedule_map:
                    self.selected_schedules.append(self._schedule_map[name])

        if not self.selected_schedules:
            show_warning("Предупреждение", "Выберите хотя бы одну ведомость")
            return

        self.DialogResult = DialogResult.OK
        self.Close()

    def get_selected(self):
        """Получить выбранные спецификации."""
        return self.selected_schedules


def is_collapsible_column(column_heading):
    """
    Проверить, является ли столбец сворачиваемым.
    Сворачиваемые столбцы: диаметры (Ø...) и столбцы с "Итого".
    """
    if not column_heading:
        return False
    heading = column_heading.strip()

    if DIAMETER_PATTERN.match(heading):
        return True

    if u"итого" in heading.lower():
        return True

    return False


def get_column_total(schedule, column_index):
    """
    Получить сумму значений в столбце спецификации.
    column_index - индекс ВИДИМОГО столбца в таблице.
    """
    try:
        table_data = schedule.GetTableData()
        section_data = table_data.GetSectionData(SectionType.Body)

        total = 0.0
        rows = section_data.NumberOfRows
        cols = section_data.NumberOfColumns

        if column_index >= cols:
            return None

        for row in range(rows):
            cell_text = schedule.GetCellText(SectionType.Body, row, column_index)
            if cell_text:
                cell_text = cell_text.strip().replace(',', '.').replace(' ', '')
                if cell_text and cell_text != '-' and cell_text != '':
                    try:
                        value = float(cell_text)
                        total += value
                    except ValueError:
                        # Нечисловое значение - пропускаем (ожидаемо для текстовых ячеек)
                        continue

        return total
    except Exception as ex:
        return None


def expand_schedule(schedule):
    """
    Развернуть спецификацию - показать скрытые столбцы с данными.
    Возвращает (shown_count, total_hidden_columns, error_message).
    """
    try:
        schedule_def = schedule.Definition
        field_count = schedule_def.GetFieldCount()

        # Собираем скрытые поля с диаметрами/итого
        hidden_fields = []
        for i in range(field_count):
            field = schedule_def.GetField(i)
            if field.IsHidden:
                heading = field.ColumnHeading
                if is_collapsible_column(heading):
                    hidden_fields.append((i, heading, field))

        if not hidden_fields:
            return 0, 0, None

        shown_count = 0

        # Для каждого скрытого поля: показываем, проверяем данные
        for field_index, heading, field in hidden_fields:
            # Показываем поле
            field.IsHidden = False

            # Теперь нужно найти его индекс в таблице
            # Пересчитываем видимые столбцы
            visible_column_index = 0
            for j in range(field_count):
                f = schedule_def.GetField(j)
                if j == field_index:
                    break
                if not f.IsHidden:
                    visible_column_index += 1

            # Проверяем сумму
            total = get_column_total(schedule, visible_column_index)

            if total is not None and abs(total) < 0.001:
                # Данных нет - скрываем обратно
                field.IsHidden = True
            else:
                # Данные есть - оставляем видимым
                shown_count += 1

        return shown_count, len(hidden_fields), None
    except Exception as ex:
        return 0, 0, str(ex)


def find_vrs_schedules():
    """
    Найти все ведомости расхода стали в проекте.
    """
    collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
    vrs_schedules = []

    for schedule in collector:
        name = schedule.Name
        if "Ведомость расхода стали" in name or "ВРС" in name.upper():
            if "template" not in name.lower() and "шаблон" not in name.lower():
                vrs_schedules.append(schedule)

    vrs_schedules.sort(key=lambda x: x.Name)
    return vrs_schedules


def main():
    """Основная функция."""
    vrs_schedules = find_vrs_schedules()

    if not vrs_schedules:
        show_info(
            "Информация",
            "В проекте не найдено ведомостей расхода стали.\n\n"
            "Поиск выполнялся по названиям, содержащим:\n"
            "- 'Ведомость расхода стали'\n"
            "- 'ВРС'"
        )
        return

    dialog = ExpandVRSDialog(vrs_schedules)
    result = dialog.ShowDialog()

    if result != DialogResult.OK:
        return

    selected = dialog.get_selected()
    if not selected:
        return

    with revit.Transaction("Развернуть ВРС"):
        total_shown = 0
        total_hidden = 0
        processed = 0
        errors = []
        details_lines = []

        for schedule in selected:
            shown, hidden_count, error = expand_schedule(schedule)

            if error:
                errors.append("{}: {}".format(schedule.Name, error))
            else:
                processed += 1
                total_shown += shown
                total_hidden += hidden_count
                if shown > 0:
                    details_lines.append("{}: показано {} из {} скрытых столбцов".format(
                        schedule.Name, shown, hidden_count))
                elif hidden_count > 0:
                    details_lines.append("{}: нет данных в скрытых столбцах".format(schedule.Name))
                else:
                    details_lines.append("{}: нет скрытых столбцов".format(schedule.Name))

        if processed > 0:
            message = "Обработано ведомостей: {}\nПоказано столбцов: {} из {}".format(
                processed, total_shown, total_hidden)

            details = "\n".join(details_lines)
            if errors:
                details += "\n\nОшибки:\n" + "\n".join(errors)

            if total_shown > 0:
                show_success("Развернуть ВРС", message, details=details)
            else:
                show_info("Развернуть ВРС",
                         "Нет скрытых столбцов с данными.\n"
                         "Все скрытые столбцы пусты.",
                         details=details)
        else:
            show_error("Ошибка",
                      "Не удалось обработать ни одну ведомость",
                      details="\n".join(errors) if errors else None)


if __name__ == "__main__":
    main()
