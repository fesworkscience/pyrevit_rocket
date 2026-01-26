# -*- coding: utf-8 -*-
"""
Свернуть ВРС - скрытие пустых столбцов в Ведомостях расхода стали.
Скрывает столбцы с диаметрами арматуры, где все значения равны 0.
"""

__title__ = "Свернуть\nВРС"
__author__ = "CPSK"
__doc__ = "Скрывает пустые столбцы с диаметрами в Ведомостях расхода стали"

import clr
import os
import sys
import re

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, CheckedListBox, Panel, ProgressBar,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, AnchorStyles, BorderStyle, FlatStyle,
    Application
)
from System.Drawing import (
    Point, Size, Font, FontStyle, Color,
    ContentAlignment, Drawing2D
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
# Поддерживает: Ø3, Ø4, Ø4.5, Ø9, Ø28, Ø32, Ø36, Ø40 и т.д.
DIAMETER_PATTERN = re.compile(r'^[OØ]\s*\d+(\.\d+)?\s*$', re.IGNORECASE)


class CollapseVRSDialog(Form):
    """Диалог выбора ведомостей расхода стали для сворачивания."""

    # Цветовая схема
    ACCENT_COLOR = Color.FromArgb(0, 122, 204)  # Синий акцент
    BG_COLOR = Color.FromArgb(250, 250, 250)
    HEADER_BG = Color.FromArgb(0, 122, 204)
    HEADER_FG = Color.White

    def __init__(self, schedules):
        self.all_schedules = schedules
        self.selected_schedules = []
        self._schedule_map = {}  # Маппинг имя -> schedule
        self._loading = True  # Флаг загрузки данных
        self.setup_form()
        self.load_schedules()
        self._loading = False

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Свернуть ВРС"
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
        title_label.Text = "Свернуть ведомости расхода стали"
        title_label.Location = Point(20, 12)
        title_label.Size = Size(500, 24)
        title_label.Font = Font("Segoe UI", 14, FontStyle.Bold)
        title_label.ForeColor = self.HEADER_FG
        header_panel.Controls.Add(title_label)

        subtitle_label = Label()
        subtitle_label.Text = "Скрытие пустых столбцов с диаметрами арматуры"
        subtitle_label.Location = Point(20, 36)
        subtitle_label.Size = Size(500, 18)
        subtitle_label.Font = Font("Segoe UI", 9)
        subtitle_label.ForeColor = Color.FromArgb(200, 220, 240)
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
        btn_ok.Text = "Свернуть"
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

        # Обработчик изменения выбора
        self.schedule_list.ItemCheck += self.on_item_check

    def load_schedules(self):
        """Загрузка списка спецификаций."""
        self.schedule_list.Items.Clear()
        self._schedule_map.clear()
        for schedule in self.all_schedules:
            # Используем имя как строку напрямую (IronPython совместимость)
            name = schedule.Name
            self._schedule_map[name] = schedule
            self.schedule_list.Items.Add(name)
            # Выбрать все по умолчанию
            self.schedule_list.SetItemChecked(self.schedule_list.Items.Count - 1, True)
        self.update_count()

    def update_count(self):
        """Обновить счётчик выбранных."""
        checked = self.schedule_list.CheckedItems.Count
        total = self.schedule_list.Items.Count
        self.count_label.Text = "Выбрано: {} из {}".format(checked, total)

    def on_item_check(self, sender, args):
        """Обработчик изменения выбора элемента."""
        # Пропускаем во время загрузки данных
        if self._loading:
            return
        # Используем BeginInvoke для обновления после изменения состояния
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


def is_diameter_column(column_heading):
    """
    Проверить, является ли заголовок столбца столбцом диаметра.
    Ищем паттерны: Ø3, Ø4, Ø4.5, Ø5, ..., Ø40 и т.д.
    """
    if not column_heading:
        return False
    heading = column_heading.strip()
    return DIAMETER_PATTERN.match(heading) is not None


def is_collapsible_column(column_heading):
    """
    Проверить, является ли столбец сворачиваемым.
    Сворачиваемые столбцы: диаметры (Ø...) и столбцы с "Итого".
    """
    if not column_heading:
        return False
    heading = column_heading.strip()

    # Проверяем диаметры
    if DIAMETER_PATTERN.match(heading):
        return True

    # Проверяем "Итого" (в любом регистре)
    if u"итого" in heading.lower():
        return True

    return False


def get_column_total(schedule, column_index):
    """
    Получить сумму значений в столбце спецификации.
    column_index - индекс ВИДИМОГО столбца в таблице (не field_index!)
    Возвращает сумму или None если не удалось получить.
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
            try:
                cell_text = schedule.GetCellText(SectionType.Body, row, column_index)
                if cell_text:
                    # Убираем пробелы и заменяем запятую на точку
                    cell_text = cell_text.strip().replace(',', '.').replace(' ', '')
                    if cell_text and cell_text != '-' and cell_text != '':
                        try:
                            value = float(cell_text)
                            total += value
                        except ValueError:
                            pass
            except:
                pass

        return total
    except Exception as ex:
        return None


def analyze_schedule_columns(schedule):
    """
    Анализировать столбцы спецификации.
    Возвращает список (column_index, column_heading, is_empty, field) для сворачиваемых столбцов.
    Сворачиваемые: диаметры (Ø...) и "Итого".

    ВАЖНО: column_index - это индекс ВИДИМОГО столбца в таблице,
    который отличается от field_index в Definition (скрытые поля сдвигают нумерацию).
    """
    results = []
    try:
        schedule_def = schedule.Definition
        field_count = schedule_def.GetFieldCount()

        # Счётчик видимых столбцов (индекс в таблице)
        visible_column_index = 0

        for i in range(field_count):
            field = schedule_def.GetField(i)

            # Пропускаем скрытые поля - они не отображаются в таблице
            if field.IsHidden:
                continue

            heading = field.ColumnHeading

            if is_collapsible_column(heading):
                # Используем visible_column_index для чтения данных из таблицы
                total = get_column_total(schedule, visible_column_index)
                is_empty = (total is not None and abs(total) < 0.001)
                results.append((visible_column_index, heading, is_empty, field))

            # Увеличиваем счётчик видимых столбцов
            visible_column_index += 1

    except Exception as ex:
        pass

    return results


def collapse_schedule(schedule):
    """
    Свернуть спецификацию - скрыть пустые столбцы с диаметрами.
    Возвращает (hidden_count, total_diameter_columns, error_message).
    """
    try:
        columns = analyze_schedule_columns(schedule)
        if not columns:
            return 0, 0, None

        hidden_count = 0
        schedule_def = schedule.Definition

        for field_index, heading, is_empty, field in columns:
            if is_empty:
                try:
                    field.IsHidden = True
                    hidden_count += 1
                except:
                    pass

        return hidden_count, len(columns), None
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
            # Исключаем шаблоны (обычно содержат "template" или "шаблон")
            if "template" not in name.lower() and "шаблон" not in name.lower():
                vrs_schedules.append(schedule)

    # Сортируем по имени
    vrs_schedules.sort(key=lambda x: x.Name)
    return vrs_schedules


def main():
    """Основная функция."""
    # Найти ведомости расхода стали
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

    # Показать диалог выбора
    dialog = CollapseVRSDialog(vrs_schedules)
    result = dialog.ShowDialog()

    if result != DialogResult.OK:
        return

    selected = dialog.get_selected()
    if not selected:
        return

    # Обработать выбранные ведомости
    with revit.Transaction("Свернуть ВРС"):
        total_hidden = 0
        total_columns = 0
        processed = 0
        errors = []
        details_lines = []

        for schedule in selected:
            hidden, columns, error = collapse_schedule(schedule)

            if error:
                errors.append("{}: {}".format(schedule.Name, error))
            else:
                processed += 1
                total_hidden += hidden
                total_columns += columns
                if hidden > 0:
                    details_lines.append("{}: скрыто {} из {} столбцов".format(
                        schedule.Name, hidden, columns))
                else:
                    details_lines.append("{}: нет пустых столбцов".format(schedule.Name))

        # Показать результат
        if processed > 0:
            message = "Обработано ведомостей: {}\nСкрыто столбцов: {} из {}".format(
                processed, total_hidden, total_columns)

            details = "\n".join(details_lines)
            if errors:
                details += "\n\nОшибки:\n" + "\n".join(errors)

            if total_hidden > 0:
                show_success("Свернуть ВРС", message, details=details)
            else:
                show_info("Свернуть ВРС",
                         "Пустых столбцов с диаметрами не найдено.\n"
                         "Все столбцы содержат данные.",
                         details=details)
        else:
            show_error("Ошибка",
                      "Не удалось обработать ни одну ведомость",
                      details="\n".join(errors) if errors else None)


if __name__ == "__main__":
    main()
