# -*- coding: utf-8 -*-
"""
Ведомость сварных швов
Единый интерфейс для работы со сварными швами в модели КМ

Автор: Никита Савков
"""

__title__ = "Сварные швы"
__author__ = "Никита Савков"
__doc__ = "Ведомость сварных швов с возможностью выделения в модели и экспорта в Excel"

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Windows.Forms import (
    Form, DataGridView, DataGridViewSelectionMode, DataGridViewAutoSizeColumnsMode,
    Button, Label, Panel, DockStyle, AnchorStyles, BorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    SaveFileDialog, DialogResult, FormStartPosition, FlatStyle
)
from System.Drawing import Size, Point, Color, Font, FontStyle
from System.Collections.Generic import List

from pyrevit import revit, DB, script
from pyrevit.revit import doc, uidoc

import codecs
import datetime
import os
import sys

# Добавляем lib в путь для импорта cpsk_notify
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from cpsk_notify import show_error, show_warning, show_info, show_success

# ============================================================================
# СБОР ДАННЫХ О СВАРНЫХ ШВАХ
# ============================================================================

def get_param_value(element, param_name, convert_to_mm=False, as_string=False):
    """Получить значение параметра по имени"""
    try:
        param = element.LookupParameter(param_name)
        if param and param.HasValue:
            # Если нужно текстовое представление (для enum-параметров)
            if as_string:
                return param.AsValueString() or str(param.AsInteger()) if param.StorageType == DB.StorageType.Integer else param.AsValueString()
            
            if param.StorageType == DB.StorageType.Double:
                val = param.AsDouble()
                return val * 304.8 if convert_to_mm else val
            elif param.StorageType == DB.StorageType.Integer:
                # Для Integer пробуем сначала AsValueString (enum)
                val_str = param.AsValueString()
                if val_str:
                    return val_str
                return param.AsInteger()
            elif param.StorageType == DB.StorageType.String:
                return param.AsString()
            elif param.StorageType == DB.StorageType.ElementId:
                return param.AsValueString()
    except Exception as e:
        show_warning("Параметр", "Не удалось прочитать параметр: {}".format(param_name), blocking=False, auto_close=1)
    return None


def get_builtin_param(element, builtin_param, convert_to_mm=False, as_string=False):
    """Получить значение встроенного параметра"""
    try:
        param = element.get_Parameter(builtin_param)
        if param and param.HasValue:
            # Если нужно текстовое представление (для enum-параметров)
            if as_string:
                return param.AsValueString()
            
            if param.StorageType == DB.StorageType.Double:
                val = param.AsDouble()
                return val * 304.8 if convert_to_mm else val
            elif param.StorageType == DB.StorageType.Integer:
                # Для Integer пробуем сначала AsValueString (enum)
                val_str = param.AsValueString()
                if val_str:
                    return val_str
                return param.AsInteger()
            elif param.StorageType == DB.StorageType.String:
                return param.AsString()
            elif param.StorageType == DB.StorageType.ElementId:
                return param.AsValueString()
    except Exception as e:
        show_warning("Параметр", "Не удалось прочитать встроенный параметр", blocking=False, auto_close=1)
    return None


def collect_welds():
    """Собрать все сварные швы из модели"""
    welds = []
    
    # Сбор по категории OST_StructConnectionWelds
    try:
        collector = DB.FilteredElementCollector(doc)\
            .OfCategory(DB.BuiltInCategory.OST_StructConnectionWelds)\
            .WhereElementIsNotElementType()
        
        for weld in collector:
            # Получаем параметры из скриншота
            weld_data = {
                'id': weld.Id.IntegerValue,
                'element': weld,
                # Основной тип (С14 - со ступенчатым скосом...)
                'main_type': get_param_value(weld, "Основной тип") or 
                            get_builtin_param(weld, DB.BuiltInParameter.STEEL_ELEM_WELD_MAIN_TYPE) or "-",
                # Основная толщина (катет)
                'thickness': get_param_value(weld, "Основная толщина", convert_to_mm=True) or
                            get_builtin_param(weld, DB.BuiltInParameter.STEEL_ELEM_WELD_MAIN_THICKNESS, convert_to_mm=True),
                # Длина
                'length': get_param_value(weld, "Длина", convert_to_mm=True) or
                         get_builtin_param(weld, DB.BuiltInParameter.CURVE_ELEM_LENGTH, convert_to_mm=True),
                # Местоположение (На заводе / На площадке)
                'location': get_param_value(weld, "Местоположение") or
                           get_builtin_param(weld, DB.BuiltInParameter.STEEL_ELEM_WELD_LOCATION) or "-",
                # Сплошной
                'continuous': get_param_value(weld, "Сплошной"),
                # Шаг
                'pitch': get_param_value(weld, "Шаг", convert_to_mm=True),
                # Форма поверхности (ГОСТ)
                'surface_form': get_param_value(weld, "Форма поверхности") or "-",
                # Двухсторонний тип
                'double_type': get_param_value(weld, "Двухсторонний тип") or "-",
                # Двухсторонняя толщина
                'double_thickness': get_param_value(weld, "Двухсторонняя толщина", convert_to_mm=True),
            }
            
            # Преобразуем Сплошной в текст
            if weld_data['continuous'] is not None:
                weld_data['continuous'] = "Да" if weld_data['continuous'] else "Нет"
            else:
                weld_data['continuous'] = "-"
            
            welds.append(weld_data)
            
    except Exception as e:
        show_warning("Ошибка сбора", "Ошибка сбора по категории: {}".format(str(e)))

    # Также ищем в соединениях
    try:
        conn_collector = DB.FilteredElementCollector(doc)\
            .OfClass(DB.Structure.StructuralConnectionHandler)\
            .WhereElementIsNotElementType()
        
        for connection in conn_collector:
            try:
                subelements = connection.GetSubelements()
                weld_num = 0
                
                for subelem in subelements:
                    try:
                        # Проверяем наличие параметра толщины сварного шва
                        weld_thickness_id = DB.ElementId(DB.BuiltInParameter.STEEL_ELEM_WELD_MAIN_THICKNESS)
                        if subelem.HasParameter(weld_thickness_id):
                            weld_num += 1
                            thickness_param = subelem.GetParameter(weld_thickness_id)
                            thickness = thickness_param.AsDouble() * 304.8 if thickness_param else None
                            
                            weld_data = {
                                'id': connection.Id.IntegerValue,
                                'element': connection,
                                'main_type': "Шов #{} в соединении".format(weld_num),
                                'thickness': thickness,
                                'length': None,
                                'location': "На заводе",
                                'continuous': "-",
                                'pitch': None,
                                'surface_form': "-",
                                'double_type': "-",
                                'double_thickness': None,
                            }
                            welds.append(weld_data)
                    except Exception as e:
                        show_warning("Ошибка", "Ошибка обработки подэлемента", blocking=False, auto_close=1)
                        continue
            except Exception as e:
                show_warning("Ошибка", "Ошибка обработки соединения", blocking=False, auto_close=1)
                continue
    except Exception as e:
        show_warning("Ошибка соединений", "Ошибка сбора соединений: {}".format(str(e)))

    return welds


# ============================================================================
# ФОРМА ИНТЕРФЕЙСА
# ============================================================================

class WeldManagerForm(Form):
    """Форма управления сварными швами"""
    
    def __init__(self):
        self.welds = []
        self.setup_form()
        self.load_data()
    
    def setup_form(self):
        """Настройка формы"""
        self.Text = "Ведомость сварных швов — GIP GROUP"
        self.Size = Size(1100, 600)
        self.MinimumSize = Size(900, 400)
        self.StartPosition = FormStartPosition.CenterScreen
        self.BackColor = Color.White
        
        # Заголовок
        self.header = Label()
        self.header.Text = "Ведомость сварных швов"
        self.header.Font = Font("Segoe UI", 14, FontStyle.Bold)
        self.header.Location = Point(15, 10)
        self.header.Size = Size(400, 30)
        self.header.ForeColor = Color.FromArgb(41, 128, 185)
        self.Controls.Add(self.header)
        
        # Счётчик
        self.counter_label = Label()
        self.counter_label.Text = "Найдено: 0 швов"
        self.counter_label.Font = Font("Segoe UI", 10)
        self.counter_label.Location = Point(15, 42)
        self.counter_label.Size = Size(300, 20)
        self.counter_label.ForeColor = Color.Gray
        self.Controls.Add(self.counter_label)
        
        # Таблица
        self.grid = DataGridView()
        self.grid.Location = Point(15, 70)
        self.grid.Size = Size(1050, 430)
        self.grid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self.grid.MultiSelect = False
        self.grid.ReadOnly = True
        self.grid.AllowUserToAddRows = False
        self.grid.AllowUserToDeleteRows = False
        self.grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.grid.BackgroundColor = Color.White
        self.grid.BorderStyle = BorderStyle.FixedSingle
        self.grid.RowHeadersVisible = False
        self.grid.Font = Font("Segoe UI", 9)
        
        # Колонки
        self.grid.Columns.Add("num", "№")
        self.grid.Columns.Add("id", "ID")
        self.grid.Columns.Add("main_type", "Тип шва")
        self.grid.Columns.Add("thickness", "Толщина (мм)")
        self.grid.Columns.Add("length", "Длина (мм)")
        self.grid.Columns.Add("location", "Местоположение")
        self.grid.Columns.Add("continuous", "Сплошной")
        self.grid.Columns.Add("pitch", "Шаг (мм)")
        self.grid.Columns.Add("surface_form", "Форма поверхности")
        self.grid.Columns.Add("double_type", "Двухсторонний")
        
        # Ширина колонок
        self.grid.Columns["num"].Width = 40
        self.grid.Columns["id"].Width = 70
        self.grid.Columns["main_type"].Width = 200
        self.grid.Columns["thickness"].Width = 90
        self.grid.Columns["length"].Width = 90
        self.grid.Columns["location"].Width = 100
        self.grid.Columns["continuous"].Width = 70
        self.grid.Columns["pitch"].Width = 70
        self.grid.Columns["surface_form"].Width = 120
        self.grid.Columns["double_type"].Width = 100
        
        # Событие выбора строки
        self.grid.SelectionChanged += self.on_selection_changed
        self.grid.CellDoubleClick += self.on_cell_double_click
        
        self.Controls.Add(self.grid)
        
        # Панель кнопок
        self.button_panel = Panel()
        self.button_panel.Location = Point(15, 510)
        self.button_panel.Size = Size(1050, 45)
        self.button_panel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.Controls.Add(self.button_panel)
        
        # Кнопка "Выделить в модели"
        self.btn_select = Button()
        self.btn_select.Text = "Выделить в модели"
        self.btn_select.Size = Size(150, 35)
        self.btn_select.Location = Point(0, 5)
        self.btn_select.BackColor = Color.FromArgb(41, 128, 185)
        self.btn_select.ForeColor = Color.White
        self.btn_select.FlatStyle = FlatStyle.Flat  # Flat
        self.btn_select.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.btn_select.Click += self.on_select_click
        self.button_panel.Controls.Add(self.btn_select)
        
        # Кнопка "Выделить все"
        self.btn_select_all = Button()
        self.btn_select_all.Text = "Выделить все"
        self.btn_select_all.Size = Size(120, 35)
        self.btn_select_all.Location = Point(160, 5)
        self.btn_select_all.BackColor = Color.FromArgb(52, 152, 219)
        self.btn_select_all.ForeColor = Color.White
        self.btn_select_all.FlatStyle = FlatStyle.Flat
        self.btn_select_all.Font = Font("Segoe UI", 9)
        self.btn_select_all.Click += self.on_select_all_click
        self.button_panel.Controls.Add(self.btn_select_all)
        
        # Кнопка "Экспорт в Excel"
        self.btn_export = Button()
        self.btn_export.Text = "Экспорт в Excel"
        self.btn_export.Size = Size(130, 35)
        self.btn_export.Location = Point(290, 5)
        self.btn_export.BackColor = Color.FromArgb(39, 174, 96)
        self.btn_export.ForeColor = Color.White
        self.btn_export.FlatStyle = FlatStyle.Flat
        self.btn_export.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.btn_export.Click += self.on_export_click
        self.button_panel.Controls.Add(self.btn_export)
        
        # Кнопка "Обновить"
        self.btn_refresh = Button()
        self.btn_refresh.Text = "Обновить"
        self.btn_refresh.Size = Size(100, 35)
        self.btn_refresh.Location = Point(430, 5)
        self.btn_refresh.BackColor = Color.FromArgb(149, 165, 166)
        self.btn_refresh.ForeColor = Color.White
        self.btn_refresh.FlatStyle = FlatStyle.Flat
        self.btn_refresh.Font = Font("Segoe UI", 9)
        self.btn_refresh.Click += self.on_refresh_click
        self.button_panel.Controls.Add(self.btn_refresh)
        
        # Подсказка
        self.hint_label = Label()
        self.hint_label.Text = "💡 Двойной клик — показать элемент в модели"
        self.hint_label.Font = Font("Segoe UI", 8)
        self.hint_label.Location = Point(550, 15)
        self.hint_label.Size = Size(300, 20)
        self.hint_label.ForeColor = Color.Gray
        self.button_panel.Controls.Add(self.hint_label)
    
    def load_data(self):
        """Загрузка данных в таблицу"""
        self.grid.Rows.Clear()
        self.welds = collect_welds()
        
        for i, weld in enumerate(self.welds, 1):
            self.grid.Rows.Add(
                str(i),
                str(weld['id']),
                str(weld['main_type']),
                self.format_value(weld['thickness']),
                self.format_value(weld['length']),
                str(weld['location']),
                str(weld['continuous']),
                self.format_value(weld['pitch']),
                str(weld['surface_form']),
                str(weld['double_type'])
            )
        
        self.counter_label.Text = "Найдено: {} швов".format(len(self.welds))
        
        if not self.welds:
            show_info(
                "Результат поиска",
                "Сварные швы не найдены в модели.\n\n"
                "Убедитесь, что:\n"
                "• В модели есть элементы категории 'Сварные швы'\n"
                "• Используются соединения со вкладки 'Сталь'"
            )
    
    def format_value(self, val):
        """Форматирование значения"""
        if val is None:
            return "-"
        if isinstance(val, float):
            return "{:.1f}".format(val)
        return str(val)
    
    def get_selected_weld(self):
        """Получить выбранный сварной шов"""
        if self.grid.SelectedRows.Count > 0:
            row_index = self.grid.SelectedRows[0].Index
            if row_index < len(self.welds):
                return self.welds[row_index]
        return None
    
    def on_selection_changed(self, sender, args):
        """Обработка изменения выбора"""
        weld = self.get_selected_weld()
        if weld:
            self.btn_select.Enabled = True
        else:
            self.btn_select.Enabled = False
    
    def on_cell_double_click(self, sender, args):
        """Двойной клик — показать элемент"""
        weld = self.get_selected_weld()
        if weld and weld.get('element'):
            try:
                element_id = DB.ElementId(weld['id'])
                uidoc.ShowElements(element_id)
                uidoc.Selection.SetElementIds(List[DB.ElementId]([element_id]))
            except Exception as e:
                show_error("Ошибка", "Не удалось показать элемент", details=str(e))

    def on_select_click(self, sender, args):
        """Выделить выбранный элемент"""
        weld = self.get_selected_weld()
        if weld:
            try:
                element_id = DB.ElementId(weld['id'])
                uidoc.Selection.SetElementIds(List[DB.ElementId]([element_id]))
                show_success("Готово", "Элемент ID {} выделен в модели".format(weld['id']))
            except Exception as e:
                show_error("Ошибка", "Не удалось выделить элемент", details=str(e))

    def on_select_all_click(self, sender, args):
        """Выделить все сварные швы"""
        if not self.welds:
            return
        
        try:
            ids = list(set([w['id'] for w in self.welds]))
            element_ids = List[DB.ElementId]([DB.ElementId(i) for i in ids])
            uidoc.Selection.SetElementIds(element_ids)
            show_success("Готово", "Выделено {} элементов".format(len(ids)))
        except Exception as e:
            show_error("Ошибка", "Не удалось выделить элементы", details=str(e))

    def on_export_click(self, sender, args):
        """Экспорт в Excel (CSV)"""
        if not self.welds:
            show_warning("Внимание", "Нет данных для экспорта")
            return
        
        # Диалог сохранения
        dialog = SaveFileDialog()
        dialog.Filter = "CSV файл (*.csv)|*.csv"
        dialog.Title = "Сохранить ведомость сварных швов"
        dialog.FileName = "{}_Сварные_швы_{}.csv".format(
            doc.Title.replace(" ", "_").replace(".", "_"),
            datetime.datetime.now().strftime("%Y%m%d_%H%M")
        )
        
        if doc.PathName:
            dialog.InitialDirectory = os.path.dirname(doc.PathName)
        
        if dialog.ShowDialog() == DialogResult.OK:
            try:
                self.export_to_csv(dialog.FileName)
                show_success(
                    "Экспорт завершён",
                    "Файл сохранён:\n{}\n\nОткройте в Excel, разделитель — точка с запятой (;)".format(dialog.FileName)
                )

                # Предложение открыть файл
                if MessageBox.Show("Открыть файл?", "Открыть",
                                  MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes:
                    os.startfile(dialog.FileName)

            except Exception as e:
                show_error("Ошибка", "Ошибка экспорта", details=str(e))
    
    def export_to_csv(self, filepath):
        """Экспорт данных в CSV"""
        with codecs.open(filepath, 'w', 'utf-8-sig') as f:
            # Заголовок
            f.write("№;ID;Тип шва;Толщина (мм);Длина (мм);Местоположение;Сплошной;Шаг (мм);Форма поверхности;Двухсторонний\n")
            
            for i, w in enumerate(self.welds, 1):
                row = [
                    str(i),
                    str(w['id']),
                    str(w['main_type']),
                    self.format_value(w['thickness']),
                    self.format_value(w['length']),
                    str(w['location']),
                    str(w['continuous']),
                    self.format_value(w['pitch']),
                    str(w['surface_form']),
                    str(w['double_type'])
                ]
                f.write(";".join(row) + "\n")
            
            # Статистика
            f.write("\n")
            f.write("ИТОГО сварных швов:;{}\n".format(len(self.welds)))
            
            # По толщинам
            thickness_stats = {}
            for w in self.welds:
                t = w.get('thickness')
                if t is not None:
                    key = round(t, 1)
                    thickness_stats[key] = thickness_stats.get(key, 0) + 1
            
            if thickness_stats:
                f.write("\nСтатистика по толщинам:\n")
                f.write("Толщина (мм);Количество\n")
                for t in sorted(thickness_stats.keys()):
                    f.write("{:.1f};{}\n".format(t, thickness_stats[t]))
            
            # По местоположению
            loc_stats = {}
            for w in self.welds:
                loc = w.get('location', '-')
                loc_stats[loc] = loc_stats.get(loc, 0) + 1
            
            if loc_stats:
                f.write("\nПо местоположению:\n")
                f.write("Местоположение;Количество\n")
                for loc, count in loc_stats.items():
                    f.write("{};{}\n".format(loc, count))
    
    def on_refresh_click(self, sender, args):
        """Обновить данные"""
        self.load_data()


# ============================================================================
# ЗАПУСК
# ============================================================================

def main():
    form = WeldManagerForm()
    form.ShowDialog()

if __name__ == "__main__":
    main()
