# -*- coding: utf-8 -*-
"""
Заполнение параметров типа значениями из IDS.
Выбор типа → выбор параметра Revit → выбор параметра IDS → выпадающий список.
"""

__title__ = "Заполнить\nтипы"
__author__ = "CPSK"

import clr
import os
import sys
import re
import codecs

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, ComboBox, ListBox, TextBox,
    FormStartPosition, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    DialogResult, OpenFileDialog, GroupBox
)
from System.Drawing import Point, Size, Color

from pyrevit import revit, script

# Добавляем lib в путь для импорта cpsk_config
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Проверка окружения
from cpsk_config import require_environment
if not require_environment():
    sys.exit()

from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, BuiltInCategory,
    StorageType, Element
)

doc = revit.doc
output = script.get_output()


# === ПАРСЕР IDS ===

def parse_ids_for_values(ids_path):
    """
    Парсить IDS и вернуть параметры с допустимыми значениями.
    Возвращает dict: param_name -> {allowed_values: [...], instructions: str}
    """
    result = {}

    try:
        with codecs.open(ids_path, 'r', 'utf-8') as f:
            content = f.read()
    except:
        try:
            with codecs.open(ids_path, 'r', 'utf-8-sig') as f:
                content = f.read()
        except:
            return result

    property_pattern = r'<property([^>]*)>(.*?)</property>'
    properties = re.findall(property_pattern, content, re.DOTALL)

    for prop_match in properties:
        prop_attrs = prop_match[0]
        prop_body = prop_match[1]

        name_match = re.search(r'<baseName>.*?<simpleValue>([^<]+)</simpleValue>', prop_body, re.DOTALL)
        if not name_match:
            continue
        param_name = name_match.group(1).strip()

        allowed_values = []
        value_block = re.search(r'<value>(.*?)</value>', prop_body, re.DOTALL)
        if value_block:
            enum_values = re.findall(r'<xs:enumeration value="([^"]+)"', value_block.group(1))
            for val in enum_values:
                if not val.upper().startswith("IFC"):
                    allowed_values.append(val)

        if allowed_values:
            instr_match = re.search(r'instructions="([^"]*)"', prop_attrs)
            instructions = instr_match.group(1) if instr_match else ""
            instructions = instructions.replace("&#xA;", "\n").replace("&quot;", '"')

            if param_name not in result:
                result[param_name] = {
                    'allowed_values': allowed_values,
                    'instructions': instructions
                }
            else:
                for val in allowed_values:
                    if val not in result[param_name]['allowed_values']:
                        result[param_name]['allowed_values'].append(val)

    return result


def get_all_types(doc):
    """Получить все типы из документа."""
    types = []

    categories = [
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_Floors,
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_StructuralFoundation,
        BuiltInCategory.OST_Rebar,
        BuiltInCategory.OST_Stairs,
        BuiltInCategory.OST_StairsRailing,
        BuiltInCategory.OST_Roofs,
        BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_Ceilings,
    ]

    for cat in categories:
        try:
            collector = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsElementType()
            for t in collector:
                try:
                    name = Element.Name.GetValue(t)
                    types.append((name, t))
                except:
                    pass
        except:
            pass

    types.sort(key=lambda x: x[0].lower())
    return types


def get_type_params(elem_type):
    """Получить редактируемые параметры типа."""
    params = []
    if elem_type is None:
        return params

    for p in elem_type.Parameters:
        if p.IsReadOnly:
            continue
        try:
            name = p.Definition.Name
            current_value = ""
            if p.HasValue:
                if p.StorageType == StorageType.String:
                    current_value = p.AsString() or ""
                elif p.StorageType == StorageType.Integer:
                    current_value = str(p.AsInteger())
                elif p.StorageType == StorageType.Double:
                    current_value = str(round(p.AsDouble(), 4))
            params.append((name, p, current_value))
        except:
            pass

    params.sort(key=lambda x: x[0].lower())
    return params


# === ГЛАВНАЯ ФОРМА ===

class FillTypeParamsForm(Form):
    """Форма заполнения параметров типа."""

    def __init__(self):
        self.ids_path = None
        self.ids_data = {}
        self.all_types = []
        self.filtered_types = []
        self.selected_type = None
        self.type_params = []
        self.filtered_revit_params = []
        self.ids_params = []
        self.filtered_ids_params = []
        self.selected_revit_param = None
        self.selected_ids_param = None
        self.setup_form()
        self.load_types()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Заполнение параметров типа из IDS"
        self.Width = 850
        self.Height = 700
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False

        y = 15

        # === IDS файл ===
        lbl_ids = Label()
        lbl_ids.Text = "IDS файл:"
        lbl_ids.Location = Point(15, y)
        lbl_ids.AutoSize = True
        self.Controls.Add(lbl_ids)

        y += 20
        self.txt_ids = TextBox()
        self.txt_ids.Location = Point(15, y)
        self.txt_ids.Width = 720
        self.txt_ids.ReadOnly = True
        self.Controls.Add(self.txt_ids)

        btn_browse = Button()
        btn_browse.Text = "Обзор..."
        btn_browse.Location = Point(745, y - 2)
        btn_browse.Width = 80
        btn_browse.Click += self.on_browse_ids
        self.Controls.Add(btn_browse)

        y += 35

        # === Выбор типа ===
        grp_type = GroupBox()
        grp_type.Text = "1. Выберите тип"
        grp_type.Location = Point(15, y)
        grp_type.Size = Size(810, 130)

        lbl_search_type = Label()
        lbl_search_type.Text = "Поиск типа:"
        lbl_search_type.Location = Point(10, 22)
        lbl_search_type.AutoSize = True
        grp_type.Controls.Add(lbl_search_type)

        self.txt_search_type = TextBox()
        self.txt_search_type.Location = Point(90, 20)
        self.txt_search_type.Width = 300
        self.txt_search_type.TextChanged += self.on_search_type_changed
        grp_type.Controls.Add(self.txt_search_type)

        self.lst_types = ListBox()
        self.lst_types.Location = Point(10, 50)
        self.lst_types.Size = Size(790, 70)
        self.lst_types.SelectedIndexChanged += self.on_type_selected
        grp_type.Controls.Add(self.lst_types)

        self.Controls.Add(grp_type)

        y += 140

        # === Параметры Revit ===
        grp_revit = GroupBox()
        grp_revit.Text = "2. Параметр Revit (параметры выбранного типа)"
        grp_revit.Location = Point(15, y)
        grp_revit.Size = Size(400, 200)

        lbl_search_revit = Label()
        lbl_search_revit.Text = "Поиск:"
        lbl_search_revit.Location = Point(10, 22)
        lbl_search_revit.AutoSize = True
        grp_revit.Controls.Add(lbl_search_revit)

        self.txt_search_revit = TextBox()
        self.txt_search_revit.Location = Point(60, 20)
        self.txt_search_revit.Width = 325
        self.txt_search_revit.TextChanged += self.on_search_revit_changed
        grp_revit.Controls.Add(self.txt_search_revit)

        self.lst_revit_params = ListBox()
        self.lst_revit_params.Location = Point(10, 50)
        self.lst_revit_params.Size = Size(380, 140)
        self.lst_revit_params.SelectedIndexChanged += self.on_revit_param_selected
        grp_revit.Controls.Add(self.lst_revit_params)

        self.Controls.Add(grp_revit)

        # === Параметры IDS ===
        grp_ids = GroupBox()
        grp_ids.Text = "3. Параметр IDS (с допустимыми значениями)"
        grp_ids.Location = Point(425, y)
        grp_ids.Size = Size(400, 200)

        lbl_search_ids = Label()
        lbl_search_ids.Text = "Поиск:"
        lbl_search_ids.Location = Point(10, 22)
        lbl_search_ids.AutoSize = True
        grp_ids.Controls.Add(lbl_search_ids)

        self.txt_search_ids = TextBox()
        self.txt_search_ids.Location = Point(60, 20)
        self.txt_search_ids.Width = 325
        self.txt_search_ids.TextChanged += self.on_search_ids_changed
        grp_ids.Controls.Add(self.txt_search_ids)

        self.lst_ids_params = ListBox()
        self.lst_ids_params.Location = Point(10, 50)
        self.lst_ids_params.Size = Size(380, 140)
        self.lst_ids_params.SelectedIndexChanged += self.on_ids_param_selected
        grp_ids.Controls.Add(self.lst_ids_params)

        self.Controls.Add(grp_ids)

        y += 210

        # === Значение ===
        grp_value = GroupBox()
        grp_value.Text = "4. Выберите значение"
        grp_value.Location = Point(15, y)
        grp_value.Size = Size(810, 150)

        lbl_revit = Label()
        lbl_revit.Text = "Параметр Revit:"
        lbl_revit.Location = Point(10, 25)
        lbl_revit.AutoSize = True
        grp_value.Controls.Add(lbl_revit)

        self.lbl_selected_revit = Label()
        self.lbl_selected_revit.Text = "(не выбран)"
        self.lbl_selected_revit.Location = Point(120, 25)
        self.lbl_selected_revit.Size = Size(280, 20)
        self.lbl_selected_revit.ForeColor = Color.DarkBlue
        grp_value.Controls.Add(self.lbl_selected_revit)

        lbl_ids_param = Label()
        lbl_ids_param.Text = "Параметр IDS:"
        lbl_ids_param.Location = Point(420, 25)
        lbl_ids_param.AutoSize = True
        grp_value.Controls.Add(lbl_ids_param)

        self.lbl_selected_ids = Label()
        self.lbl_selected_ids.Text = "(не выбран)"
        self.lbl_selected_ids.Location = Point(520, 25)
        self.lbl_selected_ids.Size = Size(280, 20)
        self.lbl_selected_ids.ForeColor = Color.DarkGreen
        grp_value.Controls.Add(self.lbl_selected_ids)

        lbl_current = Label()
        lbl_current.Text = "Текущее значение:"
        lbl_current.Location = Point(10, 50)
        lbl_current.AutoSize = True
        grp_value.Controls.Add(lbl_current)

        self.lbl_current_value = Label()
        self.lbl_current_value.Text = ""
        self.lbl_current_value.Location = Point(130, 50)
        self.lbl_current_value.Size = Size(200, 20)
        self.lbl_current_value.ForeColor = Color.Gray
        grp_value.Controls.Add(self.lbl_current_value)

        lbl_new = Label()
        lbl_new.Text = "Новое значение:"
        lbl_new.Location = Point(350, 50)
        lbl_new.AutoSize = True
        grp_value.Controls.Add(lbl_new)

        self.cmb_value = ComboBox()
        self.cmb_value.Location = Point(460, 47)
        self.cmb_value.Width = 250
        self.cmb_value.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        grp_value.Controls.Add(self.cmb_value)

        self.lbl_values_count = Label()
        self.lbl_values_count.Text = ""
        self.lbl_values_count.Location = Point(720, 50)
        self.lbl_values_count.Size = Size(80, 20)
        self.lbl_values_count.ForeColor = Color.Gray
        grp_value.Controls.Add(self.lbl_values_count)

        lbl_instr = Label()
        lbl_instr.Text = "Описание:"
        lbl_instr.Location = Point(10, 80)
        lbl_instr.AutoSize = True
        grp_value.Controls.Add(lbl_instr)

        self.lbl_instructions = Label()
        self.lbl_instructions.Text = ""
        self.lbl_instructions.Location = Point(10, 100)
        self.lbl_instructions.Size = Size(790, 45)
        self.lbl_instructions.ForeColor = Color.DarkGray
        grp_value.Controls.Add(self.lbl_instructions)

        self.Controls.Add(grp_value)

        y += 160

        # === Кнопки ===
        self.btn_apply = Button()
        self.btn_apply.Text = "Применить"
        self.btn_apply.Location = Point(620, y)
        self.btn_apply.Width = 120
        self.btn_apply.Height = 30
        self.btn_apply.Enabled = False
        self.btn_apply.Click += self.on_apply
        self.Controls.Add(self.btn_apply)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(750, y)
        btn_close.Width = 75
        btn_close.Height = 30
        btn_close.Click += self.on_close
        self.Controls.Add(btn_close)

    def load_types(self):
        """Загрузить все типы."""
        self.all_types = get_all_types(doc)
        self.filtered_types = self.all_types[:]
        self.update_types_list()

    def update_types_list(self):
        """Обновить список типов."""
        self.lst_types.Items.Clear()
        for name, _ in self.filtered_types:
            self.lst_types.Items.Add(name)

    def on_search_type_changed(self, sender, args):
        """Фильтровать типы."""
        search = self.txt_search_type.Text.lower().strip()
        if not search:
            self.filtered_types = self.all_types[:]
        else:
            self.filtered_types = [
                (name, t) for name, t in self.all_types
                if search in name.lower()
            ]
        self.update_types_list()

    def on_type_selected(self, sender, args):
        """При выборе типа - загрузить его параметры."""
        idx = self.lst_types.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_types):
            return

        _, self.selected_type = self.filtered_types[idx]
        self.load_type_params()

    def load_type_params(self):
        """Загрузить параметры выбранного типа."""
        self.type_params = get_type_params(self.selected_type)
        self.filtered_revit_params = self.type_params[:]
        self.update_revit_list()
        self.clear_selection()

    def update_revit_list(self):
        """Обновить список параметров Revit."""
        self.lst_revit_params.Items.Clear()
        for name, _, current in self.filtered_revit_params:
            display = name
            if current:
                display = "{} [{}]".format(name, current[:20])
            self.lst_revit_params.Items.Add(display)

    def on_search_revit_changed(self, sender, args):
        """Фильтровать параметры Revit."""
        search = self.txt_search_revit.Text.lower().strip()
        if not search:
            self.filtered_revit_params = self.type_params[:]
        else:
            self.filtered_revit_params = [
                (name, p, cur) for name, p, cur in self.type_params
                if search in name.lower()
            ]
        self.update_revit_list()

    def on_revit_param_selected(self, sender, args):
        """При выборе параметра Revit."""
        idx = self.lst_revit_params.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_revit_params):
            return

        name, param, current = self.filtered_revit_params[idx]
        self.selected_revit_param = (name, param)
        self.lbl_selected_revit.Text = name
        self.lbl_current_value.Text = current if current else "(пусто)"
        self.update_apply_button()

    def on_browse_ids(self, sender, args):
        """Выбор IDS файла."""
        dialog = OpenFileDialog()
        dialog.Title = "Выберите IDS файл"
        dialog.Filter = "IDS files (*.ids)|*.ids|XML files (*.xml)|*.xml|All files (*.*)|*.*"
        if dialog.ShowDialog() == DialogResult.OK:
            self.ids_path = dialog.FileName
            self.txt_ids.Text = self.ids_path
            self.load_ids()

    def load_ids(self):
        """Загрузить IDS."""
        try:
            self.ids_data = parse_ids_for_values(self.ids_path)
            self.ids_params = [(name, info) for name, info in self.ids_data.items()]
            self.ids_params.sort(key=lambda x: x[0].lower())
            self.filtered_ids_params = self.ids_params[:]
            self.update_ids_list()

            MessageBox.Show(
                "IDS загружен: {} параметров с допустимыми значениями".format(len(self.ids_params)),
                "IDS", MessageBoxButtons.OK, MessageBoxIcon.Information)
        except Exception as e:
            MessageBox.Show(
                "Ошибка загрузки IDS: {}".format(str(e)),
                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)

    def update_ids_list(self):
        """Обновить список параметров IDS."""
        self.lst_ids_params.Items.Clear()
        for name, info in self.filtered_ids_params:
            count = len(info.get('allowed_values', []))
            self.lst_ids_params.Items.Add("{} ({} знач.)".format(name, count))

    def on_search_ids_changed(self, sender, args):
        """Фильтровать параметры IDS."""
        search = self.txt_search_ids.Text.lower().strip()
        if not search:
            self.filtered_ids_params = self.ids_params[:]
        else:
            self.filtered_ids_params = [
                (name, info) for name, info in self.ids_params
                if search in name.lower()
            ]
        self.update_ids_list()

    def on_ids_param_selected(self, sender, args):
        """При выборе параметра IDS - заполнить выпадающий список."""
        idx = self.lst_ids_params.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_ids_params):
            return

        name, info = self.filtered_ids_params[idx]
        self.selected_ids_param = name
        self.lbl_selected_ids.Text = name

        # Заполнить выпадающий список
        self.cmb_value.Items.Clear()
        allowed = info.get('allowed_values', [])
        instructions = info.get('instructions', '')

        self.cmb_value.Items.Add("")
        for val in allowed:
            self.cmb_value.Items.Add(val)
        self.cmb_value.SelectedIndex = 0

        self.lbl_values_count.Text = "{} знач.".format(len(allowed))
        self.lbl_instructions.Text = instructions

        self.update_apply_button()

    def clear_selection(self):
        """Сбросить выбор."""
        self.selected_revit_param = None
        self.selected_ids_param = None
        self.lbl_selected_revit.Text = "(не выбран)"
        self.lbl_selected_ids.Text = "(не выбран)"
        self.lbl_current_value.Text = ""
        self.cmb_value.Items.Clear()
        self.lbl_values_count.Text = ""
        self.lbl_instructions.Text = ""
        self.btn_apply.Enabled = False

    def update_apply_button(self):
        """Активировать кнопку применения."""
        self.btn_apply.Enabled = (
            self.selected_revit_param is not None and
            self.selected_ids_param is not None and
            self.cmb_value.Items.Count > 1
        )

    def on_apply(self, sender, args):
        """Применить выбранное значение."""
        if self.selected_revit_param is None:
            return

        new_value = self.cmb_value.SelectedItem
        if new_value is None:
            return

        param_name, param = self.selected_revit_param

        t = Transaction(doc, "Заполнить параметр типа из IDS")
        t.Start()

        try:
            if param.StorageType == StorageType.String:
                param.Set(str(new_value))
            elif param.StorageType == StorageType.Integer:
                param.Set(int(new_value) if new_value else 0)
            elif param.StorageType == StorageType.Double:
                param.Set(float(new_value) if new_value else 0.0)

            t.Commit()

            # Обновить отображение
            self.lbl_current_value.Text = str(new_value) if new_value else "(пусто)"

            # Обновить список параметров
            self.load_type_params()

            MessageBox.Show(
                "Параметр '{}' установлен: {}".format(param_name, new_value),
                "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information)

        except Exception as e:
            t.RollBack()
            MessageBox.Show(
                "Ошибка: {}".format(str(e)),
                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error)

    def on_close(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    form = FillTypeParamsForm()
    form.ShowDialog()
