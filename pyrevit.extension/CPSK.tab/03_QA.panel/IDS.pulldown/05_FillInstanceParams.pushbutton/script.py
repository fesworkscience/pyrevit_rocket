#! python3
# -*- coding: utf-8 -*-
"""
Заполнение параметров экземпляров значениями из IDS.
Выбор элементов → выбор параметра Revit → выбор параметра IDS → выпадающий список.
"""

__title__ = "Заполнить\nэкземпляры"
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
    DialogResult, OpenFileDialog, GroupBox
)
from System.Drawing import Point, Size, Color

from pyrevit import revit, script

# Добавляем lib в путь для импорта cpsk_config
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Импорт модулей из lib
from cpsk_notify import show_error, show_warning, show_info, show_success, show_confirm
from cpsk_auth import require_auth
from cpsk_config import require_environment

# Проверка авторизации
if not require_auth():
    sys.exit()

# Проверка окружения
if not require_environment():
    sys.exit()

from Autodesk.Revit.DB import Transaction, StorageType
from Autodesk.Revit.UI.Selection import ObjectType

doc = revit.doc
uidoc = revit.uidoc
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


def get_selected_elements():
    """Получить выбранные элементы."""
    selection = uidoc.Selection.GetElementIds()
    elements = [doc.GetElement(eid) for eid in selection]
    return elements


def get_common_params(elements):
    """Получить общие параметры для всех элементов (только редактируемые)."""
    if not elements:
        return []

    # Собрать параметры первого элемента
    first_params = {}
    for p in elements[0].Parameters:
        if p.IsReadOnly:
            continue
        try:
            name = p.Definition.Name
            first_params[name] = p.StorageType
        except:
            pass

    # Проверить что параметры есть во всех элементах
    common = []
    for name, storage_type in first_params.items():
        is_common = True
        for elem in elements[1:]:
            param = elem.LookupParameter(name)
            if param is None or param.IsReadOnly:
                is_common = False
                break
        if is_common:
            common.append(name)

    common.sort(key=lambda x: x.lower())
    return common


# === ГЛАВНАЯ ФОРМА ===

class FillInstanceParamsForm(Form):
    """Форма заполнения параметров экземпляров."""

    def __init__(self, elements):
        self.elements = elements
        self.ids_path = None
        self.ids_data = {}
        self.common_params = []
        self.ids_params = []  # [(name, info)]
        self.selected_revit_param = None
        self.selected_ids_param = None
        self.setup_form()
        self.load_params()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Заполнение параметров экземпляров из IDS"
        self.Width = 800
        self.Height = 620
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False

        y = 15

        # === Статус выбора ===
        self.lbl_selection = Label()
        self.lbl_selection.Text = "Выбрано элементов: {}".format(len(self.elements))
        self.lbl_selection.Location = Point(15, y)
        self.lbl_selection.Size = Size(400, 20)
        self.lbl_selection.ForeColor = Color.DarkBlue
        self.Controls.Add(self.lbl_selection)

        y += 25

        # === IDS файл ===
        lbl_ids = Label()
        lbl_ids.Text = "IDS файл:"
        lbl_ids.Location = Point(15, y)
        lbl_ids.AutoSize = True
        self.Controls.Add(lbl_ids)

        y += 20
        self.txt_ids = TextBox()
        self.txt_ids.Location = Point(15, y)
        self.txt_ids.Width = 670
        self.txt_ids.ReadOnly = True
        self.Controls.Add(self.txt_ids)

        btn_browse = Button()
        btn_browse.Text = "Обзор..."
        btn_browse.Location = Point(695, y - 2)
        btn_browse.Width = 80
        btn_browse.Click += self.on_browse_ids
        self.Controls.Add(btn_browse)

        y += 35

        # === Параметры Revit ===
        grp_revit = GroupBox()
        grp_revit.Text = "1. Параметр Revit (общий для выбранных элементов)"
        grp_revit.Location = Point(15, y)
        grp_revit.Size = Size(370, 220)

        lbl_search_revit = Label()
        lbl_search_revit.Text = "Поиск:"
        lbl_search_revit.Location = Point(10, 22)
        lbl_search_revit.AutoSize = True
        grp_revit.Controls.Add(lbl_search_revit)

        self.txt_search_revit = TextBox()
        self.txt_search_revit.Location = Point(60, 20)
        self.txt_search_revit.Width = 295
        self.txt_search_revit.TextChanged += self.on_search_revit_changed
        grp_revit.Controls.Add(self.txt_search_revit)

        self.lst_revit_params = ListBox()
        self.lst_revit_params.Location = Point(10, 50)
        self.lst_revit_params.Size = Size(350, 160)
        self.lst_revit_params.SelectedIndexChanged += self.on_revit_param_selected
        grp_revit.Controls.Add(self.lst_revit_params)

        self.Controls.Add(grp_revit)

        # === Параметры IDS ===
        grp_ids = GroupBox()
        grp_ids.Text = "2. Параметр IDS (с допустимыми значениями)"
        grp_ids.Location = Point(395, y)
        grp_ids.Size = Size(380, 220)

        lbl_search_ids = Label()
        lbl_search_ids.Text = "Поиск:"
        lbl_search_ids.Location = Point(10, 22)
        lbl_search_ids.AutoSize = True
        grp_ids.Controls.Add(lbl_search_ids)

        self.txt_search_ids = TextBox()
        self.txt_search_ids.Location = Point(60, 20)
        self.txt_search_ids.Width = 305
        self.txt_search_ids.TextChanged += self.on_search_ids_changed
        grp_ids.Controls.Add(self.txt_search_ids)

        self.lst_ids_params = ListBox()
        self.lst_ids_params.Location = Point(10, 50)
        self.lst_ids_params.Size = Size(360, 160)
        self.lst_ids_params.SelectedIndexChanged += self.on_ids_param_selected
        grp_ids.Controls.Add(self.lst_ids_params)

        self.Controls.Add(grp_ids)

        y += 230

        # === Значение ===
        grp_value = GroupBox()
        grp_value.Text = "3. Выберите значение"
        grp_value.Location = Point(15, y)
        grp_value.Size = Size(760, 150)

        lbl_revit = Label()
        lbl_revit.Text = "Параметр Revit:"
        lbl_revit.Location = Point(10, 25)
        lbl_revit.AutoSize = True
        grp_value.Controls.Add(lbl_revit)

        self.lbl_selected_revit = Label()
        self.lbl_selected_revit.Text = "(не выбран)"
        self.lbl_selected_revit.Location = Point(120, 25)
        self.lbl_selected_revit.Size = Size(250, 20)
        self.lbl_selected_revit.ForeColor = Color.DarkBlue
        grp_value.Controls.Add(self.lbl_selected_revit)

        lbl_ids_param = Label()
        lbl_ids_param.Text = "Параметр IDS:"
        lbl_ids_param.Location = Point(380, 25)
        lbl_ids_param.AutoSize = True
        grp_value.Controls.Add(lbl_ids_param)

        self.lbl_selected_ids = Label()
        self.lbl_selected_ids.Text = "(не выбран)"
        self.lbl_selected_ids.Location = Point(480, 25)
        self.lbl_selected_ids.Size = Size(260, 20)
        self.lbl_selected_ids.ForeColor = Color.DarkGreen
        grp_value.Controls.Add(self.lbl_selected_ids)

        lbl_value = Label()
        lbl_value.Text = "Значение:"
        lbl_value.Location = Point(10, 55)
        lbl_value.AutoSize = True
        grp_value.Controls.Add(lbl_value)

        self.cmb_value = ComboBox()
        self.cmb_value.Location = Point(80, 52)
        self.cmb_value.Width = 300
        self.cmb_value.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        grp_value.Controls.Add(self.cmb_value)

        self.lbl_values_count = Label()
        self.lbl_values_count.Text = ""
        self.lbl_values_count.Location = Point(390, 55)
        self.lbl_values_count.Size = Size(150, 20)
        self.lbl_values_count.ForeColor = Color.Gray
        grp_value.Controls.Add(self.lbl_values_count)

        lbl_instr = Label()
        lbl_instr.Text = "Описание:"
        lbl_instr.Location = Point(10, 85)
        lbl_instr.AutoSize = True
        grp_value.Controls.Add(lbl_instr)

        self.lbl_instructions = Label()
        self.lbl_instructions.Text = ""
        self.lbl_instructions.Location = Point(10, 105)
        self.lbl_instructions.Size = Size(740, 40)
        self.lbl_instructions.ForeColor = Color.DarkGray
        grp_value.Controls.Add(self.lbl_instructions)

        self.Controls.Add(grp_value)

        y += 160

        # === Кнопки ===
        self.btn_apply = Button()
        self.btn_apply.Text = "Применить к {} элементам".format(len(self.elements))
        self.btn_apply.Location = Point(530, y)
        self.btn_apply.Width = 160
        self.btn_apply.Height = 30
        self.btn_apply.Enabled = False
        self.btn_apply.Click += self.on_apply
        self.Controls.Add(self.btn_apply)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(700, y)
        btn_close.Width = 75
        btn_close.Height = 30
        btn_close.Click += self.on_close
        self.Controls.Add(btn_close)

    def load_params(self):
        """Загрузить общие параметры элементов."""
        self.common_params = get_common_params(self.elements)
        self.filtered_revit_params = self.common_params[:]
        self.update_revit_list()

    def update_revit_list(self):
        """Обновить список параметров Revit."""
        self.lst_revit_params.Items.Clear()
        for name in self.filtered_revit_params:
            self.lst_revit_params.Items.Add(name)

    def on_search_revit_changed(self, sender, args):
        """Фильтровать параметры Revit."""
        search = self.txt_search_revit.Text.lower().strip()
        if not search:
            self.filtered_revit_params = self.common_params[:]
        else:
            self.filtered_revit_params = [
                name for name in self.common_params
                if search in name.lower()
            ]
        self.update_revit_list()

    def on_revit_param_selected(self, sender, args):
        """При выборе параметра Revit."""
        idx = self.lst_revit_params.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_revit_params):
            return

        self.selected_revit_param = self.filtered_revit_params[idx]
        self.lbl_selected_revit.Text = self.selected_revit_param
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
            # Преобразовать в список для отображения
            self.ids_params = [(name, info) for name, info in self.ids_data.items()]
            self.ids_params.sort(key=lambda x: x[0].lower())
            self.filtered_ids_params = self.ids_params[:]
            self.update_ids_list()

            show_info("IDS", "IDS загружен: {} параметров с допустимыми значениями".format(len(self.ids_params)))
        except Exception as e:
            show_error("Ошибка", "Ошибка загрузки IDS",
                       details=str(e))

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

        self.cmb_value.Items.Add("")  # пустое значение
        for val in allowed:
            self.cmb_value.Items.Add(val)
        self.cmb_value.SelectedIndex = 0

        self.lbl_values_count.Text = "{} значений".format(len(allowed))
        self.lbl_instructions.Text = instructions

        self.update_apply_button()

    def update_apply_button(self):
        """Активировать кнопку применения."""
        self.btn_apply.Enabled = (
            self.selected_revit_param is not None and
            self.selected_ids_param is not None and
            self.cmb_value.Items.Count > 1
        )

    def on_apply(self, sender, args):
        """Применить значение ко всем выбранным элементам."""
        if not self.selected_revit_param:
            return

        new_value = self.cmb_value.SelectedItem
        if new_value is None:
            return

        t = Transaction(doc, "Заполнить параметры экземпляров из IDS")
        t.Start()

        try:
            updated = 0
            errors = []

            for elem in self.elements:
                param = elem.LookupParameter(self.selected_revit_param)
                if param is None or param.IsReadOnly:
                    continue

                try:
                    if param.StorageType == StorageType.String:
                        param.Set(str(new_value))
                        updated += 1
                    elif param.StorageType == StorageType.Integer:
                        if new_value:
                            param.Set(int(new_value))
                        else:
                            param.Set(0)
                        updated += 1
                    elif param.StorageType == StorageType.Double:
                        if new_value:
                            param.Set(float(new_value))
                        else:
                            param.Set(0.0)
                        updated += 1
                except Exception as e:
                    errors.append(str(e))

            t.Commit()

            details = ""
            if errors:
                details = "Ошибки: {}".format(len(errors))

            if updated > 0:
                show_success("Результат", "Обновлено элементов: {}".format(updated),
                             details=details if details else None)
            else:
                show_warning("Результат", "Элементы не обновлены",
                             details=details if details else None)

        except Exception as e:
            t.RollBack()
            show_error("Ошибка", "Ошибка применения значения",
                       details=str(e))

    def on_close(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    # Получить выбранные элементы
    elements = get_selected_elements()

    if not elements:
        # Предложить выбрать элементы
        if show_confirm("Выбор элементов", "Элементы не выбраны. Выбрать сейчас?"):
            try:
                selection = uidoc.Selection
                refs = selection.PickObjects(ObjectType.Element, "Выберите элементы")
                elements = [doc.GetElement(ref.ElementId) for ref in refs]
            except:
                # Пользователь отменил выбор
                elements = []

    if not elements:
        show_warning("Внимание", "Элементы не выбраны. Операция отменена.")
    else:
        form = FillInstanceParamsForm(elements)
        form.ShowDialog()
