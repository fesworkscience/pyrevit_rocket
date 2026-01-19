# -*- coding: utf-8 -*-
"""
Заполнить параметры ФОП плейсхолдерами - тестовая команда.

Загружает ФОП файл, находит все элементы с этими параметрами
и заполняет их тестовыми значениями (строка/число/bool).
"""

__title__ = "Заполнить\nФОП"
__author__ = "CPSK"

import clr
import os
import sys
import codecs

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, Panel, CheckBox, ListBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, OpenFileDialog, ProgressBar,
    SelectionMode, GroupBox, MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import revit, script

# Добавляем lib в путь для импорта
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Импорт модулей из lib
from cpsk_notify import show_error, show_warning, show_info, show_success
from cpsk_auth import require_auth
from cpsk_logger import Logger

# Проверка авторизации
if not require_auth():
    sys.exit()

# Инициализация логгера
SCRIPT_NAME = "FillFOPPlaceholders"
Logger.init(SCRIPT_NAME)
Logger.info(SCRIPT_NAME, "Скрипт запущен")

# Revit API
from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, BuiltInCategory,
    StorageType, ElementId
)

# ParameterType для определения YesNo (булевых) параметров
# В Revit 2022+ deprecated, но всё ещё работает для чтения
HAS_PARAMETER_TYPE = False
RevitParameterType = None
try:
    from Autodesk.Revit.DB import ParameterType as RevitParameterType
    HAS_PARAMETER_TYPE = True
except ImportError:
    pass

# === НАСТРОЙКИ ===

doc = revit.doc
output = script.get_output()

# Плейсхолдеры по типам данных
PLACEHOLDER_STRING = "TEST_VALUE"
PLACEHOLDER_INT = 123
PLACEHOLDER_DOUBLE = 1.5
PLACEHOLDER_ELEMENTID = ElementId.InvalidElementId

# Категории элементов для поиска
ELEMENT_CATEGORIES = [
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_Stairs,
    BuiltInCategory.OST_Ramps,
    BuiltInCategory.OST_StairsRailing,
    BuiltInCategory.OST_Rebar,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_Roofs,
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_StructConnections,
]


# === ПАРСЕР ФОП ===

class FOPParser:
    """Парсер файла общих параметров Revit."""

    def __init__(self, fop_path):
        self.fop_path = fop_path
        self.groups = {}  # id -> name
        self.parameters = []  # list of param dicts

    def parse(self):
        """Парсить ФОП файл."""
        with codecs.open(self.fop_path, 'r', 'utf-16') as f:
            lines = f.readlines()

        current_section = None

        for line in lines:
            line = line.strip()
            if not line or line.startswith('#'):
                continue

            if line.startswith('*'):
                current_section = line[1:].split('\t')[0]
                continue

            parts = line.split('\t')

            if current_section == 'GROUP' and len(parts) >= 3:
                if parts[0] == 'GROUP':
                    group_id = parts[1]
                    group_name = parts[2]
                    self.groups[group_id] = group_name

            elif current_section == 'PARAM' and len(parts) >= 6:
                if parts[0] == 'PARAM':
                    param = {
                        'guid': parts[1],
                        'name': parts[2],
                        'datatype': parts[3],
                        'group_id': parts[5],
                        'group_name': self.groups.get(parts[5], ""),
                        'description': parts[7] if len(parts) > 7 else ""
                    }
                    self.parameters.append(param)

        return self


def get_project_info():
    """Получить элемент Project Information."""
    try:
        collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectInformation)
        project_info = list(collector)
        if project_info:
            Logger.info(SCRIPT_NAME, "Project Information найден")
            return project_info[0]
    except Exception as e:
        Logger.warning(SCRIPT_NAME, "Ошибка получения Project Information: {}".format(str(e)))
    return None


def get_all_elements():
    """Получить все элементы из указанных категорий."""
    Logger.info(SCRIPT_NAME, "Сбор элементов из проекта...")

    all_elements = []

    # Добавляем Project Information
    project_info = get_project_info()
    if project_info:
        all_elements.append(project_info)

    for bic in ELEMENT_CATEGORIES:
        try:
            collector = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
            elements = list(collector)
            if elements:
                Logger.debug(SCRIPT_NAME, "  {}: {} элементов".format(bic, len(elements)))
                all_elements.extend(elements)
        except Exception as e:
            Logger.warning(SCRIPT_NAME, "  Ошибка категории {}: {}".format(bic, str(e)))
            continue

    Logger.info(SCRIPT_NAME, "Всего элементов: {}".format(len(all_elements)))
    return all_elements


def is_yes_no_parameter(param):
    """Проверить, является ли параметр булевым (YesNo)."""
    try:
        definition = param.Definition
        if definition is None:
            return False

        # Revit 2022+ - используем GetDataType()
        if hasattr(definition, 'GetDataType'):
            data_type = definition.GetDataType()
            # ForgeTypeId для YesNo: autodesk.spec:spec.bool
            type_id = str(data_type.TypeId) if hasattr(data_type, 'TypeId') else str(data_type)
            if 'bool' in type_id.lower() or 'yesno' in type_id.lower():
                return True

        # Revit 2021 и раньше - используем ParameterType
        if HAS_PARAMETER_TYPE and RevitParameterType is not None and hasattr(definition, 'ParameterType'):
            param_type = definition.ParameterType
            # pylint: disable=no-member
            if param_type == RevitParameterType.YesNo:
                return True

    except Exception:
        pass

    return False


def set_parameter_placeholder(param):
    """Установить плейсхолдер для параметра в зависимости от типа."""
    if param.IsReadOnly:
        return False, "readonly"

    storage = param.StorageType

    try:
        if storage == StorageType.String:
            param.Set(PLACEHOLDER_STRING)
            return True, "string"
        elif storage == StorageType.Integer:
            # Проверяем, булевый ли это параметр (YesNo)
            if is_yes_no_parameter(param):
                param.Set(1)  # True
                return True, "bool"
            else:
                param.Set(PLACEHOLDER_INT)
                return True, "int"
        elif storage == StorageType.Double:
            param.Set(PLACEHOLDER_DOUBLE)
            return True, "double"
        elif storage == StorageType.ElementId:
            # ElementId параметры обычно ссылки - пропускаем
            return False, "elementid"
        else:
            return False, "unknown"
    except Exception as e:
        return False, str(e)


def fill_parameters(elements, param_names):
    """Заполнить параметры плейсхолдерами."""
    Logger.log_separator(SCRIPT_NAME, "ЗАПОЛНЕНИЕ ПАРАМЕТРОВ")

    stats = {
        'filled': 0,
        'skipped_readonly': 0,
        'skipped_elementid': 0,
        'skipped_notfound': 0,
        'errors': 0
    }

    filled_params = set()

    for elem in elements:
        for param_name in param_names:
            try:
                param = elem.LookupParameter(param_name)
                if param is None:
                    continue

                success, reason = set_parameter_placeholder(param)

                if success:
                    stats['filled'] += 1
                    filled_params.add(param_name)
                elif reason == "readonly":
                    stats['skipped_readonly'] += 1
                elif reason == "elementid":
                    stats['skipped_elementid'] += 1
                else:
                    stats['errors'] += 1
                    Logger.warning(SCRIPT_NAME, "Ошибка параметра '{}': {}".format(param_name, reason))

            except Exception as e:
                stats['errors'] += 1
                Logger.error(SCRIPT_NAME, "Ошибка элемента {}: {}".format(elem.Id.IntegerValue, str(e)))
                continue

    Logger.info(SCRIPT_NAME, "Заполнено: {}".format(stats['filled']))
    Logger.info(SCRIPT_NAME, "Пропущено (readonly): {}".format(stats['skipped_readonly']))
    Logger.info(SCRIPT_NAME, "Пропущено (ElementId): {}".format(stats['skipped_elementid']))
    Logger.info(SCRIPT_NAME, "Ошибок: {}".format(stats['errors']))
    Logger.info(SCRIPT_NAME, "Уникальных параметров заполнено: {}".format(len(filled_params)))

    return stats, filled_params


# === ГЛАВНОЕ ОКНО ===

class FillFOPForm(Form):
    """Диалог заполнения параметров ФОП плейсхолдерами."""

    def __init__(self):
        self.fop_path = None
        self.parser = None
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Заполнить параметры ФОП плейсхолдерами"
        self.Width = 500
        self.Height = 450
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # === Выбор ФОП файла ===
        lbl_fop = Label()
        lbl_fop.Text = "ФОП файл:"
        lbl_fop.Location = Point(15, y)
        lbl_fop.AutoSize = True
        self.Controls.Add(lbl_fop)

        y += 20
        self.txt_fop = System.Windows.Forms.TextBox()
        self.txt_fop.Location = Point(15, y)
        self.txt_fop.Width = 380
        self.txt_fop.ReadOnly = True
        self.Controls.Add(self.txt_fop)

        self.btn_browse = Button()
        self.btn_browse.Text = "Обзор..."
        self.btn_browse.Location = Point(405, y - 2)
        self.btn_browse.Width = 70
        self.btn_browse.Click += self.on_browse
        self.Controls.Add(self.btn_browse)

        y += 35

        # === Список параметров ===
        grp_params = GroupBox()
        grp_params.Text = "Параметры из ФОП (будут заполнены)"
        grp_params.Location = Point(15, y)
        grp_params.Size = Size(460, 250)

        self.lst_params = ListBox()
        self.lst_params.Location = Point(10, 20)
        self.lst_params.Size = Size(440, 220)
        self.lst_params.SelectionMode = SelectionMode.None
        grp_params.Controls.Add(self.lst_params)

        self.Controls.Add(grp_params)

        y += 260

        # === Статус ===
        self.lbl_status = Label()
        self.lbl_status.Text = "Выберите ФОП файл"
        self.lbl_status.Location = Point(15, y)
        self.lbl_status.Size = Size(350, 20)
        self.Controls.Add(self.lbl_status)

        y += 30

        # === Кнопки ===
        self.btn_fill = Button()
        self.btn_fill.Text = "Заполнить"
        self.btn_fill.Location = Point(300, y)
        self.btn_fill.Width = 90
        self.btn_fill.Height = 30
        self.btn_fill.Enabled = False
        self.btn_fill.Click += self.on_fill
        self.Controls.Add(self.btn_fill)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(400, y)
        btn_close.Width = 75
        btn_close.Height = 30
        btn_close.Click += self.on_close
        self.Controls.Add(btn_close)

    def on_browse(self, sender, args):
        """Выбор ФОП файла."""
        dialog = OpenFileDialog()
        dialog.Title = "Выберите ФОП файл"
        dialog.Filter = "Shared Parameters (*.txt)|*.txt|Все файлы (*.*)|*.*"
        if dialog.ShowDialog() == DialogResult.OK:
            self.fop_path = dialog.FileName
            self.txt_fop.Text = self.fop_path
            self.load_fop()

    def load_fop(self):
        """Загрузить и парсить ФОП файл."""
        Logger.log_separator(SCRIPT_NAME, "Загрузка ФОП файла")
        Logger.file_opened(SCRIPT_NAME, self.fop_path, "ФОП файл")

        try:
            self.parser = FOPParser(self.fop_path)
            self.parser.parse()

            Logger.info(SCRIPT_NAME, "ФОП успешно распарсен: {} параметров".format(len(self.parser.parameters)))

            # Заполнить список параметров
            self.lst_params.Items.Clear()
            for param in self.parser.parameters:
                display = "{} ({})".format(param['name'], param['datatype'])
                self.lst_params.Items.Add(display)

            self.lbl_status.Text = "Загружено: {} параметров".format(len(self.parser.parameters))
            self.lbl_status.ForeColor = Color.DarkGreen
            self.btn_fill.Enabled = True

        except Exception as e:
            Logger.error(SCRIPT_NAME, "Ошибка загрузки ФОП: {}".format(str(e)), exc_info=True)
            self.lbl_status.Text = "Ошибка: {}".format(str(e))
            self.lbl_status.ForeColor = Color.Red
            show_error("Ошибка", "Не удалось загрузить ФОП файл", details=str(e))

    def on_fill(self, sender, args):
        """Заполнить параметры плейсхолдерами."""
        if not self.parser or not self.parser.parameters:
            show_warning("Внимание", "Нет параметров для заполнения")
            return

        Logger.log_separator(SCRIPT_NAME, "НАЧАЛО ЗАПОЛНЕНИЯ")

        # Получить имена параметров
        param_names = [p['name'] for p in self.parser.parameters]
        Logger.info(SCRIPT_NAME, "Параметров для заполнения: {}".format(len(param_names)))

        # Получить элементы
        elements = get_all_elements()
        if not elements:
            Logger.warning(SCRIPT_NAME, "Нет элементов в проекте")
            show_warning("Внимание", "Нет элементов в проекте для заполнения")
            return

        # Транзакция
        t = Transaction(doc, "Заполнить параметры ФОП плейсхолдерами")
        t.Start()
        Logger.debug(SCRIPT_NAME, "Транзакция запущена")

        try:
            stats, filled_params = fill_parameters(elements, param_names)
            t.Commit()
            Logger.debug(SCRIPT_NAME, "Транзакция завершена (Commit)")

            # Результат
            Logger.log_separator(SCRIPT_NAME, "РЕЗУЛЬТАТ")
            Logger.result(SCRIPT_NAME, stats['filled'] > 0,
                          "Заполнено {} значений в {} параметрах".format(
                              stats['filled'], len(filled_params)))

            self.lbl_status.Text = "Заполнено: {} значений".format(stats['filled'])
            self.lbl_status.ForeColor = Color.Green if stats['filled'] > 0 else Color.Orange

            details = "Заполнено значений: {}\n".format(stats['filled'])
            details += "Уникальных параметров: {}\n".format(len(filled_params))
            details += "Пропущено (readonly): {}\n".format(stats['skipped_readonly'])
            details += "Пропущено (ElementId): {}\n".format(stats['skipped_elementid'])
            details += "Ошибок: {}".format(stats['errors'])

            if stats['filled'] > 0:
                show_success("Готово", "Параметры заполнены плейсхолдерами", details=details)
            else:
                show_warning("Внимание", "Ни один параметр не был заполнен", details=details)

        except Exception as e:
            t.RollBack()
            Logger.error(SCRIPT_NAME, "Транзакция откачена: {}".format(str(e)), exc_info=True)
            show_error("Ошибка", "Ошибка заполнения параметров", details=str(e))

    def on_close(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    form = FillFOPForm()
    form.ShowDialog()
