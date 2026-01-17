# -*- coding: utf-8 -*-
"""
ФОП в проект - добавление параметров из ФОП файла в проект Revit.

Парсит IDS файл для определения категорий и Instance/Type,
затем добавляет параметры из ФОП в проект.
"""

__title__ = "ФОП в\nпроект"
__author__ = "CPSK"

import clr
import os
import sys
import codecs

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, Panel, CheckBox, ComboBox, ListBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, OpenFileDialog,
    SelectionMode, GroupBox, ScrollBars
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import revit, forms, script

# Добавляем lib и support_files в путь для импорта
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
SUPPORT_DIR = os.path.join(EXTENSION_DIR, "support_files")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)
if SUPPORT_DIR not in sys.path:
    sys.path.insert(0, SUPPORT_DIR)

# Импорт модулей из lib
from cpsk_notify import show_error, show_warning, show_info, show_success
from cpsk_auth import require_auth
from cpsk_config import require_environment
from cpsk_logger import Logger

# Проверка авторизации
if not require_auth():
    sys.exit()

# Проверка окружения
if not require_environment():
    sys.exit()

# Инициализация логгера (очищает лог при каждом запуске)
SCRIPT_NAME = "FOPtoProject"
Logger.init(SCRIPT_NAME)
Logger.info(SCRIPT_NAME, "Скрипт запущен")

# Revit API
from Autodesk.Revit.DB import (
    Transaction, BuiltInCategory, Category,
    ExternalDefinitionCreationOptions,
    CategorySet
)

# BuiltInParameterGroup - разные версии Revit
try:
    from Autodesk.Revit.DB import BuiltInParameterGroup
except ImportError:
    # Revit 2024+ использует GroupTypeId
    from Autodesk.Revit.DB import GroupTypeId as BuiltInParameterGroup

# Импорт IFC маппингов из support_files
from ifc_mappings import IFC_TO_REVIT_CATEGORY_IDS

# === НАСТРОЙКИ ===

doc = revit.doc
app = revit.doc.Application
output = script.get_output()

# IFC_TO_REVIT_CATEGORIES теперь импортируется из ifc_mappings.py
# Формат: IFC_CLASS -> [(русское_имя, builtin_category_id), ...]
IFC_TO_REVIT_CATEGORIES = IFC_TO_REVIT_CATEGORY_IDS

# Все категории для UI
ALL_CATEGORIES = [
    ("Стены", BuiltInCategory.OST_Walls),
    ("Перекрытия", BuiltInCategory.OST_Floors),
    ("Колонны несущие", BuiltInCategory.OST_StructuralColumns),
    ("Каркас несущий", BuiltInCategory.OST_StructuralFraming),
    ("Фундамент несущий", BuiltInCategory.OST_StructuralFoundation),
    ("Лестницы", BuiltInCategory.OST_Stairs),
    ("Пандусы", BuiltInCategory.OST_Ramps),
    ("Ограждения", BuiltInCategory.OST_StairsRailing),
    ("Арматура несущая", BuiltInCategory.OST_Rebar),
    ("Обобщенные модели", BuiltInCategory.OST_GenericModel),
    ("Крыши", BuiltInCategory.OST_Roofs),
    ("Потолки", BuiltInCategory.OST_Ceilings),
    ("Несущие соединения", BuiltInCategory.OST_StructConnections),
]

# PropertySet -> Type параметр (не Instance)
TYPE_PROPERTY_SETS = [
    "Характеристики бетона",
    "Характеристики стали",
    "Характеристики древесины",
    "Характеристики раствора и камня",
    "Характеристики арматуры",
    "Pset_ConcreteElementGeneral",
    "Pset_ReinforcingBarBendingsBECCommon",
]

# === ДИНАМИЧЕСКОЕ ПОЛУЧЕНИЕ ГРУПП ПАРАМЕТРОВ ===

# Маппинг несоответствий API label -> UI label
# Revit API иногда возвращает label отличный от того что показывает UI
LABEL_FIXES = {
    #    "Общие": "Прочее",  # GroupTypeId.General в UI называется "Прочее"
}


def get_forge_type_id(g):
    """Получить строковый ID из ForgeTypeId объекта."""
    try:
        # Revit 2024+ ForgeTypeId имеет свойство TypeId
        return g.TypeId
    except:
        pass
    try:
        # Или можно попробовать repr
        s = repr(g)
        if "ForgeTypeId" in s:
            return s
        return str(g)
    except:
        return str(id(g))


def fix_label(label):
    """Исправить label если есть известное несоответствие API/UI."""
    return LABEL_FIXES.get(label, label)


def get_all_parameter_groups():
    """
    Динамически получить все группы параметров из Revit.
    Возвращает список кортежей: (локализованное_имя, group_object)

    ВАЖНО: Включает ВСЕ группы, даже с пустым label.
    """
    groups = []
    seen_ids = set()  # Для избежания дубликатов

    # Пробуем Revit 2024+ API
    try:
        from Autodesk.Revit.DB import ParameterUtils, LabelUtils

        all_groups = ParameterUtils.GetAllBuiltInGroups()
        for g in all_groups:
            # Получаем уникальный ID для ForgeTypeId
            gid = get_forge_type_id(g)
            if gid in seen_ids:
                continue
            seen_ids.add(gid)

            try:
                label = LabelUtils.GetLabelForGroup(g)
            except:
                label = ""

            # Если label пустой - генерируем читаемое имя из ID
            if not label:
                # autodesk.parameter.group.general -> General
                label = gid.replace("autodesk.parameter.group.", "")
                label = label.replace("autodesk.spec.aec.", "")
                label = label.replace(".", " ").replace("_", " ").title()

            # Исправить известные несоответствия API/UI
            label = fix_label(label)

            groups.append((label, g))

        if groups:
            # Сортировать по имени (регистронезависимо)
            groups.sort(key=lambda x: x[0].lower() if x[0] else "")
            return groups
    except Exception as e:
        # Логируем ошибку для отладки
        pass

    # Revit 2023 и раньше - используем BuiltInParameterGroup enum
    try:
        from Autodesk.Revit.DB import LabelUtils as LU

        # Пробуем импортировать BuiltInParameterGroup напрямую
        try:
            from Autodesk.Revit.DB import BuiltInParameterGroup as BIPG
        except ImportError:
            BIPG = None

        if BIPG is not None:
            # Получить все значения enum
            all_values = System.Enum.GetValues(BIPG)
            for g in all_values:
                gid = str(g)
                if gid in seen_ids or gid == "INVALID":
                    continue

                try:
                    if int(g) == -1:
                        continue
                except:
                    pass

                seen_ids.add(gid)

                try:
                    label = LU.GetLabelFor(g)
                except:
                    label = ""

                # Если label пустой - используем имя enum
                if not label:
                    label = gid.replace("PG_", "").replace("_", " ").title()

                groups.append((label, g))

            if groups:
                # Сортировать по имени
                groups.sort(key=lambda x: x[0])
                return groups
    except:
        pass

    # Fallback - базовый список (используем глобальный импорт)
    # BuiltInParameterGroup уже импортирован в начале файла (или GroupTypeId для 2024+)
    fallback = []
    try:
        fallback.append(("Данные", BuiltInParameterGroup.PG_DATA))
    except:
        pass
    try:
        fallback.append(("Идентификация", BuiltInParameterGroup.PG_IDENTITY_DATA))
    except:
        pass
    try:
        fallback.append(("Прочее", BuiltInParameterGroup.PG_GENERAL))
    except:
        pass

    if fallback:
        return fallback

    # Если совсем ничего не работает - вернём пустой список
    return []


# Кэш групп (загружается один раз)
_PARAMETER_GROUPS_CACHE = None

def get_parameter_groups():
    """Получить группы параметров (с кэшированием)."""
    global _PARAMETER_GROUPS_CACHE
    if _PARAMETER_GROUPS_CACHE is None:
        _PARAMETER_GROUPS_CACHE = get_all_parameter_groups()
    return _PARAMETER_GROUPS_CACHE


# === ПАРСЕР IDS (простой XML) ===

def parse_ids_simple(ids_path):
    """
    Простой парсер IDS без внешних зависимостей.
    Возвращает dict: param_name -> {
        ifc_classes: [...],
        property_set: str,
        is_type: bool,
        allowed_values: [...],  # допустимые значения из enumeration
        instructions: str,      # описание/инструкция
        data_type: str          # тип данных (IFCTEXT, IFCREAL и т.д.)
    }
    """
    result = {}

    try:
        with codecs.open(ids_path, 'r', 'utf-8') as f:
            content = f.read()
    except:
        # Попробовать другие кодировки
        try:
            with codecs.open(ids_path, 'r', 'utf-8-sig') as f:
                content = f.read()
        except:
            return result

    # Найти все specification блоки
    import re

    # Паттерн для specification с applicability и requirements
    spec_pattern = r'<specification[^>]*>(.*?)</specification>'
    specs = re.findall(spec_pattern, content, re.DOTALL)

    for spec in specs:
        # Найти IFC классы в applicability
        ifc_classes = []

        # simpleValue
        simple_matches = re.findall(r'<simpleValue>(IFC\w+)</simpleValue>', spec)
        ifc_classes.extend(simple_matches)

        # enumeration value для IFC классов
        enum_matches = re.findall(r'<xs:enumeration value="(IFC\w+)"', spec)
        ifc_classes.extend(enum_matches)

        # Убрать дубликаты
        ifc_classes = list(set(ifc_classes))

        # Найти все property в requirements (включая атрибуты)
        property_pattern = r'<property([^>]*)>(.*?)</property>'
        properties = re.findall(property_pattern, spec, re.DOTALL)

        for prop_match in properties:
            prop_attrs = prop_match[0]  # атрибуты тега <property>
            prop_body = prop_match[1]   # содержимое тега

            # instructions из атрибута property
            instr_match = re.search(r'instructions="([^"]*)"', prop_attrs)
            instructions_attr = instr_match.group(1) if instr_match else ""
            # Декодируем HTML entities
            instructions_attr = instructions_attr.replace("&#xA;", "\n").replace("&quot;", '"')

            # dataType из атрибута property
            dtype_match = re.search(r'dataType="([^"]+)"', prop_attrs)
            data_type = dtype_match.group(1) if dtype_match else ""

            # PropertySet
            pset_match = re.search(r'<propertySet>.*?<simpleValue>([^<]+)</simpleValue>', prop_body, re.DOTALL)
            property_set = pset_match.group(1) if pset_match else ""

            # Имя параметра
            name_match = re.search(r'<baseName>.*?<simpleValue>([^<]+)</simpleValue>', prop_body, re.DOTALL)
            if not name_match:
                continue
            param_name = name_match.group(1).strip()

            # Допустимые значения из enumeration (НЕ IFC классы)
            allowed_values = []
            value_block = re.search(r'<value>(.*?)</value>', prop_body, re.DOTALL)
            if value_block:
                # Ищем xs:enumeration value, исключая IFC классы
                enum_values = re.findall(r'<xs:enumeration value="([^"]+)"', value_block.group(1))
                for val in enum_values:
                    # Исключаем IFC классы (начинаются с IFC)
                    if not val.upper().startswith("IFC"):
                        allowed_values.append(val)

            # Определить Instance/Type
            is_type = False
            for type_pset in TYPE_PROPERTY_SETS:
                if type_pset.lower() in property_set.lower():
                    is_type = True
                    break

            # Сохранить
            if param_name not in result:
                result[param_name] = {
                    'ifc_classes': ifc_classes,
                    'property_set': property_set,
                    'is_type': is_type,
                    'allowed_values': allowed_values,
                    'instructions': instructions_attr,
                    'data_type': data_type
                }
            else:
                # Добавить IFC классы
                for ifc in ifc_classes:
                    if ifc not in result[param_name]['ifc_classes']:
                        result[param_name]['ifc_classes'].append(ifc)
                # Добавить допустимые значения
                for val in allowed_values:
                    if val not in result[param_name].get('allowed_values', []):
                        if 'allowed_values' not in result[param_name]:
                            result[param_name]['allowed_values'] = []
                        result[param_name]['allowed_values'].append(val)

    return result


def get_revit_categories_for_param(ids_info):
    """Получить категории Revit для параметра на основе IFC классов."""
    categories = []
    seen = set()

    for ifc_class in ids_info.get('ifc_classes', []):
        ifc_upper = ifc_class.upper()
        if ifc_upper in IFC_TO_REVIT_CATEGORIES:
            for cat_tuple in IFC_TO_REVIT_CATEGORIES[ifc_upper]:
                if cat_tuple[0] not in seen:
                    categories.append(cat_tuple)
                    seen.add(cat_tuple[0])

    return categories


# === ПАРСЕР ФОП ===

class FOPParser:
    """Парсер файла общих параметров Revit."""

    def __init__(self, fop_path):
        self.fop_path = fop_path
        self.groups = {}  # id -> name
        self.parameters = []  # list of param dicts

    def parse(self):
        """Парсить ФОП файл."""
        # Читаем файл в UTF-16
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
                # GROUP    ID    NAME
                if parts[0] == 'GROUP':
                    group_id = parts[1]
                    group_name = parts[2]
                    self.groups[group_id] = group_name

            elif current_section == 'PARAM' and len(parts) >= 6:
                # PARAM    GUID    NAME    DATATYPE    DATACATEGORY    GROUP    VISIBLE    DESCRIPTION...
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


# === ГЛАВНОЕ ОКНО ===

class FOPtoProjectForm(Form):
    """Диалог добавления параметров из ФОП в проект."""

    def __init__(self):
        self.fop_path = None
        self.ids_path = None
        self.parser = None
        self.ids_data = {}  # param_name -> ids_info
        self.prefix = ""    # Префикс параметров
        self.selected_params = []
        self.selected_categories = []
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "ФОП в проект - Добавление параметров"
        self.Width = 950
        self.Height = 750
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # === Выбор IDS файла ===
        lbl_ids = Label()
        lbl_ids.Text = "IDS файл (обязательно):"
        lbl_ids.Location = Point(15, y)
        lbl_ids.AutoSize = True
        self.Controls.Add(lbl_ids)

        y += 20
        self.txt_ids = System.Windows.Forms.TextBox()
        self.txt_ids.Location = Point(15, y)
        self.txt_ids.Width = 820
        self.txt_ids.ReadOnly = True
        self.Controls.Add(self.txt_ids)

        self.btn_browse_ids = Button()
        self.btn_browse_ids.Text = "Обзор..."
        self.btn_browse_ids.Location = Point(845, y - 2)
        self.btn_browse_ids.Width = 80
        self.btn_browse_ids.Click += self.on_browse_ids
        self.Controls.Add(self.btn_browse_ids)

        y += 35

        # === Выбор ФОП файла ===
        lbl_fop = Label()
        lbl_fop.Text = "ФОП файл (обязательно):"
        lbl_fop.Location = Point(15, y)
        lbl_fop.AutoSize = True
        self.Controls.Add(lbl_fop)

        y += 20
        self.txt_fop = System.Windows.Forms.TextBox()
        self.txt_fop.Location = Point(15, y)
        self.txt_fop.Width = 820
        self.txt_fop.ReadOnly = True
        self.Controls.Add(self.txt_fop)

        self.btn_browse = Button()
        self.btn_browse.Text = "Обзор..."
        self.btn_browse.Location = Point(845, y - 2)
        self.btn_browse.Width = 80
        self.btn_browse.Click += self.on_browse
        self.Controls.Add(self.btn_browse)

        y += 35

        # === Префикс параметров ===
        lbl_prefix = Label()
        lbl_prefix.Text = "Префикс параметров (если использовался при генерации ФОП):"
        lbl_prefix.Location = Point(15, y)
        lbl_prefix.AutoSize = True
        self.Controls.Add(lbl_prefix)

        y += 20
        self.txt_prefix = System.Windows.Forms.TextBox()
        self.txt_prefix.Location = Point(15, y)
        self.txt_prefix.Width = 150
        self.txt_prefix.Text = ""
        self.txt_prefix.TextChanged += self.on_prefix_changed
        self.Controls.Add(self.txt_prefix)

        lbl_prefix_hint = Label()
        lbl_prefix_hint.Text = "Например: ЦГЭ (без подчёркивания)"
        lbl_prefix_hint.Location = Point(175, y + 3)
        lbl_prefix_hint.AutoSize = True
        lbl_prefix_hint.ForeColor = Color.Gray
        self.Controls.Add(lbl_prefix_hint)

        self.btn_reload = Button()
        self.btn_reload.Text = "Обновить"
        self.btn_reload.Location = Point(400, y - 2)
        self.btn_reload.Width = 80
        self.btn_reload.Click += self.on_reload_fop
        self.Controls.Add(self.btn_reload)

        y += 35

        # === Список параметров (предпросмотр) ===
        grp_params = GroupBox()
        grp_params.Text = "Параметры из ФОП (все будут добавлены)"
        grp_params.Location = Point(15, y)
        grp_params.Size = Size(300, 280)

        self.lst_params = ListBox()
        self.lst_params.Location = Point(10, 20)
        self.lst_params.Size = Size(280, 250)
        self.lst_params.SelectedIndexChanged += self.on_param_selected
        grp_params.Controls.Add(self.lst_params)

        self.Controls.Add(grp_params)

        # === Список категорий (предпросмотр из IDS) ===
        grp_cats = GroupBox()
        grp_cats.Text = "Категории Revit (из IDS)"
        grp_cats.Location = Point(325, y)
        grp_cats.Size = Size(300, 280)

        self.lst_cats = ListBox()
        self.lst_cats.Location = Point(10, 20)
        self.lst_cats.Size = Size(280, 250)
        # Режим без выбора (только просмотр)
        self.lst_cats.SelectionMode = getattr(SelectionMode, "None")

        # Категории будут заполняться автоматически из IDS
        grp_cats.Controls.Add(self.lst_cats)
        self.Controls.Add(grp_cats)

        # === Допустимые значения из IDS ===
        grp_values = GroupBox()
        grp_values.Text = "Допустимые значения (IDS)"
        grp_values.Location = Point(635, y)
        grp_values.Size = Size(290, 280)

        self.lst_values = ListBox()
        self.lst_values.Location = Point(10, 20)
        self.lst_values.Size = Size(270, 200)
        grp_values.Controls.Add(self.lst_values)

        # Инструкция/описание параметра
        self.lbl_instructions = Label()
        self.lbl_instructions.Text = ""
        self.lbl_instructions.Location = Point(10, 225)
        self.lbl_instructions.Size = Size(270, 45)
        self.lbl_instructions.ForeColor = Color.DarkGray
        grp_values.Controls.Add(self.lbl_instructions)

        self.Controls.Add(grp_values)

        y += 290

        # === Рекомендация из IDS ===
        self.lbl_recommendation = Label()
        self.lbl_recommendation.Text = ""
        self.lbl_recommendation.Location = Point(15, y)
        self.lbl_recommendation.Size = Size(910, 40)
        self.lbl_recommendation.ForeColor = Color.DarkBlue
        self.Controls.Add(self.lbl_recommendation)

        y += 45

        # === Опции ===
        lbl_group = Label()
        lbl_group.Text = "Группа параметров в Revit:"
        lbl_group.Location = Point(15, y + 5)
        lbl_group.AutoSize = True
        self.Controls.Add(lbl_group)

        self.cmb_group = ComboBox()
        self.cmb_group.Location = Point(180, y + 2)
        self.cmb_group.Width = 280
        self.cmb_group.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self.cmb_group.MaxDropDownItems = 25

        # Динамически получить все группы параметров из Revit
        self.param_groups = get_parameter_groups()
        for group_name, _ in self.param_groups:
            self.cmb_group.Items.Add(group_name)

        # Выбрать "Данные" по умолчанию
        for i, (name, _) in enumerate(self.param_groups):
            if "Данные" in name or "Data" in name:
                self.cmb_group.SelectedIndex = i
                break
        else:
            self.cmb_group.SelectedIndex = 0

        self.Controls.Add(self.cmb_group)

        self.chk_instance = CheckBox()
        self.chk_instance.Text = "Параметр экземпляра (не типа)"
        self.chk_instance.Location = Point(450, y + 3)
        self.chk_instance.AutoSize = True
        self.chk_instance.Checked = True
        self.Controls.Add(self.chk_instance)

        y += 35

        # === Статус ===
        self.lbl_status = Label()
        self.lbl_status.Text = "Выберите IDS и ФОП файлы"
        self.lbl_status.Location = Point(15, y)
        self.lbl_status.Size = Size(600, 20)
        self.Controls.Add(self.lbl_status)

        # === Кнопки действий ===
        self.btn_add = Button()
        self.btn_add.Text = "Добавить параметры"
        self.btn_add.Location = Point(720, y - 5)
        self.btn_add.Width = 130
        self.btn_add.Height = 30
        self.btn_add.Enabled = False
        self.btn_add.Click += self.on_add_params
        self.Controls.Add(self.btn_add)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(860, y - 5)
        btn_close.Width = 65
        btn_close.Height = 30
        btn_close.Click += self.on_close
        self.Controls.Add(btn_close)

    def on_browse_ids(self, sender, args):
        """Выбор IDS файла."""
        dialog = OpenFileDialog()
        dialog.Title = "Выберите IDS файл"
        dialog.Filter = "IDS files (*.ids)|*.ids|XML files (*.xml)|*.xml|Все файлы (*.*)|*.*"
        if dialog.ShowDialog() == DialogResult.OK:
            self.ids_path = dialog.FileName
            self.txt_ids.Text = self.ids_path
            self.load_ids()

    def load_ids(self):
        """Загрузить и парсить IDS файл."""
        Logger.log_separator(SCRIPT_NAME, "Загрузка IDS файла")
        Logger.file_opened(SCRIPT_NAME, self.ids_path, "IDS файл")

        try:
            self.ids_data = parse_ids_simple(self.ids_path)
            count = len(self.ids_data)

            Logger.info(SCRIPT_NAME, "IDS успешно загружен: {} параметров".format(count))
            Logger.data(SCRIPT_NAME, "Параметры IDS", list(self.ids_data.keys()))

            self.lbl_status.Text = "IDS: {} параметров найдено".format(count)
            self.lbl_status.ForeColor = Color.DarkGreen
            self.update_add_button_state()

        except Exception as e:
            Logger.error(SCRIPT_NAME, "Ошибка загрузки IDS: {}".format(str(e)), exc_info=True)
            self.lbl_status.Text = "Ошибка IDS: {}".format(str(e))
            self.lbl_status.ForeColor = Color.Red

    def update_add_button_state(self):
        """Обновить состояние кнопки 'Добавить' в зависимости от загруженных файлов."""
        # Кнопка активна только когда загружены оба файла
        self.btn_add.Enabled = bool(self.ids_path and self.fop_path and self.parser and self.ids_data)

    def on_browse(self, sender, args):
        """Выбор ФОП файла."""
        dialog = OpenFileDialog()
        dialog.Title = "Выберите ФОП файл"
        dialog.Filter = "Shared Parameters (*.txt)|*.txt|Все файлы (*.*)|*.*"
        if dialog.ShowDialog() == DialogResult.OK:
            self.fop_path = dialog.FileName
            self.txt_fop.Text = self.fop_path
            self.load_fop()

    def get_ids_name(self, fop_name):
        """
        Получить имя параметра для поиска в IDS (без префикса).
        ФОП: 'ЦГЭ_Класс прочности' -> IDS: 'Класс прочности'
        Также убираем лишние пробелы в начале/конце.
        """
        prefix = self.txt_prefix.Text.strip()
        # Убираем лишние пробелы из имени параметра
        name = fop_name.strip()
        if prefix and name.startswith(prefix + "_"):
            return name[len(prefix) + 1:].strip()
        return name.strip()

    def on_prefix_changed(self, sender, args):
        """При изменении префикса обновить статус."""
        prefix = self.txt_prefix.Text.strip()
        if prefix:
            self.lbl_status.Text = "Префикс: '{}'. Нажмите 'Обновить' для перезагрузки".format(prefix)
            self.lbl_status.ForeColor = Color.DarkBlue

    def on_reload_fop(self, sender, args):
        """Перезагрузить ФОП с учётом нового префикса."""
        if self.fop_path:
            self.load_fop()

    def load_fop(self):
        """Загрузить и парсить ФОП файл."""
        Logger.log_separator(SCRIPT_NAME, "Загрузка ФОП файла")
        Logger.file_opened(SCRIPT_NAME, self.fop_path, "ФОП файл")

        prefix = self.txt_prefix.Text.strip()
        Logger.debug(SCRIPT_NAME, "Префикс для сопоставления: '{}'".format(prefix if prefix else "(нет)"))

        try:
            self.parser = FOPParser(self.fop_path)
            self.parser.parse()

            Logger.info(SCRIPT_NAME, "ФОП успешно распарсен: {} параметров".format(len(self.parser.parameters)))

            # Заполнить список параметров
            self.lst_params.Items.Clear()
            matched_count = 0
            matched_params = []
            unmatched_params = []

            for param in self.parser.parameters:
                fop_name = param['name']
                ids_name = self.get_ids_name(fop_name)

                # Пометить, если есть в IDS (с учётом префикса)
                marker = ""
                if ids_name in self.ids_data:
                    marker = " [IDS]"
                    matched_count += 1
                    matched_params.append(fop_name)
                else:
                    unmatched_params.append(fop_name)

                display = "{}{}".format(fop_name, marker)
                self.lst_params.Items.Add(display)

            Logger.info(SCRIPT_NAME, "Сопоставление: {} совпадают с IDS, {} без совпадения".format(
                matched_count, len(unmatched_params)))

            if matched_params:
                Logger.data(SCRIPT_NAME, "Параметры с совпадением в IDS", matched_params)
            if unmatched_params:
                Logger.data(SCRIPT_NAME, "Параметры БЕЗ совпадения в IDS", unmatched_params)

            status = "ФОП: {} параметров".format(len(self.parser.parameters))
            if prefix:
                status += " (префикс: {})".format(prefix)
            if self.ids_data:
                status += ", {} совпадают с IDS".format(matched_count)

            self.lbl_status.Text = status
            self.lbl_status.ForeColor = Color.Black
            self.update_add_button_state()

        except Exception as e:
            Logger.error(SCRIPT_NAME, "Ошибка загрузки ФОП: {}".format(str(e)), exc_info=True)
            self.lbl_status.Text = "Ошибка ФОП: {}".format(str(e))
            self.lbl_status.ForeColor = Color.Red

    def on_param_selected(self, sender, args):
        """При выборе параметра показать категории и допустимые значения из IDS."""
        if self.lst_params.SelectedIndex < 0:
            return
        if not self.parser:
            return

        idx = self.lst_params.SelectedIndex
        param = self.parser.parameters[idx]
        fop_name = param['name']
        ids_name = self.get_ids_name(fop_name)

        # Очистить список категорий
        self.lst_cats.Items.Clear()

        if ids_name in self.ids_data:
            info = self.ids_data[ids_name]

            # Категории - показать в списке категорий
            cats = get_revit_categories_for_param(info)
            for cat_name, _ in sorted(cats, key=lambda x: x[0]):
                self.lst_cats.Items.Add(cat_name)

            cat_names = [c[0] for c in cats] if cats else ["(не определено)"]

            # Instance/Type
            param_type = "ПО ТИПУ" if info.get('is_type') else "ПО ЭКЗЕМПЛЯРУ"

            # PropertySet
            pset = info.get('property_set', '')

            # Допустимые значения
            allowed = info.get('allowed_values', [])

            # Формируем рекомендацию
            rec_parts = []
            rec_parts.append("{} | Категории: {}".format(param_type, ", ".join(cat_names)))

            if allowed:
                # Показываем первые 5 значений
                if len(allowed) <= 5:
                    values_str = ", ".join(allowed)
                else:
                    values_str = ", ".join(allowed[:5]) + "... (+{})".format(len(allowed) - 5)
                rec_parts.append("Значения: {}".format(values_str))

            rec = " | ".join(rec_parts)

            # Добавляем имя параметра
            if fop_name != ids_name:
                rec = "IDS: '{}' -> ФОП: '{}' | {}".format(ids_name, fop_name, rec)
            else:
                rec = "'{}': {}".format(ids_name, rec)

            self.lbl_recommendation.Text = rec

            # Обновить список допустимых значений (если есть)
            self.update_allowed_values(allowed, info.get('instructions', ''))
        else:
            self.lbl_recommendation.Text = "Параметр '{}' не найден в IDS".format(ids_name)
            self.update_allowed_values([], "")

    def update_allowed_values(self, values, instructions):
        """Обновить список допустимых значений."""
        if hasattr(self, 'lst_values'):
            self.lst_values.Items.Clear()
            if values:
                for v in values:
                    self.lst_values.Items.Add(v)
            if hasattr(self, 'lbl_instructions'):
                self.lbl_instructions.Text = instructions if instructions else ""

    def get_param_group(self):
        """Получить группу параметров Revit."""
        index = self.cmb_group.SelectedIndex
        if index < 0 or index >= len(self.param_groups):
            index = 0

        # Возвращаем напрямую объект группы (уже получен из Revit)
        group_name, group_obj = self.param_groups[index]
        return group_obj

    def log_mismatches(self):
        """Логировать несовпадения между IDS и ФОП."""
        if not self.ids_data or not self.parser:
            return

        Logger.log_separator(SCRIPT_NAME, "Анализ несовпадений IDS и ФОП")

        # Создаём множество нормализованных имён из ФОП
        fop_names_normalized = set()
        for p in self.parser.parameters:
            fop_names_normalized.add(self.get_ids_name(p['name']))

        # IDS параметры без соответствия в ФОП
        ids_only = [name for name in self.ids_data.keys() if name not in fop_names_normalized]

        # ФОП параметры без соответствия в IDS
        fop_only = []
        for p in self.parser.parameters:
            ids_name = self.get_ids_name(p['name'])
            if ids_name not in self.ids_data:
                fop_only.append(p['name'])

        if ids_only:
            Logger.warning(SCRIPT_NAME, "Параметры IDS без соответствия в ФОП ({} шт):".format(len(ids_only)))
            for name in ids_only:
                Logger.warning(SCRIPT_NAME, "  - {}".format(name))

        if fop_only:
            Logger.warning(SCRIPT_NAME, "Параметры ФОП без соответствия в IDS ({} шт):".format(len(fop_only)))
            for name in fop_only:
                Logger.warning(SCRIPT_NAME, "  - {}".format(name))

        if not ids_only and not fop_only:
            Logger.info(SCRIPT_NAME, "Все параметры IDS и ФОП совпадают")

        return ids_only, fop_only

    def on_add_params(self, sender, args):
        """Добавить все параметры из ФОП в проект."""
        Logger.log_separator(SCRIPT_NAME, "ДОБАВЛЕНИЕ ПАРАМЕТРОВ В ПРОЕКТ")

        # Проверка обязательных файлов
        if not self.ids_path or not self.ids_data:
            Logger.error(SCRIPT_NAME, "IDS файл не выбран или не загружен")
            show_error("Ошибка", "Необходимо выбрать IDS файл",
                       details="IDS файл обязателен для определения категорий параметров.")
            return

        if not self.fop_path or not self.parser:
            Logger.error(SCRIPT_NAME, "ФОП файл не выбран или не загружен")
            show_error("Ошибка", "Необходимо выбрать ФОП файл",
                       details="ФОП файл содержит определения параметров для добавления в проект.")
            return

        Logger.info(SCRIPT_NAME, "IDS: {}".format(self.ids_path))
        Logger.info(SCRIPT_NAME, "ФОП: {}".format(self.fop_path))

        # Логируем несовпадения
        self.log_mismatches()

        # Добавляем ВСЕ параметры из списка
        if not self.parser.parameters:
            Logger.warning(SCRIPT_NAME, "Нет параметров для добавления")
            show_warning("Внимание", "Нет параметров в ФОП файле")
            return

        all_indices = list(range(len(self.parser.parameters)))
        all_names = [self.parser.parameters[i]['name'] for i in all_indices]
        Logger.info(SCRIPT_NAME, "Будет добавлено {} параметров".format(len(all_indices)))
        Logger.data(SCRIPT_NAME, "Параметры для добавления", all_names)

        # Проверить наличие IDS данных (категории берутся для каждого параметра отдельно)
        if not self.ids_data:
            Logger.error(SCRIPT_NAME, "IDS данные не загружены")
            show_warning("Внимание", "Загрузите IDS файл",
                         details="Для определения категорий каждого параметра необходим IDS файл.")
            return

        Logger.info(SCRIPT_NAME, "IDS данные загружены: {} параметров".format(len(self.ids_data)))

        # Открыть файл определений
        Logger.debug(SCRIPT_NAME, "Открытие ФОП файла через Revit API...")
        try:
            app.SharedParametersFilename = self.fop_path
            def_file = app.OpenSharedParameterFile()
            if def_file is None:
                Logger.error(SCRIPT_NAME, "Revit вернул None при открытии ФОП")
                show_error("Ошибка", "Не удалось открыть ФОП файл")
                return
            Logger.debug(SCRIPT_NAME, "ФОП файл успешно открыт через Revit API")
        except Exception as e:
            Logger.error(SCRIPT_NAME, "Ошибка открытия ФОП через Revit API: {}".format(str(e)), exc_info=True)
            show_error("Ошибка", "Ошибка открытия ФОП",
                       details=str(e))
            return

        # Вспомогательная функция для создания CategorySet для конкретного параметра
        def create_category_set_for_param(param_categories):
            """Создать CategorySet из списка категорий [(name, id), ...]"""
            cat_set = CategorySet()
            failed = []
            for cat_name, cat_id in param_categories:
                try:
                    bic = BuiltInCategory(cat_id)
                    cat = Category.GetCategory(doc, bic)
                    if cat:
                        if cat.AllowsBoundParameters:
                            cat_set.Insert(cat)
                        else:
                            failed.append(cat_name)
                            Logger.debug(SCRIPT_NAME, "    {} НЕ ДОПУСКАЕТ привязку".format(cat_name))
                    else:
                        failed.append(cat_name)
                        Logger.debug(SCRIPT_NAME, "    {} НЕ НАЙДЕНА".format(cat_name))
                except Exception as e:
                    failed.append(cat_name)
                    Logger.debug(SCRIPT_NAME, "    {} ОШИБКА: {}".format(cat_name, str(e)))
            return cat_set, failed

        Logger.info(SCRIPT_NAME, "Каждый параметр будет привязан к своим категориям из IDS")

        # Параметры
        param_group = self.get_param_group()
        is_instance = self.chk_instance.Checked

        Logger.info(SCRIPT_NAME, "Тип привязки: {}".format("Экземпляр (Instance)" if is_instance else "Тип (Type)"))
        Logger.info(SCRIPT_NAME, "Группа параметров: {}".format(self.cmb_group.SelectedItem))

        # Собрать имена существующих параметров в проекте
        existing_param_names = set()
        bindings_map = doc.ParameterBindings
        it = bindings_map.ForwardIterator()
        while it.MoveNext():
            definition = it.Key
            existing_param_names.add(definition.Name)
        Logger.info(SCRIPT_NAME, "Существующих параметров в проекте: {}".format(len(existing_param_names)))

        Logger.log_separator(SCRIPT_NAME, "Начало транзакции")

        added_count = 0
        errors = []

        # Транзакция
        t = Transaction(doc, "Добавить параметры из ФОП")
        t.Start()
        Logger.debug(SCRIPT_NAME, "Транзакция запущена")

        try:
            for idx in all_indices:
                param = self.parser.parameters[idx]
                param_name = param['name']
                group_name = param['group_name']

                Logger.debug(SCRIPT_NAME, "Обработка: {} (группа: {})".format(param_name, group_name))

                # Получить категории для ЭТОГО параметра из IDS
                ids_name = self.get_ids_name(param_name)
                param_categories = []
                if ids_name in self.ids_data:
                    param_categories = get_revit_categories_for_param(self.ids_data[ids_name])
                    ifc_classes = self.ids_data[ids_name].get('ifc_classes', [])
                    Logger.debug(SCRIPT_NAME, "  IFC классы: {}".format(", ".join(ifc_classes)))
                    Logger.debug(SCRIPT_NAME, "  Категории: {}".format(", ".join([c[0] for c in param_categories])))
                else:
                    Logger.warning(SCRIPT_NAME, "  Параметр '{}' не найден в IDS (искали '{}')".format(param_name, ids_name))
                    errors.append("Не в IDS: {}".format(param_name))
                    continue

                # Создать CategorySet для ЭТОГО параметра
                cat_set, failed_cats = create_category_set_for_param(param_categories)
                if cat_set.IsEmpty:
                    Logger.warning(SCRIPT_NAME, "  ПРОПУСК: нет категорий для привязки")
                    errors.append("Нет категорий: {}".format(param_name))
                    continue

                # Найти определение в ФОП
                ext_def = None
                for grp in def_file.Groups:
                    if grp.Name == group_name:
                        ext_def = grp.Definitions.get_Item(param_name)
                        break

                if ext_def is None:
                    errors.append("Не найден в ФОП: {}".format(param_name))
                    Logger.warning(SCRIPT_NAME, "  ПРОПУСК: параметр не найден в ФОП файле")
                    continue

                # Проверить, не существует ли уже (по имени)
                if param_name in existing_param_names:
                    errors.append("Уже существует: {}".format(param_name))
                    Logger.info(SCRIPT_NAME, "  ПРОПУСК: параметр '{}' уже существует в проекте".format(param_name))
                    continue

                # Создать привязку с категориями ЭТОГО параметра
                if is_instance:
                    new_binding = app.Create.NewInstanceBinding(cat_set)
                else:
                    new_binding = app.Create.NewTypeBinding(cat_set)

                # Добавить параметр (с обработкой ошибок для каждого параметра)
                try:
                    if doc.ParameterBindings.Insert(ext_def, new_binding, param_group):
                        added_count += 1
                        cat_names = [c[0] for c in param_categories]
                        Logger.info(SCRIPT_NAME, "  ДОБАВЛЕН: {} -> [{}]".format(param_name, ", ".join(cat_names)))
                    else:
                        errors.append("Ошибка добавления: {}".format(param_name))
                        Logger.error(SCRIPT_NAME, "  ОШИБКА: не удалось добавить параметр")
                except Exception as bind_err:
                    errors.append("{}: {}".format(param_name, str(bind_err)))
                    Logger.warning(SCRIPT_NAME, "  ОШИБКА ПРИВЯЗКИ: {}".format(str(bind_err)))

            t.Commit()
            Logger.debug(SCRIPT_NAME, "Транзакция завершена (Commit)")

        except Exception as e:
            t.RollBack()
            Logger.error(SCRIPT_NAME, "Транзакция откачена (RollBack): {}".format(str(e)), exc_info=True)
            show_error("Ошибка", "Ошибка добавления параметров",
                       details=str(e))
            return

        # Результат
        Logger.log_separator(SCRIPT_NAME, "РЕЗУЛЬТАТ")
        Logger.result(SCRIPT_NAME, added_count > 0,
                      "Добавлено: {}, Ошибок: {}".format(added_count, len(errors)),
                      errors if errors else None)

        details = ""
        if errors:
            details = "Проблемы:\n" + "\n".join(errors[:10])
            if len(errors) > 10:
                details += "\n...и ещё {}".format(len(errors) - 10)

        self.lbl_status.Text = "Добавлено: {}".format(added_count)
        self.lbl_status.ForeColor = Color.Green if added_count > 0 else Color.Red

        if added_count > 0:
            show_success("Результат", "Добавлено параметров: {}".format(added_count),
                         details=details if details else None)
        else:
            show_warning("Результат", "Параметры не добавлены",
                         details=details if details else None)

    def on_close(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    form = FOPtoProjectForm()
    form.ShowDialog()
