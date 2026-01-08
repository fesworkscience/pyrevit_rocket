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
    DialogResult, OpenFileDialog, CheckedListBox,
    SelectionMode, GroupBox, ScrollBars
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import revit, forms, script

# Добавляем lib в путь для импорта cpsk_config
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Импорт модулей из lib
from cpsk_notify import show_error, show_warning, show_info, show_success
from cpsk_auth import require_auth
from cpsk_config import require_environment

# Проверка авторизации
if not require_auth():
    sys.exit()

# Проверка окружения
if not require_environment():
    sys.exit()

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

# === НАСТРОЙКИ ===

doc = revit.doc
app = revit.doc.Application
output = script.get_output()

# IFC класс -> Категории Revit
IFC_TO_REVIT_CATEGORIES = {
    "IFCWALL": [("Стены", BuiltInCategory.OST_Walls)],
    "IFCWALLSTANDARDCASE": [("Стены", BuiltInCategory.OST_Walls)],
    "IFCSLAB": [("Перекрытия", BuiltInCategory.OST_Floors)],
    "IFCCOLUMN": [("Колонны несущие", BuiltInCategory.OST_StructuralColumns)],
    "IFCBEAM": [("Каркас несущий", BuiltInCategory.OST_StructuralFraming)],
    "IFCMEMBER": [("Каркас несущий", BuiltInCategory.OST_StructuralFraming)],
    "IFCPLATE": [("Каркас несущий", BuiltInCategory.OST_StructuralFraming)],
    "IFCFOOTING": [("Фундамент несущий", BuiltInCategory.OST_StructuralFoundation)],
    "IFCPILE": [("Фундамент несущий", BuiltInCategory.OST_StructuralFoundation)],
    "IFCSTAIR": [("Лестницы", BuiltInCategory.OST_Stairs)],
    "IFCSTAIRFLIGHT": [("Лестницы", BuiltInCategory.OST_Stairs)],
    "IFCRAMP": [("Пандусы", BuiltInCategory.OST_Ramps)],
    "IFCRAMPFLIGHT": [("Пандусы", BuiltInCategory.OST_Ramps)],
    "IFCRAILING": [("Ограждения", BuiltInCategory.OST_StairsRailing)],
    "IFCREINFORCINGBAR": [("Арматура несущая", BuiltInCategory.OST_Rebar)],
    "IFCELEMENTASSEMBLY": [("Обобщенные модели", BuiltInCategory.OST_GenericModel)],
    "IFCBUILDINGELEMENTPROXY": [("Обобщенные модели", BuiltInCategory.OST_GenericModel)],
    "IFCROOF": [("Крыши", BuiltInCategory.OST_Roofs)],
    "IFCCOVERING": [("Полы", BuiltInCategory.OST_Floors), ("Потолки", BuiltInCategory.OST_Ceilings)],
    "IFCMECHANICALFASTENERTYPE": [("Несущие соединения", BuiltInCategory.OST_StructConnections)],
    "IFCDISCRETEACCESSORY": [("Несущие соединения", BuiltInCategory.OST_StructConnections)],
    "IFCBUILDING": [],  # Параметры проекта, не элемента
    "IFCMATERIAL": [],  # Материалы - отдельная обработка
}

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
        lbl_ids.Text = "IDS файл (опционально, для автоопределения категорий):"
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

        # === Список параметров ===
        grp_params = GroupBox()
        grp_params.Text = "Параметры из ФОП (выберите для добавления)"
        grp_params.Location = Point(15, y)
        grp_params.Size = Size(300, 280)

        self.lst_params = CheckedListBox()
        self.lst_params.Location = Point(10, 20)
        self.lst_params.Size = Size(280, 250)
        self.lst_params.CheckOnClick = True
        self.lst_params.SelectedIndexChanged += self.on_param_selected
        grp_params.Controls.Add(self.lst_params)

        self.Controls.Add(grp_params)

        # === Список категорий ===
        grp_cats = GroupBox()
        grp_cats.Text = "Категории Revit (куда добавить)"
        grp_cats.Location = Point(325, y)
        grp_cats.Size = Size(300, 280)

        self.lst_cats = CheckedListBox()
        self.lst_cats.Location = Point(10, 20)
        self.lst_cats.Size = Size(280, 250)
        self.lst_cats.CheckOnClick = True

        # Заполнить категории
        for cat_name, cat_id in ALL_CATEGORIES:
            self.lst_cats.Items.Add(cat_name, False)

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

        # === Кнопки выбора ===
        btn_select_all_params = Button()
        btn_select_all_params.Text = "Все параметры"
        btn_select_all_params.Location = Point(15, y)
        btn_select_all_params.Width = 100
        btn_select_all_params.Click += self.on_select_all_params
        self.Controls.Add(btn_select_all_params)

        btn_apply_ids = Button()
        btn_apply_ids.Text = "Применить из IDS"
        btn_apply_ids.Location = Point(125, y)
        btn_apply_ids.Width = 110
        btn_apply_ids.Click += self.on_apply_ids_categories
        self.Controls.Add(btn_apply_ids)

        btn_select_all_cats = Button()
        btn_select_all_cats.Text = "Все категории"
        btn_select_all_cats.Location = Point(325, y)
        btn_select_all_cats.Width = 100
        btn_select_all_cats.Click += self.on_select_all_cats
        self.Controls.Add(btn_select_all_cats)

        btn_clear_cats = Button()
        btn_clear_cats.Text = "Очистить"
        btn_clear_cats.Location = Point(435, y)
        btn_clear_cats.Width = 80
        btn_clear_cats.Click += self.on_clear_cats
        self.Controls.Add(btn_clear_cats)

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
        try:
            self.ids_data = parse_ids_simple(self.ids_path)
            count = len(self.ids_data)
            self.lbl_status.Text = "IDS: {} параметров найдено".format(count)
            self.lbl_status.ForeColor = Color.DarkGreen
        except Exception as e:
            self.lbl_status.Text = "Ошибка IDS: {}".format(str(e))
            self.lbl_status.ForeColor = Color.Red

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
        """
        prefix = self.txt_prefix.Text.strip()
        if prefix and fop_name.startswith(prefix + "_"):
            return fop_name[len(prefix) + 1:]
        return fop_name

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
        try:
            self.parser = FOPParser(self.fop_path)
            self.parser.parse()

            # Заполнить список параметров
            self.lst_params.Items.Clear()
            matched_count = 0

            for param in self.parser.parameters:
                fop_name = param['name']
                ids_name = self.get_ids_name(fop_name)

                # Пометить, если есть в IDS (с учётом префикса)
                marker = ""
                if ids_name in self.ids_data:
                    marker = " [IDS]"
                    matched_count += 1

                display = "{}{}".format(fop_name, marker)
                self.lst_params.Items.Add(display, False)

            prefix = self.txt_prefix.Text.strip()
            status = "ФОП: {} параметров".format(len(self.parser.parameters))
            if prefix:
                status += " (префикс: {})".format(prefix)
            if self.ids_data:
                status += ", {} совпадают с IDS".format(matched_count)

            self.lbl_status.Text = status
            self.lbl_status.ForeColor = Color.Black
            self.btn_add.Enabled = True

        except Exception as e:
            self.lbl_status.Text = "Ошибка ФОП: {}".format(str(e))
            self.lbl_status.ForeColor = Color.Red

    def on_param_selected(self, sender, args):
        """При выборе параметра показать рекомендацию из IDS."""
        if self.lst_params.SelectedIndex < 0:
            return
        if not self.parser:
            return

        idx = self.lst_params.SelectedIndex
        param = self.parser.parameters[idx]
        fop_name = param['name']
        ids_name = self.get_ids_name(fop_name)

        if ids_name in self.ids_data:
            info = self.ids_data[ids_name]

            # Категории
            cats = get_revit_categories_for_param(info)
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

    def on_apply_ids_categories(self, sender, args):
        """Применить категории из IDS для выбранных параметров."""
        if not self.ids_data or not self.parser:
            show_warning("Внимание", "Сначала загрузите IDS и ФОП файлы")
            return

        # Собрать все категории для выбранных параметров
        selected_cats = set()
        is_type_count = 0
        is_instance_count = 0

        for i in range(self.lst_params.Items.Count):
            if self.lst_params.GetItemChecked(i):
                param = self.parser.parameters[i]
                fop_name = param['name']
                ids_name = self.get_ids_name(fop_name)  # Убираем префикс для поиска в IDS

                if ids_name in self.ids_data:
                    info = self.ids_data[ids_name]
                    cats = get_revit_categories_for_param(info)
                    for cat_name, _ in cats:
                        selected_cats.add(cat_name)

                    if info.get('is_type'):
                        is_type_count += 1
                    else:
                        is_instance_count += 1

        if not selected_cats:
            show_warning("Внимание", "Выберите параметры, которые есть в IDS",
                         details="Проверьте правильность префикса.")
            return

        # Отметить категории
        for i in range(self.lst_cats.Items.Count):
            cat_name = str(self.lst_cats.Items[i])
            self.lst_cats.SetItemChecked(i, cat_name in selected_cats)

        # Установить Instance/Type по большинству
        if is_type_count > is_instance_count:
            self.chk_instance.Checked = False
            msg = "Применены категории из IDS. Рекомендация: ПО ТИПУ (большинство параметров)"
        else:
            self.chk_instance.Checked = True
            msg = "Применены категории из IDS. Рекомендация: ПО ЭКЗЕМПЛЯРУ"

        self.lbl_status.Text = msg
        self.lbl_status.ForeColor = Color.DarkGreen

    def on_select_all_params(self, sender, args):
        """Выбрать все параметры."""
        for i in range(self.lst_params.Items.Count):
            self.lst_params.SetItemChecked(i, True)

    def on_select_all_cats(self, sender, args):
        """Выбрать все категории."""
        for i in range(self.lst_cats.Items.Count):
            self.lst_cats.SetItemChecked(i, True)

    def on_clear_cats(self, sender, args):
        """Очистить выбор категорий."""
        for i in range(self.lst_cats.Items.Count):
            self.lst_cats.SetItemChecked(i, False)

    def get_param_group(self):
        """Получить группу параметров Revit."""
        index = self.cmb_group.SelectedIndex
        if index < 0 or index >= len(self.param_groups):
            index = 0

        # Возвращаем напрямую объект группы (уже получен из Revit)
        group_name, group_obj = self.param_groups[index]
        return group_obj

    def on_add_params(self, sender, args):
        """Добавить выбранные параметры в проект."""
        # Получить выбранные параметры
        selected_indices = []
        for i in range(self.lst_params.Items.Count):
            if self.lst_params.GetItemChecked(i):
                selected_indices.append(i)

        if not selected_indices:
            show_warning("Внимание", "Выберите параметры для добавления")
            return

        # Получить выбранные категории
        selected_cat_indices = []
        for i in range(self.lst_cats.Items.Count):
            if self.lst_cats.GetItemChecked(i):
                selected_cat_indices.append(i)

        if not selected_cat_indices:
            show_warning("Внимание", "Выберите категории для добавления параметров")
            return

        # Открыть файл определений
        try:
            app.SharedParametersFilename = self.fop_path
            def_file = app.OpenSharedParameterFile()
            if def_file is None:
                show_error("Ошибка", "Не удалось открыть ФОП файл")
                return
        except Exception as e:
            show_error("Ошибка", "Ошибка открытия ФОП",
                       details=str(e))
            return

        # Создать CategorySet
        cat_set = CategorySet()
        for idx in selected_cat_indices:
            cat_name, cat_id = ALL_CATEGORIES[idx]
            try:
                cat = Category.GetCategory(doc, cat_id)
                if cat:
                    cat_set.Insert(cat)
            except:
                pass

        if cat_set.IsEmpty:
            show_error("Ошибка", "Не удалось получить категории")
            return

        # Параметры
        param_group = self.get_param_group()
        is_instance = self.chk_instance.Checked

        added_count = 0
        errors = []

        # Транзакция
        t = Transaction(doc, "Добавить параметры из ФОП")
        t.Start()

        try:
            for idx in selected_indices:
                param = self.parser.parameters[idx]
                param_name = param['name']
                group_name = param['group_name']

                # Найти определение в ФОП
                ext_def = None
                for grp in def_file.Groups:
                    if grp.Name == group_name:
                        ext_def = grp.Definitions.get_Item(param_name)
                        break

                if ext_def is None:
                    errors.append("Не найден: {}".format(param_name))
                    continue

                # Проверить, не существует ли уже
                binding = doc.ParameterBindings.get_Item(ext_def)
                if binding:
                    errors.append("Уже существует: {}".format(param_name))
                    continue

                # Создать привязку
                if is_instance:
                    new_binding = app.Create.NewInstanceBinding(cat_set)
                else:
                    new_binding = app.Create.NewTypeBinding(cat_set)

                # Добавить параметр
                if doc.ParameterBindings.Insert(ext_def, new_binding, param_group):
                    added_count += 1
                else:
                    errors.append("Ошибка добавления: {}".format(param_name))

            t.Commit()

        except Exception as e:
            t.RollBack()
            show_error("Ошибка", "Ошибка добавления параметров",
                       details=str(e))
            return

        # Результат
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
