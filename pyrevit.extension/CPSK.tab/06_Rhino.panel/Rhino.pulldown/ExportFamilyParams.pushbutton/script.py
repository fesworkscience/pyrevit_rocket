# -*- coding: utf-8 -*-
"""Выгрузка параметров семейств в YAML файл для интеграции с Rhino."""

__title__ = "Export\nParams"
__author__ = "CPSK"

# 1. Сначала import clr и стандартные модули
import clr
import os
import sys
import codecs
from datetime import datetime

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

# 2. System и WinForms импорты
import System
from System.Windows.Forms import (
    Form, Label, Button, Panel, CheckedListBox,
    TreeView, TreeNode, TreeViewAction,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, FolderBrowserDialog, AnchorStyles,
    CheckBox, Padding, BorderStyle
)
from System.Drawing import Point, Size, Font, FontStyle

# 3. Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# 4. Проверка авторизации
from cpsk_auth import require_auth
from cpsk_notify import show_error, show_success, show_warning, show_info
if not require_auth():
    sys.exit()

# 5. pyrevit и Revit API
from pyrevit import revit

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    FamilyInstance,
    FamilySymbol,
    BuiltInParameter,
    StorageType,
    ExternalDefinition,
    InternalDefinition,
    SharedParameterElement
)

doc = revit.doc


def get_storage_type_name(storage_type):
    """Получить название типа хранения параметра."""
    if storage_type == StorageType.String:
        return "String"
    elif storage_type == StorageType.Integer:
        return "Integer"
    elif storage_type == StorageType.Double:
        return "Double"
    elif storage_type == StorageType.ElementId:
        return "ElementId"
    return "None"


def get_param_value(param):
    """Получить значение параметра в зависимости от типа."""
    if param is None:
        return None

    storage_type = param.StorageType

    if storage_type == StorageType.String:
        return param.AsString()
    elif storage_type == StorageType.Integer:
        return param.AsInteger()
    elif storage_type == StorageType.Double:
        value = param.AsDouble()
        if value != 0:
            return round(value * 304.8, 2)
        return value
    elif storage_type == StorageType.ElementId:
        eid = param.AsElementId()
        if eid.IntegerValue > 0:
            elem = doc.GetElement(eid)
            if elem and hasattr(elem, 'Name'):
                return elem.Name
            return str(eid.IntegerValue)
        return str(eid.IntegerValue)
    return None


def get_param_description(param):
    """Получить описание параметра."""
    try:
        param_def = param.Definition
        if param_def is None:
            return ""

        # Для общих (shared) параметров - описание в ExternalDefinition
        if param.IsShared:
            # Пробуем получить SharedParameterElement
            try:
                guid = param.GUID
                if guid:
                    # Ищем SharedParameterElement по GUID
                    collector = FilteredElementCollector(doc).OfClass(SharedParameterElement)
                    for sp_elem in collector:
                        try:
                            sp_def = sp_elem.GetDefinition()
                            if sp_def and hasattr(sp_def, 'GUID') and sp_def.GUID == guid:
                                # Нашли! Пробуем получить описание
                                if hasattr(sp_def, 'Description'):
                                    return sp_def.Description or ""
                        except:
                            pass
            except:
                pass

        # Пробуем получить описание напрямую из Definition
        if hasattr(param_def, 'Description'):
            desc = param_def.Description
            if desc:
                return desc

        return ""
    except:
        return ""


def extract_params(element):
    """Извлечь все параметры из элемента."""
    params_list = []
    if element is None:
        return params_list

    for param in element.Parameters:
        try:
            param_def = param.Definition
            if param_def is None:
                continue

            param_info = {
                "name": param_def.Name,
                "storage_type": get_storage_type_name(param.StorageType),
                "is_read_only": param.IsReadOnly,
                "has_value": param.HasValue,
                "value": get_param_value(param) if param.HasValue else None,
                "description": get_param_description(param)
            }

            try:
                param_info["group"] = str(param_def.ParameterGroup)
            except:
                param_info["group"] = "Unknown"

            # Добавляем информацию о том, общий ли это параметр
            try:
                param_info["is_shared"] = param.IsShared
            except:
                param_info["is_shared"] = False

            params_list.append(param_info)
        except:
            pass

    params_list.sort(key=lambda x: x["name"])
    return params_list


def get_family_types(family_symbol):
    """Получить все типоразмеры семейства."""
    types_list = []
    if family_symbol is None or family_symbol.Family is None:
        return types_list

    family = family_symbol.Family
    symbol_ids = family.GetFamilySymbolIds()

    for symbol_id in symbol_ids:
        symbol = doc.GetElement(symbol_id)
        if symbol:
            type_name = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if type_name:
                types_list.append(type_name.AsString())
            else:
                types_list.append(str(symbol.Id.IntegerValue))

    types_list.sort()
    return types_list


def build_family_instances_cache():
    """Построить кэш экземпляров по Family Id для быстрого поиска."""
    cache = {}  # family_id -> first instance
    collector = FilteredElementCollector(doc).OfClass(FamilyInstance).WhereElementIsNotElementType()
    for instance in collector:
        try:
            family_id = instance.Symbol.Family.Id.IntegerValue
            if family_id not in cache:
                cache[family_id] = instance
        except:
            pass
    return cache


def collect_families_by_category():
    """Собрать ВСЕ загруженные семейства, сгруппированные по категориям."""
    # Кэш экземпляров для получения параметров экземпляра
    instances_cache = build_family_instances_cache()

    # Собираем все FamilySymbol (типоразмеры) - включая неиспользуемые!
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    all_symbols = list(collector)

    categories = {}

    for symbol in all_symbols:
        try:
            # Пропускаем системные семейства без Family
            family = symbol.Family
            if family is None:
                continue

            # Получаем категорию
            category = symbol.Category
            if category is None:
                category = family.FamilyCategory
            category_name = category.Name if category else "Unknown"

            if category_name not in categories:
                categories[category_name] = []

            family_name = family.Name
            type_name = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            type_name_str = type_name.AsString() if type_name else "Unknown"

            # Ищем любой экземпляр этого семейства для параметров экземпляра
            family_id = family.Id.IntegerValue
            sample_instance = instances_cache.get(family_id)

            categories[category_name].append({
                "symbol_id": symbol.Id.IntegerValue,
                "unique_id": symbol.UniqueId,
                "family_id": family_id,
                "family_unique_id": family.UniqueId,
                "family_name": family_name,
                "type_name": type_name_str,
                "display_name": "{} : {}".format(family_name, type_name_str),
                "sample_instance": sample_instance,  # Для получения параметров экземпляра
                "symbol": symbol
            })
        except:
            pass

    # Сортируем внутри категорий
    for cat in categories:
        categories[cat].sort(key=lambda x: x["display_name"])

    return categories


def yaml_escape(value):
    """Экранировать строку для YAML."""
    if value is None:
        return "null"

    s = str(value)
    if any(c in s for c in [':', '#', '[', ']', '{', '}', ',', '&', '*', '!', '|', '>', "'", '"', '%', '@', '`', '\n', '\r']):
        s = s.replace('\\', '\\\\').replace('"', '\\"')
        return '"{}"'.format(s)
    if s == "" or s.startswith(' ') or s.endswith(' '):
        return '"{}"'.format(s)
    return s


def write_yaml(filepath, data):
    """Записать данные в YAML файл с индексом для быстрого поиска."""
    lines = []
    index_entries = []

    # Заголовок
    lines.append("# ============================================")
    lines.append("# ПАРАМЕТРЫ СЕМЕЙСТВ REVIT")
    lines.append("# Файл для интеграции с Rhino/Grasshopper")
    lines.append("# ============================================")
    lines.append("#")
    lines.append("# Дата: {}".format(data["created"]))
    lines.append("# Проект: {}".format(data["project_name"]))
    lines.append("# Категорий: {}".format(len(data["categories"])))
    lines.append("# Семейств: {}".format(data["families_count"]))
    lines.append("#")
    lines.append("# ТЕРМИНОЛОГИЯ:")
    lines.append("# - family_name: Имя семейства (шаблон, например 'КЖ_Колонна')")
    lines.append("# - type_name: Имя типоразмера (вариант семейства, например '400x400')")
    lines.append("# - unique_id: UniqueId типоразмера (GUID, стабильный идентификатор)")
    lines.append("# - family_unique_id: UniqueId семейства (GUID родительского семейства)")
    lines.append("# - available_types: Все доступные типоразмеры данного семейства")
    lines.append("# - type_parameters: Параметры типа - привязаны к типоразмеру,")
    lines.append("#                    общие для всех экземпляров этого типоразмера")
    lines.append("# - instance_parameters: Параметры экземпляра - уникальны для каждого")
    lines.append("#                        размещённого объекта, заполняются при создании")
    lines.append("#")
    lines.append("# АТРИБУТЫ ПАРАМЕТРА:")
    lines.append("# - name: Имя параметра")
    lines.append("# - type: Тип данных (String, Integer, Double, ElementId)")
    lines.append("# - readonly: true/false - можно ли изменять параметр")
    lines.append("# - shared: true - общий параметр (из ФОП)")
    lines.append("# - description: Описание параметра (как заполнять)")
    lines.append("# - value: Текущее значение (если есть)")
    lines.append("#")
    lines.append("# НАВИГАЦИЯ:")
    lines.append("# - Используйте Ctrl+F для поиска по имени")
    lines.append("# - Индекс категорий в начале файла")
    lines.append("# - Каждая категория имеет маркер: ### CATEGORY: Имя ###")
    lines.append("# - Каждое семейство имеет маркер: ## FAMILY: Имя ##")
    lines.append("#")
    lines.append("# ============================================")
    lines.append("")

    # Метаданные
    lines.append("meta:")
    lines.append("  created: {}".format(yaml_escape(data["created"])))
    lines.append("  project_name: {}".format(yaml_escape(data["project_name"])))
    lines.append("  project_path: {}".format(yaml_escape(data["project_path"])))
    lines.append("  categories_count: {}".format(len(data["categories"])))
    lines.append("  families_count: {}".format(data["families_count"]))
    lines.append("")

    # Плейсхолдер для индекса (заполним позже)
    index_start_line = len(lines)
    lines.append("# INDEX (line numbers for quick navigation):")
    lines.append("index:")

    # Плейсхолдер - потом заменим
    index_placeholder_start = len(lines)

    lines.append("")
    lines.append("# ============================================")
    lines.append("# ДАННЫЕ СЕМЕЙСТВ")
    lines.append("# ============================================")
    lines.append("")
    lines.append("categories:")

    # Сортируем категории по алфавиту
    sorted_categories = sorted(data["categories"].keys())

    for category_name in sorted_categories:
        families = data["categories"][category_name]

        # Маркер категории для поиска
        category_line = len(lines) + 1  # +1 т.к. нумерация с 1
        lines.append("")
        lines.append("  # ### CATEGORY: {} ###".format(category_name))
        lines.append("  {}:".format(yaml_escape(category_name)))

        index_entries.append({
            "type": "category",
            "name": category_name,
            "line": category_line,
            "count": len(families)
        })

        for family in families:
            # Маркер семейства для поиска
            family_line = len(lines) + 1
            family_display = "{} : {}".format(family["family_name"], family["name"])

            lines.append("")
            lines.append("    # ## FAMILY: {} ##".format(family_display))
            lines.append("    - family_name: {}".format(yaml_escape(family["family_name"])))
            lines.append("      type_name: {}".format(yaml_escape(family["name"])))
            lines.append("      unique_id: {}".format(yaml_escape(family.get("unique_id", ""))))
            lines.append("      family_unique_id: {}".format(yaml_escape(family.get("family_unique_id", ""))))

            index_entries.append({
                "type": "family",
                "name": family_display,
                "line": family_line,
                "category": category_name
            })

            # Instance parameters
            lines.append("      instance_parameters:")
            if family["instance_parameters"]:
                for param in family["instance_parameters"]:
                    lines.append("        - name: {}".format(yaml_escape(param["name"])))
                    lines.append("          type: {}".format(param["storage_type"]))
                    lines.append("          readonly: {}".format(str(param["is_read_only"]).lower()))
                    if param.get("is_shared"):
                        lines.append("          shared: true")
                    if param.get("description"):
                        lines.append("          description: {}".format(yaml_escape(param["description"])))
                    if param["value"] is not None:
                        lines.append("          value: {}".format(yaml_escape(param["value"])))
            else:
                lines.append("        []")

            # Type parameters
            lines.append("      type_parameters:")
            if family["type_parameters"]:
                for param in family["type_parameters"]:
                    lines.append("        - name: {}".format(yaml_escape(param["name"])))
                    lines.append("          type: {}".format(param["storage_type"]))
                    lines.append("          readonly: {}".format(str(param["is_read_only"]).lower()))
                    if param.get("is_shared"):
                        lines.append("          shared: true")
                    if param.get("description"):
                        lines.append("          description: {}".format(yaml_escape(param["description"])))
                    if param["value"] is not None:
                        lines.append("          value: {}".format(yaml_escape(param["value"])))
            else:
                lines.append("        []")

            # Available types
            lines.append("      available_types:")
            if family["available_types"]:
                for type_name in family["available_types"]:
                    lines.append("        - {}".format(yaml_escape(type_name)))
            else:
                lines.append("        []")

    # Теперь формируем индекс
    index_lines = []
    index_lines.append("  # Категории:")
    for entry in index_entries:
        if entry["type"] == "category":
            index_lines.append("  # - {} (line {}, {} families)".format(
                entry["name"], entry["line"], entry["count"]))

    index_lines.append("  #")
    index_lines.append("  # Семейства (формат: Категория > Семейство : Тип @ строка):")

    current_category = None
    for entry in index_entries:
        if entry["type"] == "family":
            if entry["category"] != current_category:
                current_category = entry["category"]
                index_lines.append("  # [{}]".format(current_category))
            index_lines.append("  #   - {} @ line {}".format(entry["name"], entry["line"]))

    # Вставляем индекс в нужное место
    final_lines = lines[:index_placeholder_start] + index_lines + lines[index_placeholder_start:]

    with codecs.open(filepath, 'w', 'utf-8') as f:
        f.write('\n'.join(final_lines))


class FamilySelectionForm(Form):
    """Форма выбора семейств для выгрузки."""

    def __init__(self, categories):
        self.categories = categories
        self.selected_families = []
        self.save_path = None
        self.setup_form()

    def setup_form(self):
        self.Text = "Выгрузка параметров семейств"
        self.Width = 700
        self.Height = 600
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MinimumSize = Size(500, 400)

        # Верхняя панель с чекбоксами
        top_panel = Panel()
        top_panel.Dock = DockStyle.Top
        top_panel.Height = 60
        top_panel.Padding = Padding(10)

        self.chk_select_all = CheckBox()
        self.chk_select_all.Text = "Выбрать все"
        self.chk_select_all.Location = Point(10, 10)
        self.chk_select_all.Width = 150
        self.chk_select_all.CheckedChanged += self.on_select_all

        lbl_info = Label()
        lbl_info.Text = "Отметьте категории или разверните для выбора отдельных семейств:"
        lbl_info.Location = Point(10, 35)
        lbl_info.Width = 500

        top_panel.Controls.Add(self.chk_select_all)
        top_panel.Controls.Add(lbl_info)

        # TreeView для выбора
        self.tree = TreeView()
        self.tree.Dock = DockStyle.Fill
        self.tree.CheckBoxes = True
        self.tree.AfterCheck += self.on_node_check

        # Заполняем дерево
        total_families = 0
        sorted_categories = sorted(self.categories.keys())

        for category_name in sorted_categories:
            families = self.categories[category_name]
            cat_node = TreeNode("{} ({})".format(category_name, len(families)))
            cat_node.Tag = {"type": "category", "name": category_name}

            for family in families:
                family_node = TreeNode(family["display_name"])
                family_node.Tag = {"type": "family", "data": family}
                cat_node.Nodes.Add(family_node)
                total_families += 1

            self.tree.Nodes.Add(cat_node)

        # Нижняя панель с кнопками
        bottom_panel = Panel()
        bottom_panel.Dock = DockStyle.Bottom
        bottom_panel.Height = 100
        bottom_panel.Padding = Padding(10)

        # Путь сохранения
        lbl_path = Label()
        lbl_path.Text = "Папка сохранения:"
        lbl_path.Location = Point(10, 10)
        lbl_path.Width = 100

        self.txt_path = Label()
        self.txt_path.Text = self.get_default_path()
        self.txt_path.Location = Point(10, 30)
        self.txt_path.Width = 480
        self.txt_path.BorderStyle = BorderStyle.FixedSingle

        btn_browse = Button()
        btn_browse.Text = "Обзор..."
        btn_browse.Location = Point(500, 25)
        btn_browse.Width = 80
        btn_browse.Click += self.on_browse

        # Кнопки OK/Cancel
        btn_ok = Button()
        btn_ok.Text = "Выгрузить"
        btn_ok.Location = Point(500, 65)
        btn_ok.Width = 80
        btn_ok.Click += self.on_ok

        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(410, 65)
        btn_cancel.Width = 80
        btn_cancel.Click += self.on_cancel

        # Счетчик выбранных
        self.lbl_count = Label()
        self.lbl_count.Text = "Выбрано: 0 из {}".format(total_families)
        self.lbl_count.Location = Point(10, 70)
        self.lbl_count.Width = 200

        bottom_panel.Controls.Add(lbl_path)
        bottom_panel.Controls.Add(self.txt_path)
        bottom_panel.Controls.Add(btn_browse)
        bottom_panel.Controls.Add(btn_ok)
        bottom_panel.Controls.Add(btn_cancel)
        bottom_panel.Controls.Add(self.lbl_count)

        # Добавляем контролы в форму (Fill первым!)
        self.Controls.Add(self.tree)
        self.Controls.Add(bottom_panel)
        self.Controls.Add(top_panel)

    def get_default_path(self):
        """Получить путь по умолчанию."""
        if doc.PathName:
            return os.path.dirname(doc.PathName)
        return os.path.expanduser("~\\Desktop")

    def on_select_all(self, sender, args):
        """Выбрать/снять все."""
        checked = self.chk_select_all.Checked
        for cat_node in self.tree.Nodes:
            cat_node.Checked = checked
            for family_node in cat_node.Nodes:
                family_node.Checked = checked
        self.update_count()

    def on_node_check(self, sender, args):
        """Обработка изменения чекбокса."""
        node = args.Node
        if args.Action == TreeViewAction.Unknown:
            return

        # Если это категория - отмечаем все дочерние
        if node.Tag and node.Tag.get("type") == "category":
            for child in node.Nodes:
                child.Checked = node.Checked

        # Если это семейство - проверяем родителя
        if node.Parent:
            all_checked = True
            for sibling in node.Parent.Nodes:
                if not sibling.Checked:
                    all_checked = False
                    break
            # Не меняем родителя программно, чтобы избежать рекурсии

        self.update_count()

    def update_count(self):
        """Обновить счетчик выбранных."""
        count = 0
        total = 0
        for cat_node in self.tree.Nodes:
            for family_node in cat_node.Nodes:
                total += 1
                if family_node.Checked:
                    count += 1
        self.lbl_count.Text = "Выбрано: {} из {}".format(count, total)

    def on_browse(self, sender, args):
        """Выбор папки."""
        dialog = FolderBrowserDialog()
        dialog.Description = "Выберите папку для сохранения YAML файла"
        dialog.SelectedPath = self.txt_path.Text

        if dialog.ShowDialog() == DialogResult.OK:
            self.txt_path.Text = dialog.SelectedPath

    def on_ok(self, sender, args):
        """Подтверждение выбора."""
        self.selected_families = []

        for cat_node in self.tree.Nodes:
            for family_node in cat_node.Nodes:
                if family_node.Checked:
                    self.selected_families.append(family_node.Tag["data"])

        if not self.selected_families:
            show_warning("Выбор", "Выберите хотя бы одно семейство для выгрузки")
            return

        self.save_path = self.txt_path.Text
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        """Отмена."""
        self.DialogResult = DialogResult.Cancel
        self.Close()


def process_families(selected_families):
    """Обработать выбранные семейства."""
    families_by_category = {}

    # Кэш параметров экземпляра по family_id (одинаковы для всех типоразмеров)
    instance_params_cache = {}

    for family_data in selected_families:
        try:
            sample_instance = family_data.get("sample_instance")
            symbol = family_data["symbol"]
            family_id = family_data.get("family_id", 0)

            # Получаем категорию из symbol
            category = symbol.Category
            if category is None and symbol.Family:
                category = symbol.Family.FamilyCategory
            category_name = category.Name if category else "Unknown"

            if category_name not in families_by_category:
                families_by_category[category_name] = []

            # Параметры экземпляра - берём из кэша или извлекаем
            if family_id in instance_params_cache:
                instance_params = instance_params_cache[family_id]
            elif sample_instance is not None:
                instance_params = extract_params(sample_instance)
                instance_params_cache[family_id] = instance_params
            else:
                instance_params = []
                instance_params_cache[family_id] = []

            family_info = {
                "name": family_data["type_name"],
                "family_name": family_data["family_name"],
                "unique_id": symbol.UniqueId,
                "family_unique_id": family_data.get("family_unique_id", ""),
                "instance_parameters": instance_params,
                "type_parameters": extract_params(symbol),
                "available_types": get_family_types(symbol)
            }

            families_by_category[category_name].append(family_info)
        except Exception as e:
            show_warning("Ошибка", "Не удалось обработать семейство: {}".format(family_data.get("display_name", "Unknown")))

    # Сортируем внутри категорий
    for cat in families_by_category:
        families_by_category[cat].sort(key=lambda x: "{} : {}".format(x["family_name"], x["name"]))

    return families_by_category


def main():
    """Основная функция."""
    # Проверяем, сохранён ли проект
    if not doc.PathName:
        show_warning("Проект не сохранён", "Сохраните проект перед выгрузкой параметров")
        return

    # Собираем семейства по категориям
    show_info("Сбор данных", "Анализ семейств в проекте...")

    categories = collect_families_by_category()

    if not categories:
        show_warning("Нет данных", "В проекте не найдено семейств для выгрузки")
        return

    # Показываем форму выбора
    form = FamilySelectionForm(categories)
    result = form.ShowDialog()

    if result != DialogResult.OK:
        return

    selected_families = form.selected_families
    save_path = form.save_path

    if not selected_families:
        return

    # Обрабатываем семейства
    show_info("Обработка", "Извлечение параметров из {} семейств...".format(len(selected_families)))

    families_by_category = process_families(selected_families)

    # Считаем общее количество
    total_count = sum(len(fams) for fams in families_by_category.values())

    # Формируем имя файла
    project_name = os.path.splitext(os.path.basename(doc.PathName))[0]
    yaml_filename = "{}_families.yaml".format(project_name)
    yaml_path = os.path.join(save_path, yaml_filename)

    # Формируем данные
    yaml_data = {
        "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "project_name": project_name,
        "project_path": doc.PathName,
        "families_count": total_count,
        "categories": families_by_category
    }

    # Записываем YAML
    try:
        write_yaml(yaml_path, yaml_data)
        show_success(
            "Готово",
            "Выгружено {} семейств из {} категорий".format(total_count, len(families_by_category)),
            details="Файл: {}".format(yaml_path)
        )
    except Exception as e:
        show_error("Ошибка записи", "Не удалось сохранить файл", details=str(e))


if __name__ == "__main__":
    main()
