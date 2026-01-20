# -*- coding: utf-8 -*-
"""
IDS Checker - проверка параметров всех элементов на виде Revit на соответствие требованиям IDS.
Генерирует HTML-отчёт с древовидной структурой результатов.
"""

__title__ = "Проверка\nIDS"
__author__ = "CPSK"

import clr
import os
import sys
import datetime
import codecs

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System.Xml')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, ProgressBar,
    FormStartPosition, FormBorderStyle, GroupBox,
    Application
)
from System.Drawing import Point, Size, Color
from System.Xml import XmlDocument, XmlNamespaceManager

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
from cpsk_notify import show_error, show_warning, show_info
from cpsk_auth import require_auth
from cpsk_logger import Logger

# Импорт маппинга IFC -> Revit категорий
from ifc_mappings import IFC_TO_REVIT_CATEGORY_IDS

# Имя скрипта для логирования
SCRIPT_NAME = "IDSChecker"

# Проверка авторизации
if not require_auth():
    sys.exit()

# Инициализация логгера
Logger.init(SCRIPT_NAME)
Logger.info(SCRIPT_NAME, "Скрипт запущен")

# Revit API
from Autodesk.Revit.DB import FilteredElementCollector, ElementId, BuiltInCategory

# === НАСТРОЙКИ ===

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


# === ОБРАТНЫЙ МАППИНГ КАТЕГОРИЙ ===

class ReverseCategoryMapper:
    """Обратный маппинг BuiltInCategory ID -> IFC классы."""

    def __init__(self):
        self._reverse_map = {}
        self._build_reverse_map()

    def _build_reverse_map(self):
        """Построить обратный маппинг из IFC_TO_REVIT_CATEGORY_IDS."""
        for ifc_class, categories in IFC_TO_REVIT_CATEGORY_IDS.items():
            # Пропускаем TYPE классы - нас интересуют только экземпляры
            if ifc_class.endswith("TYPE"):
                continue
            for cat_name, cat_id in categories:
                if cat_id not in self._reverse_map:
                    self._reverse_map[cat_id] = []
                if ifc_class not in self._reverse_map[cat_id]:
                    self._reverse_map[cat_id].append(ifc_class)

    def get_ifc_classes(self, category_id):
        """Получить список IFC классов для данной категории Revit."""
        return self._reverse_map.get(category_id, [])

    def get_all_category_ids(self):
        """Получить все ID категорий, которые имеют маппинг."""
        return list(self._reverse_map.keys())


# === ПАРСЕР ФАЙЛА МЭППИНГА ПАРАМЕТРОВ ===

class MappingParser:
    """Парсер файла мэппинга параметров IFC -> Revit."""

    def __init__(self):
        self.param_mapping = {}  # IFC_param_name -> Revit_param_name
        self.property_sets = {}  # PropertySet name -> list of params

    def parse(self, path):
        """Парсить файл мэппинга."""
        if not os.path.exists(path):
            raise IOError("Файл мэппинга не найден: {}".format(path))

        current_pset = None

        with codecs.open(path, 'r', 'utf-8') as f:
            for line in f:
                line = line.rstrip('\r\n')

                # Пустая строка - конец PropertySet
                if not line.strip():
                    current_pset = None
                    continue

                # Строка PropertySet
                if line.startswith("PropertySet:"):
                    parts = line.split('\t')
                    if len(parts) >= 2:
                        current_pset = parts[1].strip()
                        self.property_sets[current_pset] = []
                    continue

                # Строка параметра (начинается с табуляции)
                if line.startswith('\t') and current_pset:
                    parts = line.split('\t')
                    # Формат: \t<IFC Name>\t<Type>\t<Revit Param Name>
                    if len(parts) >= 4:
                        ifc_name = parts[1].strip()
                        revit_name = parts[3].strip()
                        if ifc_name and revit_name:
                            self.param_mapping[ifc_name] = revit_name
                            self.property_sets[current_pset].append({
                                "ifc_name": ifc_name,
                                "revit_name": revit_name
                            })

        return self

    def get_revit_param_name(self, ifc_name):
        """Получить имя параметра Revit по имени IFC."""
        return self.param_mapping.get(ifc_name, ifc_name)


# === ПАРСЕР IDS ===

class IDSParser:
    """Парсер IDS файла для извлечения типов и параметров (baseName!)."""

    _XmlDocument = XmlDocument
    _XmlNamespaceManager = XmlNamespaceManager

    def __init__(self, ids_path):
        self.ids_path = ids_path
        self.specifications = []
        self.entity_params = {}  # Entity -> [params]
        self.entity_predefined_types = {}  # Entity -> [predefinedTypes]

    def parse(self):
        """Парсить IDS файл."""
        doc = self._XmlDocument()
        doc.Load(self.ids_path)

        nsm = self._XmlNamespaceManager(doc.NameTable)
        nsm.AddNamespace("ids", "http://standards.buildingsmart.org/IDS")
        nsm.AddNamespace("xs", "http://www.w3.org/2001/XMLSchema")

        spec_nodes = doc.SelectNodes("//ids:specification", nsm)

        for spec_node in spec_nodes:
            spec = self._parse_specification(spec_node, nsm)
            self.specifications.append(spec)

        # Собрать параметры по типам
        self._collect_entity_params()

        return self

    def _parse_specification(self, spec_node, nsm):
        """Парсить одну спецификацию."""
        spec = {
            "name": spec_node.GetAttribute("name") or "",
            "applicability": [],
            "predefinedTypes": [],
            "requirements": []
        }

        # Applicability - к каким элементам применяется
        app_nodes = spec_node.SelectNodes("ids:applicability/ids:entity/ids:name", nsm)
        for app_node in app_nodes:
            entities = self._get_all_node_values(app_node, nsm)
            for entity in entities:
                if entity:
                    upper_entity = entity.upper()
                    if not upper_entity.endswith("TYPE"):
                        spec["applicability"].append(upper_entity)

        # PredefinedType из applicability
        pred_node = spec_node.SelectSingleNode("ids:applicability/ids:entity/ids:predefinedType", nsm)
        if pred_node:
            pred_values = self._get_all_enum_values(pred_node, nsm)
            if pred_values:
                filtered = [v for v in pred_values if v not in ("USERDEFINED", "NOTDEFINED")]
                spec["predefinedTypes"] = filtered

        # Requirements - параметры
        req_nodes = spec_node.SelectNodes("ids:requirements/ids:property", nsm)
        for req_node in req_nodes:
            prop = self._parse_property(req_node, nsm)
            if prop:
                spec["requirements"].append(prop)

        return spec

    def _get_all_enum_values(self, node, nsm):
        """Получить все значения enumeration из узла."""
        values = []
        for child in node.ChildNodes:
            if child.LocalName == "restriction":
                for enum_node in child.ChildNodes:
                    if enum_node.LocalName == "enumeration":
                        val = enum_node.GetAttribute("value")
                        if val:
                            values.append(val)
        return values if values else None

    def _get_node_value(self, node, nsm):
        """Получить значение узла."""
        values = self._get_all_node_values(node, nsm)
        return values[0] if values else None

    def _get_all_node_values(self, node, nsm):
        """Получить все значения узла."""
        if node is None:
            return []

        values = []

        simple = node.SelectSingleNode("ids:simpleValue", nsm)
        if simple:
            return [simple.InnerText]

        for child in node.ChildNodes:
            if child.LocalName == "restriction":
                for enum_node in child.ChildNodes:
                    if enum_node.LocalName == "enumeration":
                        val = enum_node.GetAttribute("value")
                        if val:
                            values.append(val)

        if values:
            return values

        if node.InnerText:
            return [node.InnerText]

        return []

    def _parse_property(self, prop_node, nsm):
        """Парсить требование к параметру."""
        prop = {
            "propertySet": "",
            "baseName": "",
            "dataType": "IFCTEXT",
            "cardinality": "required",
            "enumeration": None
        }

        pset_node = prop_node.SelectSingleNode("ids:propertySet", nsm)
        if pset_node:
            prop["propertySet"] = self._get_node_value(pset_node, nsm) or ""

        name_node = prop_node.SelectSingleNode("ids:baseName", nsm)
        if name_node:
            prop["baseName"] = self._get_node_value(name_node, nsm) or ""

        datatype_attr = prop_node.GetAttribute("dataType")
        if datatype_attr:
            prop["dataType"] = datatype_attr.upper()

        cardinality = prop_node.GetAttribute("cardinality")
        if cardinality:
            prop["cardinality"] = cardinality

        enum_nodes = prop_node.SelectNodes("ids:value/ids:restriction/ids:enumeration", nsm)
        if enum_nodes and enum_nodes.Count > 0:
            enums = []
            for enum_node in enum_nodes:
                val = enum_node.GetAttribute("value")
                if val:
                    enums.append(val)
            if enums:
                prop["enumeration"] = enums

        return prop if prop["baseName"] else None

    def _collect_entity_params(self):
        """Собрать параметры по типам IFC."""
        for spec in self.specifications:
            for entity in spec["applicability"]:
                if entity not in self.entity_params:
                    self.entity_params[entity] = []
                if entity not in self.entity_predefined_types:
                    self.entity_predefined_types[entity] = []

                for ptype in spec.get("predefinedTypes", []):
                    if ptype not in self.entity_predefined_types[entity]:
                        self.entity_predefined_types[entity].append(ptype)

                for prop in spec["requirements"]:
                    exists = False
                    for existing in self.entity_params[entity]:
                        if existing["baseName"] == prop["baseName"]:
                            exists = True
                            break
                    if not exists:
                        self.entity_params[entity].append(prop)

    def get_entities(self):
        """Получить список типов IFC."""
        return sorted(self.entity_params.keys())

    def get_params_for_entity(self, entity):
        """Получить параметры для указанного типа."""
        return self.entity_params.get(entity, [])

    def get_params_for_entities(self, entities):
        """Получить объединённые параметры для нескольких типов IFC."""
        all_params = []
        seen_names = set()

        for entity in entities:
            params = self.entity_params.get(entity, [])
            for param in params:
                if param["baseName"] not in seen_names:
                    all_params.append(param)
                    seen_names.add(param["baseName"])

        return all_params


# === ПРОВЕРКА ЭЛЕМЕНТОВ ===

def get_param_value(param):
    """Получить значение параметра как строку."""
    if param is None:
        return ""
    if not param.HasValue:
        return ""

    storage = param.StorageType
    if str(storage) == "String":
        return param.AsString() or ""
    elif str(storage) == "Integer":
        return str(param.AsInteger())
    elif str(storage) == "Double":
        return str(round(param.AsDouble(), 4))
    elif str(storage) == "ElementId":
        eid = param.AsElementId()
        if eid and eid.IntegerValue > 0:
            el = doc.GetElement(eid)
            if el:
                return el.Name if hasattr(el, 'Name') else str(eid.IntegerValue)
        return ""
    return ""


class ViewElementChecker:
    """Проверка всех элементов на виде."""

    def __init__(self, ids_parser, mapping_parser, category_mapper):
        self.ids_parser = ids_parser
        self.mapping_parser = mapping_parser
        self.category_mapper = category_mapper
        self._doc = doc

    def check_view(self, view, progress_callback=None):
        """
        Проверить все элементы на виде.
        Возвращает словарь: category_name -> [element_results]
        """
        results = {}

        # Собрать все элементы на виде
        collector = FilteredElementCollector(self._doc, view.Id)
        all_elements = list(collector.WhereElementIsNotElementType().ToElements())

        total = len(all_elements)
        processed = 0

        for elem in all_elements:
            if elem.Category is None:
                processed += 1
                continue

            cat_id = elem.Category.Id.IntegerValue
            cat_name = elem.Category.Name

            # Получить IFC классы для этой категории
            ifc_classes = self.category_mapper.get_ifc_classes(cat_id)

            if not ifc_classes:
                processed += 1
                continue

            # Получить требования из IDS для этих классов
            required_params = self.ids_parser.get_params_for_entities(ifc_classes)

            if not required_params:
                processed += 1
                continue

            # Проверить элемент
            elem_result = self._check_element(elem, required_params)

            if cat_name not in results:
                results[cat_name] = []

            results[cat_name].append(elem_result)

            processed += 1
            if progress_callback:
                progress_callback(processed, total)

        return results

    def _check_element(self, element, required_params):
        """Проверить один элемент на соответствие требованиям."""
        elem_id = element.Id.IntegerValue
        elem_name = element.Name if hasattr(element, 'Name') else "Без имени"

        param_results = []
        passed = 0
        failed = 0
        optional_missing = 0

        for prop in required_params:
            ifc_name = prop["baseName"]
            # Получить имя параметра в Revit через мэппинг
            revit_name = self.mapping_parser.get_revit_param_name(ifc_name)

            is_required = prop["cardinality"] == "required"

            # Поиск параметра
            param = element.LookupParameter(revit_name)
            has_param = param is not None
            value = get_param_value(param) if has_param else ""

            # Определить статус
            if has_param and value:
                status = "ok"
                passed += 1
            elif has_param and not value:
                if is_required:
                    status = "empty"
                    failed += 1
                else:
                    status = "optional_empty"
                    optional_missing += 1
            else:
                if is_required:
                    status = "missing"
                    failed += 1
                else:
                    status = "optional_missing"
                    optional_missing += 1

            param_results.append({
                "ifc_name": ifc_name,
                "revit_name": revit_name,
                "has_param": has_param,
                "value": value,
                "is_required": is_required,
                "status": status
            })

        return {
            "id": elem_id,
            "name": elem_name,
            "params": param_results,
            "passed": passed,
            "failed": failed,
            "optional_missing": optional_missing,
            "total": len(required_params)
        }


# === ГЕНЕРАТОР HTML ОТЧЁТА ===

class HTMLReportGenerator:
    """Генератор HTML отчёта."""

    def __init__(self):
        self.css = """
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        h1 {
            color: #333;
            border-bottom: 2px solid #4CAF50;
            padding-bottom: 10px;
        }
        .info {
            background-color: #e7f3fe;
            border-left: 4px solid #2196F3;
            padding: 10px;
            margin-bottom: 20px;
        }
        .summary {
            background-color: #fff;
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .summary-item {
            display: inline-block;
            margin-right: 30px;
            padding: 5px 10px;
            border-radius: 4px;
        }
        .summary-ok { background-color: #dff0d8; color: #3c763d; }
        .summary-fail { background-color: #f2dede; color: #a94442; }
        .summary-total { background-color: #d9edf7; color: #31708f; }
        details {
            margin: 5px 0;
            padding: 5px;
            border: 1px solid #ddd;
            border-radius: 4px;
            background-color: #fff;
        }
        details details {
            margin-left: 20px;
            background-color: #fafafa;
        }
        details details details {
            background-color: #fff;
            border: none;
            border-left: 2px solid #ddd;
        }
        summary {
            cursor: pointer;
            padding: 8px;
            font-weight: bold;
            outline: none;
        }
        summary:hover {
            background-color: #f0f0f0;
        }
        .category-summary {
            font-size: 16px;
        }
        .element-summary {
            font-size: 14px;
        }
        .ok { color: #3c763d; }
        .fail { color: #a94442; }
        .optional { color: #8a6d3b; }
        .param-list {
            list-style: none;
            padding-left: 20px;
            margin: 5px 0;
        }
        .param-item {
            padding: 3px 0;
            font-size: 13px;
        }
        .param-name { font-weight: bold; }
        .param-value { color: #666; font-style: italic; }
        .status-icon { margin-right: 5px; }
        """

    def generate(self, results, view_name):
        """Генерировать HTML отчёт."""
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Подсчёт общей статистики
        total_elements = 0
        total_ok = 0
        total_fail = 0

        for cat_name, elements in results.items():
            for elem in elements:
                total_elements += 1
                if elem["failed"] == 0:
                    total_ok += 1
                else:
                    total_fail += 1

        html = []
        html.append("<!DOCTYPE html>")
        html.append("<html>")
        html.append("<head>")
        html.append('<meta charset="utf-8">')
        html.append("<title>IDS Check Report</title>")
        html.append("<style>{}</style>".format(self.css))
        html.append("</head>")
        html.append("<body>")

        html.append("<h1>IDS Check Report</h1>")

        html.append('<div class="info">')
        html.append("<strong>Вид:</strong> {} | ".format(view_name))
        html.append("<strong>Дата:</strong> {} | ".format(now))
        html.append("<strong>Категорий:</strong> {}".format(len(results)))
        html.append("</div>")

        html.append('<div class="summary">')
        html.append('<span class="summary-item summary-total">Всего элементов: {}</span>'.format(total_elements))
        html.append('<span class="summary-item summary-ok">Прошли проверку: {}</span>'.format(total_ok))
        html.append('<span class="summary-item summary-fail">Не прошли: {}</span>'.format(total_fail))
        html.append("</div>")

        # Сортируем категории по имени
        for cat_name in sorted(results.keys()):
            elements = results[cat_name]
            cat_ok = sum(1 for e in elements if e["failed"] == 0)
            cat_fail = len(elements) - cat_ok

            cat_class = "ok" if cat_fail == 0 else "fail"

            html.append("<details>")
            html.append('<summary class="category-summary {}">'.format(cat_class))
            html.append("{} ({} элементов) - ".format(cat_name, len(elements)))
            html.append('<span class="ok">{} OK</span>, '.format(cat_ok))
            html.append('<span class="fail">{} Failed</span>'.format(cat_fail))
            html.append("</summary>")

            # Сортируем элементы: сначала с ошибками
            sorted_elements = sorted(elements, key=lambda e: (e["failed"] == 0, e["id"]))

            for elem in sorted_elements:
                elem_class = "ok" if elem["failed"] == 0 else "fail"
                status_icon = "&#10004;" if elem["failed"] == 0 else "&#10008;"

                html.append("<details>")
                html.append('<summary class="element-summary {}">'.format(elem_class))
                html.append('<span class="status-icon">{}</span>'.format(status_icon))
                html.append("ID: {} - {} ".format(elem["id"], elem["name"]))
                html.append("({}/{} параметров)".format(elem["passed"], elem["total"]))
                html.append("</summary>")

                html.append('<ul class="param-list">')
                # Сортируем параметры: сначала с ошибками
                sorted_params = sorted(elem["params"], key=lambda p: (p["status"] == "ok", p["revit_name"]))

                for param in sorted_params:
                    if param["status"] == "ok":
                        p_class = "ok"
                        p_icon = "&#10004;"
                    elif param["status"] in ("missing", "empty"):
                        p_class = "fail"
                        p_icon = "&#10008;"
                    else:
                        p_class = "optional"
                        p_icon = "&#9888;"

                    html.append('<li class="param-item {}">'.format(p_class))
                    html.append('<span class="status-icon">{}</span>'.format(p_icon))
                    html.append('<span class="param-name">{}</span>: '.format(param["revit_name"]))
                    if param["value"]:
                        html.append('<span class="param-value">"{}"</span>'.format(param["value"]))
                    else:
                        html.append('<span class="param-value">(пусто)</span>')
                    html.append("</li>")

                html.append("</ul>")
                html.append("</details>")

            html.append("</details>")

        html.append("</body>")
        html.append("</html>")

        return "\n".join(html)

    def save(self, path, content):
        """Сохранить HTML в файл."""
        with codecs.open(path, 'w', 'utf-8') as f:
            f.write(content)


# === ГЛАВНОЕ ОКНО ===

class IDSCheckerForm(Form):
    """Диалог проверки параметров элементов на соответствие IDS."""

    _show_error = staticmethod(show_error)
    _show_warning = staticmethod(show_warning)
    _show_info = staticmethod(show_info)

    def __init__(self):
        self.ids_path = None
        self.mapping_path = None
        self.ids_parser = None
        self.mapping_parser = None
        self.category_mapper = ReverseCategoryMapper()
        self.check_results = None
        self.html_content = None

        self._forms = forms
        self._os = os
        self._doc = doc
        self._Color = Color

        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "IDS Checker - Проверка элементов на виде"
        self.Width = 550
        self.Height = 340
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = True
        self.TopMost = True

        y = 15

        # === IDS файл ===
        grp_ids = GroupBox()
        grp_ids.Text = "IDS файл"
        grp_ids.Location = Point(15, y)
        grp_ids.Size = Size(505, 55)

        self.txt_ids = TextBox()
        self.txt_ids.Location = Point(15, 22)
        self.txt_ids.Width = 395
        self.txt_ids.ReadOnly = True
        grp_ids.Controls.Add(self.txt_ids)

        self.btn_browse_ids = Button()
        self.btn_browse_ids.Text = "Обзор..."
        self.btn_browse_ids.Location = Point(420, 20)
        self.btn_browse_ids.Width = 75
        self.btn_browse_ids.Click += self.on_browse_ids
        grp_ids.Controls.Add(self.btn_browse_ids)

        self.Controls.Add(grp_ids)
        y += 65

        # === Файл мэппинга ===
        grp_mapping = GroupBox()
        grp_mapping.Text = "Файл мэппинга параметров"
        grp_mapping.Location = Point(15, y)
        grp_mapping.Size = Size(505, 55)

        self.txt_mapping = TextBox()
        self.txt_mapping.Location = Point(15, 22)
        self.txt_mapping.Width = 395
        self.txt_mapping.ReadOnly = True
        grp_mapping.Controls.Add(self.txt_mapping)

        self.btn_browse_mapping = Button()
        self.btn_browse_mapping.Text = "Обзор..."
        self.btn_browse_mapping.Location = Point(420, 20)
        self.btn_browse_mapping.Width = 75
        self.btn_browse_mapping.Click += self.on_browse_mapping
        grp_mapping.Controls.Add(self.btn_browse_mapping)

        self.Controls.Add(grp_mapping)
        y += 65

        # === Прогресс и статус ===
        grp_progress = GroupBox()
        grp_progress.Text = "Проверка"
        grp_progress.Location = Point(15, y)
        grp_progress.Size = Size(505, 85)

        self.progress_bar = ProgressBar()
        self.progress_bar.Location = Point(15, 25)
        self.progress_bar.Size = Size(480, 23)
        self.progress_bar.Minimum = 0
        self.progress_bar.Maximum = 100
        grp_progress.Controls.Add(self.progress_bar)

        self.lbl_status = Label()
        self.lbl_status.Text = "Загрузите IDS и файл мэппинга для начала"
        self.lbl_status.Location = Point(15, 55)
        self.lbl_status.Size = Size(480, 20)
        grp_progress.Controls.Add(self.lbl_status)

        self.Controls.Add(grp_progress)
        y += 95

        # === Кнопки ===
        self.btn_check = Button()
        self.btn_check.Text = "Проверить вид"
        self.btn_check.Location = Point(15, y)
        self.btn_check.Size = Size(120, 30)
        self.btn_check.Click += self.on_check_view
        self.btn_check.Enabled = False
        self.Controls.Add(self.btn_check)

        self.btn_save = Button()
        self.btn_save.Text = "Сохранить отчёт"
        self.btn_save.Location = Point(145, y)
        self.btn_save.Size = Size(120, 30)
        self.btn_save.Click += self.on_save_report
        self.btn_save.Enabled = False
        self.Controls.Add(self.btn_save)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(445, y)
        btn_close.Size = Size(75, 30)
        btn_close.Click += self.on_close
        self.Controls.Add(btn_close)

    def on_browse_ids(self, sender, args):
        """Выбор IDS файла."""
        selected = self._forms.pick_file(
            file_ext='ids',
            init_dir=SCRIPT_DIR if self._os.path.exists(SCRIPT_DIR) else None,
            title="Выберите IDS файл"
        )

        if selected:
            self.ids_path = selected
            self.txt_ids.Text = self.ids_path
            Logger.file_opened(SCRIPT_NAME, self.ids_path, "IDS файл")

            try:
                self.ids_parser = IDSParser(self.ids_path)
                self.ids_parser.parse()

                entities = self.ids_parser.get_entities()
                Logger.info(SCRIPT_NAME, "IDS загружен. IFC типов: {}".format(len(entities)))
                Logger.debug(SCRIPT_NAME, "IFC типы: {}".format(", ".join(entities[:10])))

                self.lbl_status.Text = "IDS загружен. Типов: {}".format(len(entities))
                self.lbl_status.ForeColor = self._Color.Black

                self._update_check_button()

            except Exception as e:
                Logger.error(SCRIPT_NAME, "Ошибка парсинга IDS: {}".format(str(e)), exc_info=True)
                self.lbl_status.Text = "Ошибка парсинга IDS: {}".format(str(e))
                self.lbl_status.ForeColor = self._Color.Red
                self.ids_parser = None

    def on_browse_mapping(self, sender, args):
        """Выбор файла мэппинга."""
        selected = self._forms.pick_file(
            file_ext='txt',
            init_dir=SCRIPT_DIR if self._os.path.exists(SCRIPT_DIR) else None,
            title="Выберите файл мэппинга параметров"
        )

        if selected:
            self.mapping_path = selected
            self.txt_mapping.Text = self.mapping_path
            Logger.file_opened(SCRIPT_NAME, self.mapping_path, "Файл мэппинга параметров")

            try:
                self.mapping_parser = MappingParser()
                self.mapping_parser.parse(self.mapping_path)

                Logger.info(SCRIPT_NAME, "Мэппинг загружен. Параметров: {}".format(
                    len(self.mapping_parser.param_mapping)))
                Logger.debug(SCRIPT_NAME, "PropertySets: {}".format(
                    ", ".join(self.mapping_parser.property_sets.keys())))

                self.lbl_status.Text = "Мэппинг загружен. Параметров: {}".format(
                    len(self.mapping_parser.param_mapping))
                self.lbl_status.ForeColor = self._Color.Black

                self._update_check_button()

            except Exception as e:
                Logger.error(SCRIPT_NAME, "Ошибка парсинга мэппинга: {}".format(str(e)), exc_info=True)
                self.lbl_status.Text = "Ошибка парсинга мэппинга: {}".format(str(e))
                self.lbl_status.ForeColor = self._Color.Red
                self.mapping_parser = None

    def _update_check_button(self):
        """Обновить состояние кнопки проверки."""
        self.btn_check.Enabled = (self.ids_parser is not None and
                                   self.mapping_parser is not None)

    def on_check_view(self, sender, args):
        """Проверить все элементы на активном виде."""
        if not self.ids_parser or not self.mapping_parser:
            self._show_warning("Ошибка", "Сначала загрузите IDS и файл мэппинга")
            return

        active_view = self._doc.ActiveView
        if active_view is None:
            self._show_warning("Ошибка", "Нет активного вида")
            return

        Logger.log_separator(SCRIPT_NAME, "Начало проверки вида")
        Logger.info(SCRIPT_NAME, "Активный вид: {}".format(active_view.Name))

        self.lbl_status.Text = "Проверка..."
        self.lbl_status.ForeColor = self._Color.Black
        self.progress_bar.Value = 0
        Application.DoEvents()

        try:
            checker = ViewElementChecker(
                self.ids_parser,
                self.mapping_parser,
                self.category_mapper
            )

            def progress_callback(current, total):
                if total > 0:
                    percent = int(current * 100 / total)
                    self.progress_bar.Value = percent
                    Application.DoEvents()

            self.check_results = checker.check_view(active_view, progress_callback)

            # Подсчёт статистики
            total_elements = 0
            total_ok = 0
            total_fail = 0

            for cat_name, elements in self.check_results.items():
                for elem in elements:
                    total_elements += 1
                    if elem["failed"] == 0:
                        total_ok += 1
                    else:
                        total_fail += 1

            self.progress_bar.Value = 100

            # Логирование результатов
            Logger.info(SCRIPT_NAME, "Проверка завершена")
            Logger.info(SCRIPT_NAME, "  Категорий: {}".format(len(self.check_results)))
            Logger.info(SCRIPT_NAME, "  Элементов: {}".format(total_elements))
            Logger.info(SCRIPT_NAME, "  Прошли проверку: {}".format(total_ok))
            Logger.info(SCRIPT_NAME, "  Не прошли: {}".format(total_fail))

            # Логирование по категориям
            for cat_name, elements in sorted(self.check_results.items()):
                cat_ok = sum(1 for e in elements if e["failed"] == 0)
                cat_fail = len(elements) - cat_ok
                Logger.debug(SCRIPT_NAME, "    {}: {} элементов ({} OK, {} Failed)".format(
                    cat_name, len(elements), cat_ok, cat_fail))

            if total_elements == 0:
                Logger.warning(SCRIPT_NAME, "Нет элементов для проверки на этом виде")
                self.lbl_status.Text = "Нет элементов для проверки на этом виде"
                self.lbl_status.ForeColor = self._Color.Orange
            else:
                self.lbl_status.Text = "Проверено: {} элементов. OK: {}, Ошибок: {}".format(
                    total_elements, total_ok, total_fail)
                if total_fail > 0:
                    self.lbl_status.ForeColor = self._Color.Red
                else:
                    self.lbl_status.ForeColor = self._Color.Green

            self.btn_save.Enabled = total_elements > 0

            # Генерируем HTML
            generator = HTMLReportGenerator()
            self.html_content = generator.generate(self.check_results, active_view.Name)

        except Exception as e:
            Logger.error(SCRIPT_NAME, "Ошибка проверки: {}".format(str(e)), exc_info=True)
            self.lbl_status.Text = "Ошибка проверки: {}".format(str(e))
            self.lbl_status.ForeColor = self._Color.Red

    def on_save_report(self, sender, args):
        """Сохранить HTML отчёт."""
        if not self.html_content:
            self._show_warning("Ошибка", "Сначала выполните проверку")
            return

        # Выбор места сохранения
        save_path = self._forms.save_file(
            file_ext='html',
            default_name='ids_check_report.html',
            title="Сохранить отчёт"
        )

        if save_path:
            try:
                generator = HTMLReportGenerator()
                generator.save(save_path, self.html_content)

                Logger.file_saved(SCRIPT_NAME, save_path, "HTML отчёт проверки IDS")
                Logger.log_separator(SCRIPT_NAME, "Итоги")
                Logger.info(SCRIPT_NAME, "IDS файл: {}".format(self.ids_path))
                Logger.info(SCRIPT_NAME, "Файл мэппинга: {}".format(self.mapping_path))
                Logger.info(SCRIPT_NAME, "HTML отчёт: {}".format(save_path))

                self.lbl_status.Text = "Отчёт сохранён: {}".format(save_path)
                self.lbl_status.ForeColor = self._Color.Green

                # Открыть отчёт в браузере
                self._os.startfile(save_path)

            except Exception as e:
                Logger.error(SCRIPT_NAME, "Ошибка сохранения отчёта: {}".format(str(e)), exc_info=True)
                self.lbl_status.Text = "Ошибка сохранения: {}".format(str(e))
                self.lbl_status.ForeColor = self._Color.Red

    def on_close(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    try:
        form = IDSCheckerForm()
        form.ShowDialog()
        Logger.info(SCRIPT_NAME, "Скрипт завершён")
    except Exception as e:
        error_msg = str(e)
        Logger.critical(SCRIPT_NAME, "Критическая ошибка: {}".format(error_msg))
        if "NoneType" in error_msg and "Add" in error_msg:
            show_warning("Требуется перезапуск Revit", "Ошибка pyRevit!",
                         details="Это известная проблема с телеметрией pyRevit.\n\n"
                                 "Решение: Полностью перезапустите Revit\n"
                                 "(закройте и откройте заново).\n\n"
                                 "НЕ используйте кнопку 'Reload' в pyRevit!")
        else:
            show_error("Ошибка", "Ошибка выполнения скрипта",
                       details="{}\n\nЕсли ошибка повторяется после Reload pyRevit,\n"
                               "попробуйте полностью перезапустить Revit.".format(error_msg))
