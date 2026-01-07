# -*- coding: utf-8 -*-
"""
IDS Checker - проверка модели Revit на соответствие требованиям IDS.
Вкладка 1: Проверка параметров выбранного экземпляра
Вкладка 2: Полная проверка через IFC экспорт
"""

__title__ = "Проверка\nIDS"
__author__ = "CPSK"

import clr
import os
import sys
import json
import tempfile
import subprocess
import codecs

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System.Xml')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, Panel, ComboBox, ListBox, ListView,
    ListViewItem, ColumnHeader, View, CheckBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    DialogResult, ProgressBar, ProgressBarStyle, GroupBox, TabControl, TabPage
)
from System.Drawing import Point, Size, Color, Font, FontStyle
from System.Xml import XmlDocument, XmlNamespaceManager

from pyrevit import revit, forms, script

# Добавляем lib в путь для импорта cpsk_config
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Проверка окружения
from cpsk_config import require_environment, get_venv_python
if not require_environment():
    sys.exit()

# Revit API
from Autodesk.Revit.DB import Transaction, FilteredElementCollector, BuiltInParameter, ElementId, TemporaryViewMode
from System.Collections.Generic import List

# === НАСТРОЙКИ ===

doc = revit.doc
uidoc = revit.uidoc
app = doc.Application
output = script.get_output()

IDS_CHECKER_SCRIPT = os.path.join(LIB_DIR, "ids_checker.py")

# Используем Python из установленного окружения
VENV_PYTHON = get_venv_python()


# === ПАРСЕР IDS ===

class IDSParser:
    """Парсер IDS файла для извлечения типов и параметров (baseName!)."""

    # Атрибуты класса - сохраняем ссылки на уровне класса
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
            "predefinedTypes": [],  # predefinedType из applicability
            "requirements": []  # Параметры (property/baseName) из requirements
        }

        # Applicability - к каким элементам применяется (entity/name)
        # Может быть несколько entity в enumeration (IFCWALL, IFCWALLTYPE)
        app_nodes = spec_node.SelectNodes("ids:applicability/ids:entity/ids:name", nsm)
        for app_node in app_nodes:
            entities = self._get_all_node_values(app_node, nsm)
            for entity in entities:
                if entity:
                    # Фильтруем *TYPE классы
                    upper_entity = entity.upper()
                    if not upper_entity.endswith("TYPE"):
                        spec["applicability"].append(upper_entity)

        # PredefinedType из applicability
        pred_node = spec_node.SelectSingleNode("ids:applicability/ids:entity/ids:predefinedType", nsm)
        if pred_node:
            pred_values = self._get_all_enum_values(pred_node, nsm)
            if pred_values:
                # Фильтруем USERDEFINED и NOTDEFINED
                filtered = [v for v in pred_values if v not in ("USERDEFINED", "NOTDEFINED")]
                spec["predefinedTypes"] = filtered

        # Requirements - параметры (property с baseName!)
        req_nodes = spec_node.SelectNodes("ids:requirements/ids:property", nsm)
        for req_node in req_nodes:
            prop = self._parse_property(req_node, nsm)
            if prop:
                spec["requirements"].append(prop)

        return spec

    def _get_all_enum_values(self, node, nsm):
        """Получить все значения enumeration из узла."""
        values = []
        # Ищем xs:enumeration через все дочерние элементы
        for child in node.ChildNodes:
            if child.LocalName == "restriction":
                for enum_node in child.ChildNodes:
                    if enum_node.LocalName == "enumeration":
                        val = enum_node.GetAttribute("value")
                        if val:
                            values.append(val)
        return values if values else None

    def _get_node_value(self, node, nsm):
        """Получить значение узла (simpleValue или первое enumeration)."""
        values = self._get_all_node_values(node, nsm)
        return values[0] if values else None

    def _get_all_node_values(self, node, nsm):
        """Получить все значения узла (simpleValue или все enumeration)."""
        if node is None:
            return []

        values = []

        # Попробовать simpleValue
        simple = node.SelectSingleNode("ids:simpleValue", nsm)
        if simple:
            return [simple.InnerText]

        # Попробовать enumeration через ChildNodes (xs:restriction/xs:enumeration)
        for child in node.ChildNodes:
            if child.LocalName == "restriction":
                for enum_node in child.ChildNodes:
                    if enum_node.LocalName == "enumeration":
                        val = enum_node.GetAttribute("value")
                        if val:
                            values.append(val)

        if values:
            return values

        # Просто InnerText
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

        # PropertySet
        pset_node = prop_node.SelectSingleNode("ids:propertySet", nsm)
        if pset_node:
            prop["propertySet"] = self._get_node_value(pset_node, nsm) or ""

        # BaseName (имя параметра)
        name_node = prop_node.SelectSingleNode("ids:baseName", nsm)
        if name_node:
            prop["baseName"] = self._get_node_value(name_node, nsm) or ""

        # DataType
        datatype_attr = prop_node.GetAttribute("dataType")
        if datatype_attr:
            prop["dataType"] = datatype_attr.upper()

        # Cardinality (required/optional)
        cardinality = prop_node.GetAttribute("cardinality")
        if cardinality:
            prop["cardinality"] = cardinality

        # Enumeration (допустимые значения)
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
        """Собрать параметры и predefinedTypes по типам IFC."""
        for spec in self.specifications:
            for entity in spec["applicability"]:
                # Инициализация
                if entity not in self.entity_params:
                    self.entity_params[entity] = []
                if entity not in self.entity_predefined_types:
                    self.entity_predefined_types[entity] = []

                # Собрать predefinedTypes
                for ptype in spec.get("predefinedTypes", []):
                    if ptype not in self.entity_predefined_types[entity]:
                        self.entity_predefined_types[entity].append(ptype)

                # Собрать параметры (property/baseName)
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
        """Получить параметры (baseName) для указанного типа."""
        return self.entity_params.get(entity, [])

    def get_predefined_types_for_entity(self, entity):
        """Получить predefinedTypes для указанного типа."""
        return self.entity_predefined_types.get(entity, [])


# === ПРОВЕРКА ПАРАМЕТРОВ ЭЛЕМЕНТА ===

def check_element_params(element, required_params, prefix=""):
    """
    Проверить наличие параметров у элемента.
    Возвращает список (param_name, revit_name, has_param, value).
    """
    results = []

    for prop in required_params:
        base_name = prop["baseName"]

        # Формируем имя параметра в Revit с учётом префикса
        if prefix:
            revit_name = "{}_{}".format(prefix, base_name)
        else:
            revit_name = base_name

        # Пробуем найти параметр
        has_param = False
        value = ""

        # Сначала пробуем с префиксом
        param = element.LookupParameter(revit_name)
        if param:
            has_param = True
            value = get_param_value(param)
        else:
            # Пробуем без префикса
            param = element.LookupParameter(base_name)
            if param:
                has_param = True
                value = get_param_value(param)
                revit_name = base_name

        results.append({
            "ids_name": base_name,
            "revit_name": revit_name,
            "has_param": has_param,
            "value": value,
            "required": prop["cardinality"] == "required",
            "enumeration": prop.get("enumeration")
        })

    return results


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


# === IFC КОНФИГУРАЦИИ ===

def get_ifc_configurations(_doc=None):
    """Получить список конфигураций IFC экспорта."""
    import json as _json
    _doc = _doc or doc

    configs = [
        "IFC4 Reference View",
        "IFC4 Design Transfer View",
        "IFC 2x3 Coordination View 2.0",
        "IFC 2x3 Coordination View",
    ]

    try:
        from Autodesk.Revit.DB import FilteredElementCollector as FEC
        from Autodesk.Revit.DB.ExtensibleStorage import DataStorage, Schema

        storages = FEC(_doc).OfClass(DataStorage).ToElements()

        for storage in storages:
            guids = storage.GetEntitySchemaGuids()
            for guid in guids:
                schema = Schema.Lookup(guid)
                if schema and "IFC" in schema.SchemaName.upper():
                    entity = storage.GetEntity(schema)
                    if entity and entity.IsValid():
                        for field in schema.ListFields():
                            if "map" in field.FieldName.lower():
                                try:
                                    value = entity.Get[str](field)
                                    if value:
                                        data = _json.loads(value)
                                        if isinstance(data, dict) and "Name" in data:
                                            name = data["Name"]
                                            if name and name not in configs:
                                                configs.append(name)
                                except:
                                    pass
    except:
        pass

    return configs


def get_ifc_config_data(doc, config_name):
    """Получить полные данные конфигурации IFC экспорта."""
    import json as _json
    try:
        from Autodesk.Revit.DB import FilteredElementCollector as FEC
        from Autodesk.Revit.DB.ExtensibleStorage import DataStorage, Schema

        storages = FEC(doc).OfClass(DataStorage).ToElements()
        for storage in storages:
            guids = storage.GetEntitySchemaGuids()
            for guid in guids:
                schema = Schema.Lookup(guid)
                if schema and "IFC" in schema.SchemaName.upper():
                    entity = storage.GetEntity(schema)
                    if entity and entity.IsValid():
                        for field in schema.ListFields():
                            if "map" in field.FieldName.lower():
                                value = entity.Get[str](field)
                                if value:
                                    data = _json.loads(value)
                                    if isinstance(data, dict):
                                        if data.get("Name") == config_name:
                                            return data
    except:
        pass
    return None


def get_ifc_version_from_user_config(doc, config_name):
    """Получить версию IFC из пользовательской конфигурации."""
    try:
        from Autodesk.Revit.DB import IFCVersion

        version_map = {
            0: IFCVersion.IFC2x2,
            1: IFCVersion.IFC2x3,
            2: IFCVersion.IFC2x3CV2,
            3: IFCVersion.IFC4,
            20: IFCVersion.IFC4RV,
            21: IFCVersion.IFC4DTV,
        }

        data = get_ifc_config_data(doc, config_name)
        if data:
            ifc_ver = data.get("IFCVersion", 20)
            return version_map.get(ifc_ver, IFCVersion.IFC4RV)
    except:
        pass
    return None


def export_to_ifc_with_config(doc, output_path, config_name=None, _os=None, _LIB_DIR=None, _get_ifc_version=None):
    """Экспортировать модель Revit в IFC используя сохранённую конфигурацию."""
    _os = _os or os
    folder = _os.path.dirname(output_path)
    filename = _os.path.splitext(_os.path.basename(output_path))[0]

    # Лог для отладки
    log_lines = []

    try:
        from Autodesk.Revit.DB import IFCExportOptions, IFCVersion, Transaction as RevitTransaction

        # Попробуем загрузить IFC Export toolkit для полной поддержки настроек
        ifc_export_config = None
        try:
            # Ищем DLL в стандартных местах
            ifc_dll_paths = [
                r"C:\ProgramData\Autodesk\ApplicationPlugins\IFC 2025.bundle\Contents\2025\Revit.IFC.Export.dll",
                r"C:\ProgramData\Autodesk\ApplicationPlugins\IFC 2024.bundle\Contents\2024\Revit.IFC.Export.dll",
                r"C:\ProgramData\Autodesk\ApplicationPlugins\IFC 2023.bundle\Contents\2023\Revit.IFC.Export.dll",
            ]
            for dll_path in ifc_dll_paths:
                if _os.path.exists(dll_path):
                    clr.AddReferenceToFileAndPath(dll_path)
                    log_lines.append("Loaded IFC Export DLL: {}".format(dll_path))
                    break

            from Revit.IFC.Export.Utility import IFCExportConfiguration
            log_lines.append("IFC Export Configuration class loaded")

            # Загружаем конфигурацию по имени
            if config_name:
                ifc_export_config = IFCExportConfiguration.GetInSession(config_name)
                if ifc_export_config:
                    log_lines.append("Loaded IFC config by name: {}".format(config_name))
        except Exception as e:
            log_lines.append("IFC Export toolkit not available: {}".format(str(e)))

        options = IFCExportOptions()

        # Загружаем конфигурацию из сохранённых настроек
        config_data = None
        if config_name:
            config_data = get_ifc_config_data(doc, config_name)
            log_lines.append("Config '{}': {}".format(config_name, "found" if config_data else "not found"))

        if config_data:
            # Версия IFC - строим динамически чтобы избежать ошибок
            version_map = {
                0: IFCVersion.IFC2x2,
                1: IFCVersion.IFC2x3,
                2: IFCVersion.IFC2x3CV2,
                3: IFCVersion.IFC4,
                20: IFCVersion.IFC4RV,
                21: IFCVersion.IFC4DTV,
            }
            # Пробуем добавить новые версии если они доступны
            try:
                version_map[4] = IFCVersion.IFCBCA
            except:
                pass
            try:
                version_map[5] = IFCVersion.IFCCOBIE
            except:
                pass
            try:
                version_map[22] = IFCVersion.IFC2x3FM
            except:
                pass
            try:
                version_map[23] = IFCVersion.IFC2x3BFM
            except:
                pass
            try:
                version_map[24] = IFCVersion.IFC4x3
                version_map[25] = IFCVersion.IFC4x3
            except:
                # IFC4x3 недоступен - используем IFC4RV как fallback
                version_map[24] = IFCVersion.IFC4RV
                version_map[25] = IFCVersion.IFC4RV

            ifc_ver = config_data.get("IFCVersion", 20)
            if ifc_ver in version_map:
                options.FileVersion = version_map[ifc_ver]
                log_lines.append("IFC Version: {} -> {}".format(ifc_ver, version_map[ifc_ver]))
            else:
                options.FileVersion = IFCVersion.IFC4RV
                log_lines.append("IFC Version: {} NOT FOUND, using IFC4RV".format(ifc_ver))

            # Применяем только ключевые опции для маппинга
            if config_data.get("ExportUserDefinedPsets"):
                options.AddOption("ExportUserDefinedPsets", "true")
                log_lines.append("ExportUserDefinedPsets: true")

            psets_file = config_data.get("ExportUserDefinedPsetsFileName", "")
            if psets_file:
                # Проверяем существование файла маппинга
                if _os.path.exists(psets_file):
                    options.AddOption("ExportUserDefinedPsetsFileName", psets_file)
                    log_lines.append("Mapping file: {} (EXISTS)".format(psets_file))
                else:
                    log_lines.append("Mapping file: {} (NOT FOUND!)".format(psets_file))

            # Только базовые опции - UseActiveViewGeometry может вызывать ошибки
            if config_data.get("ExportIFCCommonPropertySets"):
                options.AddOption("ExportIFCCommonPropertySets", "true")
                log_lines.append("ExportIFCCommonPropertySets: true")

            log_lines.append("Output path: {}".format(output_path))
            log_lines.append("Folder: {}".format(folder))
            log_lines.append("Filename: {}".format(filename))
        else:
            # Fallback - базовые настройки
            options.FileVersion = IFCVersion.IFC4RV
            log_lines.append("Using fallback IFC4RV")

        # Сохраняем лог в папку скрипта
        log_path = r"C:\Users\feduloves\Documents\web\rhino_cpsk\pyrevit.extension\CPSK.tab\QA.panel\IDS.pulldown\06_IDSChecker.pushbutton\ifc_export_log.txt"
        import codecs as _codecs
        with _codecs.open(log_path, 'w', 'utf-8') as f:
            f.write("\n".join(log_lines))
            if config_data:
                f.write("\n\nFull config:\n")
                import json as _json
                f.write(_json.dumps(config_data, indent=2, ensure_ascii=False))

        log_lines.append("\n--- EXPORT START ---")

        # Экспорт IFC требует транзакцию
        trans = RevitTransaction(doc, "IFC Export for IDS Check")
        trans.Start()
        try:
            log_lines.append("Transaction started")
            result = doc.Export(folder, filename, options)
            log_lines.append("Export result: {}".format(result))
            trans.Commit()
            log_lines.append("Transaction committed")

            # Обновляем лог
            with _codecs.open(log_path, 'w', 'utf-8') as f:
                f.write("\n".join(log_lines))
                if config_data:
                    f.write("\n\nFull config:\n")
                    import json as _json
                    f.write(_json.dumps(config_data, indent=2, ensure_ascii=False))

            if result and _os.path.exists(output_path):
                return True, None
            else:
                return False, "Экспорт не создал файл. Лог: " + log_path
        except Exception as e:
            log_lines.append("EXPORT ERROR: {}".format(str(e)))
            # Обновляем лог с ошибкой
            with _codecs.open(log_path, 'w', 'utf-8') as f:
                f.write("\n".join(log_lines))
                if config_data:
                    f.write("\n\nFull config:\n")
                    import json as _json
                    f.write(_json.dumps(config_data, indent=2, ensure_ascii=False))
            trans.RollBack()
            return False, str(e) + ". Лог: " + log_path
    except Exception as e:
        return False, str(e)


# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===

def find_python(python_paths=None):
    """Найти CPython интерпретатор."""
    # Сначала проверяем установленное окружение
    if os.path.exists(VENV_PYTHON):
        return VENV_PYTHON
    return None


def run_ids_check(python_path, ids_path, ifc_path, report_path, _os=None, _tempfile=None, _ids_checker_script=None, standard_report_path=None):
    """Запустить проверку IDS через внешний скрипт."""
    import json as _json
    import codecs as _codecs
    import subprocess as _subprocess
    _os = _os or os
    _tempfile = _tempfile or tempfile
    _ids_checker_script = _ids_checker_script or IDS_CHECKER_SCRIPT
    try:
        params_file = _os.path.join(_tempfile.gettempdir(), "ids_check_params.json")
        result_file = _os.path.join(_tempfile.gettempdir(), "ids_check_result.json")

        params = {
            "ids_path": ids_path,
            "ifc_path": ifc_path,
            "report_path": report_path
        }

        # Добавляем путь для стандартного отчета если указан
        if standard_report_path:
            params["standard_report_path"] = standard_report_path

        with _codecs.open(params_file, 'w', 'utf-8') as f:
            _json.dump(params, f, ensure_ascii=False)

        cmd = [
            python_path,
            _ids_checker_script,
            "--json",
            params_file,
            result_file
        ]

        startupinfo = _subprocess.STARTUPINFO()
        startupinfo.dwFlags |= _subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0

        process = _subprocess.Popen(
            cmd,
            stdout=_subprocess.PIPE,
            stderr=_subprocess.PIPE,
            startupinfo=startupinfo
        )

        stdout, stderr = process.communicate()

        if _os.path.exists(result_file):
            with _codecs.open(result_file, 'r', 'utf-8') as f:
                return _json.load(f)

        output_text = stdout.decode('utf-8', errors='ignore').strip()
        if output_text:
            try:
                return _json.loads(output_text)
            except:
                return {"success": False, "errors": ["Ошибка парсинга: " + output_text]}

        return {"success": False, "errors": ["Пустой ответ: " + stderr.decode('utf-8', errors='ignore')]}

    except Exception as e:
        return {"success": False, "errors": ["Ошибка: " + str(e)]}


# === ГЛАВНОЕ ОКНО ===

class IDSCheckerForm(Form):
    """Диалог проверки IDS с вкладками."""

    # Атрибуты класса - сохраняем ссылки на уровне класса
    _ListViewItem = ListViewItem
    _FilteredElementCollector = FilteredElementCollector
    _List_ElementId = List[ElementId]
    _TemporaryViewMode = TemporaryViewMode
    _ProgressBarStyle = ProgressBarStyle
    _System = System
    _PYTHON_PATHS = PYTHON_PATHS
    _IDS_CHECKER_SCRIPT = IDS_CHECKER_SCRIPT
    _get_param_value = staticmethod(get_param_value)
    _find_python = staticmethod(find_python)
    _run_ids_check = staticmethod(run_ids_check)

    def __init__(self):
        self.ids_path = None
        self.ifc_path = None  # Путь к IFC файлу для проверки
        self.ids_parser = None
        self.report_path = None  # CPSK отчет
        self.standard_report_path = None  # Стандартный отчет ifctester
        self.result = None
        self.ifc_configs = []
        # Сохраняем ссылки на модули (для доступа внутри методов)
        self._forms = forms
        self._os = os
        self._tempfile = tempfile
        self._revit = revit
        self._LIB_DIR = LIB_DIR
        self._doc = doc
        self._uidoc = uidoc
        self._Color = Color
        self._MessageBox = MessageBox
        self._MessageBoxButtons = MessageBoxButtons
        self._MessageBoxIcon = MessageBoxIcon
        self._IDSParser = IDSParser
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "IDS Checker - Проверка модели"
        self.Width = 650
        self.Height = 660
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MaximizeBox = True
        self.MinimizeBox = True
        self.TopMost = True  # Окно поверх Revit

        # === IDS файл (общий для обеих вкладок) ===
        grp_ids = GroupBox()
        grp_ids.Text = "IDS файл"
        grp_ids.Location = Point(15, 10)
        grp_ids.Size = Size(605, 55)

        self.txt_ids = TextBox()
        self.txt_ids.Location = Point(15, 22)
        self.txt_ids.Width = 480
        self.txt_ids.ReadOnly = True
        grp_ids.Controls.Add(self.txt_ids)

        self.btn_browse = Button()
        self.btn_browse.Text = "Обзор..."
        self.btn_browse.Location = Point(505, 20)
        self.btn_browse.Width = 85
        self.btn_browse.Click += self.on_browse_click
        grp_ids.Controls.Add(self.btn_browse)

        self.Controls.Add(grp_ids)

        # === TabControl ===
        self.tabs = TabControl()
        self.tabs.Location = Point(15, 75)
        self.tabs.Size = Size(605, 495)

        # Вкладка 1: Проверка экземпляра
        self.tab_instance = TabPage()
        self.tab_instance.Text = "Проверка экземпляра"
        self.setup_tab_instance()
        self.tabs.TabPages.Add(self.tab_instance)

        # Вкладка 2: Полная проверка IFC
        self.tab_ifc = TabPage()
        self.tab_ifc.Text = "Проверка через IFC"
        self.setup_tab_ifc()
        self.tabs.TabPages.Add(self.tab_ifc)

        self.Controls.Add(self.tabs)

        # Кнопка Закрыть
        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(545, 580)
        btn_close.Size = Size(75, 28)
        btn_close.Click += self.on_close_click
        self.Controls.Add(btn_close)

    def setup_tab_instance(self):
        """Настройка вкладки проверки экземпляра."""
        y = 10

        # Строка 1: Тип IFC и PredefinedType
        lbl_entity = Label()
        lbl_entity.Text = "Тип IFC:"
        lbl_entity.Location = Point(15, y + 3)
        lbl_entity.AutoSize = True
        self.tab_instance.Controls.Add(lbl_entity)

        self.cmb_entity = ComboBox()
        self.cmb_entity.Location = Point(70, y)
        self.cmb_entity.Width = 150
        self.cmb_entity.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        self.cmb_entity.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        self.cmb_entity.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        self.cmb_entity.SelectedIndexChanged += self.on_entity_changed
        self.tab_instance.Controls.Add(self.cmb_entity)

        lbl_predefined = Label()
        lbl_predefined.Text = "Подтип:"
        lbl_predefined.Location = Point(230, y + 3)
        lbl_predefined.AutoSize = True
        self.tab_instance.Controls.Add(lbl_predefined)

        self.cmb_predefined = ComboBox()
        self.cmb_predefined.Location = Point(285, y)
        self.cmb_predefined.Width = 150
        self.cmb_predefined.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        self.cmb_predefined.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        self.cmb_predefined.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        self.tab_instance.Controls.Add(self.cmb_predefined)

        # Префикс параметров
        lbl_prefix = Label()
        lbl_prefix.Text = "Префикс:"
        lbl_prefix.Location = Point(445, y + 3)
        lbl_prefix.AutoSize = True
        self.tab_instance.Controls.Add(lbl_prefix)

        self.txt_prefix = TextBox()
        self.txt_prefix.Location = Point(500, y)
        self.txt_prefix.Width = 80
        self.txt_prefix.Text = ""
        self.txt_prefix.TextChanged += self.on_prefix_changed
        self.tab_instance.Controls.Add(self.txt_prefix)

        y += 28

        # Строка 2: Выбранный элемент и кнопки
        lbl_selected = Label()
        lbl_selected.Text = "Элемент:"
        lbl_selected.Location = Point(15, y + 3)
        lbl_selected.AutoSize = True
        self.tab_instance.Controls.Add(lbl_selected)

        self.lbl_element_info = Label()
        self.lbl_element_info.Text = "Выделите элемент в Revit"
        self.lbl_element_info.Location = Point(70, y + 3)
        self.lbl_element_info.Size = Size(350, 18)
        self.lbl_element_info.ForeColor = Color.Gray
        self.tab_instance.Controls.Add(self.lbl_element_info)

        self.btn_refresh = Button()
        self.btn_refresh.Text = "Обновить"
        self.btn_refresh.Location = Point(430, y)
        self.btn_refresh.Width = 70
        self.btn_refresh.Click += self.on_refresh_selection
        self.tab_instance.Controls.Add(self.btn_refresh)

        self.btn_check_instance = Button()
        self.btn_check_instance.Text = "Проверить"
        self.btn_check_instance.Location = Point(505, y)
        self.btn_check_instance.Width = 75
        self.btn_check_instance.Click += self.on_check_instance
        self.btn_check_instance.Enabled = False
        self.tab_instance.Controls.Add(self.btn_check_instance)

        y += 28

        # Строка 3: Заголовок списка и кнопки выбора
        lbl_params = Label()
        lbl_params.Text = "Параметры для проверки (baseName из IDS):"
        lbl_params.Location = Point(15, y)
        lbl_params.AutoSize = True
        self.tab_instance.Controls.Add(lbl_params)

        self.btn_select_all = Button()
        self.btn_select_all.Text = "Все"
        self.btn_select_all.Location = Point(455, y - 3)
        self.btn_select_all.Width = 55
        self.btn_select_all.Click += self.on_select_all
        self.tab_instance.Controls.Add(self.btn_select_all)

        self.btn_select_none = Button()
        self.btn_select_none.Text = "Снять"
        self.btn_select_none.Location = Point(515, y - 3)
        self.btn_select_none.Width = 60
        self.btn_select_none.Click += self.on_select_none
        self.tab_instance.Controls.Add(self.btn_select_none)

        y += 22

        # ListView для параметров с CheckBoxes
        self.list_params = ListView()
        self.list_params.Location = Point(15, y)
        self.list_params.Size = Size(565, 320)
        self.list_params.View = View.Details
        self.list_params.FullRowSelect = True
        self.list_params.GridLines = True
        self.list_params.CheckBoxes = True
        self.list_params.DoubleClick += self.on_list_double_click

        col1 = ColumnHeader()
        col1.Text = "baseName (IDS)"
        col1.Width = 120

        col2 = ColumnHeader()
        col2.Text = "Имя в Revit"
        col2.Width = 110

        col3 = ColumnHeader()
        col3.Text = "Обяз."
        col3.Width = 40

        col4 = ColumnHeader()
        col4.Text = "Есть"
        col4.Width = 35

        col5 = ColumnHeader()
        col5.Text = "Значение"
        col5.Width = 90

        col6 = ColumnHeader()
        col6.Text = "С парам."
        col6.Width = 60

        col7 = ColumnHeader()
        col7.Text = "Со знач."
        col7.Width = 60

        self.list_params.Columns.Add(col1)
        self.list_params.Columns.Add(col2)
        self.list_params.Columns.Add(col3)
        self.list_params.Columns.Add(col4)
        self.list_params.Columns.Add(col5)
        self.list_params.Columns.Add(col6)
        self.list_params.Columns.Add(col7)

        self.tab_instance.Controls.Add(self.list_params)

        y += 328

        # Кнопки фильтрации элементов на виде
        lbl_filter = Label()
        lbl_filter.Text = "Фильтр (2x клик на строку или кнопки):"
        lbl_filter.Location = Point(15, y + 3)
        lbl_filter.AutoSize = True
        self.tab_instance.Controls.Add(lbl_filter)

        self.btn_filter_has_param = Button()
        self.btn_filter_has_param.Text = "С парам."
        self.btn_filter_has_param.Location = Point(250, y)
        self.btn_filter_has_param.Width = 70
        self.btn_filter_has_param.Click += self.on_filter_has_param
        self.tab_instance.Controls.Add(self.btn_filter_has_param)

        self.btn_filter_by_value = Button()
        self.btn_filter_by_value.Text = "Со знач."
        self.btn_filter_by_value.Location = Point(325, y)
        self.btn_filter_by_value.Width = 70
        self.btn_filter_by_value.Click += self.on_filter_by_value
        self.tab_instance.Controls.Add(self.btn_filter_by_value)

        self.btn_filter_missing = Button()
        self.btn_filter_missing.Text = "Без парам."
        self.btn_filter_missing.Location = Point(400, y)
        self.btn_filter_missing.Width = 75
        self.btn_filter_missing.Click += self.on_filter_missing
        self.tab_instance.Controls.Add(self.btn_filter_missing)

        self.btn_show_all = Button()
        self.btn_show_all.Text = "Все"
        self.btn_show_all.Location = Point(480, y)
        self.btn_show_all.Width = 50
        self.btn_show_all.Click += self.on_show_all
        self.tab_instance.Controls.Add(self.btn_show_all)

        y += 28

        # Статус
        self.lbl_instance_status = Label()
        self.lbl_instance_status.Text = "Выберите IDS файл для начала"
        self.lbl_instance_status.Location = Point(15, y)
        self.lbl_instance_status.Size = Size(565, 20)
        self.tab_instance.Controls.Add(self.lbl_instance_status)

    def setup_tab_ifc(self):
        """Настройка вкладки полной проверки IFC."""
        y = 10

        # Информация об исправлениях
        grp_info = GroupBox()
        grp_info.Text = "CPSK: исправления ifctester (GitHub #4661)"
        grp_info.Location = Point(15, y)
        grp_info.Size = Size(560, 55)

        info_text = "Fix: Entity Restriction, *TYPE filter, USERDEFINED/NOTDEFINED filter, HTML reporter"
        lbl_info = Label()
        lbl_info.Text = info_text
        lbl_info.Location = Point(15, 18)
        lbl_info.Size = Size(420, 18)
        lbl_info.ForeColor = Color.FromArgb(80, 80, 80)
        grp_info.Controls.Add(lbl_info)

        # Кнопка-ссылка на GitHub issue
        self.btn_github = Button()
        self.btn_github.Text = "GitHub #4661"
        self.btn_github.Location = Point(445, 16)
        self.btn_github.Size = Size(100, 22)
        self.btn_github.Click += self.on_github_click
        grp_info.Controls.Add(self.btn_github)

        self.tab_ifc.Controls.Add(grp_info)

        y += 62

        # Шаг 1: Экспорт IFC
        grp_export = GroupBox()
        grp_export.Text = "Шаг 1: Экспорт IFC из Revit"
        grp_export.Location = Point(15, y)
        grp_export.Size = Size(560, 50)

        lbl_export_info = Label()
        lbl_export_info.Text = "Экспортируйте модель в IFC через стандартное окно Revit"
        lbl_export_info.Location = Point(15, 20)
        lbl_export_info.Size = Size(350, 20)
        grp_export.Controls.Add(lbl_export_info)

        self.btn_open_export = Button()
        self.btn_open_export.Text = "Экспорт IFC..."
        self.btn_open_export.Location = Point(450, 18)
        self.btn_open_export.Size = Size(95, 25)
        self.btn_open_export.Click += self.on_open_ifc_export_click
        grp_export.Controls.Add(self.btn_open_export)

        self.tab_ifc.Controls.Add(grp_export)

        y += 58

        # Шаг 2: Выбор IFC файла
        grp_ifc = GroupBox()
        grp_ifc.Text = "Шаг 2: Выберите IFC файл для проверки"
        grp_ifc.Location = Point(15, y)
        grp_ifc.Size = Size(560, 55)

        self.txt_ifc = TextBox()
        self.txt_ifc.Location = Point(15, 22)
        self.txt_ifc.Width = 440
        self.txt_ifc.ReadOnly = True
        grp_ifc.Controls.Add(self.txt_ifc)

        self.btn_browse_ifc = Button()
        self.btn_browse_ifc.Text = "Обзор..."
        self.btn_browse_ifc.Location = Point(465, 20)
        self.btn_browse_ifc.Width = 80
        self.btn_browse_ifc.Click += self.on_browse_ifc_click
        grp_ifc.Controls.Add(self.btn_browse_ifc)

        self.tab_ifc.Controls.Add(grp_ifc)

        y += 70

        # Статус и результаты
        grp_status = GroupBox()
        grp_status.Text = "Шаг 3: Результаты проверки"
        grp_status.Location = Point(15, y)
        grp_status.Size = Size(560, 130)

        self.lbl_ifc_status = Label()
        self.lbl_ifc_status.Text = "Выберите IDS и IFC файлы для начала проверки"
        self.lbl_ifc_status.Location = Point(15, 25)
        self.lbl_ifc_status.Size = Size(530, 20)
        grp_status.Controls.Add(self.lbl_ifc_status)

        self.progress = ProgressBar()
        self.progress.Location = Point(15, 50)
        self.progress.Size = Size(530, 18)
        self.progress.Visible = False
        grp_status.Controls.Add(self.progress)

        self.lbl_result = Label()
        self.lbl_result.Text = ""
        self.lbl_result.Location = Point(15, 75)
        self.lbl_result.Size = Size(530, 45)
        grp_status.Controls.Add(self.lbl_result)

        self.tab_ifc.Controls.Add(grp_status)

        y += 145

        # Чекбокс для стандартного отчета
        self.chk_standard_report = CheckBox()
        self.chk_standard_report.Text = "Также создать стандартный отчет ifctester (для сравнения)"
        self.chk_standard_report.Location = Point(15, y)
        self.chk_standard_report.Size = Size(400, 20)
        self.chk_standard_report.Checked = False
        self.tab_ifc.Controls.Add(self.chk_standard_report)

        y += 28

        # Кнопки
        self.btn_check_ifc = Button()
        self.btn_check_ifc.Text = "Запустить проверку"
        self.btn_check_ifc.Location = Point(15, y)
        self.btn_check_ifc.Size = Size(140, 32)
        self.btn_check_ifc.Click += self.on_check_ifc_click
        self.btn_check_ifc.Enabled = False
        self.tab_ifc.Controls.Add(self.btn_check_ifc)

        self.btn_open = Button()
        self.btn_open.Text = "CPSK отчет"
        self.btn_open.Location = Point(165, y)
        self.btn_open.Size = Size(100, 32)
        self.btn_open.Click += self.on_open_click
        self.btn_open.Enabled = False
        self.tab_ifc.Controls.Add(self.btn_open)

        self.btn_open_std = Button()
        self.btn_open_std.Text = "Станд. отчет"
        self.btn_open_std.Location = Point(275, y)
        self.btn_open_std.Size = Size(100, 32)
        self.btn_open_std.Click += self.on_open_std_click
        self.btn_open_std.Enabled = False
        self.tab_ifc.Controls.Add(self.btn_open_std)

        self.btn_save = Button()
        self.btn_save.Text = "Сохранить"
        self.btn_save.Location = Point(385, y)
        self.btn_save.Size = Size(90, 32)
        self.btn_save.Click += self.on_save_click
        self.btn_save.Enabled = False
        self.tab_ifc.Controls.Add(self.btn_save)

    def on_browse_click(self, sender, args):
        """Выбор IDS файла."""
        # Используем pyrevit.forms - надёжнее чем WinForms в IronPython
        selected = self._forms.pick_file(
            file_ext='ids',
            init_dir=self._LIB_DIR if self._os.path.exists(self._LIB_DIR) else None,
            title="Выберите IDS файл"
        )

        if selected:
            self.ids_path = selected
            self.txt_ids.Text = self.ids_path

            # Парсим IDS
            try:
                self.ids_parser = self._IDSParser(self.ids_path)
                self.ids_parser.parse()

                # Заполняем список типов
                self.cmb_entity.Items.Clear()
                entities = self.ids_parser.get_entities()
                for ent in entities:
                    self.cmb_entity.Items.Add(ent)

                if self.cmb_entity.Items.Count > 0:
                    self.cmb_entity.SelectedIndex = 0

                # Для вкладки экземпляра
                self.btn_check_instance.Enabled = True
                self.lbl_instance_status.Text = "IDS загружен. Типов: {}, Выберите тип и элемент".format(len(entities))
                self.lbl_instance_status.ForeColor = self._Color.Black

                # Для вкладки IFC - проверяем наличие обоих файлов
                self._update_ifc_check_state()

            except Exception as e:
                self.lbl_instance_status.Text = "Ошибка парсинга IDS: {}".format(str(e))
                self.lbl_instance_status.ForeColor = self._Color.Red

    def _update_ifc_check_state(self):
        """Обновить состояние кнопки проверки IFC."""
        if self.ids_path and self.ifc_path:
            self.btn_check_ifc.Enabled = True
            self.lbl_ifc_status.Text = "Готово к проверке. Нажмите 'Запустить проверку'"
            self.lbl_ifc_status.ForeColor = self._Color.Black
        elif self.ids_path:
            self.btn_check_ifc.Enabled = False
            self.lbl_ifc_status.Text = "IDS загружен. Выберите IFC файл"
            self.lbl_ifc_status.ForeColor = self._Color.Black
        elif self.ifc_path:
            self.btn_check_ifc.Enabled = False
            self.lbl_ifc_status.Text = "IFC выбран. Выберите IDS файл"
            self.lbl_ifc_status.ForeColor = self._Color.Black
        else:
            self.btn_check_ifc.Enabled = False
            self.lbl_ifc_status.Text = "Выберите IDS и IFC файлы для начала проверки"
            self.lbl_ifc_status.ForeColor = self._Color.Black

    def on_github_click(self, sender, args):
        """Открыть GitHub issue в браузере."""
        import webbrowser
        webbrowser.open("https://github.com/IfcOpenShell/IfcOpenShell/issues/4661")

    def on_open_ifc_export_click(self, sender, args):
        """Открыть стандартное окно экспорта IFC в Revit."""
        try:
            from Autodesk.Revit.UI import PostableCommand, RevitCommandId
            cmd_id = RevitCommandId.LookupPostableCommandId(PostableCommand.ExportIFC)
            # Закрываем модальное окно перед вызовом команды
            # иначе PostCommand выполнится только после закрытия диалога
            self.Close()
            self._uidoc.Application.PostCommand(cmd_id)
        except Exception as e:
            self._MessageBox.Show(
                "Не удалось открыть окно экспорта IFC:\n{}".format(str(e)),
                "Ошибка",
                self._MessageBoxButtons.OK,
                self._MessageBoxIcon.Error
            )

    def on_browse_ifc_click(self, sender, args):
        """Выбор IFC файла для проверки."""
        selected = self._forms.pick_file(
            file_ext='ifc',
            title="Выберите IFC файл для проверки"
        )

        if selected:
            self.ifc_path = selected
            self.txt_ifc.Text = self.ifc_path
            self._update_ifc_check_state()

    def on_entity_changed(self, sender, args):
        """Изменение выбранного типа IFC."""
        # Заполнить dropdown predefinedTypes
        self.cmb_predefined.Items.Clear()
        self.cmb_predefined.Items.Add("(Все)")  # Опция для всех подтипов

        if self.ids_parser and self.cmb_entity.SelectedIndex >= 0:
            entity = str(self.cmb_entity.SelectedItem)
            pred_types = self.ids_parser.get_predefined_types_for_entity(entity)
            for pt in pred_types:
                self.cmb_predefined.Items.Add(pt)

        if self.cmb_predefined.Items.Count > 0:
            self.cmb_predefined.SelectedIndex = 0

        self.update_params_list()

    def on_prefix_changed(self, sender, args):
        """Изменение префикса - обновить имена параметров Revit."""
        self.update_params_list()

    def on_select_all(self, sender, args):
        """Выбрать все атрибуты."""
        for item in self.list_params.Items:
            item.Checked = True

    def on_select_none(self, sender, args):
        """Снять выбор со всех атрибутов."""
        for item in self.list_params.Items:
            item.Checked = False

    def on_list_double_click(self, sender, args):
        """Двойной клик по строке - показать элементы со значением."""
        if self.list_params.SelectedItems.Count == 0:
            return

        selected_item = self.list_params.SelectedItems[0]
        param_name = selected_item.SubItems[1].Text  # Имя в Revit
        param_value = selected_item.SubItems[4].Text  # Значение

        if param_value and param_value != "-":
            self._isolate_elements_by_param(param_name, param_value, filter_mode="has_value")
        else:
            # Если нет значения - показать элементы с параметром
            self._isolate_elements_by_param(param_name, None, filter_mode="has_param")

    def on_filter_has_param(self, sender, args):
        """Показать элементы, у которых ЕСТЬ выбранный параметр."""
        if self.list_params.SelectedItems.Count == 0:
            self._MessageBox.Show("Выберите параметр в списке", "Информация",
                            self._MessageBoxButtons.OK, self._MessageBoxIcon.Information)
            return

        selected_item = self.list_params.SelectedItems[0]
        param_name = selected_item.SubItems[1].Text  # Имя в Revit

        self._isolate_elements_by_param(param_name, None, filter_mode="has_param")

    def on_filter_by_value(self, sender, args):
        """Показать элементы с определённым значением выбранного параметра."""
        if self.list_params.SelectedItems.Count == 0:
            self._MessageBox.Show("Выберите параметр в списке", "Информация",
                            self._MessageBoxButtons.OK, self._MessageBoxIcon.Information)
            return

        selected_item = self.list_params.SelectedItems[0]
        param_name = selected_item.SubItems[1].Text  # Имя в Revit
        param_value = selected_item.SubItems[4].Text  # Значение

        if param_value == "-" or not param_value:
            self._MessageBox.Show("Сначала проверьте элемент чтобы получить значение", "Информация",
                            self._MessageBoxButtons.OK, self._MessageBoxIcon.Information)
            return

        self._isolate_elements_by_param(param_name, param_value, filter_mode="has_value")

    def on_filter_missing(self, sender, args):
        """Показать элементы БЕЗ выбранного параметра."""
        if self.list_params.SelectedItems.Count == 0:
            self._MessageBox.Show("Выберите параметр в списке", "Информация",
                            self._MessageBoxButtons.OK, self._MessageBoxIcon.Information)
            return

        selected_item = self.list_params.SelectedItems[0]
        param_name = selected_item.SubItems[1].Text  # Имя в Revit

        self._isolate_elements_by_param(param_name, None, filter_mode="missing")

    def on_show_all(self, sender, args):
        """Сбросить изоляцию - показать все элементы."""
        try:
            active_view = self._doc.ActiveView
            # Временная изоляция не требует транзакции
            active_view.DisableTemporaryViewMode(self._TemporaryViewMode.TemporaryHideIsolate)
            self.lbl_instance_status.Text = "Изоляция сброшена"
            self.lbl_instance_status.ForeColor = self._Color.Black
        except Exception as e:
            self.lbl_instance_status.Text = "Ошибка: {}".format(str(e))
            self.lbl_instance_status.ForeColor = self._Color.Red

    def _isolate_elements_by_param(self, param_name, param_value, filter_mode="has_value"):
        """Изолировать элементы на виде по значению параметра."""
        try:
            active_view = self._doc.ActiveView

            # Собрать все элементы на виде
            collector = self._FilteredElementCollector(self._doc, active_view.Id)
            all_elements = collector.WhereElementIsNotElementType().ToElements()

            matching_ids = self._List_ElementId()
            checked_count = 0
            match_count = 0

            for elem in all_elements:
                if elem.Category is None:
                    continue

                param = elem.LookupParameter(param_name)

                if filter_mode == "missing":
                    # Элементы БЕЗ параметра
                    if param is None:
                        matching_ids.Add(elem.Id)
                        match_count += 1
                elif filter_mode == "has_param":
                    # Элементы с параметром (даже пустым)
                    if param is not None:
                        matching_ids.Add(elem.Id)
                        match_count += 1
                elif filter_mode == "has_value":
                    # Элементы с конкретным значением
                    if param and param.HasValue:
                        val = self._get_param_value(param)
                        if val == param_value:
                            matching_ids.Add(elem.Id)
                            match_count += 1

                checked_count += 1

            if matching_ids.Count == 0:
                self.lbl_instance_status.Text = "Не найдено элементов"
                self.lbl_instance_status.ForeColor = self._Color.Orange
                return

            # Изолировать элементы (временная изоляция не требует транзакции)
            active_view.IsolateElementsTemporary(matching_ids)

            if filter_mode == "missing":
                self.lbl_instance_status.Text = "Изолировано {} элементов без параметра '{}'".format(
                    match_count, param_name)
            elif filter_mode == "has_param":
                self.lbl_instance_status.Text = "Изолировано {} элементов с параметром '{}'".format(
                    match_count, param_name)
            else:
                self.lbl_instance_status.Text = "Изолировано {} элементов с {}='{}'".format(
                    match_count, param_name, param_value)
            self.lbl_instance_status.ForeColor = self._Color.Blue

        except Exception as e:
            self.lbl_instance_status.Text = "Ошибка: {}".format(str(e))
            self.lbl_instance_status.ForeColor = self._Color.Red

    def update_params_list(self):
        """Обновить список параметров для выбранного типа."""
        self.list_params.Items.Clear()

        if not self.ids_parser or self.cmb_entity.SelectedIndex < 0:
            return

        entity = str(self.cmb_entity.SelectedItem)
        params = self.ids_parser.get_params_for_entity(entity)
        prefix = self.txt_prefix.Text.strip()

        # Добавляем параметры (property/baseName)
        for prop in params:
            base_name = prop["baseName"]
            if prefix:
                revit_name = "{}_{}".format(prefix, base_name)
            else:
                revit_name = base_name

            is_required = prop.get("cardinality", "required") == "required"

            # 7 колонок: baseName, Имя в Revit, Обяз., Есть, Значение, С парам., Со знач.
            item = self._ListViewItem(base_name)  # Col 0: baseName (IDS)
            item.SubItems.Add(revit_name)   # Col 1: Имя в Revit
            item.SubItems.Add("Да" if is_required else "Нет")  # Col 2: Обяз.
            item.SubItems.Add("-")          # Col 3: Есть
            item.SubItems.Add("-")          # Col 4: Значение
            item.SubItems.Add("-")          # Col 5: С парам. (кол-во)
            item.SubItems.Add("-")          # Col 6: Со знач. (кол-во)
            item.Checked = True
            item.Tag = prop  # Сохраняем данные параметра напрямую

            self.list_params.Items.Add(item)

    def on_refresh_selection(self, sender, args):
        """Обновить информацию о выделенном элементе."""
        sel = self._uidoc.Selection.GetElementIds()
        if sel.Count == 0:
            self.lbl_element_info.Text = "Нет выделенных элементов"
            self.lbl_element_info.ForeColor = self._Color.Gray
            return

        if sel.Count > 1:
            self.lbl_element_info.Text = "Выделено {} элементов. Выберите один".format(sel.Count)
            self.lbl_element_info.ForeColor = self._Color.Orange
            return

        elem_id = list(sel)[0]
        elem = self._doc.GetElement(elem_id)

        if elem:
            cat_name = elem.Category.Name if elem.Category else "Без категории"
            elem_name = elem.Name if hasattr(elem, 'Name') else "Без имени"
            self.lbl_element_info.Text = "{}: {} (ID: {})".format(cat_name, elem_name, elem_id.IntegerValue)
            self.lbl_element_info.ForeColor = self._Color.Black
        else:
            self.lbl_element_info.Text = "Элемент не найден"
            self.lbl_element_info.ForeColor = self._Color.Red

    def on_check_instance(self, sender, args):
        """Проверить параметры выбранного экземпляра и подсчитать элементы."""
        if not self.ids_parser or self.cmb_entity.SelectedIndex < 0:
            self._MessageBox.Show("Сначала выберите IDS файл и тип", "Ошибка",
                            self._MessageBoxButtons.OK, self._MessageBoxIcon.Warning)
            return

        # Проверить что есть выбранные элементы
        checked_count = 0
        for item in self.list_params.Items:
            if item.Checked:
                checked_count += 1

        if checked_count == 0:
            self._MessageBox.Show("Отметьте хотя бы один параметр для проверки", "Ошибка",
                            self._MessageBoxButtons.OK, self._MessageBoxIcon.Warning)
            return

        sel = self._uidoc.Selection.GetElementIds()
        if sel.Count != 1:
            self._MessageBox.Show("Выделите один элемент в Revit", "Ошибка",
                            self._MessageBoxButtons.OK, self._MessageBoxIcon.Warning)
            return

        elem_id = list(sel)[0]
        elem = self._doc.GetElement(elem_id)

        if not elem:
            self._MessageBox.Show("Элемент не найден", "Ошибка",
                            self._MessageBoxButtons.OK, self._MessageBoxIcon.Error)
            return

        prefix = self.txt_prefix.Text.strip()

        # Собрать все элементы на виде для подсчёта
        active_view = self._doc.ActiveView
        collector = self._FilteredElementCollector(self._doc, active_view.Id)
        all_elements = list(collector.WhereElementIsNotElementType().ToElements())

        passed = 0
        failed = 0
        checked_total = 0

        # Проверяем только отмеченные элементы
        # Колонки: 0=baseName, 1=Имя в Revit, 2=Обяз., 3=Есть, 4=Значение, 5=С парам., 6=Со знач.
        for item in self.list_params.Items:
            if not item.Checked:
                # Сбросить результаты для невыбранных
                item.SubItems[3].Text = "-"
                item.SubItems[4].Text = "-"
                item.SubItems[5].Text = "-"
                item.SubItems[6].Text = "-"
                item.ForeColor = self._Color.Gray
                continue

            checked_total += 1
            prop = item.Tag
            if not prop:
                continue

            # Получаем данные параметра
            base_name = prop.get("baseName", "")
            is_required = prop.get("cardinality", "required") == "required"

            if prefix:
                revit_name = "{}_{}".format(prefix, base_name)
            else:
                revit_name = base_name

            has_param = False
            value = ""

            # Пробуем с префиксом
            param = elem.LookupParameter(revit_name)
            if param:
                has_param = True
                value = self._get_param_value(param)
            else:
                # Пробуем без префикса
                param = elem.LookupParameter(base_name)
                if param:
                    has_param = True
                    value = self._get_param_value(param)
                    revit_name = base_name

            # Обновляем имя в Revit (колонка 1)
            item.SubItems[1].Text = revit_name

            # Обновляем результаты
            if has_param:
                item.SubItems[3].Text = "Да"
                item.ForeColor = self._Color.Black
                passed += 1
            else:
                item.SubItems[3].Text = "Нет"
                if is_required:
                    failed += 1
                    item.ForeColor = self._Color.Red
                else:
                    item.ForeColor = self._Color.Orange

            item.SubItems[4].Text = value if value else "-"

            # Подсчитать элементы с параметром и со значением
            count_with_param = 0
            count_with_value = 0

            for el in all_elements:
                if el.Category is None:
                    continue
                p = el.LookupParameter(revit_name)
                if p is not None:
                    count_with_param += 1
                    if p.HasValue and value:
                        el_value = self._get_param_value(p)
                        if el_value == value:
                            count_with_value += 1

            item.SubItems[5].Text = str(count_with_param)
            item.SubItems[6].Text = str(count_with_value) if value else "-"

        # Статус
        if failed == 0:
            self.lbl_instance_status.Text = "OK: Найдено {}/{} параметров. Клик на колонки для фильтрации".format(passed, checked_total)
            self.lbl_instance_status.ForeColor = self._Color.Green
        else:
            self.lbl_instance_status.Text = "Найдено {}/{}, отсутствует {} обязательных".format(passed, checked_total, failed)
            self.lbl_instance_status.ForeColor = self._Color.Red

    def on_check_ifc_click(self, sender, args):
        """Запуск полной проверки через IFC."""
        if not self.ids_path or not self.ifc_path:
            return

        # Найти Python (inline чтобы избежать проблем с областью видимости)
        python_path = None
        for p in self._PYTHON_PATHS:
            if self._os.path.exists(p):
                python_path = p
                break

        if not python_path:
            self._MessageBox.Show(
                "Python не найден!\n\nУстановите Python 3.9+ и библиотеки:\npip install ifcopenshell ifctester",
                "Ошибка",
                self._MessageBoxButtons.OK,
                self._MessageBoxIcon.Error
            )
            return

        # Проверяем существование IFC файла
        if not self._os.path.exists(self.ifc_path):
            self._MessageBox.Show(
                "IFC файл не найден:\n{}".format(self.ifc_path),
                "Ошибка",
                self._MessageBoxButtons.OK,
                self._MessageBoxIcon.Error
            )
            return

        self.progress.Visible = True
        self.progress.Style = self._ProgressBarStyle.Marquee
        self.btn_check_ifc.Enabled = False
        self.lbl_ifc_status.Text = "Проверка IDS требований..."
        self.lbl_ifc_status.ForeColor = self._Color.Black
        self._System.Windows.Forms.Application.DoEvents()

        # Путь для отчёта - рядом с IFC файлом
        ifc_dir = self._os.path.dirname(self.ifc_path)
        ifc_name = self._os.path.splitext(self._os.path.basename(self.ifc_path))[0]
        self.report_path = self._os.path.join(ifc_dir, "{}_ids_report_CPSK.html".format(ifc_name))

        # Стандартный отчет ifctester если чекбокс включен
        self.standard_report_path = None
        if self.chk_standard_report.Checked:
            self.standard_report_path = self._os.path.join(ifc_dir, "{}_ids_report_STANDARD.html".format(ifc_name))

        self.result = self._run_ids_check(
            python_path, self.ids_path, self.ifc_path, self.report_path,
            self._os, self._tempfile, self._IDS_CHECKER_SCRIPT,
            standard_report_path=self.standard_report_path
        )

        self.progress.Visible = False
        self.btn_check_ifc.Enabled = True

        if self.result.get("success"):
            passed = self.result.get("passed_specs", 0)
            failed = self.result.get("failed_specs", 0)
            total = self.result.get("total_specs", 0)
            passed_reqs = self.result.get("passed_requirements", 0)
            failed_reqs = self.result.get("failed_requirements", 0)

            if failed == 0 and failed_reqs == 0:
                self.lbl_ifc_status.Text = "PASSED - Все проверки пройдены!"
                self.lbl_ifc_status.ForeColor = self._Color.Green
            else:
                self.lbl_ifc_status.Text = "FAILED - Есть несоответствия"
                self.lbl_ifc_status.ForeColor = self._Color.Red

            self.lbl_result.Text = (
                "Спецификации: {} всего, {} пройдено, {} не пройдено\n"
                "Требования: {} пройдено, {} не пройдено"
            ).format(total, passed, failed, passed_reqs, failed_reqs)

            self.btn_open.Enabled = True
            self.btn_save.Enabled = True
            # Включаем кнопку стандартного отчета если он был создан
            if self.standard_report_path and self._os.path.exists(self.standard_report_path):
                self.btn_open_std.Enabled = True
            else:
                self.btn_open_std.Enabled = False
        else:
            self.show_ifc_error("\n".join(self.result.get("errors", ["Неизвестная ошибка"])))

    def show_ifc_error(self, message):
        """Показать ошибку на вкладке IFC."""
        self.progress.Visible = False
        self.btn_check_ifc.Enabled = True
        self.lbl_ifc_status.Text = "ОШИБКА"
        self.lbl_ifc_status.ForeColor = self._Color.Red
        self.lbl_result.Text = message

    def on_open_click(self, sender, args):
        """Открыть CPSK HTML отчет."""
        if self.report_path and self._os.path.exists(self.report_path):
            self._os.startfile(self.report_path)

    def on_open_std_click(self, sender, args):
        """Открыть стандартный HTML отчет ifctester."""
        if self.standard_report_path and self._os.path.exists(self.standard_report_path):
            self._os.startfile(self.standard_report_path)
        else:
            self._MessageBox.Show(
                "Стандартный отчет не найден.\n\nВключите опцию 'Также создать стандартный отчет' и запустите проверку снова.",
                "Информация",
                self._MessageBoxButtons.OK,
                self._MessageBoxIcon.Information
            )

    def on_save_click(self, sender, args):
        """Сохранить HTML отчет."""
        if not self.report_path or not self._os.path.exists(self.report_path):
            return

        # Используем pyrevit.forms - надёжнее чем WinForms в IronPython
        save_path = self._forms.save_file(
            file_ext='html',
            default_name='ids_report.html',
            title="Сохранить отчет"
        )
        if save_path:
            import shutil
            shutil.copy(self.report_path, save_path)
            self._forms.alert("Отчет сохранен:\n" + save_path, title="Готово")

    def on_close_click(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    try:
        form = IDSCheckerForm()
        form.ShowDialog()  # Модальное окно - требуется для транзакций (IFC экспорт)
    except Exception as e:
        error_msg = str(e)
        if "NoneType" in error_msg and "Add" in error_msg:
            MessageBox.Show(
                "Ошибка pyRevit!\n\n"
                "Это известная проблема с телеметрией pyRevit.\n\n"
                "Решение: Полностью перезапустите Revit\n"
                "(закройте и откройте заново).\n\n"
                "НЕ используйте кнопку 'Reload' в pyRevit!",
                "Требуется перезапуск Revit",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
        else:
            MessageBox.Show(
                "Ошибка: {}\n\n"
                "Если ошибка повторяется после Reload pyRevit,\n"
                "попробуйте полностью перезапустить Revit.".format(error_msg),
                "Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )
