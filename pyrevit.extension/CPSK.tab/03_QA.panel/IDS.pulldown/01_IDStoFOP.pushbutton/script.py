# -*- coding: utf-8 -*-
"""
IDS to FOP - Создание ФОП и IFC Mapping из IDS файла.

Парсит IDS файл и генерирует:
- ФОП файл (Shared Parameters) для Revit
- IFC Parameter Mapping файл для экспорта
"""

__title__ = "IDS в\nФОП"
__author__ = "CPSK"

import clr
import os
import sys
import codecs
import uuid

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System.Xml')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, Panel, CheckBox, ComboBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    DialogResult, OpenFileDialog, SaveFileDialog, FolderBrowserDialog,
    GroupBox, RadioButton
)
from System.Drawing import Point, Size, Color, Font, FontStyle
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

# Проверка авторизации
from cpsk_auth import require_auth
if not require_auth():
    sys.exit()

# Проверка окружения
from cpsk_config import require_environment
if not require_environment():
    sys.exit()

# Логгер
from cpsk_logger import Logger

# Инициализация логгера
SCRIPT_NAME = "IDStoFOP"
Logger.init(SCRIPT_NAME)
Logger.info(SCRIPT_NAME, "Скрипт запущен")

# Импорт IFC маппингов из support_files
from ifc_mappings import IFC_TO_REVIT_TYPE


# === ПАРСЕР IDS ===

class IDSParser:
    """Парсер IDS файла."""

    def __init__(self, ids_path):
        self.ids_path = ids_path
        self.specifications = []
        self.all_properties = []       # Уникальные параметры по ИМЕНИ (для ФОП)
        self.all_properties_full = []  # ВСЕ параметры с учётом PropertySet (для Mapping/Report)
        self.property_sets = {}        # PropertySet -> [properties]
        self.property_sets_full = {}   # PropertySet -> [properties] (полный, для Mapping/Report)

    def parse(self):
        """Парсить IDS файл."""
        doc = XmlDocument()
        doc.Load(self.ids_path)

        # Namespace manager для IDS
        nsm = XmlNamespaceManager(doc.NameTable)
        nsm.AddNamespace("ids", "http://standards.buildingsmart.org/IDS")
        nsm.AddNamespace("xs", "http://www.w3.org/2001/XMLSchema")

        # Найти все specifications
        spec_nodes = doc.SelectNodes("//ids:specification", nsm)

        for spec_node in spec_nodes:
            spec = self._parse_specification(spec_node, nsm)
            self.specifications.append(spec)

        # Собрать уникальные параметры
        self._collect_unique_properties()

        return self

    def _parse_specification(self, spec_node, nsm):
        """Парсить одну спецификацию."""
        spec = {
            "name": spec_node.GetAttribute("name") or "",
            "applicability": [],
            "requirements": []
        }

        # Applicability - к каким элементам применяется
        entity_nodes = spec_node.SelectNodes("ids:applicability/ids:entity", nsm)
        for entity_node in entity_nodes:
            # Сначала пробуем simpleValue
            name_simple = entity_node.SelectSingleNode("ids:name/ids:simpleValue", nsm)
            if name_simple and name_simple.InnerText:
                entity = name_simple.InnerText.strip().upper()
                if entity and entity not in spec["applicability"]:
                    spec["applicability"].append(entity)
            else:
                # Пробуем xs:restriction/xs:enumeration
                enum_nodes = entity_node.SelectNodes("ids:name/xs:restriction/xs:enumeration", nsm)
                for enum_node in enum_nodes:
                    val = enum_node.GetAttribute("value")
                    if val:
                        entity = val.strip().upper()
                        # Пропускаем *TYPE - это типы семейств, не экземпляры
                        if not entity.endswith("TYPE") and entity not in spec["applicability"]:
                            spec["applicability"].append(entity)

        # Requirements - требуемые параметры
        req_nodes = spec_node.SelectNodes("ids:requirements/ids:property", nsm)
        for req_node in req_nodes:
            prop = self._parse_property(req_node, nsm)
            if prop:
                spec["requirements"].append(prop)

        return spec

    def _parse_property(self, prop_node, nsm):
        """Парсить требование к параметру."""
        prop = {
            "propertySet": "",
            "baseName": "",
            "dataType": "IFCTEXT",
            "cardinality": "required",
            "enumeration": None,
            "instructions": ""
        }

        # PropertySet
        pset_node = prop_node.SelectSingleNode("ids:propertySet/ids:simpleValue", nsm)
        if pset_node:
            prop["propertySet"] = pset_node.InnerText

        # BaseName (имя параметра)
        name_node = prop_node.SelectSingleNode("ids:baseName/ids:simpleValue", nsm)
        if name_node:
            prop["baseName"] = name_node.InnerText

        # DataType
        datatype_attr = prop_node.GetAttribute("dataType")
        if datatype_attr:
            prop["dataType"] = datatype_attr.upper()

        # Cardinality (required/optional)
        cardinality = prop_node.GetAttribute("cardinality")
        if cardinality:
            prop["cardinality"] = cardinality

        # Instructions (подсказка)
        instr_attr = prop_node.GetAttribute("instructions")
        if instr_attr:
            prop["instructions"] = instr_attr

        # Enumeration (допустимые значения)
        # В IDS файлах enumeration может быть с namespace xs: или ids:
        # Пробуем оба варианта
        enum_nodes = prop_node.SelectNodes("ids:value/xs:restriction/xs:enumeration", nsm)
        if not enum_nodes or enum_nodes.Count == 0:
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

    def _get_simple_value(self, node):
        """Получить простое значение из узла."""
        if node is None:
            return None
        # Попробовать simpleValue
        if node.InnerText:
            return node.InnerText
        return None

    def _collect_unique_properties(self):
        """Собрать параметры из всех спецификаций.

        Создаёт два набора:
        1. all_properties / property_sets - уникальные по ИМЕНИ (для ФОП)
        2. all_properties_full / property_sets_full - все с учётом PropertySet (для Mapping/Report)
        """
        # Для ФОП - дедупликация по имени параметра
        seen_by_name = {}  # baseName -> prop_copy
        # Для Mapping/Report - дедупликация по (propertySet, baseName)
        seen_by_pset_name = {}  # (propertySet, baseName) -> prop_copy

        duplicates_fop = 0
        duplicates_full = 0

        Logger.info(SCRIPT_NAME, "Сбор параметров из {} спецификаций".format(len(self.specifications)))

        for spec in self.specifications:
            entities = spec["applicability"]

            for prop in spec["requirements"]:
                base_name = prop["baseName"]
                pset = prop["propertySet"]

                # === 1. Для ФОП (уникальность по имени) ===
                if base_name not in seen_by_name:
                    prop_copy = dict(prop)
                    prop_copy["entities"] = list(entities)
                    seen_by_name[base_name] = prop_copy
                    self.all_properties.append(prop_copy)

                    if pset not in self.property_sets:
                        self.property_sets[pset] = []
                    self.property_sets[pset].append(prop_copy)

                    Logger.debug(SCRIPT_NAME, "  + ФОП: {} (PropertySet: {})".format(base_name, pset))
                else:
                    duplicates_fop += 1
                    existing = seen_by_name[base_name]
                    for ent in entities:
                        if ent not in existing["entities"]:
                            existing["entities"].append(ent)

                # === 2. Для Mapping/Report (уникальность по PropertySet + имя) ===
                key_full = (pset, base_name)
                if key_full not in seen_by_pset_name:
                    prop_copy_full = dict(prop)
                    prop_copy_full["entities"] = list(entities)
                    seen_by_pset_name[key_full] = prop_copy_full
                    self.all_properties_full.append(prop_copy_full)

                    if pset not in self.property_sets_full:
                        self.property_sets_full[pset] = []
                    self.property_sets_full[pset].append(prop_copy_full)
                else:
                    duplicates_full += 1
                    existing_full = seen_by_pset_name[key_full]
                    for ent in entities:
                        if ent not in existing_full["entities"]:
                            existing_full["entities"].append(ent)

        Logger.info(SCRIPT_NAME, "Параметров для ФОП (уникальных по имени): {}".format(len(self.all_properties)))
        Logger.info(SCRIPT_NAME, "Параметров для Mapping/Report (с PropertySet): {}".format(len(self.all_properties_full)))
        Logger.info(SCRIPT_NAME, "Пропущено дубликатов (ФОП): {}".format(duplicates_fop))
        Logger.info(SCRIPT_NAME, "Пропущено дубликатов (Mapping): {}".format(duplicates_full))


# === ГЕНЕРАТОР ФОП ===

class FOPGenerator:
    """Генератор файла общих параметров Revit."""

    def __init__(self, properties, property_sets, prefix=""):
        self.properties = properties
        self.property_sets = property_sets
        self.prefix = prefix.strip()

    def _apply_prefix(self, name):
        """Добавить префикс к имени параметра."""
        if self.prefix:
            return "{}_{}".format(self.prefix, name)
        return name

    def generate(self, output_path):
        """Сгенерировать ФОП файл."""
        Logger.log_separator(SCRIPT_NAME, "ГЕНЕРАЦИЯ ФОП ФАЙЛА")
        Logger.info(SCRIPT_NAME, "Путь: {}".format(output_path))
        Logger.info(SCRIPT_NAME, "Префикс: {}".format(self.prefix if self.prefix else "(нет)"))

        lines = []

        # Заголовок
        lines.append("# This is a Revit shared parameter file.")
        lines.append("# Generated from IDS file by CPSK Tools.")
        if self.prefix:
            lines.append("# Prefix: {}".format(self.prefix))
        lines.append("*META\tVERSION\tMINVERSION")
        lines.append("META\t2\t1")
        lines.append("*GROUP\tID\tNAME")

        # Группы (PropertySets)
        group_ids = {}
        for i, pset in enumerate(sorted(self.property_sets.keys()), 1):
            group_ids[pset] = i
            lines.append("GROUP\t{}\t{}".format(i, pset))

        Logger.info(SCRIPT_NAME, "Создано групп: {}".format(len(group_ids)))

        # Параметры
        lines.append("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE")

        for prop in self.properties:
            guid = str(uuid.uuid4()).upper()
            name = self._apply_prefix(prop["baseName"])
            datatype = IFC_TO_REVIT_TYPE.get(prop["dataType"], "TEXT")
            group_id = group_ids.get(prop["propertySet"], 1)

            # Описание: инструкции из IDS + enumeration
            desc_parts = []

            # Инструкции (подсказка из IDS)
            if prop.get("instructions"):
                # Убираем переносы строк для ФОП
                instr = prop["instructions"].replace("\n", " ").replace("\r", "").strip()
                desc_parts.append(instr)

            # Допустимые значения
            if prop.get("enumeration"):
                enum_str = "Values: " + ", ".join(prop["enumeration"][:10])
                if len(prop["enumeration"]) > 10:
                    enum_str += "..."
                desc_parts.append(enum_str)

            description = " | ".join(desc_parts) if desc_parts else ""

            line = "PARAM\t{guid}\t{name}\t{dtype}\t\t{group}\t1\t{desc}\t1\t0".format(
                guid=guid,
                name=name,
                dtype=datatype,
                group=group_id,
                desc=description
            )
            lines.append(line)

            Logger.debug(SCRIPT_NAME, "  Параметр: {} | {} | группа: {}".format(name, datatype, prop["propertySet"]))

        # Записать файл в UTF-16 LE с BOM (требование Revit!)
        with codecs.open(output_path, 'w', 'utf-16') as f:
            f.write("\n".join(lines))

        Logger.info(SCRIPT_NAME, "ФОП файл создан: {} параметров".format(len(self.properties)))
        return len(self.properties)


# === ГЕНЕРАТОР IFC MAPPING ===

class IFCMappingGenerator:
    """Генератор файла маппинга IFC параметров."""

    def __init__(self, properties, property_sets, prefix=""):
        self.properties = properties
        self.property_sets = property_sets
        self.prefix = prefix.strip()

    def _apply_prefix(self, name):
        """Добавить префикс к имени параметра."""
        if self.prefix:
            return "{}_{}".format(self.prefix, name)
        return name

    def generate(self, output_path):
        """
        Сгенерировать IFC Mapping файл в формате Revit.

        Формат Revit IFC Parameter Mapping:
        PropertySet:	<PsetName>	I	<IfcClass1>,<IfcClass2>,...
        	<IFC Property Name>	<Type>	<Revit Param Name>
        """
        lines = []

        # Группируем по PropertySet и собираем все IFC классы для каждого PropertySet
        pset_entities = {}  # PropertySet -> set of IFC entities
        pset_props = {}     # PropertySet -> list of properties

        for pset in self.property_sets.keys():
            pset_entities[pset] = set()
            pset_props[pset] = []

            for prop in self.property_sets[pset]:
                entities = prop.get("entities", [])
                for ent in entities:
                    # Конвертируем в правильный формат IfcXxx
                    ifc_class = ent.upper()
                    if ifc_class.startswith("IFC"):
                        # Преобразуем IFCWALL -> IfcWall
                        ifc_class = "Ifc" + ifc_class[3:].capitalize()
                        # Обработка составных имён (IFCWALLSTANDARDCASE -> IfcWallStandardCase)
                        if "STANDARDCASE" in ent.upper():
                            ifc_class = "IfcWallStandardCase"
                        elif "ELEMENTASSEMBLY" in ent.upper():
                            ifc_class = "IfcElementAssembly"
                        elif "BUILDINGELEMENTPROXY" in ent.upper():
                            ifc_class = "IfcBuildingElementProxy"
                        elif "REINFORCINGBAR" in ent.upper():
                            ifc_class = "IfcReinforcingBar"
                    pset_entities[pset].add(ifc_class)

                pset_props[pset].append(prop)

        # Генерируем файл
        for pset in sorted(self.property_sets.keys()):
            # Список IFC классов
            ifc_classes = sorted(pset_entities[pset])
            if not ifc_classes:
                ifc_classes = ["IfcBuildingElement"]

            # Строка PropertySet
            lines.append("PropertySet:\t{}\tI\t{}".format(pset, ",".join(ifc_classes)))

            # Параметры (с отступом TAB)
            for prop in pset_props[pset]:
                ifc_prop_name = prop["baseName"]
                revit_param = self._apply_prefix(prop["baseName"])
                dtype = "Text"  # Revit использует Text для большинства

                # Формат: TAB + IFC Property Name + TAB + Type + TAB + Revit Param Name
                lines.append("\t{}\t{}\t{}".format(ifc_prop_name, dtype, revit_param))

            # Пустая строка между PropertySets
            lines.append("")

        # Записать файл
        with codecs.open(output_path, 'w', 'utf-8') as f:
            f.write("\n".join(lines))

        return len(self.properties)


# === ГЕНЕРАТОР ОТЧЁТА ===

class ReportGenerator:
    """Генератор отчёта о параметрах."""

    def __init__(self, parser, prefix=""):
        self.parser = parser
        self.prefix = prefix.strip()

    def _apply_prefix(self, name):
        """Добавить префикс к имени параметра."""
        if self.prefix:
            return "{}_{}".format(self.prefix, name)
        return name

    def generate(self, output_path):
        """Сгенерировать HTML отчёт."""
        html = []
        html.append("<!DOCTYPE html>")
        html.append("<html><head>")
        html.append("<meta charset='utf-8'>")
        html.append("<title>IDS Parameters Report</title>")
        html.append("<style>")
        html.append("body { font-family: Arial, sans-serif; margin: 20px; }")
        html.append("h1 { color: #333; }")
        html.append("h2 { color: #666; border-bottom: 1px solid #ccc; }")
        html.append("table { border-collapse: collapse; width: 100%; margin: 10px 0; }")
        html.append("th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }")
        html.append("th { background: #f5f5f5; }")
        html.append(".required { color: #c00; }")
        html.append(".optional { color: #999; }")
        html.append(".fop-param { font-family: monospace; background: #f0f0f0; }")
        html.append("</style>")
        html.append("</head><body>")

        html.append("<h1>IDS Parameters Report</h1>")
        if self.prefix:
            html.append("<p>Prefix: <strong>{}</strong></p>".format(self.prefix))
        html.append("<p>Total PropertySets: {}</p>".format(len(self.parser.property_sets_full)))
        html.append("<p>Total Parameters: {}</p>".format(len(self.parser.all_properties_full)))

        # Таблица по PropertySets (используем полный список с учётом PropertySet)
        for pset in sorted(self.parser.property_sets_full.keys()):
            props = self.parser.property_sets_full[pset]
            html.append("<h2>{} ({} params)</h2>".format(pset, len(props)))
            html.append("<table>")
            html.append("<tr><th>IDS Parameter</th><th>FOP Parameter</th><th>Type</th><th>Required</th><th>Entities</th><th>Allowed Values</th></tr>")

            for prop in props:
                req_class = "required" if prop["cardinality"] == "required" else "optional"
                req_text = "Yes" if prop["cardinality"] == "required" else "No"
                entities = ", ".join(prop.get("entities", [])[:3])
                if len(prop.get("entities", [])) > 3:
                    entities += "..."

                # Допустимые значения из enumeration
                enum_list = prop.get("enumeration") or []
                if enum_list:
                    values = ", ".join(enum_list)
                else:
                    values = "-"

                # Имя параметра в ФОП (с префиксом)
                fop_param = self._apply_prefix(prop["baseName"])

                html.append("<tr>")
                html.append("<td>{}</td>".format(prop["baseName"]))
                html.append("<td class='fop-param'>{}</td>".format(fop_param))
                html.append("<td>{}</td>".format(prop["dataType"]))
                html.append("<td class='{}'>{}</td>".format(req_class, req_text))
                html.append("<td>{}</td>".format(entities))
                html.append("<td>{}</td>".format(values))
                html.append("</tr>")

            html.append("</table>")

        html.append("</body></html>")

        with codecs.open(output_path, 'w', 'utf-8') as f:
            f.write("\n".join(html))


# === ГЛАВНОЕ ОКНО ===

class IDStoFOPForm(Form):
    """Диалог конвертации IDS в ФОП."""

    def __init__(self):
        self.ids_path = None
        self.output_folder = None
        self.result = None
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "IDS в ФОП - Создание параметров"
        self.Width = 550
        self.Height = 440
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

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
        self.txt_ids.Width = 420
        self.txt_ids.ReadOnly = True
        self.Controls.Add(self.txt_ids)

        self.btn_ids = Button()
        self.btn_ids.Text = "Обзор..."
        self.btn_ids.Location = Point(445, y - 2)
        self.btn_ids.Width = 80
        self.btn_ids.Click += self.on_browse_ids
        self.Controls.Add(self.btn_ids)

        # === Папка вывода ===
        y += 35
        lbl_output = Label()
        lbl_output.Text = "Папка для сохранения:"
        lbl_output.Location = Point(15, y)
        lbl_output.AutoSize = True
        self.Controls.Add(lbl_output)

        y += 20
        self.txt_output = TextBox()
        self.txt_output.Location = Point(15, y)
        self.txt_output.Width = 420
        self.txt_output.ReadOnly = True
        self.Controls.Add(self.txt_output)

        self.btn_output = Button()
        self.btn_output.Text = "Обзор..."
        self.btn_output.Location = Point(445, y - 2)
        self.btn_output.Width = 80
        self.btn_output.Click += self.on_browse_output
        self.Controls.Add(self.btn_output)

        # === Префикс параметров ===
        y += 35
        lbl_prefix = Label()
        lbl_prefix.Text = "Префикс параметров (опционально):"
        lbl_prefix.Location = Point(15, y)
        lbl_prefix.AutoSize = True
        self.Controls.Add(lbl_prefix)

        y += 20
        self.txt_prefix = TextBox()
        self.txt_prefix.Location = Point(15, y)
        self.txt_prefix.Width = 200
        self.txt_prefix.Text = ""
        self.Controls.Add(self.txt_prefix)

        lbl_prefix_hint = Label()
        lbl_prefix_hint.Text = "Например: ЦГЭ, МОЭСК, Заказчик1"
        lbl_prefix_hint.Location = Point(225, y + 3)
        lbl_prefix_hint.AutoSize = True
        lbl_prefix_hint.ForeColor = Color.Gray
        self.Controls.Add(lbl_prefix_hint)

        # === Опции генерации ===
        y += 35
        grp_options = GroupBox()
        grp_options.Text = "Генерировать"
        grp_options.Location = Point(15, y)
        grp_options.Size = Size(510, 80)

        self.chk_fop = CheckBox()
        self.chk_fop.Text = "ФОП файл (Shared Parameters)"
        self.chk_fop.Location = Point(15, 22)
        self.chk_fop.AutoSize = True
        self.chk_fop.Checked = True
        grp_options.Controls.Add(self.chk_fop)

        self.chk_mapping = CheckBox()
        self.chk_mapping.Text = "IFC Mapping файл"
        self.chk_mapping.Location = Point(15, 45)
        self.chk_mapping.AutoSize = True
        self.chk_mapping.Checked = True
        grp_options.Controls.Add(self.chk_mapping)

        self.chk_report = CheckBox()
        self.chk_report.Text = "HTML отчёт"
        self.chk_report.Location = Point(280, 22)
        self.chk_report.AutoSize = True
        self.chk_report.Checked = True
        grp_options.Controls.Add(self.chk_report)

        self.Controls.Add(grp_options)

        # === Статус ===
        y += 95
        self.lbl_status = Label()
        self.lbl_status.Text = "Выберите IDS файл и папку для сохранения"
        self.lbl_status.Location = Point(15, y)
        self.lbl_status.Size = Size(510, 40)
        self.Controls.Add(self.lbl_status)

        # === Кнопки ===
        y += 50
        self.btn_generate = Button()
        self.btn_generate.Text = "Генерировать"
        self.btn_generate.Location = Point(15, y)
        self.btn_generate.Width = 120
        self.btn_generate.Height = 30
        self.btn_generate.Enabled = False
        self.btn_generate.Click += self.on_generate
        self.Controls.Add(self.btn_generate)

        self.btn_open_folder = Button()
        self.btn_open_folder.Text = "Открыть папку"
        self.btn_open_folder.Location = Point(145, y)
        self.btn_open_folder.Width = 110
        self.btn_open_folder.Height = 30
        self.btn_open_folder.Enabled = False
        self.btn_open_folder.Click += self.on_open_folder
        self.Controls.Add(self.btn_open_folder)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(445, y)
        btn_close.Width = 80
        btn_close.Height = 30
        btn_close.Click += self.on_close
        self.Controls.Add(btn_close)

    def on_browse_ids(self, sender, args):
        """Выбор IDS файла."""
        dialog = OpenFileDialog()
        dialog.Title = "Выберите IDS файл"
        dialog.Filter = "IDS файлы (*.ids)|*.ids|XML файлы (*.xml)|*.xml|Все файлы (*.*)|*.*"
        if dialog.ShowDialog() == DialogResult.OK:
            self.ids_path = dialog.FileName
            self.txt_ids.Text = self.ids_path
            self.update_state()

    def on_browse_output(self, sender, args):
        """Выбор папки вывода."""
        dialog = FolderBrowserDialog()
        dialog.Description = "Выберите папку для сохранения файлов"
        if dialog.ShowDialog() == DialogResult.OK:
            self.output_folder = dialog.SelectedPath
            self.txt_output.Text = self.output_folder
            self.update_state()

    def update_state(self):
        """Обновить состояние кнопок."""
        can_generate = bool(self.ids_path and self.output_folder)
        self.btn_generate.Enabled = can_generate
        if can_generate:
            self.lbl_status.Text = "Готово к генерации"
            self.lbl_status.ForeColor = Color.Black

    def on_generate(self, sender, args):
        """Генерация файлов."""
        try:
            Logger.log_separator(SCRIPT_NAME, "ГЕНЕРАЦИЯ ФАЙЛОВ")
            Logger.info(SCRIPT_NAME, "IDS файл: {}".format(self.ids_path))
            Logger.info(SCRIPT_NAME, "Папка вывода: {}".format(self.output_folder))

            self.lbl_status.Text = "Парсинг IDS файла..."
            self.lbl_status.ForeColor = Color.Black
            System.Windows.Forms.Application.DoEvents()

            # Парсинг IDS
            Logger.info(SCRIPT_NAME, "Парсинг IDS файла...")
            parser = IDSParser(self.ids_path)
            parser.parse()

            if not parser.all_properties:
                Logger.warning(SCRIPT_NAME, "В IDS файле не найдено параметров!")
                self.lbl_status.Text = "В IDS файле не найдено параметров!"
                self.lbl_status.ForeColor = Color.Red
                return

            results = []
            base_name = os.path.splitext(os.path.basename(self.ids_path))[0]
            prefix = self.txt_prefix.Text.strip()

            # Добавить префикс к имени файла если указан
            if prefix:
                file_prefix = "{}_{}".format(prefix, base_name)
                Logger.info(SCRIPT_NAME, "Префикс файлов: {}".format(file_prefix))
            else:
                file_prefix = base_name

            # ФОП - используем дедуплицированные по имени (all_properties)
            if self.chk_fop.Checked:
                fop_path = os.path.join(self.output_folder, file_prefix + "_SharedParams.txt")
                gen = FOPGenerator(parser.all_properties, parser.property_sets, prefix)
                count = gen.generate(fop_path)
                results.append("ФОП: {} параметров".format(count))

            # IFC Mapping - используем ПОЛНЫЙ список (all_properties_full)
            if self.chk_mapping.Checked:
                map_path = os.path.join(self.output_folder, file_prefix + "_IFCMapping.txt")
                gen = IFCMappingGenerator(parser.all_properties_full, parser.property_sets_full, prefix)
                count = gen.generate(map_path)
                results.append("IFC Mapping: {} параметров".format(count))
                Logger.info(SCRIPT_NAME, "IFC Mapping создан: {} параметров".format(count))

            # HTML Report
            if self.chk_report.Checked:
                report_path = os.path.join(self.output_folder, file_prefix + "_Report.html")
                gen = ReportGenerator(parser, prefix)
                gen.generate(report_path)
                results.append("HTML отчёт создан")
                Logger.info(SCRIPT_NAME, "HTML отчёт создан: {}".format(report_path))

            Logger.log_separator(SCRIPT_NAME, "ИТОГИ")
            Logger.info(SCRIPT_NAME, "Генерация завершена успешно")
            Logger.info(SCRIPT_NAME, "Результаты: {}".format(", ".join(results)))
            Logger.info(SCRIPT_NAME, "Лог: {}".format(Logger.get_log_path()))

            self.lbl_status.Text = "Готово! " + ", ".join(results)
            self.lbl_status.ForeColor = Color.Green
            self.btn_open_folder.Enabled = True

        except Exception as e:
            Logger.error(SCRIPT_NAME, "Ошибка генерации: {}".format(str(e)))
            self.lbl_status.Text = "Ошибка: {}".format(str(e))
            self.lbl_status.ForeColor = Color.Red

    def on_open_folder(self, sender, args):
        """Открыть папку с результатами."""
        if self.output_folder:
            os.startfile(self.output_folder)

    def on_close(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    form = IDStoFOPForm()
    form.ShowDialog()
    Logger.info(SCRIPT_NAME, "Скрипт завершён")
