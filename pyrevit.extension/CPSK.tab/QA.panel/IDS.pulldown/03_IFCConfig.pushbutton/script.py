# -*- coding: utf-8 -*-
"""
IFC Config Manager - управление конфигурациями IFC экспорта.

Позволяет:
- Сохранить конфигурацию из Revit в JSON
- Загрузить конфигурацию из JSON
- Автоматически создать конфигурацию с файлом маппинга
"""

__title__ = "Конфигурация\nIFC"
__author__ = "CPSK"

import clr
import os
import json
import codecs

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, Panel, ComboBox, CheckBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    DialogResult, OpenFileDialog, SaveFileDialog,
    GroupBox, TextBox, ListBox, SelectionMode
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import revit, forms, script

# Revit API
from Autodesk.Revit.DB import Transaction, FilteredElementCollector

# === НАСТРОЙКИ ===

doc = revit.doc
app = doc.Application
output = script.get_output()

SCRIPT_DIR = os.path.dirname(__file__)
# IFCConfig.pushbutton -> IDS.pulldown -> QA.panel -> CPSK.tab -> pyrevit.extension -> lib
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")

# Путь к файлу конфигурации по умолчанию
DEFAULT_CONFIG_PATH = os.path.join(LIB_DIR, "CPSK_IFC_Config.json")

# Лог файл для отладки
LOG_FILE = os.path.join(LIB_DIR, "ifc_config_debug.log")


def log(message):
    """Записать сообщение в лог файл."""
    try:
        with codecs.open(LOG_FILE, 'a', 'utf-8') as f:
            f.write(message + "\n")
    except:
        pass


def clear_log():
    """Очистить лог файл."""
    try:
        with codecs.open(LOG_FILE, 'w', 'utf-8') as f:
            f.write("=== IFC Config Debug Log ===\n")
    except:
        pass


# === ЧТЕНИЕ/ЗАПИСЬ КОНФИГУРАЦИЙ IFC ===

def get_ifc_storage_info():
    """
    Получить информацию о DataStorage, Schema и Field для IFC конфигураций.
    Возвращает (storage, schema, field, current_data) или (None, None, None, None).
    """
    try:
        from Autodesk.Revit.DB.ExtensibleStorage import DataStorage, Schema

        storages = FilteredElementCollector(doc).OfClass(DataStorage).ToElements()

        for storage in storages:
            try:
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
                                            data = json.loads(value)
                                            return storage, schema, field, data
                                    except:
                                        pass
            except:
                pass
    except Exception as e:
        output.print_md("Ошибка: {}".format(str(e)))

    return None, None, None, None


def save_ifc_config_to_storage(new_config):
    """
    Сохранить новую конфигурацию IFC в ExtensibleStorage документа.
    Каждая конфигурация хранится в ОТДЕЛЬНОМ DataStorage элементе!
    Возвращает (success, error_message).
    """
    clear_log()
    log("== save_ifc_config_to_storage ==")
    log("Каждая конфигурация = отдельный DataStorage!")

    try:
        from Autodesk.Revit.DB.ExtensibleStorage import DataStorage, Schema, SchemaBuilder, AccessLevel, Entity
        from System import Guid

        config_name = new_config.get("Name", "")
        log("- config_name: {}".format(config_name))

        if not config_name:
            log("ОШИБКА: Нет имени конфигурации")
            return False, "Конфигурация должна иметь имя (Name)"

        # Проверить, не существует ли уже конфигурация с таким именем
        log("")
        log("Шаг 1: Проверка существующих конфигураций...")
        existing_configs = get_all_ifc_configs()
        for name, cfg, schema_name in existing_configs:
            if name == config_name:
                log("ОШИБКА: Конфигурация '{}' уже существует!".format(config_name))
                return False, "Конфигурация '{}' уже существует!".format(config_name)
        log("- Существующих конфигураций: {}".format(len(existing_configs)))
        log("- Конфигурация '{}' не существует, можно создавать".format(config_name))

        # Найти существующую схему IFC
        log("")
        log("Шаг 2: Поиск схемы IFCExportConfigurationMap...")
        existing_storage, existing_schema, existing_field, _ = get_ifc_storage_info()

        if not existing_schema:
            log("ОШИБКА: Схема IFC не найдена!")
            return False, "Схема IFC не найдена. Сначала создайте конфигурацию вручную в диалоге экспорта IFC."

        log("- Найдена схема: {}".format(existing_schema.SchemaName))
        log("- GUID схемы: {}".format(existing_schema.GUID))
        log("- Поле: {}".format(existing_field.FieldName if existing_field else None))

        # Сериализовать конфигурацию в JSON
        config_json = json.dumps(new_config, ensure_ascii=False)
        log("")
        log("Шаг 3: JSON создан, длина: {} символов".format(len(config_json)))
        log("JSON preview (first 300 chars):")
        log(config_json[:300])

        # Создать новый DataStorage и записать конфигурацию
        log("")
        log("Шаг 4: Создание нового DataStorage...")

        trans = Transaction(doc, "Create IFC Configuration: {}".format(config_name))
        trans.Start()
        log("- Транзакция начата")

        try:
            # Создать новый DataStorage
            new_storage = DataStorage.Create(doc)
            log("- Новый DataStorage создан: {}".format(new_storage.Id))

            # Создать Entity с существующей схемой
            entity = Entity(existing_schema)
            log("- Entity создан со схемой {}".format(existing_schema.SchemaName))

            # Записать JSON в поле
            entity.Set[str](existing_field, config_json)
            log("- JSON записан в поле {}".format(existing_field.FieldName))

            # Установить Entity в Storage
            new_storage.SetEntity(entity)
            log("- Entity установлен в DataStorage")

            # Установить имя DataStorage (для идентификации)
            try:
                new_storage.Name = "IFCExportConfiguration_{}".format(config_name)
                log("- Имя DataStorage: {}".format(new_storage.Name))
            except:
                log("- Не удалось установить имя DataStorage")

            trans.Commit()
            log("- Транзакция завершена (Commit)")

            # Проверка после сохранения
            log("")
            log("Шаг 5: Проверка после сохранения...")
            verify_configs = get_all_ifc_configs()
            log("- Всего конфигураций: {}".format(len(verify_configs)))

            found = False
            for name, cfg, schema_name in verify_configs:
                if name == config_name:
                    found = True
                    log("УСПЕХ: Конфигурация '{}' найдена!".format(config_name))
                    break

            if not found:
                log("ВНИМАНИЕ: Конфигурация '{}' НЕ найдена после сохранения!".format(config_name))

            log("")
            log("== ЗАВЕРШЕНО ==")
            return True, None

        except Exception as e:
            trans.RollBack()
            log("ОШИБКА при создании: {}".format(str(e)))
            import traceback
            log(traceback.format_exc())
            return False, "Ошибка создания: {}".format(str(e))

    except Exception as e:
        import traceback
        log("ИСКЛЮЧЕНИЕ: {}".format(str(e)))
        log(traceback.format_exc())
        return False, "Ошибка: {}".format(str(e))


def get_ifc_config_from_storage(config_name=None):
    """
    Получить конфигурацию IFC из ExtensibleStorage документа.
    Возвращает dict с настройками конфигурации.
    """
    try:
        from Autodesk.Revit.DB.ExtensibleStorage import DataStorage, Schema

        storages = FilteredElementCollector(doc).OfClass(DataStorage).ToElements()

        for storage in storages:
            try:
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
                                            data = json.loads(value)

                                            # Если ищем конкретную конфигурацию
                                            if config_name:
                                                if isinstance(data, dict):
                                                    if data.get("Name") == config_name:
                                                        return data
                                                    # Поиск в словаре конфигураций
                                                    for key, cfg in data.items():
                                                        if isinstance(cfg, dict) and cfg.get("Name") == config_name:
                                                            return cfg
                                            else:
                                                return data
                                    except:
                                        pass
            except:
                pass
    except Exception as e:
        output.print_md("Ошибка чтения ExtensibleStorage: {}".format(str(e)))

    return None


def get_all_ifc_configs():
    """
    Получить все конфигурации IFC из ExtensibleStorage.
    Возвращает список (name, config_dict).
    """
    configs = []

    try:
        from Autodesk.Revit.DB.ExtensibleStorage import DataStorage, Schema

        storages = FilteredElementCollector(doc).OfClass(DataStorage).ToElements()

        for storage in storages:
            try:
                guids = storage.GetEntitySchemaGuids()
                for guid in guids:
                    schema = Schema.Lookup(guid)
                    if schema:
                        schema_name = schema.SchemaName
                        entity = storage.GetEntity(schema)
                        if entity and entity.IsValid():
                            for field in schema.ListFields():
                                if "map" in field.FieldName.lower():
                                    try:
                                        value = entity.Get[str](field)
                                        if value:
                                            data = json.loads(value)

                                            # Одна конфигурация
                                            if isinstance(data, dict) and "Name" in data:
                                                name = data.get("Name", "Unnamed")
                                                configs.append((name, data, schema_name))
                                            # Словарь конфигураций
                                            elif isinstance(data, dict):
                                                for key, cfg in data.items():
                                                    if isinstance(cfg, dict) and "Name" in cfg:
                                                        name = cfg.get("Name", key)
                                                        configs.append((name, cfg, schema_name))
                                            # Список конфигураций
                                            elif isinstance(data, list):
                                                for cfg in data:
                                                    if isinstance(cfg, dict) and "Name" in cfg:
                                                        name = cfg.get("Name", "Unnamed")
                                                        configs.append((name, cfg, schema_name))
                                    except:
                                        pass
            except:
                pass
    except Exception as e:
        output.print_md("Ошибка: {}".format(str(e)))

    return configs


def save_config_to_json(config_data, output_path):
    """
    Сохранить конфигурацию в JSON файл.
    """
    try:
        with codecs.open(output_path, 'w', 'utf-8') as f:
            json.dump(config_data, f, ensure_ascii=False, indent=2)
        return True, None
    except Exception as e:
        return False, str(e)


def load_config_from_json(json_path):
    """
    Загрузить конфигурацию из JSON файла.
    """
    try:
        with codecs.open(json_path, 'r', 'utf-8') as f:
            return json.load(f), None
    except Exception as e:
        return None, str(e)


def create_default_config(mapping_file_path=None):
    """
    Создать конфигурацию по умолчанию для CPSK.

    Поля конфигурации IFC (на основе IFC Exporter):
    - Name: имя конфигурации
    - IFCVersion: версия IFC (20 = IFC4RV, 21 = IFC4DTV, 1 = IFC2x3)
    - ExchangeRequirement: 0-4
    - IFCFileType: 0 = IFC, 1 = IFCXML, 2 = IFCZIP
    - SpaceBoundaries: 0-2
    - ExportBaseQuantities: bool
    - SplitWallsAndColumns: bool
    - ExportUserDefinedPsets: bool
    - ExportUserDefinedPsetsFileName: путь к файлу маппинга
    - и другие...
    """
    config = {
        "Name": "CPSK_Export",
        "IFCVersion": 20,  # IFC4 Reference View
        "ExchangeRequirement": 3,  # Design Transfer
        "IFCFileType": 0,  # IFC
        "SpaceBoundaries": 0,
        "ActivePhaseId": -1,
        "ExportBaseQuantities": True,
        "SplitWallsAndColumns": False,
        "VisibleElementsOfCurrentView": False,
        "Use2DRoomBoundaryForVolume": False,
        "UseFamilyAndTypeNameForReference": True,
        "Export2DElements": False,
        "ExportPartsAsBuildingElements": False,
        "ExportBoundingBox": False,
        "ExportSolidModelRep": False,
        "ExportSchedulesAsPsets": False,
        "ExportUserDefinedPsets": True,
        "ExportUserDefinedPsetsFileName": mapping_file_path or "",
        "ExportLinkedFiles": False,
        "IncludeSiteElevation": True,
        "UseActiveViewGeometry": False,
        "ExportSpecificSchedules": False,
        "TessellationLevelOfDetail": 0.5,
        "StoreIFCGUID": True,
        "ExportRoomsInView": False,
        "UseOnlyTriangulation": False,
        "UseTypeNameOnlyForIfcType": False,
        "UseVisibleRevitNameAsEntityName": False,
        "COBieCompanyInfo": "",
        "COBieProjectInfo": "",
        "IncludeSteelElements": True,
        "GeoRefCRSName": "",
        "GeoRefCRSDesc": "",
        "GeoRefEPSGCode": "",
        "GeoRefGeodeticDatum": "",
        "GeoRefMapUnit": "",
        "ExcludeFilter": "",
        "SitePlacement": 0
    }

    return config


# === ФОРМА УПРАВЛЕНИЯ КОНФИГУРАЦИЯМИ ===

class IFCConfigForm(Form):
    """Диалог управления конфигурациями IFC."""

    def __init__(self):
        self.configs = []
        self.selected_config = None
        self.base_config = None  # Базовая конфигурация для создания новой
        self.base_config_path = None
        self.setup_form()
        self.load_configs()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Управление конфигурациями IFC"
        self.Width = 700
        self.Height = 530
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # === Конфигурации в документе ===
        grp_configs = GroupBox()
        grp_configs.Text = "Конфигурации IFC в документе"
        grp_configs.Location = Point(15, y)
        grp_configs.Size = Size(655, 180)

        self.lst_configs = ListBox()
        self.lst_configs.Location = Point(15, 20)
        self.lst_configs.Size = Size(400, 120)
        self.lst_configs.SelectionMode = SelectionMode.One
        self.lst_configs.SelectedIndexChanged += self.on_config_selected
        grp_configs.Controls.Add(self.lst_configs)

        btn_refresh = Button()
        btn_refresh.Text = "Обновить"
        btn_refresh.Location = Point(430, 20)
        btn_refresh.Size = Size(100, 28)
        btn_refresh.Click += self.on_refresh
        grp_configs.Controls.Add(btn_refresh)

        btn_save = Button()
        btn_save.Text = "Сохранить в JSON"
        btn_save.Location = Point(430, 55)
        btn_save.Size = Size(130, 28)
        btn_save.Click += self.on_save_config
        grp_configs.Controls.Add(btn_save)

        btn_details = Button()
        btn_details.Text = "Показать детали"
        btn_details.Location = Point(430, 90)
        btn_details.Size = Size(130, 28)
        btn_details.Click += self.on_show_details
        grp_configs.Controls.Add(btn_details)

        self.lbl_config_status = Label()
        self.lbl_config_status.Text = ""
        self.lbl_config_status.Location = Point(15, 148)
        self.lbl_config_status.Size = Size(620, 20)
        self.lbl_config_status.ForeColor = Color.Gray
        grp_configs.Controls.Add(self.lbl_config_status)

        self.Controls.Add(grp_configs)

        y += 195

        # === Создание новой конфигурации на основе существующей ===
        grp_create = GroupBox()
        grp_create.Text = "Создать новую конфигурацию (на основе существующей)"
        grp_create.Location = Point(15, y)
        grp_create.Size = Size(655, 195)

        # Базовая конфигурация
        lbl_base = Label()
        lbl_base.Text = "Базовая конфигурация (JSON):"
        lbl_base.Location = Point(15, 22)
        lbl_base.AutoSize = True
        grp_create.Controls.Add(lbl_base)

        self.txt_base_config = TextBox()
        self.txt_base_config.Location = Point(15, 42)
        self.txt_base_config.Width = 520
        self.txt_base_config.ReadOnly = True
        grp_create.Controls.Add(self.txt_base_config)

        btn_browse_base = Button()
        btn_browse_base.Text = "Загрузить..."
        btn_browse_base.Location = Point(545, 40)
        btn_browse_base.Size = Size(95, 26)
        btn_browse_base.Click += self.on_browse_base_config
        grp_create.Controls.Add(btn_browse_base)

        self.lbl_base_status = Label()
        self.lbl_base_status.Text = "Загрузите базовую конфигурацию"
        self.lbl_base_status.Location = Point(15, 70)
        self.lbl_base_status.Size = Size(620, 18)
        self.lbl_base_status.ForeColor = Color.Gray
        grp_create.Controls.Add(self.lbl_base_status)

        # Новое имя конфигурации
        lbl_config_name = Label()
        lbl_config_name.Text = "Новое имя:"
        lbl_config_name.Location = Point(15, 95)
        lbl_config_name.AutoSize = True
        grp_create.Controls.Add(lbl_config_name)

        self.txt_config_name = TextBox()
        self.txt_config_name.Location = Point(100, 92)
        self.txt_config_name.Width = 200
        self.txt_config_name.Text = "CPSK_Export"
        grp_create.Controls.Add(self.txt_config_name)

        # Новый файл маппинга
        lbl_mapping = Label()
        lbl_mapping.Text = "Новый файл маппинга:"
        lbl_mapping.Location = Point(15, 125)
        lbl_mapping.AutoSize = True
        grp_create.Controls.Add(lbl_mapping)

        self.txt_mapping = TextBox()
        self.txt_mapping.Location = Point(15, 145)
        self.txt_mapping.Width = 520

        # Попробовать найти существующий файл маппинга
        default_mapping = os.path.join(LIB_DIR, "CPSK_PropertySet_Mapping.txt")
        if os.path.exists(default_mapping):
            self.txt_mapping.Text = default_mapping
        grp_create.Controls.Add(self.txt_mapping)

        btn_browse_mapping = Button()
        btn_browse_mapping.Text = "Обзор..."
        btn_browse_mapping.Location = Point(545, 143)
        btn_browse_mapping.Size = Size(95, 26)
        btn_browse_mapping.Click += self.on_browse_mapping
        grp_create.Controls.Add(btn_browse_mapping)

        self.Controls.Add(grp_create)

        y += 210

        # === Кнопки создания ===
        self.btn_create = Button()
        self.btn_create.Text = "Создать в Revit"
        self.btn_create.Location = Point(15, y)
        self.btn_create.Size = Size(130, 32)
        self.btn_create.Click += self.on_create_config
        self.btn_create.Enabled = False
        self.Controls.Add(self.btn_create)

        lbl_hint = Label()
        lbl_hint.Text = "(создаёт конфигурацию в наборах экспорта IFC документа)"
        lbl_hint.Location = Point(155, y + 8)
        lbl_hint.AutoSize = True
        lbl_hint.ForeColor = Color.Gray
        self.Controls.Add(lbl_hint)

        y += 45

        # === Кнопки ===
        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(575, y)
        btn_close.Size = Size(95, 30)
        btn_close.Click += self.on_close
        self.Controls.Add(btn_close)

    def load_configs(self):
        """Загрузить конфигурации из документа."""
        self.lst_configs.Items.Clear()
        self.configs = get_all_ifc_configs()

        for name, cfg, schema in self.configs:
            version = cfg.get("IFCVersion", "?")
            version_str = self.get_version_name(version)
            display = "{} ({}) [{}]".format(name, version_str, schema[:20])
            self.lst_configs.Items.Add(display)

        if self.configs:
            self.lbl_config_status.Text = "Найдено {} конфигураций".format(len(self.configs))
            self.lbl_config_status.ForeColor = Color.DarkGreen
        else:
            self.lbl_config_status.Text = "Конфигурации не найдены"
            self.lbl_config_status.ForeColor = Color.Gray

    def get_version_name(self, version_code):
        """Получить имя версии IFC по коду."""
        version_map = {
            0: "IFC2x2",
            1: "IFC2x3",
            2: "IFC2x3 CV2",
            3: "IFC4",
            20: "IFC4 RV",
            21: "IFC4 DTV",
            22: "IFC4x3"
        }
        return version_map.get(version_code, "IFC {}".format(version_code))

    def on_config_selected(self, sender, args):
        """При выборе конфигурации."""
        idx = self.lst_configs.SelectedIndex
        if idx >= 0 and idx < len(self.configs):
            name, cfg, schema = self.configs[idx]
            self.selected_config = cfg

    def on_refresh(self, sender, args):
        """Обновить список конфигураций."""
        self.load_configs()

    def on_save_config(self, sender, args):
        """Сохранить выбранную конфигурацию в JSON."""
        if not self.selected_config:
            MessageBox.Show(
                "Выберите конфигурацию для сохранения",
                "Внимание",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return

        dialog = SaveFileDialog()
        dialog.Title = "Сохранить конфигурацию IFC"
        dialog.Filter = "JSON files (*.json)|*.json"

        name = self.selected_config.get("Name", "config")
        dialog.FileName = "{}_ifc_config.json".format(name.replace(" ", "_"))
        dialog.InitialDirectory = LIB_DIR

        if dialog.ShowDialog() == DialogResult.OK:
            success, error = save_config_to_json(self.selected_config, dialog.FileName)
            if success:
                MessageBox.Show(
                    "Конфигурация сохранена:\n{}".format(dialog.FileName),
                    "Готово",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )
            else:
                MessageBox.Show(
                    "Ошибка: {}".format(error),
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                )

    def on_show_details(self, sender, args):
        """Показать детали конфигурации."""
        if not self.selected_config:
            MessageBox.Show(
                "Выберите конфигурацию",
                "Внимание",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return

        output.print_md("# Детали конфигурации IFC")
        output.print_md("---")

        for key, value in sorted(self.selected_config.items()):
            if isinstance(value, bool):
                val_str = "Да" if value else "Нет"
            elif isinstance(value, str) and len(value) > 50:
                val_str = value[:50] + "..."
            else:
                val_str = str(value)
            output.print_md("- **{}**: {}".format(key, val_str))

        MessageBox.Show(
            "Детали выведены в окно pyRevit",
            "Готово",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )

    def on_browse_mapping(self, sender, args):
        """Выбрать файл маппинга."""
        dialog = OpenFileDialog()
        dialog.Title = "Выберите файл маппинга PropertySet"
        dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        dialog.InitialDirectory = LIB_DIR

        if dialog.ShowDialog() == DialogResult.OK:
            self.txt_mapping.Text = dialog.FileName

    def on_browse_base_config(self, sender, args):
        """Загрузить базовую конфигурацию из JSON."""
        dialog = OpenFileDialog()
        dialog.Title = "Выберите базовую конфигурацию IFC (JSON)"
        dialog.Filter = "JSON files (*.json)|*.json"
        dialog.InitialDirectory = LIB_DIR

        if dialog.ShowDialog() == DialogResult.OK:
            config, error = load_config_from_json(dialog.FileName)
            if config:
                self.base_config = config
                self.base_config_path = dialog.FileName
                self.txt_base_config.Text = dialog.FileName

                # Показать информацию о загруженной конфигурации
                name = config.get("Name", "Без имени")
                version = config.get("IFCVersion", "?")
                version_str = self.get_version_name(version)
                old_mapping = config.get("ExportUserDefinedPsetsFileName", "")

                self.lbl_base_status.Text = "Загружено: {} ({})".format(name, version_str)
                self.lbl_base_status.ForeColor = Color.DarkGreen

                # Установить имя по умолчанию
                self.txt_config_name.Text = name + "_new"

                # Включить кнопку создания
                self.btn_create.Enabled = True

                # Показать детали в output
                output.print_md("# Загружена базовая конфигурация")
                output.print_md("---")
                output.print_md("**Файл:** {}".format(dialog.FileName))
                output.print_md("**Имя:** {}".format(name))
                output.print_md("**Версия IFC:** {}".format(version_str))
                if old_mapping:
                    output.print_md("**Текущий маппинг:** {}".format(old_mapping))
                output.print_md("")
                output.print_md("Теперь укажите новое имя и файл маппинга.")
            else:
                MessageBox.Show(
                    "Ошибка загрузки: {}".format(error),
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                )

    def on_create_config(self, sender, args):
        """Создать новую конфигурацию на основе базовой и сохранить в Revit."""
        if not self.base_config:
            MessageBox.Show(
                "Сначала загрузите базовую конфигурацию!",
                "Внимание",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return

        mapping_path = self.txt_mapping.Text.strip()
        config_name = self.txt_config_name.Text.strip() or "CPSK_Export"

        if mapping_path and not os.path.exists(mapping_path):
            result = MessageBox.Show(
                "Файл маппинга не существует:\n{}\n\nПродолжить без файла маппинга?".format(mapping_path),
                "Внимание",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            )
            if result != DialogResult.Yes:
                return
            mapping_path = ""

        # Копировать базовую конфигурацию и изменить только нужные поля
        import copy
        config = copy.deepcopy(self.base_config)

        # Изменить только имя и путь к маппингу
        old_name = config.get("Name", "")
        old_mapping = config.get("ExportUserDefinedPsetsFileName", "")

        config["Name"] = config_name
        config["ExportUserDefinedPsetsFileName"] = mapping_path
        if mapping_path:
            config["ExportUserDefinedPsets"] = True

        # Сохранить в Revit ExtensibleStorage
        success, error = save_ifc_config_to_storage(config)

        if success:
            msg = "Конфигурация '{}' создана в Revit!\n\n".format(config_name)
            msg += "Изменения относительно базовой '{}':\n".format(old_name)
            msg += "- Имя: {}\n".format(config_name)
            msg += "- Маппинг: {}\n\n".format(mapping_path or "(нет)")
            msg += "Конфигурация доступна в диалоге экспорта IFC."

            MessageBox.Show(
                msg,
                "Готово",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

            # Показать в output
            output.print_md("# Создана конфигурация IFC в Revit")
            output.print_md("---")
            output.print_md("**Имя:** {}".format(config_name))
            output.print_md("**Базовая:** {}".format(old_name))
            output.print_md("")
            output.print_md("## Изменённые поля:")
            output.print_md("- **Name:** {} -> {}".format(old_name, config_name))
            output.print_md("- **ExportUserDefinedPsetsFileName:** {} -> {}".format(
                old_mapping or "(пусто)", mapping_path or "(пусто)"
            ))
            output.print_md("")
            output.print_md("Конфигурация сохранена в документе Revit.")

            # Обновить список конфигураций
            self.load_configs()
        else:
            MessageBox.Show(
                "Ошибка создания конфигурации:\n{}".format(error),
                "Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

    def on_close(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===

if __name__ == "__main__":
    if doc is None:
        forms.alert("Откройте документ Revit", title="Ошибка")
    else:
        form = IFCConfigForm()
        form.ShowDialog()
