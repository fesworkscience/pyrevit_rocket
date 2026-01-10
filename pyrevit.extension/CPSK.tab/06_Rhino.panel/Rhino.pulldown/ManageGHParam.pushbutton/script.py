# -*- coding: utf-8 -*-
"""Управление shared параметром CPSK_gh_временный_элемент_будет_удален."""

__title__ = "GH\nParam"
__author__ = "CPSK"

# 1. Сначала import clr и стандартные модули
import clr
import os
import sys
import codecs

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

# 2. System и WinForms импорты
import System
from System.Windows.Forms import (
    Form, Label, Button, Panel, CheckedListBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, OpenFileDialog, SaveFileDialog,
    AnchorStyles, CheckBox, Padding, GroupBox,
    MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Point, Size, Font, FontStyle, Color

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
    BuiltInCategory,
    Transaction,
    CategorySet,
    InstanceBinding,
    TypeBinding,
    ExternalDefinitionCreationOptions,
    Category,
    SpecTypeId,
    GroupTypeId
)

doc = revit.doc
app = doc.Application

# Константы
PARAM_NAME = "CPSK_gh_временный_элемент_будет_удален"
PARAM_GROUP_NAME = "CPSK"
SHARED_PARAM_FILENAME = "CPSK_SharedParameters.txt"


def get_default_shared_param_path():
    """Получить путь по умолчанию для файла общих параметров."""
    if doc.PathName:
        return os.path.join(os.path.dirname(doc.PathName), SHARED_PARAM_FILENAME)
    return os.path.join(os.path.expanduser("~"), "Documents", SHARED_PARAM_FILENAME)


def create_shared_param_file(filepath):
    """Создать новый файл общих параметров."""
    try:
        # Создаём пустой файл с правильной структурой
        with codecs.open(filepath, 'w', 'utf-8') as f:
            f.write("# This is a Revit shared parameter file.\n")
            f.write("# Do not edit manually.\n")
            f.write("*META\tVERSION\tMINVERSION\n")
            f.write("META\t2\t1\n")
            f.write("*GROUP\tID\tNAME\n")
            f.write("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\n")
        return True
    except Exception as e:
        show_error("Ошибка", "Не удалось создать файл параметров", details=str(e))
        return False


def get_or_create_shared_param_file(filepath):
    """Получить или создать файл общих параметров."""
    if not os.path.exists(filepath):
        if not create_shared_param_file(filepath):
            return None

    try:
        app.SharedParametersFilename = filepath
        return app.OpenSharedParameterFile()
    except Exception as e:
        show_error("Ошибка", "Не удалось открыть файл параметров", details=str(e))
        return None


def get_or_create_param_group(shared_file, group_name):
    """Получить или создать группу параметров."""
    # Ищем существующую группу
    for group in shared_file.Groups:
        if group.Name == group_name:
            return group

    # Создаём новую группу
    return shared_file.Groups.Create(group_name)


def get_or_create_param_definition(param_group, param_name):
    """Получить или создать определение параметра."""
    # Ищем существующий параметр
    for definition in param_group.Definitions:
        if definition.Name == param_name:
            return definition

    # Создаём новый параметр (YesNo = Boolean)
    # В Revit 2022+ используем SpecTypeId вместо ParameterType
    try:
        # Revit 2022+
        options = ExternalDefinitionCreationOptions(param_name, SpecTypeId.Boolean.YesNo)
    except:
        # Fallback для старых версий
        from Autodesk.Revit.DB import ParameterType
        options = ExternalDefinitionCreationOptions(param_name, ParameterType.YesNo)

    options.Description = "Временный элемент для Grasshopper (будет удалён)"
    options.UserModifiable = True
    options.Visible = True

    return param_group.Definitions.Create(options)


def get_model_categories():
    """Получить категории модели с количеством элементов."""
    categories_info = []

    # Список категорий для работы (проверенные для Revit 2024)
    target_categories = [
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_StructuralFoundation,
        BuiltInCategory.OST_Floors,
        BuiltInCategory.OST_Walls,
        BuiltInCategory.OST_Roofs,
        BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_Columns,
        BuiltInCategory.OST_Stairs,
        BuiltInCategory.OST_Ramps,
        BuiltInCategory.OST_Railings,
        BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_Windows,
        BuiltInCategory.OST_Furniture,
        BuiltInCategory.OST_Casework,
        BuiltInCategory.OST_Ceilings,
        BuiltInCategory.OST_CurtainWallPanels,
        BuiltInCategory.OST_CurtainWallMullions,
        BuiltInCategory.OST_MechanicalEquipment,
        BuiltInCategory.OST_PlumbingFixtures,
        BuiltInCategory.OST_ElectricalEquipment,
        BuiltInCategory.OST_ElectricalFixtures,
        BuiltInCategory.OST_LightingFixtures,
        BuiltInCategory.OST_SpecialityEquipment,
        BuiltInCategory.OST_Entourage,
        BuiltInCategory.OST_Parking,
        BuiltInCategory.OST_Planting,
        BuiltInCategory.OST_Site,
        BuiltInCategory.OST_Topography,
        BuiltInCategory.OST_Mass,
        BuiltInCategory.OST_Parts,
        BuiltInCategory.OST_Rebar,
        BuiltInCategory.OST_FabricAreas,
        BuiltInCategory.OST_StructuralTruss,
        BuiltInCategory.OST_StructuralStiffener,
        BuiltInCategory.OST_StructConnections,
    ]

    for bic in target_categories:
        try:
            category = Category.GetCategory(doc, bic)
            if category is None:
                continue

            # Считаем элементы
            collector = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
            count = collector.GetElementCount()

            if count > 0 or True:  # Показываем все категории, даже пустые
                categories_info.append({
                    "category": category,
                    "bic": bic,
                    "name": category.Name,
                    "count": count
                })
        except:
            pass

    # Сортируем по имени
    categories_info.sort(key=lambda x: x["name"])
    return categories_info


def check_param_in_category(category):
    """Проверить, добавлен ли параметр в категорию."""
    try:
        binding_map = doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        iterator.Reset()

        while iterator.MoveNext():
            definition = iterator.Key
            if definition.Name == PARAM_NAME:
                binding = iterator.Current
                if hasattr(binding, 'Categories'):
                    for cat in binding.Categories:
                        if cat.Id.IntegerValue == category.Id.IntegerValue:
                            return True
        return False
    except:
        return False


def get_param_status():
    """Получить общий статус параметра в проекте."""
    try:
        binding_map = doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        iterator.Reset()

        while iterator.MoveNext():
            definition = iterator.Key
            if definition.Name == PARAM_NAME:
                return "added"
        return "not_added"
    except:
        return "unknown"


def add_param_to_categories(param_definition, categories):
    """Добавить параметр к категориям."""
    try:
        # Создаём CategorySet
        cat_set = app.Create.NewCategorySet()
        for cat_info in categories:
            cat_set.Insert(cat_info["category"])

        # Создаём InstanceBinding
        binding = app.Create.NewInstanceBinding(cat_set)

        # Проверяем, существует ли уже привязка
        binding_map = doc.ParameterBindings
        existing_definition = None

        iterator = binding_map.ForwardIterator()
        iterator.Reset()
        while iterator.MoveNext():
            if iterator.Key.Name == PARAM_NAME:
                existing_definition = iterator.Key
                break

        if existing_definition:
            # Обновляем существующую привязку
            # Получаем текущие категории и добавляем новые
            current_binding = binding_map.get_Item(existing_definition)
            if current_binding and hasattr(current_binding, 'Categories'):
                for cat in current_binding.Categories:
                    cat_set.Insert(cat)

            new_binding = app.Create.NewInstanceBinding(cat_set)
            return binding_map.ReInsert(existing_definition, new_binding, GroupTypeId.Data)
        else:
            # Добавляем новую привязку
            return binding_map.Insert(param_definition, binding, GroupTypeId.Data)

    except Exception as e:
        show_error("Ошибка", "Не удалось добавить параметр", details=str(e))
        return False


def remove_param_from_categories(categories):
    """Удалить параметр из указанных категорий."""
    try:
        binding_map = doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        iterator.Reset()

        existing_definition = None
        current_binding = None

        while iterator.MoveNext():
            if iterator.Key.Name == PARAM_NAME:
                existing_definition = iterator.Key
                current_binding = iterator.Current
                break

        if not existing_definition or not current_binding:
            return True  # Параметр не существует

        # Получаем категории для удаления
        remove_cat_ids = set(cat_info["category"].Id.IntegerValue for cat_info in categories)

        # Создаём новый CategorySet без удаляемых категорий
        new_cat_set = app.Create.NewCategorySet()
        remaining_count = 0

        if hasattr(current_binding, 'Categories'):
            for cat in current_binding.Categories:
                if cat.Id.IntegerValue not in remove_cat_ids:
                    new_cat_set.Insert(cat)
                    remaining_count += 1

        if remaining_count == 0:
            # Удаляем параметр полностью
            return binding_map.Remove(existing_definition)
        else:
            # Обновляем привязку с оставшимися категориями
            new_binding = app.Create.NewInstanceBinding(new_cat_set)
            return binding_map.ReInsert(existing_definition, new_binding, GroupTypeId.Data)

    except Exception as e:
        show_error("Ошибка", "Не удалось удалить параметр", details=str(e))
        return False


class ManageParamForm(Form):
    """Форма управления параметром."""

    def __init__(self, categories_info, shared_param_path):
        self.categories_info = categories_info
        self.shared_param_path = shared_param_path
        self.result_action = None
        self.selected_categories = []
        self.setup_form()

    def setup_form(self):
        self.Text = "Управление параметром Grasshopper"
        self.Width = 550
        self.Height = 550
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MinimumSize = Size(450, 400)

        # Верхняя панель - информация
        top_panel = Panel()
        top_panel.Dock = DockStyle.Top
        top_panel.Height = 100
        top_panel.Padding = Padding(10)

        lbl_param = Label()
        lbl_param.Text = "Параметр: {}".format(PARAM_NAME)
        lbl_param.Location = Point(10, 10)
        lbl_param.Width = 500
        lbl_param.Font = Font(lbl_param.Font, FontStyle.Bold)

        # Статус параметра
        status = get_param_status()
        status_text = "Добавлен в проект" if status == "added" else "Не добавлен"
        status_color = Color.Green if status == "added" else Color.Red

        self.lbl_status = Label()
        self.lbl_status.Text = "Статус: {}".format(status_text)
        self.lbl_status.Location = Point(10, 35)
        self.lbl_status.Width = 300
        self.lbl_status.ForeColor = status_color

        lbl_file = Label()
        lbl_file.Text = "Файл ФОП: {}".format(os.path.basename(self.shared_param_path))
        lbl_file.Location = Point(10, 55)
        lbl_file.Width = 400

        btn_change_file = Button()
        btn_change_file.Text = "Изменить..."
        btn_change_file.Location = Point(420, 50)
        btn_change_file.Width = 90
        btn_change_file.Click += self.on_change_file

        top_panel.Controls.Add(lbl_param)
        top_panel.Controls.Add(self.lbl_status)
        top_panel.Controls.Add(lbl_file)
        top_panel.Controls.Add(btn_change_file)

        # Средняя панель - список категорий
        middle_panel = Panel()
        middle_panel.Dock = DockStyle.Fill
        middle_panel.Padding = Padding(10)

        lbl_categories = Label()
        lbl_categories.Text = "Категории (выберите для добавления/удаления параметра):"
        lbl_categories.Dock = DockStyle.Top
        lbl_categories.Height = 20

        # Кнопки выбора
        select_panel = Panel()
        select_panel.Dock = DockStyle.Top
        select_panel.Height = 30

        btn_select_all = Button()
        btn_select_all.Text = "Выбрать все"
        btn_select_all.Location = Point(0, 2)
        btn_select_all.Width = 100
        btn_select_all.Click += self.on_select_all

        btn_select_none = Button()
        btn_select_none.Text = "Снять все"
        btn_select_none.Location = Point(110, 2)
        btn_select_none.Width = 100
        btn_select_none.Click += self.on_select_none

        btn_select_with_elements = Button()
        btn_select_with_elements.Text = "С элементами"
        btn_select_with_elements.Location = Point(220, 2)
        btn_select_with_elements.Width = 100
        btn_select_with_elements.Click += self.on_select_with_elements

        select_panel.Controls.Add(btn_select_all)
        select_panel.Controls.Add(btn_select_none)
        select_panel.Controls.Add(btn_select_with_elements)

        # Список категорий
        self.checklist = CheckedListBox()
        self.checklist.Dock = DockStyle.Fill
        self.checklist.CheckOnClick = True

        for cat_info in self.categories_info:
            has_param = check_param_in_category(cat_info["category"])
            status_mark = "[+]" if has_param else "[-]"
            display = "{} {} ({} эл.)".format(status_mark, cat_info["name"], cat_info["count"])
            self.checklist.Items.Add(display)

        middle_panel.Controls.Add(self.checklist)
        middle_panel.Controls.Add(select_panel)
        middle_panel.Controls.Add(lbl_categories)

        # Нижняя панель - кнопки действий
        bottom_panel = Panel()
        bottom_panel.Dock = DockStyle.Bottom
        bottom_panel.Height = 80
        bottom_panel.Padding = Padding(10)

        lbl_legend = Label()
        lbl_legend.Text = "[+] = параметр добавлен, [-] = параметр отсутствует"
        lbl_legend.Location = Point(10, 5)
        lbl_legend.Width = 350
        lbl_legend.ForeColor = Color.Gray

        btn_add = Button()
        btn_add.Text = "Добавить параметр"
        btn_add.Location = Point(10, 35)
        btn_add.Width = 150
        btn_add.Height = 30
        btn_add.Click += self.on_add

        btn_remove = Button()
        btn_remove.Text = "Удалить параметр"
        btn_remove.Location = Point(170, 35)
        btn_remove.Width = 150
        btn_remove.Height = 30
        btn_remove.Click += self.on_remove

        btn_cancel = Button()
        btn_cancel.Text = "Закрыть"
        btn_cancel.Location = Point(420, 35)
        btn_cancel.Width = 90
        btn_cancel.Height = 30
        btn_cancel.Click += self.on_cancel

        bottom_panel.Controls.Add(lbl_legend)
        bottom_panel.Controls.Add(btn_add)
        bottom_panel.Controls.Add(btn_remove)
        bottom_panel.Controls.Add(btn_cancel)

        # Добавляем контролы (Fill первым!)
        self.Controls.Add(middle_panel)
        self.Controls.Add(bottom_panel)
        self.Controls.Add(top_panel)

    def get_selected_categories(self):
        """Получить выбранные категории."""
        selected = []
        for i in range(self.checklist.Items.Count):
            if self.checklist.GetItemChecked(i):
                selected.append(self.categories_info[i])
        return selected

    def on_select_all(self, sender, args):
        for i in range(self.checklist.Items.Count):
            self.checklist.SetItemChecked(i, True)

    def on_select_none(self, sender, args):
        for i in range(self.checklist.Items.Count):
            self.checklist.SetItemChecked(i, False)

    def on_select_with_elements(self, sender, args):
        for i in range(self.checklist.Items.Count):
            has_elements = self.categories_info[i]["count"] > 0
            self.checklist.SetItemChecked(i, has_elements)

    def on_change_file(self, sender, args):
        dialog = OpenFileDialog()
        dialog.Title = "Выберите файл общих параметров"
        dialog.Filter = "Shared Parameters (*.txt)|*.txt|All files (*.*)|*.*"
        dialog.InitialDirectory = os.path.dirname(self.shared_param_path)

        if dialog.ShowDialog() == DialogResult.OK:
            self.shared_param_path = dialog.FileName
            # Обновляем метку
            for ctrl in self.Controls:
                if isinstance(ctrl, Panel):
                    for subctrl in ctrl.Controls:
                        if isinstance(subctrl, Label) and "Файл ФОП:" in subctrl.Text:
                            subctrl.Text = "Файл ФОП: {}".format(os.path.basename(self.shared_param_path))

    def on_add(self, sender, args):
        self.selected_categories = self.get_selected_categories()
        if not self.selected_categories:
            show_warning("Выбор", "Выберите хотя бы одну категорию")
            return

        self.result_action = "add"
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_remove(self, sender, args):
        self.selected_categories = self.get_selected_categories()
        if not self.selected_categories:
            show_warning("Выбор", "Выберите хотя бы одну категорию")
            return

        self.result_action = "remove"
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


def main():
    """Основная функция."""
    # Получаем путь к файлу общих параметров
    shared_param_path = get_default_shared_param_path()

    # Если файл не существует и нет текущего файла - предлагаем создать или выбрать
    current_shared_file = None
    try:
        current_shared_file = app.OpenSharedParameterFile()
    except:
        pass

    if current_shared_file is None and not os.path.exists(shared_param_path):
        # Предлагаем создать новый или выбрать существующий
        dialog = SaveFileDialog()
        dialog.Title = "Создать или выбрать файл общих параметров"
        dialog.Filter = "Shared Parameters (*.txt)|*.txt"
        dialog.FileName = SHARED_PARAM_FILENAME
        dialog.InitialDirectory = os.path.dirname(shared_param_path) if os.path.dirname(shared_param_path) else os.path.expanduser("~")

        if dialog.ShowDialog() != DialogResult.OK:
            return

        shared_param_path = dialog.FileName
    elif current_shared_file is not None:
        # Используем текущий файл
        shared_param_path = app.SharedParametersFilename

    # Собираем категории
    categories_info = get_model_categories()

    if not categories_info:
        show_warning("Нет данных", "Не найдено категорий для добавления параметра")
        return

    # Показываем форму
    form = ManageParamForm(categories_info, shared_param_path)
    result = form.ShowDialog()

    if result != DialogResult.OK:
        return

    action = form.result_action
    selected_categories = form.selected_categories
    shared_param_path = form.shared_param_path

    if not selected_categories:
        return

    # Выполняем действие
    with revit.Transaction("CPSK: {} параметр GH".format("Добавить" if action == "add" else "Удалить")):
        if action == "add":
            # Получаем или создаём файл параметров
            shared_file = get_or_create_shared_param_file(shared_param_path)
            if shared_file is None:
                return

            # Получаем или создаём группу
            param_group = get_or_create_param_group(shared_file, PARAM_GROUP_NAME)
            if param_group is None:
                show_error("Ошибка", "Не удалось создать группу параметров")
                return

            # Получаем или создаём определение параметра
            param_def = get_or_create_param_definition(param_group, PARAM_NAME)
            if param_def is None:
                show_error("Ошибка", "Не удалось создать определение параметра")
                return

            # Добавляем параметр к категориям
            if add_param_to_categories(param_def, selected_categories):
                cat_names = ", ".join(c["name"] for c in selected_categories[:5])
                if len(selected_categories) > 5:
                    cat_names += " и ещё {}".format(len(selected_categories) - 5)
                show_success(
                    "Параметр добавлен",
                    "Добавлен в {} категорий".format(len(selected_categories)),
                    details="Категории: {}".format(cat_names)
                )
            else:
                show_error("Ошибка", "Не удалось добавить параметр")

        elif action == "remove":
            if remove_param_from_categories(selected_categories):
                cat_names = ", ".join(c["name"] for c in selected_categories[:5])
                if len(selected_categories) > 5:
                    cat_names += " и ещё {}".format(len(selected_categories) - 5)
                show_success(
                    "Параметр удалён",
                    "Удалён из {} категорий".format(len(selected_categories)),
                    details="Категории: {}".format(cat_names)
                )
            else:
                show_error("Ошибка", "Не удалось удалить параметр")


if __name__ == "__main__":
    main()
