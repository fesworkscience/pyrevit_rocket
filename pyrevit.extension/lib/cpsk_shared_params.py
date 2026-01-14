# -*- coding: utf-8 -*-
"""
Модуль для работы с Shared Parameters для Smart Openings.

Управляет параметром CPSK_RebarCutData - JSON массив UniqueId отверстий.
"""

import os
import json
import codecs

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Transaction,
    ExternalDefinitionCreationOptions
)

# Совместимость с разными версиями Revit
try:
    # Revit 2022+
    from Autodesk.Revit.DB import SpecTypeId, GroupTypeId
    USE_NEW_API = True
except ImportError:
    # Revit 2021 и ранее
    from Autodesk.Revit.DB import ParameterType, BuiltInParameterGroup
    USE_NEW_API = False

# Имя параметра
REBAR_CUT_DATA_PARAM = "CPSK_RebarCutData"


def get_shared_param_file_path():
    """Получить путь к файлу Shared Parameters."""
    lib_dir = os.path.dirname(__file__)
    extension_dir = os.path.dirname(lib_dir)
    return os.path.join(extension_dir, "CPSK_SharedParams.txt")


def ensure_shared_param_file(app):
    """
    Убедиться что файл Shared Parameters существует и настроен.

    Args:
        app: Revit Application

    Returns:
        str: Путь к файлу
    """
    file_path = get_shared_param_file_path()

    # Создать файл если не существует
    if not os.path.exists(file_path):
        with codecs.open(file_path, 'w', 'utf-8') as f:
            f.write("# CPSK Shared Parameters\n")
            f.write("# Do not edit manually.\n")
            f.write("*META\tVERSION\tMINVERSION\n")
            f.write("META\t2\t1\n")
            f.write("*GROUP\tID\tNAME\n")
            f.write("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\n")

    # Установить как текущий файл параметров
    app.SharedParametersFilename = file_path

    return file_path


def get_or_create_shared_param(doc, app, force_recreate=False):
    """
    Получить или создать Shared Parameter CPSK_RebarCutData.

    Args:
        doc: Revit Document
        app: Revit Application
        force_recreate: Пересоздать параметр если он скрыт

    Returns:
        Definition или None
    """
    # Убедиться что файл существует
    ensure_shared_param_file(app)

    # Открыть файл параметров
    def_file = app.OpenSharedParameterFile()
    if def_file is None:
        return None

    # Найти или создать группу
    group = None
    for g in def_file.Groups:
        if g.Name == "CPSK_Rebar":
            group = g
            break

    if group is None:
        group = def_file.Groups.Create("CPSK_Rebar")

    # Найти существующее определение параметра
    definition = None
    for d in group.Definitions:
        if d.Name == REBAR_CUT_DATA_PARAM:
            definition = d
            break

    # Проверить видимость существующего параметра
    if definition is not None:
        if not definition.Visible and force_recreate:
            # Параметр скрыт - нужно пересоздать
            # Удаляем старое определение из файла ФОП
            # К сожалению, API не позволяет удалить определение напрямую
            # Но мы можем пересоздать файл без этого параметра
            definition = None  # Будет создан заново ниже
        elif definition.Visible:
            return definition
        else:
            # Параметр скрыт, но force_recreate=False
            return definition

    if definition is None:
        # Создать новое определение
        try:
            # Revit 2022+
            options = ExternalDefinitionCreationOptions(REBAR_CUT_DATA_PARAM, SpecTypeId.String.Text)
        except Exception:
            # Fallback для старых версий
            options = ExternalDefinitionCreationOptions(REBAR_CUT_DATA_PARAM, ParameterType.Text)

        options.Visible = True  # Видимый для пользователя
        options.UserModifiable = False  # Но не редактируемый вручную
        options.Description = "JSON array of opening UniqueIds for Smart Openings"
        definition = group.Definitions.Create(options)

    return definition


def bind_param_to_rebar_category(doc, app, definition):
    """
    Привязать параметр к категории Structural Rebar.

    Создаёт параметр проекта (Project Parameter) типа Instance
    в группе Data.

    Args:
        doc: Revit Document
        app: Revit Application
        definition: Definition параметра

    Returns:
        bool: Успех операции
    """
    # Проверить, не привязан ли уже
    binding_map = doc.ParameterBindings
    iterator = binding_map.ForwardIterator()
    iterator.Reset()

    while iterator.MoveNext():
        if iterator.Key.Name == REBAR_CUT_DATA_PARAM:
            # Уже привязан
            return True

    # Создать CategorySet через app.Create (правильный способ!)
    cat_set = app.Create.NewCategorySet()

    # Добавить категорию Rebar
    rebar_cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Rebar)
    if rebar_cat is None:
        return False
    cat_set.Insert(rebar_cat)

    # Также добавим категорию Area Reinforcement и Path Reinforcement для полноты
    try:
        area_rein_cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_AreaRein)
        if area_rein_cat is not None:
            cat_set.Insert(area_rein_cat)
    except Exception:
        pass

    try:
        path_rein_cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_PathRein)
        if path_rein_cat is not None:
            cat_set.Insert(path_rein_cat)
    except Exception:
        pass

    # Создать Instance binding через app.Create (правильный способ!)
    binding = app.Create.NewInstanceBinding(cat_set)

    # Привязать с группой Data
    try:
        # Revit 2022+
        return binding_map.Insert(definition, binding, GroupTypeId.Data)
    except Exception:
        # Fallback для старых версий
        return binding_map.Insert(definition, binding, BuiltInParameterGroup.PG_DATA)


def ensure_rebar_cut_param(doc, app):
    """
    Убедиться что параметр CPSK_RebarCutData существует и привязан.

    Должен вызываться внутри транзакции!

    Args:
        doc: Revit Document
        app: Revit Application

    Returns:
        bool: Успех операции
    """
    definition = get_or_create_shared_param(doc, app)
    if definition is None:
        return False

    return bind_param_to_rebar_category(doc, app, definition)


def get_rebar_cut_data(rebar):
    """
    Получить список UniqueId отверстий из параметра стержня.

    Args:
        rebar: Rebar element

    Returns:
        list: Список UniqueId строк или пустой список
    """
    param = rebar.LookupParameter(REBAR_CUT_DATA_PARAM)
    if param is None:
        return []

    value = param.AsString()
    if not value:
        return []

    try:
        data = json.loads(value)
        if isinstance(data, list):
            return data
        return []
    except Exception:
        return []


def set_rebar_cut_data(rebar, opening_guids):
    """
    Записать список UniqueId отверстий в параметр стержня.

    Args:
        rebar: Rebar element
        opening_guids: list of UniqueId strings

    Returns:
        bool: Успех операции
    """
    param = rebar.LookupParameter(REBAR_CUT_DATA_PARAM)
    if param is None:
        # Попробуем найти параметр по всем параметрам элемента
        for p in rebar.Parameters:
            if p.Definition.Name == REBAR_CUT_DATA_PARAM:
                param = p
                break

    if param is None:
        # Логируем для отладки
        import codecs
        log_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "log_param.log")
        try:
            with codecs.open(log_path, 'a', 'utf-8') as f:
                f.write("set_rebar_cut_data: param NOT FOUND on rebar {}\n".format(rebar.Id.IntegerValue))
                f.write("  Available params: ")
                param_names = [p.Definition.Name for p in rebar.Parameters]
                f.write(", ".join(param_names[:20]) + "\n")
        except Exception:
            pass
        return False

    if not opening_guids:
        # Очистить параметр
        param.Set("")
        return True

    try:
        value = json.dumps(opening_guids)
        param.Set(value)
        return True
    except Exception as e:
        # Логируем ошибку
        import codecs
        log_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "log_param.log")
        try:
            with codecs.open(log_path, 'a', 'utf-8') as f:
                f.write("set_rebar_cut_data: ERROR setting param on rebar {}: {}\n".format(
                    rebar.Id.IntegerValue, str(e)))
        except Exception:
            pass
        return False


def add_opening_to_rebar(rebar, opening_guid):
    """
    Добавить UniqueId отверстия в список стержня.

    Args:
        rebar: Rebar element
        opening_guid: UniqueId отверстия

    Returns:
        bool: Успех операции
    """
    current = get_rebar_cut_data(rebar)
    if opening_guid not in current:
        current.append(opening_guid)
    return set_rebar_cut_data(rebar, current)


def remove_opening_from_rebar(rebar, opening_guid):
    """
    Удалить UniqueId отверстия из списка стержня.

    Args:
        rebar: Rebar element
        opening_guid: UniqueId отверстия

    Returns:
        bool: Успех операции
    """
    current = get_rebar_cut_data(rebar)
    if opening_guid in current:
        current.remove(opening_guid)
    return set_rebar_cut_data(rebar, current)


def find_rebars_by_opening_guid(doc, opening_guid):
    """
    Найти все стержни, связанные с данным отверстием.

    Args:
        doc: Revit Document
        opening_guid: UniqueId отверстия

    Returns:
        list: Список Rebar элементов
    """
    result = []

    collector = FilteredElementCollector(doc).OfCategory(
        BuiltInCategory.OST_Rebar
    ).WhereElementIsNotElementType()

    for rebar in collector:
        guids = get_rebar_cut_data(rebar)
        if opening_guid in guids:
            result.append(rebar)

    return result


def check_shared_param_exists(doc):
    """
    Проверить существует ли параметр CPSK_RebarCutData в проекте.

    Args:
        doc: Revit Document

    Returns:
        tuple (exists: bool, is_bound: bool)
    """
    # Проверить привязку в документе
    binding_map = doc.ParameterBindings
    iterator = binding_map.ForwardIterator()
    iterator.Reset()

    while iterator.MoveNext():
        if iterator.Key.Name == REBAR_CUT_DATA_PARAM:
            return True, True

    return False, False


def check_param_visibility(app):
    """
    Проверить видимость параметра в файле ФОП.

    Args:
        app: Revit Application

    Returns:
        tuple (exists_in_file: bool, is_visible: bool)
    """
    file_path = get_shared_param_file_path()
    if not os.path.exists(file_path):
        return False, False

    # Установить файл и открыть
    app.SharedParametersFilename = file_path
    def_file = app.OpenSharedParameterFile()
    if def_file is None:
        return False, False

    # Найти группу
    for g in def_file.Groups:
        if g.Name == "CPSK_Rebar":
            # Найти параметр
            for d in g.Definitions:
                if d.Name == REBAR_CUT_DATA_PARAM:
                    return True, d.Visible

    return False, False


def remove_param_from_file():
    """
    Удалить параметр CPSK_RebarCutData из файла ФОП.

    Это нужно для пересоздания параметра с правильной видимостью.

    Returns:
        bool: Успех операции
    """
    file_path = get_shared_param_file_path()
    if not os.path.exists(file_path):
        return True  # Файла нет - нечего удалять

    try:
        # Читаем файл
        with codecs.open(file_path, 'r', 'utf-8') as f:
            lines = f.readlines()

        # Фильтруем строки - убираем строку с нашим параметром
        new_lines = []
        for line in lines:
            # Строка параметра содержит имя параметра во втором столбце (после GUID)
            if REBAR_CUT_DATA_PARAM not in line:
                new_lines.append(line)

        # Записываем обратно
        with codecs.open(file_path, 'w', 'utf-8') as f:
            f.writelines(new_lines)

        return True
    except Exception:
        return False


def get_shared_param_info(doc, app):
    """
    Получить информацию о параметре ФОП.

    Args:
        doc: Revit Document
        app: Revit Application

    Returns:
        dict with keys:
            - exists: bool
            - is_bound: bool
            - file_path: str
            - visible: bool
            - exists_in_file: bool
            - message: str
    """
    info = {
        'exists': False,
        'is_bound': False,
        'file_path': get_shared_param_file_path(),
        'visible': True,
        'exists_in_file': False,
        'message': ''
    }

    # Проверить файл ФОП
    if not os.path.exists(info['file_path']):
        info['message'] = "Файл ФОП не существует, будет создан"
        return info

    # Проверить привязку в документе
    exists, is_bound = check_shared_param_exists(doc)
    info['exists'] = exists
    info['is_bound'] = is_bound

    # Проверить видимость в файле ФОП
    exists_in_file, is_visible = check_param_visibility(app)
    info['exists_in_file'] = exists_in_file
    info['visible'] = is_visible

    if is_bound:
        if not is_visible:
            info['message'] = "Параметр {} привязан, но СКРЫТ! Требуется пересоздание.".format(REBAR_CUT_DATA_PARAM)
        else:
            info['message'] = "Параметр {} привязан и видим".format(REBAR_CUT_DATA_PARAM)
    elif exists_in_file:
        if not is_visible:
            info['message'] = "Параметр {} скрыт в файле ФОП. Требуется пересоздание.".format(REBAR_CUT_DATA_PARAM)
        else:
            info['message'] = "Параметр {} есть в файле, но не привязан".format(REBAR_CUT_DATA_PARAM)
    else:
        info['message'] = "Параметр {} не существует, будет создан".format(REBAR_CUT_DATA_PARAM)

    return info


def ensure_rebar_cut_param_with_info(doc, app, fix_visibility=True, force_rebind=False):
    """
    Убедиться что параметр существует и вернуть информацию.

    Должен вызываться внутри транзакции!

    Args:
        doc: Revit Document
        app: Revit Application
        fix_visibility: Исправить видимость если параметр скрыт
        force_rebind: Принудительно пересоздать привязку (для смены группы параметров)

    Returns:
        tuple (success: bool, message: str, was_created: bool)
    """
    was_created = False

    # Проверить существует ли уже и привязан ли
    exists, is_bound = check_shared_param_exists(doc)

    # Проверить видимость в файле ФОП
    exists_in_file, is_visible = check_param_visibility(app)

    # Принудительная перепривязка - удаляем старую привязку
    if force_rebind and is_bound:
        binding_map = doc.ParameterBindings
        iterator = binding_map.ForwardIterator()
        iterator.Reset()
        key_to_remove = None
        while iterator.MoveNext():
            if iterator.Key.Name == REBAR_CUT_DATA_PARAM:
                key_to_remove = iterator.Key
                break
        if key_to_remove is not None:
            binding_map.Remove(key_to_remove)
        is_bound = False

    # Если параметр привязан но скрыт - это проблема
    if is_bound and exists_in_file and not is_visible:
        if fix_visibility:
            # Удалить привязку из документа
            binding_map = doc.ParameterBindings
            iterator = binding_map.ForwardIterator()
            iterator.Reset()
            key_to_remove = None
            while iterator.MoveNext():
                if iterator.Key.Name == REBAR_CUT_DATA_PARAM:
                    key_to_remove = iterator.Key
                    break
            if key_to_remove is not None:
                binding_map.Remove(key_to_remove)

            # Удалить параметр из файла ФОП
            if not remove_param_from_file():
                return False, "Не удалось удалить скрытый параметр из файла ФОП", False

            # Теперь создаём заново с правильной видимостью
            is_bound = False
            exists_in_file = False
        else:
            return False, "Параметр {} скрыт! Требуется пересоздание.".format(REBAR_CUT_DATA_PARAM), False

    # Если параметр скрыт в файле (но не привязан) - удалить и пересоздать
    if exists_in_file and not is_visible:
        if not remove_param_from_file():
            return False, "Не удалось удалить скрытый параметр из файла ФОП", False
        exists_in_file = False

    # Если всё ок - возвращаем успех
    if is_bound and is_visible and not force_rebind:
        return True, "Параметр {} уже существует и привязан (видимый)".format(REBAR_CUT_DATA_PARAM), False

    # Создать/привязать
    definition = get_or_create_shared_param(doc, app, force_recreate=True)
    if definition is None:
        return False, "Не удалось создать параметр в файле ФОП", False

    # Проверить что параметр видимый
    if not definition.Visible:
        return False, "Параметр создан но скрыт! Удалите файл ФОП вручную.", False

    success = bind_param_to_rebar_category(doc, app, definition)
    if success:
        was_created = True
        return True, "Параметр {} создан и привязан к категории Rebar (видимый)".format(REBAR_CUT_DATA_PARAM), True
    else:
        return False, "Не удалось привязать параметр к категории Rebar", False
