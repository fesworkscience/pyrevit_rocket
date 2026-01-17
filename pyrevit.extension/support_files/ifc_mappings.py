# -*- coding: utf-8 -*-
"""
IFC Mappings - Словари для маппинга IFC типов и сущностей на Revit.

Эти словари используются в командах для конвертации между IFC и Revit форматами:
- IDS в ФОП (01_IDStoFOP)
- ФОП в проект (02_FOPtoProject)
- Экспорт IFC и другие команды

Использование:
    from support_files.ifc_mappings import IFC_TO_REVIT_TYPE, IFC_TO_REVIT_CATEGORY
"""

# Маппинг IFC типов данных на типы параметров Revit
# Используется при генерации ФОП файла из IDS
IFC_TO_REVIT_TYPE = {
    # Текстовые типы
    "IFCTEXT": "TEXT",
    "IFCLABEL": "TEXT",
    "IFCIDENTIFIER": "TEXT",

    # Логические типы
    "IFCBOOLEAN": "YESNO",
    "IFCLOGICAL": "YESNO",

    # Числовые типы
    "IFCINTEGER": "INTEGER",
    "IFCREAL": "NUMBER",

    # Единицы измерения
    "IFCLENGTHMEASURE": "LENGTH",
    "IFCAREAMEASURE": "AREA",
    "IFCVOLUMEMEASURE": "VOLUME",
    "IFCPOSITIVELENGTHMEASURE": "LENGTH",
    "IFCPLANEANGLEMEASURE": "ANGLE",

    # Физические величины
    "IFCMASSMEASURE": "NUMBER",
    "IFCFORCEMEASURE": "NUMBER",
    "IFCPRESSUREMEASURE": "NUMBER",
}

# Маппинг IFC сущностей на категории Revit
# Используется для определения категорий при импорте/экспорте
IFC_TO_REVIT_CATEGORY = {
    # Стены
    "IFCWALL": "Walls",
    "IFCWALLSTANDARDCASE": "Walls",

    # Перекрытия и полы
    "IFCSLAB": "Floors",
    "IFCCOVERING": "Floors",

    # Несущие конструкции
    "IFCCOLUMN": "Structural Columns",
    "IFCBEAM": "Structural Framing",
    "IFCMEMBER": "Structural Framing",
    "IFCPLATE": "Structural Framing",

    # Фундаменты
    "IFCFOOTING": "Structural Foundations",
    "IFCPILE": "Structural Foundations",

    # Лестницы и пандусы
    "IFCSTAIR": "Stairs",
    "IFCSTAIRFLIGHT": "Stairs",
    "IFCRAMP": "Ramps",
    "IFCRAMPFLIGHT": "Ramps",

    # Ограждения
    "IFCRAILING": "Railings",

    # Кровля
    "IFCROOF": "Roofs",

    # Проект
    "IFCBUILDING": "Project Information",
    "IFCSITE": "Project Information",

    # Материалы
    "IFCMATERIAL": "Materials",

    # Армирование
    "IFCREINFORCINGBAR": "Structural Rebar",
    "IFCREINFORCINGMESH": "Structural Rebar",
    "IFCREINFORCINGELEMENT": "Structural Rebar",

    # Сборки и прочее
    "IFCELEMENTASSEMBLY": "Assemblies",
    "IFCBUILDINGELEMENTPROXY": "Generic Models",

    # Соединения
    "IFCMECHANICALFASTENER": "Structural Connections",
    "IFCDISCRETEACCESSORY": "Structural Connections",
}
