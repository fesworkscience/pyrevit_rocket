# -*- coding: utf-8 -*-
"""
CPSK Categories - список категорий Revit для работы с типами и экземплярами.
"""

from Autodesk.Revit.DB import BuiltInCategory


# Все поддерживаемые категории для работы с типами
ALL_TYPE_CATEGORIES = [
    # Архитектура
    BuiltInCategory.OST_Walls,
    BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Roofs,
    BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_Stairs,
    BuiltInCategory.OST_StairsRailing,
    BuiltInCategory.OST_Ramps,
    BuiltInCategory.OST_Columns,
    BuiltInCategory.OST_Curtain_Panels,
    # Конструкции
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_Rebar,
    # Трубы
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_FlexPipeCurves,
    BuiltInCategory.OST_PipeInsulations,
    # Воздуховоды
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_DuctInsulations,
    BuiltInCategory.OST_DuctTerminal,
    # Электрика
    BuiltInCategory.OST_CableTray,
    BuiltInCategory.OST_CableTrayFitting,
    BuiltInCategory.OST_Conduit,
    BuiltInCategory.OST_ConduitFitting,
    BuiltInCategory.OST_ElectricalEquipment,
    BuiltInCategory.OST_ElectricalFixtures,
    BuiltInCategory.OST_LightingFixtures,
    # Оборудование
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_Sprinklers,
    # Прочее
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_SpecialityEquipment,
    BuiltInCategory.OST_Furniture,
    BuiltInCategory.OST_FurnitureSystems,
]
