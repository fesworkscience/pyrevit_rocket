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
    "IFCCOUNTMEASURE": "INTEGER",
    "IFCNUMERICMEASURE": "NUMBER",

    # Единицы измерения
    "IFCLENGTHMEASURE": "LENGTH",
    "IFCAREAMEASURE": "AREA",
    "IFCVOLUMEMEASURE": "VOLUME",
    "IFCPOSITIVELENGTHMEASURE": "LENGTH",
    "IFCPLANEANGLEMEASURE": "ANGLE",
    "IFCNONNEGATIVELENGTHMEASURE": "LENGTH",
    "IFCPOSITIVEINTEGER": "INTEGER",

    # Физические величины
    "IFCMASSMEASURE": "NUMBER",
    "IFCFORCEMEASURE": "NUMBER",
    "IFCPRESSUREMEASURE": "NUMBER",
    "IFCTHERMALTRANSMITTANCEMEASURE": "NUMBER",
    "IFCPOWERMEASURE": "NUMBER",
    "IFCELECTRICCURRENTMEASURE": "NUMBER",
    "IFCELECTRICVOLTAGEMEASURE": "NUMBER",
    "IFCFREQUENCYMEASURE": "NUMBER",
    "IFCVOLUMETRICFLOWRATEMEASURE": "NUMBER",
    "IFCMASSFLOWRATEMEASURE": "NUMBER",
    "IFCTHERMODYNAMICTEMPERATUREMEASURE": "NUMBER",
    "IFCTEMPERATURERATEOFCHANGEMEASURE": "NUMBER",
    "IFCLUMINOUSFLUXMEASURE": "NUMBER",
    "IFCILLUMINANCEMEASURE": "NUMBER",
}

# Маппинг IFC сущностей на категории Revit
# Используется для определения категорий при импорте/экспорте
# Значение - название категории на английском (Revit internal name)
IFC_TO_REVIT_CATEGORY = {
    # ============================================
    # АРХИТЕКТУРА
    # ============================================

    # Стены
    "IFCWALL": "Walls",
    "IFCWALLSTANDARDCASE": "Walls",
    "IFCWALLELEMENTEDCASE": "Walls",
    "IFCWALLTYPE": "Walls",

    # Перекрытия и полы
    "IFCSLAB": "Floors",
    "IFCSLABSTANDARDCASE": "Floors",
    "IFCSLABELEMENTEDCASE": "Floors",
    "IFCSLABTYPE": "Floors",

    # Покрытия (полы, потолки, изоляция)
    "IFCCOVERING": "Floors",  # Может быть также Ceilings в зависимости от типа
    "IFCCOVERINGTYPE": "Floors",

    # Кровля
    "IFCROOF": "Roofs",
    "IFCROOFTYPE": "Roofs",

    # Потолки
    "IFCCEILING": "Ceilings",

    # Навесные стены
    "IFCCURTAINWALL": "Curtain Walls",
    "IFCCURTAINWALLTYPE": "Curtain Walls",

    # Двери
    "IFCDOOR": "Doors",
    "IFCDOORSTANDARDCASE": "Doors",
    "IFCDOORTYPE": "Doors",

    # Окна
    "IFCWINDOW": "Windows",
    "IFCWINDOWSTANDARDCASE": "Windows",
    "IFCWINDOWTYPE": "Windows",

    # Лестницы
    "IFCSTAIR": "Stairs",
    "IFCSTAIRFLIGHT": "Stairs",
    "IFCSTAIRFLIGHTTYPE": "Stairs",
    "IFCSTAIRTYPE": "Stairs",

    # Пандусы
    "IFCRAMP": "Ramps",
    "IFCRAMPFLIGHT": "Ramps",
    "IFCRAMPFLIGHTTYPE": "Ramps",
    "IFCRAMPTYPE": "Ramps",

    # Ограждения
    "IFCRAILING": "Railings",
    "IFCRAILINGTYPE": "Railings",

    # Мебель
    "IFCFURNITURE": "Furniture",
    "IFCFURNITURETYPE": "Furniture",

    # Пространства и зоны
    "IFCSPACE": "Rooms",
    "IFCSPACETYPE": "Rooms",
    "IFCSPATIALZONE": "Areas",
    "IFCBUILDINGSTOREY": "Levels",

    # ============================================
    # НЕСУЩИЕ КОНСТРУКЦИИ
    # ============================================

    # Колонны
    "IFCCOLUMN": "Structural Columns",
    "IFCCOLUMNSTANDARDCASE": "Structural Columns",
    "IFCCOLUMNTYPE": "Structural Columns",

    # Балки и каркас
    "IFCBEAM": "Structural Framing",
    "IFCBEAMSTANDARDCASE": "Structural Framing",
    "IFCBEAMTYPE": "Structural Framing",
    "IFCMEMBER": "Structural Framing",
    "IFCMEMBERSTANDARDCASE": "Structural Framing",
    "IFCMEMBERTYPE": "Structural Framing",
    "IFCPLATE": "Structural Framing",
    "IFCPLATESTANDARDCASE": "Structural Framing",
    "IFCPLATETYPE": "Structural Framing",

    # Фундаменты
    "IFCFOOTING": "Structural Foundations",
    "IFCFOOTINGTYPE": "Structural Foundations",
    "IFCPILE": "Structural Foundations",
    "IFCPILETYPE": "Structural Foundations",

    # Армирование
    "IFCREINFORCINGBAR": "Structural Rebar",
    "IFCREINFORCINGBARTYPE": "Structural Rebar",
    "IFCREINFORCINGMESH": "Structural Rebar",
    "IFCREINFORCINGMESHTYPE": "Structural Rebar",
    "IFCREINFORCINGELEMENT": "Structural Rebar",
    "IFCTENDON": "Structural Rebar",
    "IFCTENDONTYPE": "Structural Rebar",
    "IFCTENDONANCHOR": "Structural Rebar",
    "IFCTENDONANCHORTYPE": "Structural Rebar",
    "IFCREINFORCEDSOIL": "Structural Foundations",

    # Соединения
    "IFCMECHANICALFASTENER": "Structural Connections",
    "IFCMECHANICALFASTENERTYPE": "Structural Connections",
    "IFCDISCRETEACCESSORY": "Structural Connections",
    "IFCDISCRETEACCESSORYTYPE": "Structural Connections",

    # ============================================
    # ИНЖЕНЕРНЫЕ СИСТЕМЫ - ТРУБОПРОВОДЫ
    # ============================================

    # Трубы
    "IFCPIPESEGMENT": "Pipes",
    "IFCPIPESEGMENTTYPE": "Pipes",

    # Фитинги труб
    "IFCPIPEFITTING": "Pipe Fittings",
    "IFCPIPEFITTINGTYPE": "Pipe Fittings",

    # Арматура (клапаны)
    "IFCVALVE": "Pipe Accessories",
    "IFCVALVETYPE": "Pipe Accessories",

    # Насосы
    "IFCPUMP": "Mechanical Equipment",
    "IFCPUMPTYPE": "Mechanical Equipment",

    # Баки/ёмкости
    "IFCTANK": "Mechanical Equipment",
    "IFCTANKTYPE": "Mechanical Equipment",

    # Котлы
    "IFCBOILER": "Mechanical Equipment",
    "IFCBOILERTYPE": "Mechanical Equipment",

    # Теплообменники
    "IFCHEATEXCHANGER": "Mechanical Equipment",
    "IFCHEATEXCHANGERTYPE": "Mechanical Equipment",

    # Перехватчики (жироуловители, песколовки)
    "IFCINTERCEPTOR": "Plumbing Fixtures",
    "IFCINTERCEPTORTYPE": "Plumbing Fixtures",

    # Расходомеры
    "IFCFLOWMETER": "Pipe Accessories",
    "IFCFLOWMETERTYPE": "Pipe Accessories",
    "IFCFLOWINSTRUMENT": "Pipe Accessories",
    "IFCFLOWINSTRUMENTTYPE": "Pipe Accessories",

    # ============================================
    # ИНЖЕНЕРНЫЕ СИСТЕМЫ - ВОЗДУХОВОДЫ
    # ============================================

    # Воздуховоды
    "IFCDUCTSEGMENT": "Ducts",
    "IFCDUCTSEGMENTTYPE": "Ducts",

    # Фитинги воздуховодов
    "IFCDUCTFITTING": "Duct Fittings",
    "IFCDUCTFITTINGTYPE": "Duct Fittings",

    # Глушители
    "IFCDUCTSILENCER": "Duct Accessories",
    "IFCDUCTSILENCERTYPE": "Duct Accessories",

    # Воздухораспределители
    "IFCAIRTERMINAL": "Air Terminals",
    "IFCAIRTERMINALTYPE": "Air Terminals",
    "IFCAIRTERMINALBOX": "Air Terminals",
    "IFCAIRTERMINALBOXTYPE": "Air Terminals",

    # Вентиляторы
    "IFCFAN": "Mechanical Equipment",
    "IFCFANTYPE": "Mechanical Equipment",

    # Фильтры
    "IFCFILTER": "Duct Accessories",
    "IFCFILTERTYPE": "Duct Accessories",

    # Заслонки/демпферы
    "IFCDAMPER": "Duct Accessories",
    "IFCDAMPERTYPE": "Duct Accessories",

    # Увлажнители
    "IFCHUMIDIFIER": "Mechanical Equipment",
    "IFCHUMIDIFIERTYPE": "Mechanical Equipment",

    # Теплообменные элементы
    "IFCCOIL": "Mechanical Equipment",
    "IFCCOILTYPE": "Mechanical Equipment",
    "IFCCOOLEDBEAM": "Mechanical Equipment",
    "IFCCOOLEDBEAMTYPE": "Mechanical Equipment",

    # Рекуператоры
    "IFCAIRTOAIRHEATRECOVERY": "Mechanical Equipment",
    "IFCAIRTOAIRHEATRECOVERYTYPE": "Mechanical Equipment",

    # ============================================
    # ИНЖЕНЕРНЫЕ СИСТЕМЫ - ХОЛОДОСНАБЖЕНИЕ
    # ============================================

    # Чиллеры
    "IFCCHILLER": "Mechanical Equipment",
    "IFCCHILLERTYPE": "Mechanical Equipment",

    # Компрессоры
    "IFCCOMPRESSOR": "Mechanical Equipment",
    "IFCCOMPRESSORTYPE": "Mechanical Equipment",

    # Конденсаторы
    "IFCCONDENSER": "Mechanical Equipment",
    "IFCCONDENSERTYPE": "Mechanical Equipment",

    # Испарители
    "IFCEVAPORATOR": "Mechanical Equipment",
    "IFCEVAPORATORTYPE": "Mechanical Equipment",

    # Испарительные охладители
    "IFCEVAPORATIVECOOLER": "Mechanical Equipment",
    "IFCEVAPORATIVECOOLERTYPE": "Mechanical Equipment",

    # Трубчатые пучки
    "IFCTUBEBUNDLE": "Mechanical Equipment",
    "IFCTUBEBUNDLETYPE": "Mechanical Equipment",

    # ============================================
    # ИНЖЕНЕРНЫЕ СИСТЕМЫ - ГАЗОСНАБЖЕНИЕ/ТОПЛИВО
    # ============================================

    # Горелки
    "IFCBURNER": "Mechanical Equipment",
    "IFCBURNERTYPE": "Mechanical Equipment",

    # Двигатели (дизельгенераторы)
    "IFCENGINE": "Mechanical Equipment",
    "IFCENGINETYPE": "Mechanical Equipment",

    # ============================================
    # ИНЖЕНЕРНЫЕ СИСТЕМЫ - ЭЛЕКТРИКА
    # ============================================

    # Кабельные лотки
    "IFCCABLECARRIERSEGMENT": "Cable Tray",
    "IFCCABLECARRIERSEGMENTTYPE": "Cable Tray",

    # Фитинги кабельных лотков
    "IFCCABLECARRIERFITTING": "Cable Tray Fittings",
    "IFCCABLECARRIERFITTINGTYPE": "Cable Tray Fittings",

    # Кабели
    "IFCCABLESEGMENT": "Electrical Equipment",
    "IFCCABLESEGMENTTYPE": "Electrical Equipment",

    # Кабельные фитинги
    "IFCCABLEFITTING": "Electrical Equipment",
    "IFCCABLEFITTINGTYPE": "Electrical Equipment",

    # Электрощиты
    "IFCELECTRICDISTRIBUTIONBOARD": "Electrical Equipment",
    "IFCELECTRICDISTRIBUTIONBOARDTYPE": "Electrical Equipment",

    # Коробки соединительные
    "IFCJUNCTIONBOX": "Electrical Equipment",
    "IFCJUNCTIONBOXTYPE": "Electrical Equipment",

    # Светильники
    "IFCLIGHTFIXTURE": "Lighting Fixtures",
    "IFCLIGHTFIXTURETYPE": "Lighting Fixtures",

    # Розетки
    "IFCOUTLET": "Electrical Fixtures",
    "IFCOUTLETTYPE": "Electrical Fixtures",

    # Защитные устройства (автоматы)
    "IFCPROTECTIVEDEVICE": "Electrical Equipment",
    "IFCPROTECTIVEDEVICETYPE": "Electrical Equipment",

    # Выключатели
    "IFCSWITCHINGDEVICE": "Electrical Fixtures",
    "IFCSWITCHINGDEVICETYPE": "Electrical Fixtures",

    # Трансформаторы
    "IFCTRANSFORMER": "Electrical Equipment",
    "IFCTRANSFORMERTYPE": "Electrical Equipment",

    # Генераторы
    "IFCELECTRICGENERATOR": "Electrical Equipment",
    "IFCELECTRICGENERATORTYPE": "Electrical Equipment",

    # Электродвигатели
    "IFCELECTRICMOTOR": "Electrical Equipment",
    "IFCELECTRICMOTORTYPE": "Electrical Equipment",

    # Соединения двигателей
    "IFCMOTORCONNECTION": "Electrical Equipment",
    "IFCMOTORCONNECTIONTYPE": "Electrical Equipment",

    # Таймеры
    "IFCELECTRICTIMECONTROL": "Electrical Equipment",
    "IFCELECTRICTIMECONTROLTYPE": "Electrical Equipment",

    # Электроприборы
    "IFCELECTRICAPPLIANCE": "Electrical Equipment",
    "IFCELECTRICAPPLIANCETYPE": "Electrical Equipment",

    # Аккумуляторы/ИБП
    "IFCELECTRICFLOWSTORAGEDEVICE": "Electrical Equipment",
    "IFCELECTRICFLOWSTORAGEDEVICETYPE": "Electrical Equipment",

    # Солнечные панели
    "IFCSOLARDEVICE": "Electrical Equipment",
    "IFCSOLARDEVICETYPE": "Electrical Equipment",

    # ============================================
    # ИНЖЕНЕРНЫЕ СИСТЕМЫ - САНТЕХНИКА
    # ============================================

    # Сантехника (унитазы, раковины)
    "IFCSANITARYTERMINAL": "Plumbing Fixtures",
    "IFCSANITARYTERMINALTYPE": "Plumbing Fixtures",

    # Сливы/трапы
    "IFCWASTETERMINAL": "Plumbing Fixtures",
    "IFCWASTETERMINALTYPE": "Plumbing Fixtures",

    # Вентиляционные выводы канализации
    "IFCSTACKTERMINAL": "Plumbing Fixtures",
    "IFCSTACKTERMINALTYPE": "Plumbing Fixtures",

    # Противопожарное оборудование
    "IFCFIRESUPPRESSIONTERMINAL": "Sprinklers",
    "IFCFIRESUPPRESSIONTERMINALTYPE": "Sprinklers",

    # ============================================
    # АВТОМАТИКА И УПРАВЛЕНИЕ
    # ============================================

    # Приводы
    "IFCACTUATOR": "Mechanical Equipment",
    "IFCACTUATORTYPE": "Mechanical Equipment",

    # Датчики
    "IFCSENSOR": "Electrical Equipment",
    "IFCSENSORTYPE": "Electrical Equipment",

    # Контроллеры
    "IFCCONTROLLER": "Electrical Equipment",
    "IFCCONTROLLERTYPE": "Electrical Equipment",

    # Сигнализация
    "IFCALARM": "Fire Alarm Devices",
    "IFCALARMTYPE": "Fire Alarm Devices",

    # ============================================
    # СВЯЗЬ И КОММУНИКАЦИИ
    # ============================================

    # Коммуникационное оборудование
    "IFCCOMMUNICATIONSAPPLIANCE": "Communication Devices",
    "IFCCOMMUNICATIONSAPPLIANCETYPE": "Communication Devices",

    # Аудио-видео оборудование
    "IFCAUDIOVISUALAPPLIANCE": "Communication Devices",
    "IFCAUDIOVISUALAPPLIANCETYPE": "Communication Devices",

    # ============================================
    # ОБЩИЕ И СПЕЦИАЛЬНЫЕ
    # ============================================

    # Унитарное оборудование (кондиционеры, фанкойлы)
    "IFCUNITARYEQUIPMENT": "Mechanical Equipment",
    "IFCUNITARYEQUIPMENTTYPE": "Mechanical Equipment",

    # Распределительные камеры
    "IFCDISTRIBUTIONCHAMBERELEMENT": "Mechanical Equipment",
    "IFCDISTRIBUTIONCHAMBERELEMENTTYPE": "Mechanical Equipment",

    # Порты распределения
    "IFCDISTRIBUTIONPORT": "Generic Models",

    # Системы распределения
    "IFCDISTRIBUTIONSYSTEM": "Generic Models",

    # Транспортные элементы (лифты, эскалаторы)
    "IFCTRANSPORTELEMENT": "Specialty Equipment",
    "IFCTRANSPORTELEMENTTYPE": "Specialty Equipment",

    # Дымоходы
    "IFCCHIMNEY": "Generic Models",

    # Географические элементы (деревья, рельеф)
    "IFCGEOGRAPHICELEMENT": "Site",

    # Сборки
    "IFCELEMENTASSEMBLY": "Assemblies",
    "IFCELEMENTASSEMBLYTYPE": "Assemblies",

    # Прокси-элементы (обобщённые модели)
    "IFCBUILDINGELEMENTPROXY": "Generic Models",
    "IFCBUILDINGELEMENTPROXYTYPE": "Generic Models",

    # ============================================
    # ПРОЕКТ И СТРУКТУРА
    # ============================================

    # Проект и площадка
    "IFCBUILDING": "Project Information",
    "IFCSITE": "Project Information",
    "IFCSPATIALSTRUCTUREELEMENT": "Project Information",

    # Материалы
    "IFCMATERIAL": "Materials",
}

# Маппинг IFC сущностей на категории Revit с ID (BuiltInCategory int values)
# Формат: IFC_CLASS -> [(русское_имя, builtin_category_id), ...]
# Используется в скрипте FOPtoProject для добавления параметров
# ID категорий взяты из BuiltInCategory enum
IFC_TO_REVIT_CATEGORY_IDS = {
    # ============================================
    # АРХИТЕКТУРА
    # ============================================

    # Стены
    "IFCWALL": [("Стены", -2000011)],  # OST_Walls
    "IFCWALLSTANDARDCASE": [("Стены", -2000011)],
    "IFCWALLELEMENTEDCASE": [("Стены", -2000011)],
    "IFCWALLTYPE": [("Стены", -2000011)],

    # Перекрытия и полы
    "IFCSLAB": [("Перекрытия", -2000032)],  # OST_Floors
    "IFCSLABSTANDARDCASE": [("Перекрытия", -2000032)],
    "IFCSLABELEMENTEDCASE": [("Перекрытия", -2000032)],
    "IFCSLABTYPE": [("Перекрытия", -2000032)],

    # Покрытия (полы, потолки)
    "IFCCOVERING": [("Перекрытия", -2000032), ("Потолки", -2000038)],  # OST_Floors, OST_Ceilings
    "IFCCOVERINGTYPE": [("Перекрытия", -2000032), ("Потолки", -2000038)],

    # Кровля
    "IFCROOF": [("Крыши", -2000035)],  # OST_Roofs
    "IFCROOFTYPE": [("Крыши", -2000035)],

    # Потолки
    "IFCCEILING": [("Потолки", -2000038)],  # OST_Ceilings

    # Навесные стены
    "IFCCURTAINWALL": [("Витражи", -2000170)],  # OST_CurtainWallPanels
    "IFCCURTAINWALLTYPE": [("Витражи", -2000170)],

    # Двери
    "IFCDOOR": [("Двери", -2000023)],  # OST_Doors
    "IFCDOORSTANDARDCASE": [("Двери", -2000023)],
    "IFCDOORTYPE": [("Двери", -2000023)],

    # Окна
    "IFCWINDOW": [("Окна", -2000014)],  # OST_Windows
    "IFCWINDOWSTANDARDCASE": [("Окна", -2000014)],
    "IFCWINDOWTYPE": [("Окна", -2000014)],

    # Лестницы
    "IFCSTAIR": [("Лестницы", -2000120)],  # OST_Stairs
    "IFCSTAIRFLIGHT": [("Лестницы", -2000120)],
    "IFCSTAIRFLIGHTTYPE": [("Лестницы", -2000120)],
    "IFCSTAIRTYPE": [("Лестницы", -2000120)],

    # Пандусы
    "IFCRAMP": [("Пандусы", -2000180)],  # OST_Ramps
    "IFCRAMPFLIGHT": [("Пандусы", -2000180)],
    "IFCRAMPFLIGHTTYPE": [("Пандусы", -2000180)],
    "IFCRAMPTYPE": [("Пандусы", -2000180)],

    # Ограждения
    "IFCRAILING": [("Ограждения", -2000126)],  # OST_StairsRailing
    "IFCRAILINGTYPE": [("Ограждения", -2000126)],

    # Мебель
    "IFCFURNITURE": [("Мебель", -2000080)],  # OST_Furniture (правильный ID)
    "IFCFURNITURETYPE": [("Мебель", -2000080)],

    # Помещения
    "IFCSPACE": [("Помещения", -2000160)],  # OST_Rooms
    "IFCSPACETYPE": [("Помещения", -2000160)],
    "IFCSPATIALZONE": [("Зоны", -2003200)],  # OST_Areas (правильный ID)

    # ============================================
    # НЕСУЩИЕ КОНСТРУКЦИИ
    # ============================================

    # Колонны
    "IFCCOLUMN": [("Несущие колонны", -2001330)],  # OST_StructuralColumns
    "IFCCOLUMNSTANDARDCASE": [("Несущие колонны", -2001330)],
    "IFCCOLUMNTYPE": [("Несущие колонны", -2001330)],

    # Балки и каркас
    "IFCBEAM": [("Несущий каркас", -2001320)],  # OST_StructuralFraming
    "IFCBEAMSTANDARDCASE": [("Несущий каркас", -2001320)],
    "IFCBEAMTYPE": [("Несущий каркас", -2001320)],
    "IFCMEMBER": [("Несущий каркас", -2001320)],
    "IFCMEMBERSTANDARDCASE": [("Несущий каркас", -2001320)],
    "IFCMEMBERTYPE": [("Несущий каркас", -2001320)],
    "IFCPLATE": [("Несущий каркас", -2001320)],
    "IFCPLATESTANDARDCASE": [("Несущий каркас", -2001320)],
    "IFCPLATETYPE": [("Несущий каркас", -2001320)],

    # Фундаменты
    "IFCFOOTING": [("Несущие фундаменты", -2001300)],  # OST_StructuralFoundation
    "IFCFOOTINGTYPE": [("Несущие фундаменты", -2001300)],
    "IFCPILE": [("Несущие фундаменты", -2001300)],
    "IFCPILETYPE": [("Несущие фундаменты", -2001300)],

    # Армирование
    "IFCREINFORCINGBAR": [("Несущая арматура", -2009000)],  # OST_Rebar (правильный ID)
    "IFCREINFORCINGBARTYPE": [("Несущая арматура", -2009000)],
    "IFCREINFORCINGMESH": [("Несущая арматура", -2009000)],
    "IFCREINFORCINGMESHTYPE": [("Несущая арматура", -2009000)],
    "IFCREINFORCINGELEMENT": [("Несущая арматура", -2009000)],

    # Соединения
    "IFCMECHANICALFASTENER": [("Соединения несущих конструкций", -2009030)],  # OST_StructConnections (правильный ID)
    "IFCMECHANICALFASTENERTYPE": [("Соединения несущих конструкций", -2009030)],
    "IFCDISCRETEACCESSORY": [("Соединения несущих конструкций", -2009030)],
    "IFCDISCRETEACCESSORYTYPE": [("Соединения несущих конструкций", -2009030)],

    # ============================================
    # ИНЖЕНЕРНЫЕ СИСТЕМЫ - ТРУБОПРОВОДЫ
    # ============================================

    "IFCPIPESEGMENT": [("Трубы", -2008044)],  # OST_PipeCurves
    "IFCPIPESEGMENTTYPE": [("Трубы", -2008044)],
    "IFCPIPEFITTING": [("Соединительные детали трубопроводов", -2008055)],  # OST_PipeFitting
    "IFCPIPEFITTINGTYPE": [("Соединительные детали трубопроводов", -2008055)],
    "IFCVALVE": [("Арматура трубопроводов", -2008114)],  # OST_PipeAccessory
    "IFCVALVETYPE": [("Арматура трубопроводов", -2008114)],
    "IFCPUMP": [("Оборудование", -2001140)],  # OST_MechanicalEquipment
    "IFCPUMPTYPE": [("Оборудование", -2001140)],
    "IFCTANK": [("Оборудование", -2001140)],
    "IFCTANKTYPE": [("Оборудование", -2001140)],
    "IFCBOILER": [("Оборудование", -2001140)],
    "IFCBOILERTYPE": [("Оборудование", -2001140)],

    # ============================================
    # ИНЖЕНЕРНЫЕ СИСТЕМЫ - ВОЗДУХОВОДЫ
    # ============================================

    "IFCDUCTSEGMENT": [("Воздуховоды", -2008000)],  # OST_DuctCurves
    "IFCDUCTSEGMENTTYPE": [("Воздуховоды", -2008000)],
    "IFCDUCTFITTING": [("Соединительные детали воздуховодов", -2008010)],  # OST_DuctFitting
    "IFCDUCTFITTINGTYPE": [("Соединительные детали воздуховодов", -2008010)],
    "IFCAIRTERMINAL": [("Воздухораспределители", -2008013)],  # OST_DuctTerminal
    "IFCAIRTERMINALTYPE": [("Воздухораспределители", -2008013)],
    "IFCAIRTERMINALBOX": [("Воздухораспределители", -2008013)],
    "IFCAIRTERMINALBOXTYPE": [("Воздухораспределители", -2008013)],
    "IFCFAN": [("Оборудование", -2001140)],
    "IFCFANTYPE": [("Оборудование", -2001140)],
    "IFCFILTER": [("Арматура воздуховодов", -2008016)],  # OST_DuctAccessory
    "IFCFILTERTYPE": [("Арматура воздуховодов", -2008016)],
    "IFCDAMPER": [("Арматура воздуховодов", -2008016)],
    "IFCDAMPERTYPE": [("Арматура воздуховодов", -2008016)],

    # ============================================
    # ИНЖЕНЕРНЫЕ СИСТЕМЫ - ЭЛЕКТРИКА
    # ============================================

    "IFCCABLECARRIERSEGMENT": [("Кабельные лотки", -2008130)],  # OST_CableTray
    "IFCCABLECARRIERSEGMENTTYPE": [("Кабельные лотки", -2008130)],
    "IFCCABLECARRIERFITTING": [("Соединительные детали кабельных лотков", -2008126)],  # OST_CableTrayFitting (правильный ID)
    "IFCCABLECARRIERFITTINGTYPE": [("Соединительные детали кабельных лотков", -2008126)],
    "IFCCABLESEGMENT": [("Электрооборудование", -2001040)],  # OST_ElectricalEquipment (правильный ID)
    "IFCCABLESEGMENTTYPE": [("Электрооборудование", -2001040)],
    "IFCLIGHTFIXTURE": [("Осветительные приборы", -2001120)],  # OST_LightingFixtures
    "IFCLIGHTFIXTURETYPE": [("Осветительные приборы", -2001120)],
    "IFCOUTLET": [("Электрооборудование", -2001040)],
    "IFCOUTLETTYPE": [("Электрооборудование", -2001040)],
    "IFCSWITCHINGDEVICE": [("Электрооборудование", -2001040)],
    "IFCSWITCHINGDEVICETYPE": [("Электрооборудование", -2001040)],
    "IFCELECTRICDISTRIBUTIONBOARD": [("Электрооборудование", -2001040)],
    "IFCELECTRICDISTRIBUTIONBOARDTYPE": [("Электрооборудование", -2001040)],
    "IFCJUNCTIONBOX": [("Электрооборудование", -2001040)],
    "IFCJUNCTIONBOXTYPE": [("Электрооборудование", -2001040)],

    # ============================================
    # ИНЖЕНЕРНЫЕ СИСТЕМЫ - САНТЕХНИКА
    # ============================================

    "IFCSANITARYTERMINAL": [("Сантехнические приборы", -2001160)],  # OST_PlumbingFixtures (правильный ID)
    "IFCSANITARYTERMINALTYPE": [("Сантехнические приборы", -2001160)],
    "IFCWASTETERMINAL": [("Сантехнические приборы", -2001160)],
    "IFCWASTETERMINALTYPE": [("Сантехнические приборы", -2001160)],
    "IFCFIRESUPPRESSIONTERMINAL": [("Спринклеры", -2008099)],  # OST_Sprinklers (правильный ID)
    "IFCFIRESUPPRESSIONTERMINALTYPE": [("Спринклеры", -2008099)],

    # ============================================
    # АВТОМАТИКА И УПРАВЛЕНИЕ
    # ============================================

    "IFCACTUATOR": [("Оборудование", -2001140)],
    "IFCACTUATORTYPE": [("Оборудование", -2001140)],
    "IFCSENSOR": [("Электрооборудование", -2001040)],
    "IFCSENSORTYPE": [("Электрооборудование", -2001040)],
    "IFCCONTROLLER": [("Электрооборудование", -2001040)],
    "IFCCONTROLLERTYPE": [("Электрооборудование", -2001040)],
    "IFCALARM": [("Пожарная сигнализация", -2008085)],  # OST_FireAlarmDevices (правильный ID)
    "IFCALARMTYPE": [("Пожарная сигнализация", -2008085)],

    # ============================================
    # ОБЩИЕ И СПЕЦИАЛЬНЫЕ
    # ============================================

    # Сборки
    "IFCELEMENTASSEMBLY": [("Сборки", -2000267)],  # OST_Assemblies (правильный ID)
    "IFCELEMENTASSEMBLYTYPE": [("Сборки", -2000267)],

    # Обобщённые модели
    "IFCBUILDINGELEMENTPROXY": [("Обобщённые модели", -2000151)],  # OST_GenericModel
    "IFCBUILDINGELEMENTPROXYTYPE": [("Обобщённые модели", -2000151)],

    # ============================================
    # ПРОЕКТ И СТРУКТУРА
    # ============================================

    # Проект и площадка (параметры проекта)
    "IFCBUILDING": [("Сведения о проекте", -2003101)],  # OST_ProjectInformation (правильный ID)
    "IFCSITE": [("Сведения о проекте", -2003101)],
    "IFCSPATIALSTRUCTUREELEMENT": [("Сведения о проекте", -2003101)],

    # Уровни
    "IFCBUILDINGSTOREY": [("Уровни", -2000240)],  # OST_Levels

    # Материалы
    "IFCMATERIAL": [("Материалы", -2000700)],  # OST_Materials
}
