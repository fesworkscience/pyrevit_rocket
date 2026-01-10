# CPSK Tools - pyRevit Extension

pyRevit extension for industrial building automation in Revit.

## РЕЛИЗ (ОБЯЗАТЕЛЬНО!)

**НИКОГДА не создавай теги вручную!** Используй скрипт релиза:

```bash
# Patch версия (1.0.35 -> 1.0.36) - по умолчанию
powershell -ExecutionPolicy Bypass -File "release_scripts\release.ps1"

# Minor версия (1.0.35 -> 1.1.0)
powershell -ExecutionPolicy Bypass -File "release_scripts\release.ps1" -BumpType minor

# Major версия (1.0.35 -> 2.0.0)
powershell -ExecutionPolicy Bypass -File "release_scripts\release.ps1" -BumpType major
```

Скрипт автоматически:
1. Проверяет, что нет незакоммиченных изменений
2. Определяет последнюю версию из тегов
3. Увеличивает версию
4. Обновляет `version.yaml`
5. Создаёт коммит "Release vX.X.X"
6. Создаёт тег
7. Пушит в origin и github

**Перед релизом обязательно закоммить все изменения!**

### GitHub Secrets для автоматической загрузки на сервер

В GitHub репозитории нужно настроить секрет (Settings → Secrets → Actions → Repository secrets):

- `RELEASE_API_TOKEN` - JWT токен для API

**Получить токен локально:**
```powershell
powershell -ExecutionPolicy Bypass -File "get_token.ps1"
```

При релизе GitHub Actions автоматически:
1. Собирает installer
2. Создаёт GitHub Release
3. Загружает exe на `https://rocket-tools.ru/api/rocketrevit/releases/upload/`

## Структура IDS файла (Information Delivery Specification)

IDS файл содержит требования к параметрам IFC элементов.

### Структура XML:
```xml
<specification name="Таблица 4.А.11 – Имена атрибутов для IfcFooting">
  <applicability>
    <entity>
      <name>
        <xs:enumeration value="IFCFOOTING" />
      </name>
      <predefinedType>
        <xs:enumeration value="PAD_FOOTING" />
        <xs:enumeration value="STRIP_FOOTING" />
      </predefinedType>
    </entity>
  </applicability>
  <requirements>
    <property cardinality="required" dataType="IFCTEXT">
      <propertySet>
        <simpleValue>Местоположение</simpleValue>
      </propertySet>
      <baseName>
        <simpleValue>Номер корпуса</simpleValue>
      </baseName>
    </property>
  </requirements>
</specification>
```

### ВАЖНО - Терминология IDS:
- **applicability/entity/name** - тип IFC (IFCFOOTING, IFCCOLUMN, etc.)
- **applicability/entity/predefinedType** - подтип IFC (PAD_FOOTING, STRIP_FOOTING, etc.)
- **requirements/property** - требуемые параметры (в документе ЦГЭ называются "атрибутами")
- **property/baseName** - имя параметра (НЕ name, а именно baseName!)
- **property/propertySet** - группа параметров (PropertySet)
- **property/@cardinality** - обязательность (required/optional)

### Парсинг IDS:
```python
# Получить имя параметра из property
baseName_node = prop_node.SelectSingleNode("ids:baseName", nsm)
param_name = baseName_node.InnerText  # "Номер корпуса"

# Получить predefinedType из applicability
predType_node = spec_node.SelectSingleNode("ids:applicability/ids:entity/ids:predefinedType", nsm)
```

## ОБЯЗАТЕЛЬНО: Проверка кода после написания!

После написания или изменения любого pyRevit скрипта ВСЕГДА запускай универсальный чекер.

### Команда запуска (ВАЖНО - использовать абсолютные пути!)

```bash
"C:\ProgramData\miniconda3\python.exe" "C:\Users\feduloves\Documents\web\pyrevit_rocket\pyrevit.extension\lib\pyrevit_checker.py" "C:\Users\feduloves\Documents\web\pyrevit_rocket\pyrevit.extension\CPSK.tab\QA.panel\IDStoFOP.pushbutton\script.py"
```

Или для всей папки:
```bash
"C:\ProgramData\miniconda3\python.exe" "C:\Users\feduloves\Documents\web\pyrevit_rocket\pyrevit.extension\lib\pyrevit_checker.py" "C:\Users\feduloves\Documents\web\pyrevit_rocket\pyrevit.extension\CPSK.tab"
```

### ВАЖНО: Проблема с os.path.exists() на Windows

На этой системе `os.path.exists()` в Python НЕ РАБОТАЕТ для некоторых путей (особенно Desktop и пути с кириллицей). Чекер использует Windows API (`GetFileAttributesW`) через ctypes для корректной работы.

Если чекер выдаёт "не существует" для существующего файла - это баг Python, не чекера.

**Все скрипты должны проходить проверку без ошибок!**

### Чекер проверяет:
- f-строки (запрещены)
- walrus operator `:=` (запрещён)
- `open(encoding=)` (использовать `codecs.open()`)
- `Application.Run()` (использовать `ShowDialog()`)
- `async/await`, `yield from`, `nonlocal` (запрещены)
- расширенная распаковка `*rest` (запрещена)
- отсутствие `require_auth()` в скриптах кнопок
- неправильная работа с `cpsk_settings.yaml` (использовать `cpsk_config`)
- использование `MessageBox.Show` или `forms.alert` (использовать `cpsk_notify`)
- использование `output.print_md` для пользовательских сообщений (использовать `cpsk_notify`)
- `subprocess.Popen/check_output` для Python/pip без `env=` (использовать `get_clean_env()`)

## Проверка авторизации (КРИТИЧНО!)

Все кнопки (кроме Login) должны проверять авторизацию в начале скрипта.
Если пользователь не залогинен, скрипт должен показать предупреждение и завершиться.

### ПРАВИЛО: Добавляй проверку авторизации в каждый script.py!

```python
# -*- coding: utf-8 -*-
"""Описание скрипта."""

__title__ = "Название"
__author__ = "CPSK"

import os
import sys

# ... импорты clr, WinForms, pyrevit ...

# Добавляем lib в путь для импорта cpsk_auth
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Проверка авторизации
from cpsk_auth import require_auth
if not require_auth():
    sys.exit()

# ... остальной код скрипта ...
```

### Исключения:
- `Login.pushbutton` - не требует проверки (это сама кнопка входа)
- Вспомогательные модули (не `script.py`) - не проверяются

## Уведомления cpsk_notify (ОБЯЗАТЕЛЬНО!)

Все уведомления в проекте должны использовать единый модуль `cpsk_notify`.

**ЗАПРЕЩЕНО использовать:**
- `MessageBox.Show` - системные диалоги
- `forms.alert` - pyRevit алерты
- `output.print_md` для пользовательских сообщений (только для отладки!)

**ПРАВИЛЬНО:** Все сообщения пользователю через `cpsk_notify`!

### Импорт:
```python
from cpsk_notify import show_error, show_warning, show_info, show_success, show_toast
```

### Типы уведомлений:

```python
# Ошибка (красный) - для критических ошибок
show_error("Заголовок", "Сообщение", details="Подробности...")

# Предупреждение (оранжевый) - для предупреждений
show_warning("Внимание", "Сообщение", details="Подробности...")

# Информация (синий) - для информационных сообщений
show_info("Информация", "Сообщение")

# Успех (зелёный) - для подтверждения успешных операций
show_success("Готово", "Операция завершена")
```

### Параметры:

```python
show_error(
    title,          # Заголовок
    message,        # Краткое сообщение
    details=None,   # Детали (разворачиваются кнопкой)
    blocking=True,  # True=модальное окно, False=тост в углу
    auto_close=0    # Секунды до автозакрытия (0=не закрывать)
)
```

### Примеры использования:

```python
# Блокирующая ошибка с деталями (ждёт OK)
show_error(
    "Ошибка загрузки",
    "Не удалось загрузить файл",
    details="Путь: C:\\...\\file.txt\nОшибка: File not found"
)

# Неблокирующий тост успеха (автозакрытие через 3 сек)
show_success("Сохранено", "Файл успешно сохранён", blocking=False, auto_close=3)

# Или через удобную обёртку show_toast
show_toast("CPSK", "Готово к работе", notification_type="success", auto_close=5)
```

### Замена старого кода:

```python
# НЕПРАВИЛЬНО - запрещено!
MessageBox.Show("Сообщение", "Заголовок", MessageBoxButtons.OK, MessageBoxIcon.Error)
forms.alert("Сообщение", title="Заголовок")
output.print_md("## Ошибка: что-то пошло не так")  # Только для отладки!

# ПРАВИЛЬНО - используй cpsk_notify!
show_error("Заголовок", "Сообщение")
show_info("Заголовок", "Сообщение")
show_warning("Внимание", "Предупреждение")
show_success("Готово", "Операция завершена")
```

### Когда использовать output.print_md:
- Только для отладочного вывода (DEBUG режим)
- Для вывода длинных списков/таблиц данных
- Никогда для ошибок, предупреждений и пользовательских сообщений!

## ВАЖНО: Ошибка при Reload pyRevit

При использовании кнопки "Reload" в pyRevit может появиться ошибка:

```
AttributeError: 'NoneType' object has no attribute 'Add'
File "...\pyrevit\telemetry\record.py", line 35, in __setattr__
```

**Это известный баг pyRevit с телеметрией, НЕ ошибка нашего кода!**

**Решение:** Полностью перезапустите Revit (закройте и откройте заново). НЕ используйте кнопку Reload pyRevit.

## Quick Start

1. Install pyRevit (https://github.com/pyrevitlabs/pyRevit)
2. Add this folder to pyRevit extensions paths
3. Reload pyRevit or restart Revit
4. Find "CPSK" tab in Revit ribbon

## Project Structure

```
pyrevit.extension/
├── extension.json           # Extension metadata (CPython 3)
├── CPSK.tab/                 # Main ribbon tab
│   ├── Columns.panel/        # Column/grid tools
│   │   ├── CreateColumns.pushbutton/
│   │   ├── CreateAxes.pushbutton/
│   │   └── CreateLevels.pushbutton/
│   ├── Trusses.panel/        # Truss/framing tools
│   │   ├── CreateTruss.pushbutton/
│   │   ├── CreateCapitals.pushbutton/
│   │   └── CreateRoofBeams.pushbutton/
│   ├── 02_Dynamo.panel/      # Dynamo script launcher
│   │   ├── _config.yaml      # User config (recent, favorites) - в .gitignore
│   │   ├── dynamo_scripts/   # Dynamo scripts library (1000+ supported)
│   │   │   ├── Examples/     # Category folders
│   │   │   ├── KM_Documentation/
│   │   │   └── ...
│   │   └── RunScript.pushbutton/
│   ├── Geometry.panel/       # Geometry utilities
│   │   └── BoundingBox.pushbutton/
│   └── Utils.panel/          # General utilities
│       ├── SelectElements.pushbutton/
│       └── GetParams.pushbutton/
└── lib/                      # Shared libraries
    ├── cpsk_utils.py         # Core Revit API utilities
    ├── cpsk_geometry.py      # Geometry extraction
    ├── cpsk_parameters.py    # Parameter get/set
    ├── cpsk_selection.py     # Selection utilities
    └── cpsk_dynamo.py        # Dynamo script utilities
```

## Shared Libraries (lib/)

Import in any script:
```python
from cpsk_utils import get_doc, collect_elements, mm_to_feet
from cpsk_geometry import get_bounding_box, get_location
from cpsk_parameters import get_param, set_param, get_all_params
from cpsk_selection import get_selected_elements, pick_element
```

### cpsk_utils.py
- `get_doc()` - Active Revit document
- `get_uidoc()` - Active UI document
- `collect_elements(category, of_class, view_id)` - FilteredElementCollector wrapper
- `mm_to_feet(mm)` / `feet_to_mm(feet)` - Unit conversion

### cpsk_geometry.py
- `get_bounding_box(element, view)` - Element bounding box
- `get_location(element)` - Location point or curve
- `get_geometry(element, detail_level)` - Geometry options
- `get_center(bbox)` - Bounding box center point

### cpsk_parameters.py
- `get_param(element, param_name)` - Get parameter value (auto-detects type)
- `set_param(element, param_name, value)` - Set parameter value
- `get_all_params(element)` - Dict of all parameters

### cpsk_selection.py
- `get_selected_elements()` - Currently selected elements
- `select_elements(elements)` - Set selection
- `pick_element(message)` - User pick single element
- `pick_elements(message)` - User pick multiple elements

### cpsk_dynamo.py
- `DynamoScanner(folder)` - Script scanner with caching
- `run_dynamo_script(path)` - Run .dyn script
- `parse_yaml_simple(path)` / `save_yaml_simple(path, data)` - Config utilities

## Dynamo Scripts System

Scalable system for managing 1000+ Dynamo scripts.

### Folder Structure

```
CPSK.tab/02_Dynamo.panel/
├── _config.yaml           # User config (recent, favorites) - в .gitignore
├── dynamo_scripts/        # Скрипты Dynamo
│   ├── Examples/          # Category folder
│   │   └── HelloWorld.dyn
│   ├── KM_Documentation/  # Another category
│   │   └── create_km_documentation.dyn
│   └── ...
└── RunScript.pushbutton/  # Кнопка запуска
```

### Adding Scripts

1. Create category folder in `02_Dynamo.panel/dynamo_scripts/`
2. Add `.dyn` files to category folder
3. Scripts auto-appear in launcher

### Script Metadata

Read from `.dyn` JSON: `Description`, `Author`, `Name`

### Launcher Features

- **Search**: Multi-term search (name + path)
- **Categories**: Tree view from folders
- **Recent**: Last 20 scripts (auto-tracked)
- **Favorites**: User-managed list
- **Pagination**: 100 scripts/page

### Using cpsk_dynamo.py

```python
from cpsk_dynamo import DynamoScanner, run_dynamo_script

scanner = DynamoScanner("path/to/dynamo_scripts")
scripts = scanner.get_scripts_in_category("Columns")
results = scanner.search_scripts("create column")
info = scanner.get_script_info(script_path)

success, msg = run_dynamo_script(script_path)
```

### Config File (_config.yaml)

```yaml
categories:
  Columns: "Columns"
  Export: "Export/Import"

recent:
  - "Columns/Create_Columns.dyn"

favorites:
  - "Columns/Create_Columns.dyn"
```

## Adding New Command

### 1. Create Pushbutton Folder

```
CPSK.tab/YourPanel.panel/YourCommand.pushbutton/
├── script.py      # Main script
├── bundle.yaml    # Button metadata
└── icon.png       # 32x32 or 16x16 icon
```

### 2. bundle.yaml Template

```yaml
title: "Кнопка\nНазвание"   # ВАЖНО: \n в двойных кавычках!
tooltip: Описание кнопки
author: CPSK
```

**ВАЖНО:** Для переноса строки `\n` нужны двойные кавычки!

### 3. script.py Template

```python
# -*- coding: utf-8 -*-
"""Script description."""

__title__ = "Button\nTitle"
__author__ = "CPSK"

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

doc = revit.doc
output = script.get_output()

# Your code here

with revit.Transaction("Transaction Name"):
    # Revit API operations
    pass

output.print_md("## Done")
```

## Windows Forms Pattern

For custom dialogs (see CreateColumns.pushbutton/script.py):

```python
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System.Windows.Forms as WinForms
import System.Drawing as Drawing

class MyForm(WinForms.Form):
    def __init__(self):
        self.result = None
        self.setup_form()

    def setup_form(self):
        self.Text = "Dialog Title"
        self.Width = 400
        self.Height = 300
        # Add controls...

    def on_ok(self, sender, args):
        self.result = {...}  # Collect values
        self.DialogResult = WinForms.DialogResult.OK
        self.Close()

# Usage
form = MyForm()
if form.ShowDialog() == WinForms.DialogResult.OK:
    params = form.result
    # Use params...
```

## Common Revit API Patterns

### Collect Elements by Category

```python
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

collector = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_StructuralColumns)\
    .WhereElementIsNotElementType()
elements = list(collector)
```

### Collect Element Types

```python
collector = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_StructuralColumns)\
    .WhereElementIsElementType()
types = list(collector)
```

### Create Structural Column

```python
from Autodesk.Revit.DB import XYZ
from Autodesk.Revit.DB.Structure import StructuralType

point = XYZ(x_feet, y_feet, z_feet)
column = doc.Create.NewFamilyInstance(
    point,
    column_type,
    level,
    StructuralType.Column
)
```

### Create Structural Beam

```python
from Autodesk.Revit.DB import Line

line = Line.CreateBound(start_pt, end_pt)
beam = doc.Create.NewFamilyInstance(
    line,
    beam_type,
    level,
    StructuralType.Beam
)
```

### Set Level Constraints

```python
# Top level constraint
top_param = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
if top_param:
    top_param.Set(top_level_id)

# Top offset
offset_param = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
if offset_param:
    offset_param.Set(0.0)
```

## Unit Conversion

Revit uses feet internally. Always convert:

```python
def mm_to_feet(mm):
    return mm / 304.8

def feet_to_mm(feet):
    return feet * 304.8

# Usage
x_feet = mm_to_feet(12000)  # 12000mm -> feet
point = XYZ(x_feet, y_feet, z_feet)
```

## Technical Notes

- **Python Engine:** IronPython 2.7 (even if extension.json says CPython 3!)
- **Transaction Management:** Use `with revit.Transaction("Name"):` for atomic operations
- **Output:** Use `script.get_output()` for formatted output with markdown
- **Forms:** Use `pyrevit.forms` for simple dialogs, WinForms for custom
- **Error Handling:** Wrap Revit API calls in try/except

## IronPython Compatibility (IMPORTANT!)

pyRevit uses IronPython 2.7, NOT CPython 3. Follow these rules:

### Forbidden Syntax
```python
# NO f-strings
f"Value: {x}"           # ERROR
"Value: {}".format(x)   # OK

# NO walrus operator
if (n := len(a)) > 10:  # ERROR
n = len(a)              # OK
if n > 10:

# NO type hints
def foo(x: int) -> str:  # ERROR
def foo(x):              # OK
```

### WinForms Imports
Always import `System` when using WinForms:
```python
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System  # REQUIRED for System.Windows.Forms.Padding etc.
from System.Windows.Forms import (
    Form, Panel, Button, Padding  # Import Padding explicitly
)
```

### Common IronPython Errors
- `NameError: global name 'System' is not defined` - Add `import System`
- `'module' object has no attribute` - Check clr.AddReference
- `SyntaxError` on f-strings - Use `.format()` instead
- `open() got an unexpected keyword argument 'encoding'` - Use `codecs.open()`:
```python
import codecs
with codecs.open(path, 'r', 'utf-8') as f:  # OK
    content = f.read()
# НЕ: open(path, 'r', encoding='utf-8')  # ERROR
```

### Subprocess и CPython (КРИТИЧНО!)

При запуске внешних Python-скриптов через `subprocess` из IronPython возникает ошибка:
```
Failed to import encodings module
```

**Причина:** IronPython устанавливает переменные окружения (`PYTHONHOME`, `PYTHONPATH`, `IRONPYTHONPATH`), которые конфликтуют с CPython.

**Решение:** Очищать окружение перед запуском subprocess:

```python
from cpsk_config import get_clean_env

# ПРАВИЛЬНО - с очищенным окружением
process = subprocess.Popen(
    ['python', '-m', 'pip', 'install', 'package'],
    env=get_clean_env()  # Убирает PYTHONHOME, PYTHONPATH, IRONPYTHONPATH
)

# НЕПРАВИЛЬНО - вызовет "Failed to import encodings"!
process = subprocess.Popen(['python', '-m', 'pip', 'install', 'package'])
```

**Функция `get_clean_env()` из `cpsk_config.py`:**
```python
def get_clean_env():
    """Получить очищенное окружение для subprocess."""
    env = os.environ.copy()
    env.pop('PYTHONHOME', None)
    env.pop('PYTHONPATH', None)
    env.pop('IRONPYTHONPATH', None)
    return env
```

**Применять для:** любых вызовов `python`, `pip`, `venv`, `virtualenv` через subprocess.

## Работа с кодировкой UTF-8 (КРИТИЧНО!)

IronPython 2.7 НЕ поддерживает параметр `encoding` в стандартной функции `open()`.
Это приводит к кракозябрам при чтении файлов с кириллицей или другими Unicode символами.

### ПРАВИЛО: Всегда используй `codecs.open()` для текстовых файлов!

```python
import codecs

# ПРАВИЛЬНО - чтение с кодировкой
with codecs.open(path, 'r', 'utf-8') as f:
    content = f.read()

# ПРАВИЛЬНО - запись с кодировкой
with codecs.open(path, 'w', 'utf-8') as f:
    f.write(content)

# НЕПРАВИЛЬНО - вызовет ошибку в IronPython!
with open(path, 'r', encoding='utf-8') as f:  # ERROR!
    content = f.read()

# НЕПРАВИЛЬНО - кракозябры при чтении кириллицы!
with open(path, 'r') as f:  # Кракозябры!
    content = f.read()
```

### Типичные случаи, требующие codecs.open():

1. **JSON файлы с кириллицей** (например, .dyn файлы Dynamo):
```python
import codecs
import json

with codecs.open(dyn_path, 'r', 'utf-8') as f:
    data = json.load(f)
description = data.get('Description', '')  # Корректная кириллица
```

2. **YAML/XML конфигурации**:
```python
with codecs.open(config_path, 'r', 'utf-8') as f:
    content = f.read()
```

3. **Любые текстовые файлы с не-ASCII символами**:
```python
with codecs.open(txt_path, 'r', 'utf-8') as f:
    lines = f.readlines()
```

### Признаки проблемы с кодировкой:
- Кракозябры вместо кириллицы: `ÐŸÑ€Ð¸Ð²ÐµÑ‚` вместо `Привет`
- `UnicodeDecodeError` при чтении файла
- Пустые или искажённые строки в UI

## WinForms в pyRevit (КРИТИЧНО!)

### ПОРЯДОК ИМПОРТОВ - КРИТИЧЕСКИ ВАЖЕН!

Для работы наследования от Form нужен СТРОГИЙ порядок импортов (как в модуле "Запуск Dynamo"):

```python
# -*- coding: utf-8 -*-
"""Описание скрипта."""

__title__ = "Название"
__author__ = "CPSK"

# 1. СНАЧАЛА import clr!
import clr
import os
import json
import codecs
# ... другие стандартные импорты

# 2. ЗАТЕМ clr.AddReference
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

# 3. ЗАТЕМ System и WinForms импорты
import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, Panel,
    DockStyle, FormStartPosition, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    DialogResult, OpenFileDialog, SaveFileDialog
)
from System.Drawing import Point, Size, Color, Font, FontStyle

# 4. ЗАТЕМ pyrevit
from pyrevit import revit, forms, script

# 5. ЗАТЕМ Revit API (НЕ используй from ... import *)
from Autodesk.Revit.DB import Transaction

# 6. НАСТРОЙКИ
doc = revit.doc
output = script.get_output()

# 7. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
def helper_function():
    pass

# 8. КЛАСС ФОРМЫ (в конце файла!)
class MyForm(Form):
    def __init__(self):
        self.setup_form()

# 9. MAIN
if __name__ == "__main__":
    form = MyForm()
    form.ShowDialog()
```

**ВАЖНО:** Если порядок нарушен - получишь ошибку:
`Parent does not have a default constructor`

### ВСЕГДА используй стандартный стиль WinForms!

НЕ раскрашивай формы! Используй стандартный системный стиль как в модуле "Запуск Dynamo".
- НЕ используй BackColor для панелей и форм
- НЕ используй FlatStyle.Flat с цветными кнопками
- НЕ создавай "красивые" заголовки с цветным фоном
- Используй стандартные системные цвета и шрифты

```python
# НЕПРАВИЛЬНО - НЕ раскрашивай!
form.BackColor = Drawing.Color.White
header.BackColor = Drawing.Color.FromArgb(41, 128, 185)
btn.FlatStyle = WinForms.FlatStyle.Flat
btn.BackColor = Drawing.Color.FromArgb(46, 204, 113)

# ПРАВИЛЬНО - используй стандартный стиль!
form = Form()
form.Text = "My Dialog"
btn = Button()
btn.Text = "OK"
# Никаких BackColor, FlatStyle и т.п.!
```

### Наследование от Form - ПРАВИЛЬНЫЙ способ

Наследование от Form РАБОТАЕТ если импортировать Form напрямую из System.Windows.Forms:

```python
# ПРАВИЛЬНО - импорт напрямую и наследование!
from System.Windows.Forms import Form, Button, Label

class MyForm(Form):
    def __init__(self):
        self.my_data = None  # Атрибуты класса работают!
        self.setup_form()

    def setup_form(self):
        self.Text = "My Dialog"
        self.Width = 400

        self.btn = Button()
        self.btn.Text = "OK"
        self.btn.Click += self.on_click
        self.Controls.Add(self.btn)

    def on_click(self, sender, args):
        self.my_data = "result"
        self.Close()

# Использование
form = MyForm()
form.ShowDialog()
result = form.my_data
```

**ВАЖНО:** При наследовании от Form можно использовать атрибуты класса (self.xxx).
Проблемы возникают только при создании Form() как объекта без наследования.

### НИКОГДА не используй Application.Run()!

```python
# НЕПРАВИЛЬНО - вызовет ошибку!
Application.Run(form)

# ПРАВИЛЬНО - используй ShowDialog()
form = create_form()
form.ShowDialog()
```

Ошибка: `Запуск второго цикла сообщения в единичном потоке является недопустимой операцией`

### Порядок добавления контролов (Docking)

WinForms добавляет контролы в обратном z-порядке:
- Fill контрол добавляется ПЕРВЫМ (уходит на задний план)
- Top/Bottom контролы добавляются ПОСЛЕ (выходят на передний план)

```python
# ПРАВИЛЬНЫЙ порядок для панели с Top + Fill + Bottom:
panel.Controls.Add(list_box)     # Fill - ПЕРВЫЙ
panel.Controls.Add(bottom_panel) # Bottom
panel.Controls.Add(top_label)    # Top - ПОСЛЕДНИЙ

# ПРАВИЛЬНЫЙ порядок для главной формы:
self.Controls.Add(main_split)    # Fill - ПЕРВЫЙ
self.Controls.Add(bottom_panel)  # Bottom
self.Controls.Add(top_panel)     # Top - ПОСЛЕДНИЙ
```

Если контролы не видны - проверь порядок добавления!

## Icons (REQUIRED!)

Every pushbutton MUST have `icon.png`:
- Size: 32x32 or 16x16 pixels
- Format: PNG with transparency
- Location: `YourCommand.pushbutton/icon.png`

Without icon, button may not appear in ribbon!

## Debugging

```python
output = script.get_output()
output.print_md("## Debug")
output.print_md("- Value: {}".format(some_value))

# Or use forms for quick alerts
forms.alert("Debug message", title="Debug")
```

## Common Issues

### "No families loaded"
Load required Revit families before running tool:
- Structural Columns: Metric Column families
- Structural Framing: Metric Beam families

### Transaction errors
Always wrap Revit modifications in transaction:
```python
with revit.Transaction("Create Elements"):
    # modifications here
```

### Element type vs instance
- `WhereElementIsElementType()` - Family types (templates)
- `WhereElementIsNotElementType()` - Placed instances

## Работа с конфигом cpsk_settings.yaml (КРИТИЧНО!)

Все настройки проекта хранятся в едином файле `cpsk_settings.yaml` в корне проекта.
**НИКОГДА не создавай свои функции чтения/записи конфига!**

### Структура конфига:

```yaml
# cpsk_settings.yaml
auth:
  email: "user@email.com"
  token: "eyJhbGciOiJIUzI1NiIs..."
  remember: false

environment:
  python_path: "C:/Python313/python.exe"
  installed: true
  last_check: "2024-01-15 10:30:00"

notifications:
  show_startup_check: true
  show_env_warnings: true
```

### ПРАВИЛО: Всегда используй cpsk_config.py!

```python
from cpsk_config import get_setting, set_setting

# Чтение настройки (путь через точку)
token = get_setting("auth.token", "")  # второй аргумент - значение по умолчанию
email = get_setting("auth.email", "")
is_installed = get_setting("environment.installed", False)

# Запись настройки
set_setting("auth.token", "new_token_value")
set_setting("auth.email", "user@example.com")
set_setting("environment.installed", True)
```

### НЕПРАВИЛЬНО - НЕ делай так!

```python
# НЕПРАВИЛЬНО - свои функции чтения/записи!
def _read_config():
    with open("cpsk_settings.yaml") as f:
        ...

def _save_config(data):
    with open("cpsk_settings.yaml", "w") as f:
        ...

# НЕПРАВИЛЬНО - прямая работа с файлом!
import yaml
with open("cpsk_settings.yaml") as f:
    config = yaml.load(f)
```

### Добавление новой секции в конфиг:

1. Добавь секцию в `DEFAULT_SETTINGS` в `cpsk_config.py`:
```python
DEFAULT_SETTINGS = {
    "auth": {...},
    "environment": {...},
    "notifications": {...},
    "my_new_section": {  # Новая секция
        "option1": "",
        "option2": False
    }
}
```

2. Используй через get_setting/set_setting:
```python
value = get_setting("my_new_section.option1", "default")
set_setting("my_new_section.option2", True)
```

### Почему это важно:

1. **Единообразие** - все модули работают с одним форматом
2. **Не затирает данные** - при сохранении сохраняются ВСЕ секции
3. **Миграция** - легко добавлять новые настройки с defaults
4. **В .gitignore** - файл не попадает в репозиторий (токены безопасны)
