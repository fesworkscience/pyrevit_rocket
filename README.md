<div align="center">

# CPSK Tools

### Industrial Building Automation for Autodesk Revit

[![Version](https://img.shields.io/badge/version-1.0.53-blue.svg)](https://github.com/fesworkscience/pyrevit_rocket/releases)
[![pyRevit](https://img.shields.io/badge/pyRevit-5.0+-green.svg)](https://github.com/pyrevitlabs/pyRevit)
[![Revit](https://img.shields.io/badge/Revit-2022--2025-orange.svg)](https://www.autodesk.com/products/revit)
[![License](https://img.shields.io/badge/license-MIT-lightgrey.svg)](LICENSE)

[English](#features) | [Русский](#возможности)

<img src="docs/social_preview.png" alt="CPSK Tools" width="800">

</div>

---

## Features

- **Dynamo Integration** - Launch and manage Dynamo scripts directly from Revit ribbon
- **IDS Validation** - Information Delivery Specification tools for BIM quality control
- **Family Management** - Quick family insertion and parameter management
- **Specifications** - Automated schedule and specification generation
- **Rhino/Grasshopper** - Integration with Rhino.Inside.Revit
- **SLAM Tools** - Point cloud processing from LiDAR scans (iOS/Android)
- **Structural Tools** - КЖ documentation automation

## Requirements

- **Autodesk Revit** 2022, 2023, 2024 or 2025
- **pyRevit** 5.0 or higher
- **Windows** 10/11

## Installation

### Option 1: Installer (Recommended)

Download the latest installer from [Releases](https://github.com/fesworkscience/pyrevit_rocket/releases)

### Option 2: Manual Installation

1. **Install pyRevit**

   Download from [pyRevit Releases](https://github.com/pyrevitlabs/pyRevit/releases)

2. **Clone the repository**
   ```bash
   git clone https://github.com/fesworkscience/pyrevit_rocket.git
   ```

3. **Add extension to pyRevit**
   - Open Revit
   - Go to pyRevit → Settings → Custom Extension Directories
   - Add path to `pyrevit_rocket` folder
   - Restart Revit

4. **Setup environment**
   - CPSK tab → Settings → Environment
   - Click "Install Environment"

---

## Возможности

- **Интеграция с Dynamo** - Запуск и управление скриптами Dynamo из ленты Revit
- **IDS Валидация** - Инструменты проверки по Information Delivery Specification
- **Управление семействами** - Быстрая вставка семейств и управление параметрами
- **Спецификации** - Автоматизация создания ведомостей и спецификаций
- **Rhino/Grasshopper** - Интеграция с Rhino.Inside.Revit
- **SLAM инструменты** - Обработка облаков точек с LiDAR сканов (iOS/Android)
- **Инструменты КЖ** - Автоматизация документации по разделу КЖ

## Требования

- **Autodesk Revit** 2022, 2023, 2024 или 2025
- **pyRevit** 5.0 или выше
- **Windows** 10/11

## Установка

### Вариант 1: Установщик (Рекомендуется)

Скачайте последний установщик из [Releases](https://github.com/fesworkscience/pyrevit_rocket/releases)

### Вариант 2: Ручная установка

1. **Установить pyRevit**

   Скачать с [pyRevit Releases](https://github.com/pyrevitlabs/pyRevit/releases)

2. **Клонировать репозиторий**
   ```bash
   git clone https://github.com/fesworkscience/pyrevit_rocket.git
   ```

3. **Добавить расширение в pyRevit**
   - Открыть Revit
   - pyRevit → Settings → Custom Extension Directories
   - Добавить путь к папке `pyrevit_rocket`
   - Перезапустить Revit

4. **Настроить окружение**
   - Вкладка CPSK → Settings → Окружение
   - Нажать "Установить окружение"

---

## Project Structure

```
pyrevit.extension/
├── CPSK.tab/
│   ├── 01_Settings.panel/     # Login, updates, environment
│   ├── 02_Dynamo.panel/       # Dynamo script launcher
│   ├── 03_QA.panel/           # IDS validation tools
│   ├── 04_Families.panel/     # Family management
│   ├── 05_Specifications.panel/# Schedule tools
│   ├── 06_Projects.panel/     # Project utilities
│   ├── 06_Rhino.panel/        # Rhino.Inside integration
│   ├── 07_КЖ.panel/           # Structural documentation
│   └── 08_SLAM.panel/         # Point cloud tools
└── lib/                       # Shared libraries
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Links

- [GIP Group](https://gip.su) - Developer company
- [pyRevit](https://github.com/pyrevitlabs/pyRevit) - Platform

---

<div align="center">

Made with :heart: by [GIP Group](https://gip.su)

</div>
