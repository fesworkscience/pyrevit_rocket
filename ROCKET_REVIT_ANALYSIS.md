# Анализ проекта rocket-revit-master

**Тип проекта:** C# плагин для Revit (.NET Framework 4.8)
**Платформа:** Autodesk Revit 2024
**Система обновлений:** Squirrel.Windows (автообновление без перезапуска для UI)

---

## Структура проекта

```
RocketRevit.sln
├── RocketRevitMain/         - Основной плагин (Revit API)
├── RocketRevitHot/          - UI компоненты (горячая перезагрузка)
├── RocketRevitInstaller/    - Установщик Squirrel
├── RocketRevitLoader/       - Загрузчик плагина
└── RocketRevitTests/        - Тесты
```

---

## ФУНКЦИОНАЛ (отметь что переносить)

### 1. СИСТЕМА ОБНОВЛЕНИЙ
[НЕ ПЕРЕНОСИМ] **Автообновление Squirrel** - автоматические обновления плагина
    - Умное определение изменённых компонентов
    - Горячая перезагрузка UI без перезапуска Revit
    - Манифест компонентов с хешами файлов

### 2. DYNAMO СКРИПТЫ (DynamoScriptCommand)
[НЕ ПЕРЕНОСИМ] **Запуск Dynamo скриптов с сервера** - скачивание и выполнение .dyn
    - Поиск скриптов по API
    - Скачивание во временную папку
    - Запуск через Dynamo Player
    - WPF диалог выбора скрипта
    - Требует авторизации

### 3. СЕМЕЙСТВА (FamilyInsertCommand)
[ПЕРЕНОСИМ!!! Нужно] **Вставка семейств с сервера** - скачивание и загрузка .rfa
    - Поиск семейств по API
    - Скачивание во временную папку
    - Загрузка в проект Revit
    - WPF диалог с поиском
    - Требует авторизации

### 4. СПЕЦИФИКАЦИИ (SpecificationCommands)
[ ] **CreateBOMSpecificationCommand** - создание спецификации ВД
    - Копирование существующей спецификации
    - Настройка ширины столбцов (A+B=20мм, D+E=70мм)
    - Фильтр для скрытия строк

[ ] **AnalyzeSpecificationCommand** - анализ спецификаций
    - Диалог выбора спецификации
    - Отображение результатов анализа

[ ] **UpdateSpecificationCommand** - обновление спецификаций

[ ] **HideSpecificationRowsCommand** - скрытие строк спецификации

[ ] **ShowAllSpecificationRowsCommand** - показ всех строк

[ ] **FillSketchScheduleCommand** - заполнение эскизной спецификации

[ ] **UpdateSketchScheduleCommand** - обновление эскизной спецификации

[ ] **GenerateRebarShapeImagesCommand** - генерация изображений форм арматуры

[ ] **CheckFamilyCommand** - проверка семейства

[ ] **CheckProjectBeforeBOMCommand** - проверка проекта перед созданием ВД

[ ] **DeleteVDAnnotationsCommand** - удаление аннотаций ВД

[ ] **PlaceFamilyCommand** - размещение семейства

[ ] **DWGConverter + NetDxfProcessor** - работа с DWG/DXF

### 5. AR ИНТЕГРАЦИЯ
[НЕ ПЕРЕНОСИМ] **ARIntegrationCommand** - интеграция с AR
    - ARIntegrationService
    - ARIntegrationWindow (WPF)

[НЕ ПЕРЕНОСИМ] **ObservationPointsCommand** - точки наблюдения

### 6. ЧАТ СИСТЕМА
[НЕ ПЕРЕНОСИМ] **Chat** - встроенный чат для совместной работы
    - ChatWindow, ChatNotificationPopup
    - ChatWebSocket, ChatApiService
    - ChatRoom, ChatMessage модели
    - Интеграция с Revit элементами
    - Требует авторизации

### 7. АВТОРИЗАЦИЯ
[x] **StaticAuthService** - статическая аутентификация (ПЕРЕНЕСЕНО)
    - Управление сессиями
    - API токены
    - **Реализовано в:** `lib/cpsk_auth.py`

[x] **Login/Logout команды** - вход/выход (ПЕРЕНЕСЕНО)
    - **Реализовано в:** `01_Settings.panel/Login.pushbutton/script.py`

### 8. НАСТРОЙКИ
[НЕ ПЕРЕНОСИМ] **SettingsCommand** - настройки плагина
    - UnifiedSettings
    - AppSettings, ApiConfig

### 9. ПРОЕКТЫ
[ ] **RegisterProjectCommand** - регистрация проекта
[ ] **RegisterElementCommand** - регистрация элементов
[ ] **SendProjectStatisticsCommand** - отправка статистики

### 10. UI КОМПОНЕНТЫ
[НЕ ПЕРЕНОСИМ] **NotificationWindow** - уведомления (WPF)
[НЕ ПЕРЕНОСИМ] **ProgressForm** - прогресс бар
[НЕ ПЕРЕНОСИМ ] **BaseWindow** - базовое окно
[НЕ ПЕРЕНОСИМ] **Диалоги выбора** - ScheduleSelectionDialog, FamilyTypeSelectionDialog и др.

### 11. УТИЛИТЫ
[НЕ ПЕРЕНОСИМ] **CentralLogger** - централизованное логирование
[НЕ ПЕРЕНОСИМ] **DebugLogger** - отладочные логи
[x] **ApiService** - работа с API (ПЕРЕНЕСЕНО)
    - **Реализовано в:** `lib/cpsk_auth.py` (класс ApiClient)

### 12. ДОКУМЕНТАЦИЯ
[НЕ ПЕРЕНОСИМ] **README.md** - документация для пользователей и разработчиков
[НЕ ПЕРЕНОСИМ] **CLAUDE.md** - конфигурация для Claude

---

## ПРИМЕЧАНИЯ

1. **Требует бэкенд** - многие функции (Dynamo, Family, Chat) требуют серверную часть API
2. **Авторизация** - большинство функций требуют StaticAuthService.IsAuthenticated
3. **WPF vs WinForms** - проект использует WPF для UI (не совместим с pyRevit IronPython)
4. **Версия .NET** - .NET Framework 4.8 (не .NET Core)

---

## РЕКОМЕНДАЦИИ

**Легко перенести в pyRevit (IronPython):**
- Логика работы со спецификациями
- Работа с параметрами элементов
- Простые команды без сложного UI

**Сложно перенести (требует переписывание):**
- WPF диалоги -> WinForms
- async/await -> синхронный код
- API интеграция (другой подход)
- Чат и авторизация

**Не переносится:**
- Squirrel обновления (только для C# плагинов)
- Горячая перезагрузка (специфика C#)
