# -*- coding: utf-8 -*-
"""
CPSK Config - Модуль управления настройками CPSK Tools.
Работает с cpsk_settings.yaml в корне проекта.
"""

import os
import codecs
from datetime import datetime

# Пути
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.dirname(_THIS_DIR)
PROJECT_DIR = os.path.dirname(EXTENSION_DIR)  # Корень проекта (pyrevit_rocket/)
SETTINGS_FILE = os.path.join(PROJECT_DIR, "cpsk_settings.yaml")
LIB_DIR = _THIS_DIR

# Путь к venv вне OneDrive (OneDrive блокирует os.path.exists для файлов в Documents)
VENV_BASE_DIR = r"C:\cpsk_envs"
VENV_NAME = "pyrevit_rocket"

# Значения по умолчанию
# Примечание: venv_path и requirements_path фиксированы в коде (get_venv_path, get_requirements_path)
DEFAULT_SETTINGS = {
    "environment": {
        "python_path": "",
        "installed": False,
        "last_check": ""
    },
    "auth": {
        "email": "",
        "remember": False,
        "token": ""
    },
    "notifications": {
        "show_startup_check": True,
        "show_env_warnings": True
    },
    "dynamo": {
        "recent": [],
        "favorites": [],
        "run_counts": {},
        "last_runs": {}
    }
}

# Возможные пути к Python
PYTHON_SEARCH_PATHS = [
    r"C:\Python313\python.exe",
    r"C:\Python312\python.exe",
    r"C:\Python311\python.exe",
    r"C:\Python310\python.exe",
    r"C:\ProgramData\miniconda3\python.exe",
    r"C:\Users\{user}\AppData\Local\Programs\Python\Python313\python.exe",
    r"C:\Users\{user}\AppData\Local\Programs\Python\Python312\python.exe",
    r"C:\Users\{user}\AppData\Local\Programs\Python\Python311\python.exe",
]


def _simple_yaml_load(filepath):
    """Простой парсер YAML (без внешних зависимостей)."""
    if not os.path.exists(filepath):
        return {}

    result = {}
    current_section = None
    current_subsection = None

    with codecs.open(filepath, 'r', 'utf-8') as f:
        for line in f:
            # Пропуск комментариев и пустых строк
            stripped = line.strip()
            if not stripped or stripped.startswith('#'):
                continue

            # Определяем уровень отступа
            indent = len(line) - len(line.lstrip())

            # Парсим ключ: значение
            if ':' in stripped:
                key, _, value = stripped.partition(':')
                key = key.strip()
                value = value.strip()

                # Убираем кавычки
                if value.startswith('"') and value.endswith('"'):
                    value = value[1:-1]
                elif value.startswith("'") and value.endswith("'"):
                    value = value[1:-1]

                # Преобразование типов
                if value.lower() == 'true':
                    value = True
                elif value.lower() == 'false':
                    value = False
                elif value == '':
                    value = None

                if indent == 0:
                    # Секция верхнего уровня
                    if value is None or value == '':
                        result[key] = {}
                        current_section = key
                        current_subsection = None
                    else:
                        result[key] = value
                        current_section = None
                elif indent == 2 and current_section:
                    # Подсекция
                    if value is None or value == '':
                        result[current_section][key] = {}
                        current_subsection = key
                    else:
                        result[current_section][key] = value
                elif indent == 4 and current_section and current_subsection:
                    # Вложенное значение
                    if current_subsection not in result[current_section]:
                        result[current_section][current_subsection] = {}
                    result[current_section][current_subsection][key] = value

    return result


def _simple_yaml_dump(data, filepath):
    """Простой сериализатор YAML."""
    lines = ["# CPSK Tools - Настройки", "# Автоматически сгенерировано", ""]

    for section, content in data.items():
        lines.append("{}:".format(section))
        if isinstance(content, dict):
            for key, value in content.items():
                if isinstance(value, dict):
                    lines.append("  {}:".format(key))
                    for k, v in value.items():
                        lines.append("    {}: {}".format(k, _format_value(v)))
                else:
                    lines.append("  {}: {}".format(key, _format_value(value)))
        else:
            lines[-1] = "{}: {}".format(section, _format_value(content))
        lines.append("")

    with codecs.open(filepath, 'w', 'utf-8') as f:
        f.write('\n'.join(lines))


def _format_value(value):
    """Форматировать значение для YAML."""
    if value is None or value == '':
        return '""'
    if isinstance(value, bool):
        return 'true' if value else 'false'
    if isinstance(value, str):
        if ' ' in value or ':' in value or '#' in value:
            return '"{}"'.format(value)
        return value
    return str(value)


def load_settings():
    """Загрузить настройки из файла."""
    settings = dict(DEFAULT_SETTINGS)

    if os.path.exists(SETTINGS_FILE):
        try:
            loaded = _simple_yaml_load(SETTINGS_FILE)
            # Мержим с defaults
            for section in settings:
                if section in loaded:
                    if isinstance(settings[section], dict):
                        settings[section].update(loaded[section])
                    else:
                        settings[section] = loaded[section]
        except Exception:
            pass

    return settings


def save_settings(settings):
    """Сохранить настройки в файл."""
    try:
        _simple_yaml_dump(settings, SETTINGS_FILE)
        return True
    except Exception:
        return False


def get_setting(path, default=None):
    """
    Получить настройку по пути.
    Пример: get_setting("environment.venv_path")
    """
    settings = load_settings()
    keys = path.split('.')
    value = settings

    for key in keys:
        if isinstance(value, dict) and key in value:
            value = value[key]
        else:
            return default

    return value if value is not None else default


def set_setting(path, value):
    """
    Установить настройку по пути.
    Пример: set_setting("environment.python_path", "C:/Python313/python.exe")
    """
    settings = load_settings()
    keys = path.split('.')

    # Навигация к нужной секции
    current = settings
    for key in keys[:-1]:
        if key not in current:
            current[key] = {}
        current = current[key]

    current[keys[-1]] = value
    return save_settings(settings)


def get_absolute_path(relative_path):
    """Преобразовать относительный путь в абсолютный."""
    if not relative_path:
        return ""
    if os.path.isabs(relative_path):
        return relative_path
    return os.path.normpath(os.path.join(EXTENSION_DIR, relative_path))


def get_relative_path(absolute_path):
    """Преобразовать абсолютный путь в относительный (если возможно)."""
    if not absolute_path:
        return ""
    try:
        rel = os.path.relpath(absolute_path, EXTENSION_DIR)
        # Проверяем что путь действительно внутри extension
        if not rel.startswith('..'):
            return rel.replace('\\', '/')
    except ValueError:
        pass
    return absolute_path


def get_venv_path():
    """Получить абсолютный путь к venv (вне OneDrive)."""
    # Фиксированный путь - C:\cpsk_envs\pyrevit_rocket
    # Вне OneDrive чтобы избежать проблем с placeholder файлами
    return os.path.normpath(os.path.join(VENV_BASE_DIR, VENV_NAME))


def get_requirements_path():
    """Получить абсолютный путь к requirements.txt."""
    # Фиксированный путь - requirements.txt в корне проекта
    return os.path.normpath(os.path.join(PROJECT_DIR, "requirements.txt"))


def get_venv_python():
    """Получить путь к python.exe в venv."""
    venv = get_venv_path()
    return os.path.join(venv, "Scripts", "python.exe")


def get_venv_pip():
    """Получить путь к pip.exe в venv."""
    venv = get_venv_path()
    return os.path.join(venv, "Scripts", "pip.exe")


def find_system_python():
    """Найти системный Python."""
    # Сначала проверяем сохранённый путь
    saved = get_setting("environment.python_path", "")
    if saved and os.path.exists(saved):
        return saved

    # Ищем по стандартным путям
    import getpass
    user = getpass.getuser()

    for path in PYTHON_SEARCH_PATHS:
        path = path.format(user=user)
        if os.path.exists(path):
            return path

    # Пробуем через where
    try:
        import subprocess
        result = subprocess.check_output(["where", "python"], shell=True)
        paths = result.decode('utf-8', errors='ignore').strip().split('\n')
        for p in paths:
            p = p.strip()
            if os.path.exists(p) and 'WindowsApps' not in p:
                return p
    except Exception:
        pass

    return None


def get_clean_env():
    """Получить очищенное окружение для subprocess (без переменных IronPython)."""
    env = os.environ.copy()
    env.pop('PYTHONHOME', None)
    env.pop('PYTHONPATH', None)
    env.pop('IRONPYTHONPATH', None)
    return env


def get_python_version(python_path):
    """Получить версию Python."""
    if not python_path or not os.path.exists(python_path):
        return None
    try:
        import subprocess
        CREATE_NO_WINDOW = 0x08000000
        result = subprocess.check_output(
            [python_path, "--version"],
            stderr=subprocess.STDOUT,
            env=get_clean_env(),
            creationflags=CREATE_NO_WINDOW
        )
        return result.decode('utf-8', errors='ignore').strip()
    except Exception:
        return None


def parse_requirements(req_path):
    """
    Парсить requirements.txt и вернуть dict {package_name: version_spec}.
    Примеры:
        ifcopenshell>=0.8.0 -> {"ifcopenshell": ">=0.8.0"}
        requests -> {"requests": ""}
    """
    packages = {}
    if not os.path.exists(req_path):
        return packages

    with codecs.open(req_path, 'r', 'utf-8') as f:
        for line in f:
            line = line.strip()
            # Пропуск комментариев и пустых строк
            if not line or line.startswith('#'):
                continue
            # Пропуск опций (-r, -e, etc.)
            if line.startswith('-'):
                continue

            # Парсим имя пакета и версию
            # Форматы: package, package==1.0, package>=1.0, package[extra]>=1.0
            pkg_name = line
            version_spec = ""

            # Убираем extras [xxx]
            if '[' in pkg_name:
                pkg_name = pkg_name.split('[')[0]

            # Ищем спецификатор версии
            for sep in ['>=', '<=', '==', '!=', '~=', '>', '<']:
                if sep in line:
                    parts = line.split(sep, 1)
                    pkg_name = parts[0].strip()
                    if '[' in pkg_name:
                        pkg_name = pkg_name.split('[')[0]
                    version_spec = sep + parts[1].strip()
                    break

            packages[pkg_name.lower()] = version_spec

    return packages


def get_installed_packages():
    """
    Получить список установленных пакетов в venv.
    Возвращает dict {package_name: version}.
    """
    packages = {}
    pip = get_venv_pip()

    if not os.path.exists(pip):
        return packages

    try:
        import subprocess
        CREATE_NO_WINDOW = 0x08000000
        output = subprocess.check_output(
            [pip, "list", "--format=freeze"],
            creationflags=CREATE_NO_WINDOW,
            env=get_clean_env()
        )
        lines = output.decode('utf-8', errors='ignore').strip().split('\n')

        for line in lines:
            line = line.strip()
            if not line:
                continue
            if '==' in line:
                name, version = line.split('==', 1)
                packages[name.lower()] = version
            else:
                packages[line.lower()] = ""

    except Exception:
        pass

    return packages


def parse_version(version_str):
    """
    Парсить строку версии в кортеж для сравнения.
    Примеры:
        "1.2.3" -> (1, 2, 3)
        "0.8.0" -> (0, 8, 0)
        "1.0.0a1" -> (1, 0, 0, 'a1')
    """
    if not version_str:
        return (0,)

    parts = []
    current = ""

    for char in version_str:
        if char == '.':
            if current:
                # Пробуем преобразовать в число
                try:
                    parts.append(int(current))
                except ValueError:
                    parts.append(current)
                current = ""
        elif char.isdigit():
            current += char
        else:
            # Начало буквенной части (alpha, beta, etc.)
            if current:
                try:
                    parts.append(int(current))
                except ValueError:
                    parts.append(current)
                current = ""
            current += char

    if current:
        try:
            parts.append(int(current))
        except ValueError:
            parts.append(current)

    return tuple(parts) if parts else (0,)


def compare_versions(v1, v2):
    """
    Сравнить две версии.
    Возвращает: -1 если v1 < v2, 0 если v1 == v2, 1 если v1 > v2
    """
    p1 = parse_version(v1)
    p2 = parse_version(v2)

    # Выравниваем длину кортежей
    max_len = max(len(p1), len(p2))
    p1 = p1 + (0,) * (max_len - len(p1))
    p2 = p2 + (0,) * (max_len - len(p2))

    for a, b in zip(p1, p2):
        # Если оба числа - сравниваем как числа
        if isinstance(a, int) and isinstance(b, int):
            if a < b:
                return -1
            elif a > b:
                return 1
        else:
            # Строковое сравнение
            a_str = str(a)
            b_str = str(b)
            if a_str < b_str:
                return -1
            elif a_str > b_str:
                return 1

    return 0


def check_version_constraint(installed_ver, version_spec):
    """
    Проверить соответствует ли установленная версия требованию.
    Примеры:
        check_version_constraint("0.8.4", ">=0.8.0") -> True
        check_version_constraint("0.7.0", ">=0.8.0") -> False
        check_version_constraint("1.0.0", "==1.0.0") -> True
    """
    if not version_spec:
        return True

    # Определяем оператор и требуемую версию
    operators = ['>=', '<=', '==', '!=', '~=', '>', '<']
    op = None
    required_ver = version_spec

    for operator in operators:
        if version_spec.startswith(operator):
            op = operator
            required_ver = version_spec[len(operator):].strip()
            break

    if not op:
        # Нет оператора - считаем как ==
        op = '=='

    cmp_result = compare_versions(installed_ver, required_ver)

    if op == '==':
        return cmp_result == 0
    elif op == '!=':
        return cmp_result != 0
    elif op == '>=':
        return cmp_result >= 0
    elif op == '<=':
        return cmp_result <= 0
    elif op == '>':
        return cmp_result > 0
    elif op == '<':
        return cmp_result < 0
    elif op == '~=':
        # ~= означает совместимую версию (>=X.Y, <X+1.0)
        # Упрощённо: проверяем что major.minor совпадают
        inst_parts = parse_version(installed_ver)
        req_parts = parse_version(required_ver)
        if len(inst_parts) >= 2 and len(req_parts) >= 2:
            return inst_parts[0] == req_parts[0] and inst_parts[1] >= req_parts[1]
        return cmp_result >= 0

    return True


def compare_packages():
    """
    Сравнить пакеты из requirements.txt с установленными в venv.
    Возвращает dict:
        - missing: список пакетов которые есть в requirements, но не установлены
        - extra: список пакетов которые установлены, но нет в requirements
        - outdated: список пакетов с несовпадающими версиями (требуется обновление)
        - match: True если все пакеты из requirements установлены
    """
    result = {
        "missing": [],
        "extra": [],
        "outdated": [],
        "match": True
    }

    req_path = get_requirements_path()
    required = parse_requirements(req_path)
    installed = get_installed_packages()

    if not required:
        return result

    # Проверяем отсутствующие пакеты и версии
    for pkg, version_spec in required.items():
        if pkg not in installed:
            result["missing"].append(pkg)
            result["match"] = False
        elif version_spec:
            installed_ver = installed.get(pkg, "")
            if not check_version_constraint(installed_ver, version_spec):
                result["outdated"].append({
                    "package": pkg,
                    "required": version_spec,
                    "installed": installed_ver
                })
                result["match"] = False

    return result


def check_environment():
    """
    Проверить состояние окружения.
    Возвращает dict с результатами проверки.
    """
    result = {
        "venv_exists": False,
        "venv_python": None,
        "venv_version": None,
        "requirements_exists": False,
        "requirements_count": 0,
        "packages_installed": False,
        "packages_match": True,
        "missing_packages": [],
        "outdated_packages": [],
        "system_python": None,
        "system_version": None,
        "is_ready": False,
        "needs_update": False,
        "errors": []
    }

    # Проверка venv
    venv_python = get_venv_python()
    if os.path.exists(venv_python):
        result["venv_exists"] = True
        result["venv_python"] = venv_python
        result["venv_version"] = get_python_version(venv_python)

        # Проверка установленных пакетов
        try:
            installed = get_installed_packages()
            # Проверяем ключевые пакеты
            if 'ifcopenshell' in installed and 'ifctester' in installed:
                result["packages_installed"] = True

            # Сравнение с requirements.txt
            comparison = compare_packages()
            result["packages_match"] = comparison["match"]
            result["missing_packages"] = comparison["missing"]
            result["outdated_packages"] = comparison["outdated"]

            if not comparison["match"]:
                result["needs_update"] = True

        except Exception as e:
            result["errors"].append("Ошибка проверки пакетов: {}".format(str(e)))
    else:
        result["errors"].append("Виртуальное окружение не найдено")

    # Проверка requirements.txt
    req_path = get_requirements_path()
    if os.path.exists(req_path):
        result["requirements_exists"] = True
        with codecs.open(req_path, 'r', 'utf-8') as f:
            lines = [l.strip() for l in f if l.strip() and not l.startswith('#')]
            result["requirements_count"] = len(lines)
    else:
        result["errors"].append("Файл requirements.txt не найден")

    # Проверка системного Python
    sys_python = find_system_python()
    if sys_python:
        result["system_python"] = sys_python
        result["system_version"] = get_python_version(sys_python)
    else:
        result["errors"].append("Системный Python не найден")

    # Итоговый статус
    result["is_ready"] = (
        result["venv_exists"] and
        result["packages_installed"] and
        result["requirements_exists"] and
        result["packages_match"]
    )

    # Обновляем настройки
    set_setting("environment.installed", result["is_ready"])
    set_setting("environment.last_check", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    return result


# Глобальная переменная для кэширования статуса окружения в текущей сессии
_ENVIRONMENT_CHECKED = None


def is_environment_ready():
    """
    Проверка готовности окружения.
    Результат кэшируется в глобальной переменной на время сессии.
    """
    global _ENVIRONMENT_CHECKED

    # Если уже проверяли в этой сессии - возвращаем кэшированный результат
    if _ENVIRONMENT_CHECKED is not None:
        return _ENVIRONMENT_CHECKED

    # Полная проверка
    venv_python = get_venv_python()
    if not os.path.exists(venv_python):
        _ENVIRONMENT_CHECKED = False
        return False

    result = check_environment()
    _ENVIRONMENT_CHECKED = result["is_ready"]
    return _ENVIRONMENT_CHECKED


def reset_environment_cache():
    """Сбросить кэш проверки окружения (вызывать после установки)."""
    global _ENVIRONMENT_CHECKED
    _ENVIRONMENT_CHECKED = None


def require_environment(show_message=True):
    """
    Проверить окружение и показать сообщение если не готово.
    Используется в начале скриптов для блокировки выполнения.

    Возвращает True если окружение готово, False если нет.

    Пример использования:
        from cpsk_config import require_environment
        if not require_environment():
            import sys
            sys.exit()
    """
    if is_environment_ready():
        return True

    if show_message:
        details = """Для работы некоторых инструментов CPSK требуется Python окружение.

Как установить окружение:
1. Найдите панель "CPSK" на ленте Revit
2. Перейдите в Settings → Окружение
3. Нажмите кнопку "Установить окружение"
4. Дождитесь завершения установки

Окружение устанавливается один раз и сохраняется в:
{}

Если установка не удаётся, проверьте:
- Наличие интернет-соединения
- Права на запись в указанную папку
- Наличие Python 3.x в системе""".format(VENV_BASE_DIR)

        try:
            from cpsk_notify import show_warning
            show_warning(
                "Требуется настройка окружения",
                "Python окружение не установлено. Перейдите в Settings → Окружение.",
                details=details
            )
        except Exception:
            # Fallback на forms.alert
            try:
                from pyrevit import forms
                forms.alert(
                    "Окружение не установлено!\n\n"
                    "Перейдите в Settings → Окружение\n"
                    "и нажмите 'Установить окружение'.",
                    title="CPSK - Требуется настройка",
                    warn_icon=True
                )
            except Exception:
                pass

    return False
