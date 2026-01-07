# -*- coding: utf-8 -*-
"""
CPSK Config - Модуль управления настройками CPSK Tools.
Работает с cpsk_settings.yaml в корне extension.
"""

import os
import codecs
from datetime import datetime

# Пути
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
EXTENSION_DIR = os.path.dirname(_THIS_DIR)
SETTINGS_FILE = os.path.join(EXTENSION_DIR, "cpsk_settings.yaml")
LIB_DIR = _THIS_DIR

# Значения по умолчанию
DEFAULT_SETTINGS = {
    "environment": {
        "venv_path": "lib/.venv",
        "requirements_path": "lib/requirements.txt",
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
    """Получить абсолютный путь к venv."""
    rel_path = get_setting("environment.venv_path", "lib/.venv")
    return get_absolute_path(rel_path)


def get_requirements_path():
    """Получить абсолютный путь к requirements.txt."""
    rel_path = get_setting("environment.requirements_path", "lib/requirements.txt")
    return get_absolute_path(rel_path)


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


def get_python_version(python_path):
    """Получить версию Python."""
    if not python_path or not os.path.exists(python_path):
        return None
    try:
        import subprocess
        result = subprocess.check_output(
            [python_path, "--version"],
            stderr=subprocess.STDOUT
        )
        return result.decode('utf-8', errors='ignore').strip()
    except Exception:
        return None


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
        "system_python": None,
        "system_version": None,
        "is_ready": False,
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
            import subprocess
            pip = get_venv_pip()
            output = subprocess.check_output([pip, "list", "--format=freeze"])
            packages = output.decode('utf-8', errors='ignore').strip().split('\n')
            # Проверяем ключевые пакеты
            package_names = [p.split('==')[0].lower() for p in packages if p]
            if 'ifcopenshell' in package_names and 'ifctester' in package_names:
                result["packages_installed"] = True
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
        result["requirements_exists"]
    )

    # Обновляем настройки
    set_setting("environment.installed", result["is_ready"])
    set_setting("environment.last_check", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    return result
