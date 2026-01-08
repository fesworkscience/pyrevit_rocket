# -*- coding: utf-8 -*-
"""
pyRevit Universal Code Checker
Проверяет Python код на совместимость с pyRevit (IronPython 2.7)

Использование:
    python pyrevit_checker.py <file.py>
    python pyrevit_checker.py <directory>

Проверки:
    - f-строки (не поддерживаются)
    - walrus operator := (не поддерживается)
    - type hints (предупреждение)
    - subprocess timeout (не поддерживается)
    - open() с encoding (использовать codecs)
    - open() без codecs для текстовых файлов (кракозябры!)
    - Application.Run() (использовать ShowDialog())
    - async/await (не поддерживается)
    - yield from (не поддерживается)
    - nonlocal (не поддерживается)
    - расширенная распаковка *rest (не поддерживается)
    - except без show_error (предупреждение)
    - неправильная работа с cpsk_settings.yaml (использовать cpsk_config)
    - отсутствие проверки авторизации require_auth() в скриптах кнопок
    - использование MessageBox.Show или forms.alert (использовать cpsk_notify)
    - использование output.print_md для ошибок (использовать cpsk_notify)
    - отсутствие icon.png в папках .pushbutton и .pulldown
"""

import sys
import os
import re
import ctypes
from ctypes import wintypes

# Windows API для проверки существования файлов
INVALID_FILE_ATTRIBUTES = 0xFFFFFFFF
GetFileAttributesW = ctypes.windll.kernel32.GetFileAttributesW
GetFileAttributesW.argtypes = [wintypes.LPCWSTR]
GetFileAttributesW.restype = wintypes.DWORD

def win_path_exists(path):
    """Проверить существование пути через Windows API."""
    attrs = GetFileAttributesW(path)
    return attrs != INVALID_FILE_ATTRIBUTES

def win_isfile(path):
    """Проверить что путь - файл."""
    FILE_ATTRIBUTE_DIRECTORY = 0x10
    attrs = GetFileAttributesW(path)
    if attrs == INVALID_FILE_ATTRIBUTES:
        return False
    return not (attrs & FILE_ATTRIBUTE_DIRECTORY)

def win_isdir(path):
    """Проверить что путь - директория."""
    FILE_ATTRIBUTE_DIRECTORY = 0x10
    attrs = GetFileAttributesW(path)
    if attrs == INVALID_FILE_ATTRIBUTES:
        return False
    return bool(attrs & FILE_ATTRIBUTE_DIRECTORY)


class PyRevitChecker:
    """Проверка совместимости с pyRevit (IronPython 2.7)."""

    def __init__(self):
        self.errors = []
        self.warnings = []
        self.current_file = None

    def check_file(self, filepath):
        """Проверить файл."""
        self.current_file = filepath
        self.errors = []
        self.warnings = []

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception as e:
            self.errors.append("Ошибка чтения файла: {}".format(str(e)))
            return False

        lines = content.split('\n')

        # Проверки
        self.check_fstrings(lines)
        self.check_walrus_operator(lines)
        self.check_type_hints(lines)
        self.check_subprocess_timeout(lines)
        self.check_open_encoding(lines)
        self.check_open_without_codecs(lines)
        self.check_application_run(lines)
        self.check_async_await(lines)
        self.check_yield_from(lines)
        self.check_nonlocal(lines)
        self.check_unpacking(lines)
        self.check_except_without_notify(lines, content)
        self.check_config_usage(lines, content)
        self.check_require_auth(lines, content)
        self.check_notification_usage(lines, content)
        self.check_icon_exists()

        return len(self.errors) == 0

    def check_fstrings(self, lines):
        """Проверить f-строки (не поддерживаются в IronPython)."""
        # Улучшенный паттерн для f-строк
        for i, line in enumerate(lines, 1):
            code_part = line.split('#')[0]
            # f"..." или f'...'
            if re.search(r'\bf["\']', code_part):
                self.errors.append(
                    "Строка {}: f-строки не поддерживаются. Используйте .format()".format(i)
                )

    def check_walrus_operator(self, lines):
        """Проверить walrus operator := (Python 3.8+)."""
        for i, line in enumerate(lines, 1):
            code_part = line.split('#')[0]
            if re.search(r':=', code_part):
                self.errors.append(
                    "Строка {}: Walrus operator := не поддерживается.".format(i)
                )

    def check_type_hints(self, lines):
        """Проверить type hints (предупреждение)."""
        # Аннотации возвращаемого значения def foo() -> Type:
        for i, line in enumerate(lines, 1):
            if re.search(r'def\s+\w+\s*\([^)]*\)\s*->', line):
                self.warnings.append(
                    "Строка {}: Return type hints могут не работать.".format(i)
                )

    def check_subprocess_timeout(self, lines):
        """Проверить timeout в subprocess (не поддерживается)."""
        for i, line in enumerate(lines, 1):
            if re.search(r'\.communicate\s*\([^)]*timeout\s*=', line):
                self.errors.append(
                    "Строка {}: timeout в communicate() не поддерживается.".format(i)
                )
            if re.search(r'subprocess\.TimeoutExpired', line):
                self.errors.append(
                    "Строка {}: subprocess.TimeoutExpired не существует.".format(i)
                )

    def check_open_encoding(self, lines):
        """Проверить open() с encoding (использовать codecs)."""
        for i, line in enumerate(lines, 1):
            code_part = line.split('#')[0]
            if re.search(r'\bopen\s*\([^)]*encoding\s*=', code_part):
                self.errors.append(
                    "Строка {}: open(encoding=) не поддерживается. Используйте codecs.open().".format(i)
                )

    def check_open_without_codecs(self, lines):
        """Проверить open() в текстовом режиме без codecs (кракозябры!)."""
        for i, line in enumerate(lines, 1):
            code_part = line.split('#')[0]
            # Пропускаем строки с codecs.open
            if 'codecs.open' in code_part:
                continue
            # Пропускаем webbrowser.open(), os.open() и другие не-файловые open()
            if 'webbrowser.open' in code_part or 'os.open' in code_part:
                continue
            # Ищем open() в текстовом режиме: open(path, 'r'), open(path, 'w'), open(path)
            # НЕ ловим бинарный режим: 'rb', 'wb', 'ab'
            # Паттерн: open(..., 'r'...) или open(..., 'w'...) или open(...) без режима
            match = re.search(r'\bopen\s*\(\s*[^,)]+(?:\s*,\s*[\'"]([rwax][+]?)[\'"])?', code_part)
            if match:
                mode = match.group(1)
                # Если режим не указан или текстовый (r, w, a, x без b)
                if mode is None or 'b' not in mode:
                    self.warnings.append(
                        "Строка {}: open() без codecs может вызвать кракозябры. Используйте codecs.open(path, 'r', 'utf-8').".format(i)
                    )

    def check_application_run(self, lines):
        """Проверить Application.Run() (использовать ShowDialog())."""
        for i, line in enumerate(lines, 1):
            code_part = line.split('#')[0]
            if re.search(r'\bApplication\.Run\s*\(', code_part):
                self.errors.append(
                    "Строка {}: Application.Run() вызовет ошибку! Используйте form.ShowDialog().".format(i)
                )

    def check_async_await(self, lines):
        """Проверить async/await (Python 3.5+)."""
        for i, line in enumerate(lines, 1):
            code_part = line.split('#')[0]
            if re.search(r'\basync\s+def\b', code_part):
                self.errors.append(
                    "Строка {}: async def не поддерживается.".format(i)
                )
            if re.search(r'\bawait\s+', code_part):
                self.errors.append(
                    "Строка {}: await не поддерживается.".format(i)
                )

    def check_yield_from(self, lines):
        """Проверить yield from (Python 3.3+)."""
        for i, line in enumerate(lines, 1):
            code_part = line.split('#')[0]
            if re.search(r'\byield\s+from\b', code_part):
                self.errors.append(
                    "Строка {}: yield from не поддерживается.".format(i)
                )

    def check_nonlocal(self, lines):
        """Проверить nonlocal (Python 3+)."""
        for i, line in enumerate(lines, 1):
            code_part = line.split('#')[0]
            if re.search(r'\bnonlocal\s+', code_part):
                self.errors.append(
                    "Строка {}: nonlocal не поддерживается.".format(i)
                )

    def check_unpacking(self, lines):
        """Проверить расширенную распаковку *args в присваивании."""
        for i, line in enumerate(lines, 1):
            code_part = line.split('#')[0]
            # a, *rest = [1,2,3]
            if re.search(r'^\s*\w+\s*,\s*\*\w+\s*=', code_part):
                self.errors.append(
                    "Строка {}: Расширенная распаковка (*rest) не поддерживается.".format(i)
                )

    def check_except_without_notify(self, lines, content):
        """Проверить except блоки без show_error (предупреждение)."""
        # Проверяем импортирован ли show_error
        has_show_error_import = bool(re.search(r'from\s+cpsk_notify\s+import.*show_error', content))

        # Пропускаем файлы библиотек (cpsk_*.py, pyrevit_checker.py)
        filename = os.path.basename(self.current_file) if self.current_file else ""
        if filename.startswith('cpsk_') or filename == 'pyrevit_checker.py':
            return

        # Ищем блоки except
        in_except = False
        except_line = 0
        except_indent = 0
        has_notify_in_block = False

        for i, line in enumerate(lines, 1):
            stripped = line.lstrip()
            current_indent = len(line) - len(stripped)

            # Начало except блока
            if re.match(r'except\s*.*:', stripped):
                # Если был предыдущий except без notify - добавляем предупреждение
                if in_except and not has_notify_in_block:
                    self.warnings.append(
                        "Строка {}: except без show_error(). Добавьте уведомление пользователю.".format(except_line)
                    )

                in_except = True
                except_line = i
                except_indent = current_indent
                has_notify_in_block = False
                continue

            # Внутри except блока
            if in_except:
                # Проверяем выход из блока (меньший или равный отступ, не пустая строка)
                if stripped and current_indent <= except_indent and not stripped.startswith('#'):
                    # Конец except блока
                    if not has_notify_in_block:
                        self.warnings.append(
                            "Строка {}: except без show_error(). Добавьте уведомление пользователю.".format(except_line)
                        )
                    in_except = False

                    # Проверяем, может это новый except
                    if re.match(r'except\s*.*:', stripped):
                        in_except = True
                        except_line = i
                        except_indent = current_indent
                        has_notify_in_block = False

                # Проверяем наличие cpsk_notify внутри блока (ТОЛЬКО cpsk_notify!)
                if 'show_error' in line or 'show_warning' in line or 'show_info' in line or 'show_success' in line:
                    has_notify_in_block = True
                # pass, return, raise - допустимые варианты без notify
                if re.match(r'\s*(pass|return|raise)\b', line):
                    has_notify_in_block = True

        # Проверяем последний except если файл закончился
        if in_except and not has_notify_in_block:
            self.warnings.append(
                "Строка {}: except без show_error(). Добавьте уведомление пользователю.".format(except_line)
            )

    def check_config_usage(self, lines, content):
        """Проверить правильность работы с конфигом cpsk_settings.yaml."""
        # Пропускаем сам модуль cpsk_config.py
        filename = os.path.basename(self.current_file) if self.current_file else ""
        if filename == 'cpsk_config.py':
            return

        # Проверяем наличие импорта cpsk_config
        has_cpsk_config_import = bool(re.search(r'from\s+cpsk_config\s+import', content))

        # Паттерны неправильного использования конфига
        bad_patterns = [
            # Прямое чтение/запись cpsk_settings.yaml
            (r'["\']cpsk_settings\.yaml["\']', "Прямой доступ к cpsk_settings.yaml. Используйте cpsk_config."),
            # Функции _read_config, _save_config (кроме cpsk_config.py)
            (r'def\s+_read_config\s*\(', "Функция _read_config(). Используйте cpsk_config.get_setting()."),
            (r'def\s+_save_config\s*\(', "Функция _save_config(). Используйте cpsk_config.set_setting()."),
            # yaml.load/dump с settings
            (r'yaml\.(safe_)?load.*settings', "yaml.load() для настроек. Используйте cpsk_config."),
            (r'yaml\.(safe_)?dump.*settings', "yaml.dump() для настроек. Используйте cpsk_config."),
        ]

        for i, line in enumerate(lines, 1):
            code_part = line.split('#')[0]
            for pattern, message in bad_patterns:
                if re.search(pattern, code_part, re.IGNORECASE):
                    self.errors.append(
                        "Строка {}: {}".format(i, message)
                    )

        # Проверяем использование настроек без импорта cpsk_config
        settings_usage_patterns = [
            r'auth\.token',
            r'auth\.email',
            r'environment\.',
            r'notifications\.',
        ]

        if not has_cpsk_config_import:
            for i, line in enumerate(lines, 1):
                code_part = line.split('#')[0]
                for pattern in settings_usage_patterns:
                    if re.search(pattern, code_part):
                        self.warnings.append(
                            "Строка {}: Использование настроек без импорта cpsk_config.".format(i)
                        )
                        break

    def check_require_auth(self, lines, content):
        """Проверить наличие проверки авторизации в скриптах кнопок."""
        if not self.current_file:
            return

        filepath = self.current_file
        filename = os.path.basename(filepath)

        # Проверяем только script.py в .pushbutton папках
        if '.pushbutton' not in filepath:
            return

        # Проверяем только script.py (не вспомогательные модули)
        if filename != 'script.py':
            return

        # Пропускаем Login.pushbutton - ему не нужна проверка авторизации
        if 'Login.pushbutton' in filepath:
            return

        # Пропускаем файлы в lib/ (библиотеки)
        if '\\lib\\' in filepath or '/lib/' in filepath:
            return

        # Проверяем наличие импорта require_auth
        has_require_auth_import = bool(re.search(
            r'from\s+cpsk_auth\s+import.*require_auth',
            content
        ))

        # Проверяем вызов require_auth()
        has_require_auth_call = bool(re.search(
            r'require_auth\s*\(',
            content
        ))

        if not has_require_auth_import:
            self.warnings.append(
                "Скрипт кнопки без проверки авторизации. Добавьте: from cpsk_auth import require_auth"
            )
        elif not has_require_auth_call:
            self.warnings.append(
                "Импорт require_auth есть, но функция не вызывается. Добавьте проверку в начале скрипта."
            )

    def check_notification_usage(self, lines, content):
        """Проверить использование единой системы уведомлений cpsk_notify."""
        if not self.current_file:
            return

        filename = os.path.basename(self.current_file)

        # Пропускаем библиотечные файлы
        if filename.startswith('cpsk_') or filename == 'pyrevit_checker.py':
            return

        # Пропускаем startup.py (там fallback на случай если cpsk_notify недоступен)
        if filename == 'startup.py':
            return

        # Ищем использование запрещённых методов уведомлений
        for i, line in enumerate(lines, 1):
            code_part = line.split('#')[0]

            # MessageBox.Show (кроме импортов)
            if 'MessageBox.Show' in code_part and 'import' not in code_part.lower():
                self.warnings.append(
                    "Строка {}: MessageBox.Show запрещён. Используйте cpsk_notify.show_error/warning/info/success".format(i)
                )

            # forms.alert
            if 'forms.alert' in code_part and 'import' not in code_part.lower():
                self.warnings.append(
                    "Строка {}: forms.alert запрещён. Используйте cpsk_notify.show_error/warning/info/success".format(i)
                )

            # output.print_md для сообщений об ошибках - должен использоваться cpsk_notify
            # Ловим только явные ошибки: "Ошибка", "Error", "**Error**", "failed", "не удалось"
            if 'output.print_md' in code_part:
                error_keywords = ['ошибк', 'error', 'failed', 'не удалось', 'невозможно', 'exception']
                line_lower = code_part.lower()
                for keyword in error_keywords:
                    if keyword in line_lower:
                        self.warnings.append(
                            "Строка {}: output.print_md для ошибок запрещён. Используйте cpsk_notify.show_error()".format(i)
                        )
                        break

    def check_icon_exists(self):
        """Проверить наличие icon.png в папках .pushbutton и .pulldown."""
        if not self.current_file:
            return

        filepath = self.current_file
        filename = os.path.basename(filepath)
        parent_dir = os.path.dirname(filepath)
        parent_name = os.path.basename(parent_dir)

        # Проверяем только script.py в .pushbutton или .pulldown папках
        if filename != 'script.py':
            return

        if not (parent_name.endswith('.pushbutton') or parent_name.endswith('.pulldown')):
            return

        # Проверяем наличие icon.png
        icon_path = os.path.join(parent_dir, 'icon.png')
        if not win_path_exists(icon_path):
            self.errors.append(
                "Отсутствует icon.png в папке {}".format(parent_name)
            )

    def get_report(self):
        """Получить отчет."""
        return {
            "file": self.current_file,
            "errors": self.errors,
            "warnings": self.warnings,
            "passed": len(self.errors) == 0
        }

    def print_report(self):
        """Вывести отчет."""
        filename = os.path.basename(self.current_file)

        if self.errors:
            print("\n[FAILED] {}".format(filename))
            for err in self.errors:
                print("  [X] {}".format(err))
        elif self.warnings:
            print("\n[PASSED] {} (с предупреждениями)".format(filename))
            for warn in self.warnings:
                print("  [!] {}".format(warn))
        else:
            print("[PASSED] {}".format(filename))


def check_directory(directory):
    """Проверить все .py файлы в директории."""
    checker = PyRevitChecker()
    results = []

    for root, dirs, files in os.walk(directory):
        # Пропустить __pycache__ и lib (там CPython скрипты)
        dirs[:] = [d for d in dirs if d not in ('__pycache__', 'lib')]

        for filename in files:
            if filename.endswith('.py') and filename != '__init__.py':
                filepath = os.path.join(root, filename)
                checker.check_file(filepath)
                checker.print_report()
                results.append(checker.get_report())

    return results


def main():
    """Главная функция."""
    if len(sys.argv) < 2:
        print("pyRevit Universal Code Checker")
        print("=" * 40)
        print("Использование:")
        print("  python pyrevit_checker.py <file.py>")
        print("  python pyrevit_checker.py <directory>")
        sys.exit(1)

    target = sys.argv[1]

    print("\npyRevit Universal Code Checker")
    print("=" * 40)

    if win_isfile(target):
        checker = PyRevitChecker()
        checker.check_file(target)
        checker.print_report()
        report = checker.get_report()
        print("\n" + "=" * 40)
        if report["passed"]:
            print("РЕЗУЛЬТАТ: PASSED")
        else:
            print("РЕЗУЛЬТАТ: FAILED ({} ошибок)".format(len(report["errors"])))
        sys.exit(0 if report["passed"] else 1)

    elif win_isdir(target):
        results = check_directory(target)
        passed = sum(1 for r in results if r["passed"])
        failed = len(results) - passed

        print("\n" + "=" * 40)
        print("ИТОГО: {} файлов".format(len(results)))
        print("  PASSED: {}".format(passed))
        print("  FAILED: {}".format(failed))
        print("=" * 40)
        sys.exit(0 if failed == 0 else 1)

    else:
        print("Ошибка: {} не существует".format(target))
        print("win_path_exists: {}".format(win_path_exists(target)))
        sys.exit(1)


if __name__ == "__main__":
    main()
