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
