# -*- coding: utf-8 -*-
"""
Запуск Dynamo скриптов.
Поддержка 1000+ скриптов с поиском, категориями, избранным и историей.
"""

__title__ = "Запуск\nDynamo"
__author__ = "CPSK"

import clr
import os
import json
import time
import codecs
import random
import re
import webbrowser
from datetime import datetime

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, ListBox, TreeView, TreeNode,
    Panel, SplitContainer, Orientation, DockStyle,
    FormStartPosition, FormBorderStyle, SelectionMode,
    Padding, DialogResult, SendKeys, Clipboard, LinkLabel
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import revit, forms, script

# Добавляем lib в путь для импорта cpsk_auth
import sys
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Импорт модулей из lib
from cpsk_notify import show_error, show_info, show_success, show_warning, show_confirm
from cpsk_auth import require_auth

# Проверка авторизации
if not require_auth():
    sys.exit()


# === УТИЛИТЫ ДЛЯ ССЫЛОК ===

# Паттерн для поиска URL в тексте
URL_PATTERN = re.compile(r'(https?://[^\s<>"{}|\\^`\[\]]+)')

def find_urls(text):
    """Найти все URL в тексте и вернуть список (start, length, url)."""
    if not text:
        return []
    results = []
    for match in URL_PATTERN.finditer(text):
        results.append((match.start(), len(match.group(0)), match.group(0)))
    return results


# === НАСТРОЙКИ ===

# Путь к папке со скриптами - в папке 02_Dynamo.panel
PANEL_FOLDER = os.path.dirname(os.path.dirname(__file__))  # 02_Dynamo.panel
SCRIPTS_FOLDER = os.path.join(PANEL_FOLDER, "dynamo_scripts")
CONFIG_FILE = os.path.join(PANEL_FOLDER, "_config.yaml")
MAX_RECENT = 20
PAGE_SIZE = 100

# Отладка
output = script.get_output()


# === YAML ПАРСЕР ===

def parse_yaml(filepath):
    """Простой YAML парсер (списки и словари)."""
    if not os.path.exists(filepath):
        return {}

    result = {}
    current_key = None
    current_container = None  # list или dict
    container_type = None  # 'list' или 'dict'

    try:
        with codecs.open(filepath, 'r', 'utf-8') as f:
            for line in f:
                line = line.rstrip()
                if not line or line.strip().startswith('#'):
                    continue

                stripped = line.lstrip()
                indent = len(line) - len(stripped)

                if indent == 0 and ':' in stripped:
                    parts = stripped.split(':', 1)
                    key = parts[0].strip()
                    val = parts[1].strip() if len(parts) > 1 else ''

                    if val == '' or val == '[]' or val == '{}':
                        # Определим тип контейнера по первому элементу
                        if val == '{}':
                            result[key] = {}
                            container_type = 'dict'
                        else:
                            result[key] = []
                            container_type = 'list'
                        current_container = result[key]
                        current_key = key
                    else:
                        result[key] = val.strip('"\'')
                        current_key = None
                        current_container = None
                        container_type = None

                elif indent > 0 and current_key and current_container is not None:
                    if stripped.startswith('- '):
                        val = stripped[2:].strip().strip('"\'')
                        if isinstance(current_container, list):
                            current_container.append(val)
                    elif ':' in stripped:
                        # Это словарь: "key": "value"
                        parts = stripped.split(':', 1)
                        k = parts[0].strip().strip('"\'')
                        v = parts[1].strip().strip('"\'') if len(parts) > 1 else ''
                        if isinstance(current_container, list):
                            # Преобразуем в словарь
                            result[current_key] = {}
                            current_container = result[current_key]
                        current_container[k] = v
    except Exception as e:
        output.print_md("Ошибка чтения YAML: {}".format(str(e)))

    return result


def save_yaml(filepath, data):
    """Сохранить YAML файл (списки и словари)."""
    lines = []

    for key, val in data.items():
        if isinstance(val, list):
            if not val:
                lines.append("{}: []".format(key))
            else:
                lines.append("{}:".format(key))
                for item in val:
                    lines.append('  - "{}"'.format(item))
        elif isinstance(val, dict):
            if not val:
                lines.append("{}: {{}}".format(key))
            else:
                lines.append("{}:".format(key))
                for k, v in val.items():
                    lines.append('  "{}": "{}"'.format(k, v))
        else:
            lines.append('{}: "{}"'.format(key, val))

    try:
        with codecs.open(filepath, 'w', 'utf-8') as f:
            f.write('\n'.join(lines))
    except:
        pass


# === СКАНЕР СКРИПТОВ ===

class ScriptScanner:
    """Сканирует и кэширует Dynamo скрипты."""

    def __init__(self, scripts_folder):
        self.scripts_folder = scripts_folder
        self._cache = {}
        self._all_scripts = []

    def scan_categories(self):
        """Получить все папки-категории (только первый уровень)."""
        if not os.path.exists(self.scripts_folder):
            return []

        categories = []
        try:
            for name in os.listdir(self.scripts_folder):
                path = os.path.join(self.scripts_folder, name)
                if os.path.isdir(path) and not name.startswith('_'):
                    categories.append(name)
        except Exception as e:
            output.print_md("Ошибка сканирования: {}".format(str(e)))

        return sorted(categories)

    def scan_folder_tree(self, folder_path, rel_prefix=""):
        """Рекурсивно сканировать структуру папок."""
        result = []
        try:
            for name in sorted(os.listdir(folder_path)):
                path = os.path.join(folder_path, name)
                if os.path.isdir(path) and not name.startswith('_'):
                    rel_path = os.path.join(rel_prefix, name) if rel_prefix else name
                    children = self.scan_folder_tree(path, rel_path)
                    result.append({
                        'name': name,
                        'path': path,
                        'rel_path': rel_path,
                        'children': children
                    })
        except:
            pass
        return result

    def get_scripts_in_category(self, category):
        """Получить .dyn скрипты только в указанной папке (без подпапок)."""
        if category in self._cache:
            cached_time, cached_data = self._cache[category]
            if time.time() - cached_time < 30:
                return cached_data

        scripts = []
        category_path = os.path.join(self.scripts_folder, category)

        if os.path.exists(category_path):
            try:
                # Только файлы в текущей папке, без рекурсии
                for f in os.listdir(category_path):
                    if f.lower().endswith('.dyn'):
                        full_path = os.path.join(category_path, f)
                        if os.path.isfile(full_path):
                            rel_path = os.path.relpath(full_path, self.scripts_folder)
                            scripts.append({
                                'name': os.path.splitext(f)[0],
                                'path': full_path,
                                'rel_path': rel_path,
                                'category': category
                            })
            except:
                pass

        scripts.sort(key=lambda x: x['name'].lower())
        self._cache[category] = (time.time(), scripts)

        return scripts

    def get_all_scripts(self, force_rescan=False):
        """Получить все скрипты из всех папок рекурсивно."""
        if self._all_scripts and not force_rescan:
            return self._all_scripts

        all_scripts = []
        try:
            # Рекурсивно обходим все папки
            for root, dirs, files in os.walk(self.scripts_folder):
                # Пропускаем папки начинающиеся с _
                dirs[:] = [d for d in dirs if not d.startswith('_')]

                for f in files:
                    if f.lower().endswith('.dyn'):
                        full_path = os.path.join(root, f)
                        rel_path = os.path.relpath(full_path, self.scripts_folder)
                        category = os.path.dirname(rel_path) or "root"
                        all_scripts.append({
                            'name': os.path.splitext(f)[0],
                            'path': full_path,
                            'rel_path': rel_path,
                            'category': category
                        })
        except:
            pass

        all_scripts.sort(key=lambda x: x['name'].lower())
        self._all_scripts = all_scripts
        return all_scripts

    def search_scripts(self, query, scripts=None):
        """Поиск скриптов по имени."""
        if scripts is None:
            scripts = self.get_all_scripts()

        query = query.lower().strip()
        if not query:
            return scripts

        terms = query.split()
        results = []

        for s in scripts:
            name_lower = s['name'].lower()
            path_lower = s['rel_path'].lower()
            if all(term in name_lower or term in path_lower for term in terms):
                results.append(s)

        return results

    def get_script_info(self, script_path):
        """Получить метаданные из .dyn файла."""
        info = {'description': '', 'author': '', 'name': ''}

        try:
            with codecs.open(script_path, 'r', 'utf-8') as f:
                data = json.load(f)
                info['description'] = data.get('Description', '')
                info['author'] = data.get('Author', '')
                info['name'] = data.get('Name', '')
        except:
            pass

        return info

    def clear_cache(self):
        """Очистить кэш."""
        self._cache.clear()
        self._all_scripts = []


# === ДОВЕРЕННЫЕ ПАПКИ DYNAMO ===

def get_dynamo_settings_path():
    """Получить путь к настройкам Dynamo."""
    appdata = os.environ.get('APPDATA', '')
    dynamo_base = os.path.join(appdata, 'Dynamo', 'Dynamo Revit')

    # Найти последнюю версию (только числовые папки типа "2.18", "2.19")
    if os.path.exists(dynamo_base):
        versions = []
        for d in os.listdir(dynamo_base):
            dir_path = os.path.join(dynamo_base, d)
            if os.path.isdir(dir_path):
                # Проверяем, что это версия (начинается с цифры)
                if d and d[0].isdigit():
                    versions.append(d)
        if versions:
            # Сортируем версии (2.19 > 2.18 > 2.10)
            versions.sort(key=lambda x: [int(p) if p.isdigit() else 0 for p in x.split('.')], reverse=True)
            return os.path.join(dynamo_base, versions[0], 'DynamoSettings.xml')
    return None


def add_trusted_location(folder_path):
    """Добавить папку в доверенные локации Dynamo."""
    settings_path = get_dynamo_settings_path()
    if not settings_path or not os.path.exists(settings_path):
        output.print_md("Настройки Dynamo не найдены: {}".format(settings_path))
        return False

    try:
        # Читаем XML (IronPython 2.7 - используем codecs)
        with codecs.open(settings_path, 'r', 'utf-8') as f:
            content = f.read()

        # Проверяем, есть ли уже эта папка в TrustedLocations
        if folder_path in content:
            output.print_md("Папка уже в доверенных: {}".format(folder_path))
            return True

        # Ищем секцию TrustedLocations
        if '<TrustedLocations>' in content:
            # Добавляем новую локацию
            new_location = '    <string>{}</string>\n  </TrustedLocations>'.format(folder_path)
            content = content.replace('</TrustedLocations>', new_location)
        else:
            # Создаём секцию
            insert_point = '</PreferenceSettings>'
            new_section = '''  <TrustedLocations>
    <string>{}</string>
  </TrustedLocations>
</PreferenceSettings>'''.format(folder_path)
            content = content.replace(insert_point, new_section)

        # Сохраняем (IronPython 2.7 - используем codecs)
        with codecs.open(settings_path, 'w', 'utf-8') as f:
            f.write(content)

        output.print_md("Папка добавлена в доверенные: {}".format(folder_path))
        return True

    except Exception as e:
        output.print_md("Ошибка добавления в доверенные: {}".format(str(e)))
        return False


def add_dynamo_player_folder(folder_path):
    """Добавить папку в Dynamo Player."""
    settings_path = get_dynamo_settings_path()
    if not settings_path or not os.path.exists(settings_path):
        output.print_md("Настройки Dynamo не найдены")
        return False

    try:
        with codecs.open(settings_path, 'r', 'utf-8') as f:
            content = f.read()

        # Имя папки для отображения
        display_name = os.path.basename(folder_path)

        # Проверяем, есть ли уже папка в DynamoPlayerFolderGroups
        # Ищем по DisplayName, т.к. путь может быть с $USERDOC$
        if '<DisplayName>{}</DisplayName>'.format(display_name) in content:
            output.print_md("Папка уже в Dynamo Player: {}".format(display_name))
            return True

        # Уникальный ID (timestamp в миллисекундах)
        folder_id = str(int(time.time() * 1000) + random.randint(0, 999))

        # Новая запись папки
        new_folder_entry = '''
        <DynamoPlayerFolder>
          <Path>{}</Path>
          <DisplayName>{}</DisplayName>
          <Id>{}</Id>
          <IsRemovable>true</IsRemovable>
          <Order>1</Order>
          <IsValid>true</IsValid>
          <IsDefault>false</IsDefault>
        </DynamoPlayerFolder>'''.format(folder_path, display_name, folder_id)

        # Ищем место для вставки - перед </Folders>
        if '<DynamoPlayerFolderGroups>' in content and '</Folders>' in content:
            # Вставляем новую папку перед </Folders>
            content = content.replace('</Folders>', new_folder_entry + '\n      </Folders>', 1)

            with codecs.open(settings_path, 'w', 'utf-8') as f:
                f.write(content)

            output.print_md("Папка добавлена в Dynamo Player: {}".format(folder_path))
            return True

        elif '<DynamoPlayerFolderGroups />' in content or '<DynamoPlayerFolderGroups/>' in content:
            # Заменяем пустую секцию
            new_section = '''<DynamoPlayerFolderGroups>
    <DynamoPlayerFolderGroup>
      <EntryPoint>dplayer</EntryPoint>
      <Folders>
        <DynamoPlayerFolder>
          <Path>{}</Path>
          <DisplayName>{}</DisplayName>
          <Id>{}</Id>
          <IsRemovable>true</IsRemovable>
          <Order>1</Order>
          <IsValid>true</IsValid>
          <IsDefault>false</IsDefault>
        </DynamoPlayerFolder>
      </Folders>
    </DynamoPlayerFolderGroup>
  </DynamoPlayerFolderGroups>'''.format(folder_path, display_name, folder_id)

            if '<DynamoPlayerFolderGroups />' in content:
                content = content.replace('<DynamoPlayerFolderGroups />', new_section)
            else:
                content = content.replace('<DynamoPlayerFolderGroups/>', new_section)

            with codecs.open(settings_path, 'w', 'utf-8') as f:
                f.write(content)

            output.print_md("Папка добавлена в Dynamo Player: {}".format(folder_path))
            return True

        else:
            output.print_md("Секция DynamoPlayerFolderGroups не найдена")
            return False

    except Exception as e:
        output.print_md("Ошибка добавления в Dynamo Player: {}".format(str(e)))
        return False


# === ПРОВЕРКА СИНХРОНИЗАЦИИ ПАПОК ===

def get_synced_folders():
    """Получить список папок, уже добавленных в Dynamo Player."""
    settings_path = get_dynamo_settings_path()
    if not settings_path or not os.path.exists(settings_path):
        return set()

    synced = set()
    try:
        with codecs.open(settings_path, 'r', 'utf-8') as f:
            content = f.read()

        # Ищем все DisplayName в DynamoPlayerFolderGroups
        matches = re.findall(r'<DisplayName>([^<]+)</DisplayName>', content)
        for m in matches:
            synced.add(m)
    except:
        pass

    return synced


def get_all_subfolders(folder_path):
    """Рекурсивно получить все подпапки."""
    folders = []
    try:
        for name in os.listdir(folder_path):
            subfolder = os.path.join(folder_path, name)
            if os.path.isdir(subfolder) and not name.startswith('_'):
                folders.append((name, subfolder))
                # Рекурсивно добавляем вложенные папки
                folders.extend(get_all_subfolders(subfolder))
    except:
        pass
    return folders


def get_folders_to_sync():
    """Получить папки, которые нужно синхронизировать (включая вложенные)."""
    synced = get_synced_folders()
    to_sync = []

    # Получаем все подпапки рекурсивно
    all_folders = get_all_subfolders(SCRIPTS_FOLDER)

    for name, path in all_folders:
        if name not in synced:
            to_sync.append((name, path))

    return to_sync


def sync_all_folders():
    """Синхронизировать все папки с Dynamo Player."""
    to_sync = get_folders_to_sync()
    added = []

    for name, path in to_sync:
        if add_dynamo_player_folder(path):
            added.append(name)
        add_trusted_location(path)

    # Также проверяем корневую папку
    root_name = os.path.basename(SCRIPTS_FOLDER)
    synced = get_synced_folders()
    if root_name not in synced:
        if add_dynamo_player_folder(SCRIPTS_FOLDER):
            added.append(root_name)
        add_trusted_location(SCRIPTS_FOLDER)

    return added


# === ЗАПУСК DYNAMO PLAYER ===

import ctypes
from ctypes import wintypes
import subprocess

# Константы для subprocess (могут не быть в IronPython)
CREATE_NO_WINDOW = 0x08000000
STARTF_USESHOWWINDOW = 0x00000001

# Windows API для работы с окнами
user32 = ctypes.windll.user32
EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
GetWindowText = user32.GetWindowTextW
GetWindowTextLength = user32.GetWindowTextLengthW
SetForegroundWindow = user32.SetForegroundWindow
ShowWindow = user32.ShowWindow
IsWindowVisible = user32.IsWindowVisible
SW_RESTORE = 9

# Константы для SendInput
INPUT_KEYBOARD = 1
INPUT_MOUSE = 0
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
VK_CONTROL = 0x11
VK_RETURN = 0x0D
VK_TAB = 0x09

# Для получения размера окна
GetWindowRect = user32.GetWindowRect


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long)
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))
    ]


class INPUT_UNION(ctypes.Union):
    _fields_ = [
        ("ki", KEYBDINPUT),
        ("mi", MOUSEINPUT)
    ]


class INPUT(ctypes.Structure):
    _fields_ = [
        ("type", wintypes.DWORD),
        ("union", INPUT_UNION)
    ]


def click_at(x, y):
    """Кликнуть мышью в указанных координатах."""
    # Переместить курсор
    ctypes.windll.user32.SetCursorPos(x, y)
    import time as t
    t.sleep(0.1)

    # Клик
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    t.sleep(0.05)
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def send_unicode_char(char):
    """Отправить Unicode символ."""
    inputs = (INPUT * 2)()

    # Key down
    inputs[0].type = INPUT_KEYBOARD
    inputs[0].union.ki.wVk = 0
    inputs[0].union.ki.wScan = ord(char)
    inputs[0].union.ki.dwFlags = KEYEVENTF_UNICODE

    # Key up
    inputs[1].type = INPUT_KEYBOARD
    inputs[1].union.ki.wVk = 0
    inputs[1].union.ki.wScan = ord(char)
    inputs[1].union.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP

    user32.SendInput(2, ctypes.byref(inputs), ctypes.sizeof(INPUT))


def send_key(vk_code, key_up=False):
    """Отправить виртуальную клавишу."""
    inputs = (INPUT * 1)()
    inputs[0].type = INPUT_KEYBOARD
    inputs[0].union.ki.wVk = vk_code
    inputs[0].union.ki.dwFlags = KEYEVENTF_KEYUP if key_up else 0
    user32.SendInput(1, ctypes.byref(inputs), ctypes.sizeof(INPUT))


def send_tab():
    """Отправить Tab."""
    send_key(VK_TAB)
    import time as t
    t.sleep(0.05)
    send_key(VK_TAB, True)


def type_string(text):
    """Напечатать строку посимвольно."""
    import time as t
    for char in text:
        send_unicode_char(char)
        t.sleep(0.03)


def copy_to_clipboard(text):
    """Скопировать текст в буфер обмена."""
    try:
        Clipboard.SetText(text)
        return True
    except:
        return False


def find_window_by_title(title_part):
    """Найти окно по части заголовка."""
    result = []

    def callback(hwnd, lparam):
        if IsWindowVisible(hwnd):
            length = GetWindowTextLength(hwnd)
            if length > 0:
                buff = ctypes.create_unicode_buffer(length + 1)
                GetWindowText(hwnd, buff, length + 1)
                if title_part.lower() in buff.value.lower():
                    result.append(hwnd)
        return True

    EnumWindows(EnumWindowsProc(callback), 0)
    return result[0] if result else None


def activate_window(hwnd):
    """Активировать окно."""
    try:
        ShowWindow(hwnd, SW_RESTORE)
        SetForegroundWindow(hwnd)
        return True
    except:
        return False


import base64

def create_autotype_script(script_name):
    """Запустить PowerShell код для автоматического ввода текста."""
    debug_file = os.path.join(SCRIPTS_FOLDER, "_debug.txt")

    # PowerShell скрипт для отложенного ввода
    ps_content = '''
# Отладка
$debugFile = "{debug_file}"
function Log($msg) {{
    Add-Content -Path $debugFile -Value $msg -Encoding UTF8
}}

Log "=== PowerShell autotype script ==="
Log "Script name: {script_name}"

# Добавляем необходимые типы
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinAPI {{
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {{
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }}
}}
"@

# Глобальные переменные для хранения результатов поиска
$script:foundHwnd = [IntPtr]::Zero
$script:foundTitle = ""
$script:foundLeft = 0
$script:foundTop = 0
$script:foundWidth = 0
$script:foundHeight = 0

# Функция поиска окна Dynamo Player с валидными размерами
function Find-DynamoWindow {{
    $script:foundHwnd = [IntPtr]::Zero

    $callback = {{
        param($hwnd, $lparam)

        if ([WinAPI]::IsWindowVisible($hwnd)) {{
            $sb = New-Object System.Text.StringBuilder 256
            [void][WinAPI]::GetWindowText($hwnd, $sb, 256)
            $title = $sb.ToString().Trim()

            # Ищем окно с "Dynamo" в названии (но не Revit и не .dyn файлы)
            if ($title -like "*Dynamo*" -and $title -notlike "*Revit*" -and $title -notlike "*.dyn*") {{
                # Проверяем размеры окна
                $rect = New-Object WinAPI+RECT
                [WinAPI]::GetWindowRect($hwnd, [ref]$rect)
                $w = $rect.Right - $rect.Left
                $h = $rect.Bottom - $rect.Top

                if ($w -gt 100 -and $h -gt 100) {{
                    # Сохраняем всё сразу
                    $script:foundHwnd = $hwnd
                    $script:foundTitle = $title
                    $script:foundLeft = $rect.Left
                    $script:foundTop = $rect.Top
                    $script:foundWidth = $w
                    $script:foundHeight = $h
                    return $false  # Stop enumeration
                }}
            }}
        }}
        return $true  # Continue enumeration
    }}

    [WinAPI]::EnumWindows($callback, [IntPtr]::Zero)
    return $script:foundHwnd
}}

# Начальная задержка - даём Revit начать открывать окно
Start-Sleep -Milliseconds 800

# Периодически ищем окно с валидными размерами (макс 10 секунд, проверка каждые 300мс)
$maxAttempts = 30
$attempt = 0
$targetHwnd = [IntPtr]::Zero

Log "Searching for Dynamo Player window..."

while ($attempt -lt $maxAttempts) {{
    $hwnd = Find-DynamoWindow
    if ($hwnd -ne [IntPtr]::Zero) {{
        Log "Found window: '$foundTitle' after $($attempt * 300 + 800)ms"
        Log "Saved rect: Left=$foundLeft, Top=$foundTop, Width=$foundWidth, Height=$foundHeight"
        break
    }}
    Start-Sleep -Milliseconds 300
    $attempt++
}}

if ($foundHwnd -ne [IntPtr]::Zero) {{
    Log "Activating window..."
    [WinAPI]::ShowWindow($foundHwnd, 9)  # SW_RESTORE
    [WinAPI]::SetForegroundWindow($foundHwnd)
    Start-Sleep -Milliseconds 300

    # Используем сохранённые координаты
    Log "Using saved rect: Left=$foundLeft, Top=$foundTop, Width=$foundWidth, Height=$foundHeight"

    # Кликаем в поле поиска (строка "Поиск" - примерно 90px от верха)
    $clickX = $foundLeft + 300
    $clickY = $foundTop + 90

    Log "Clicking at: X=$clickX, Y=$clickY"

    [WinAPI]::SetCursorPos($clickX, $clickY)
    Start-Sleep -Milliseconds 100
    [WinAPI]::mouse_event([WinAPI]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    Start-Sleep -Milliseconds 50
    [WinAPI]::mouse_event([WinAPI]::MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    Start-Sleep -Milliseconds 300

    Log "Sending keys: {script_name}"

    # Используем SendKeys для ввода текста
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait("{script_name}")

    Log "Done!"
}} else {{
    Log "ERROR: Window not found after 10 seconds!"
}}

Log "=== Script finished ==="
'''.format(
        debug_file=debug_file.replace("\\", "\\\\"),
        script_name=script_name.replace("{", "{{").replace("}", "}}")
    )

    try:
        # Очищаем debug файл
        with codecs.open(debug_file, 'w', 'utf-8') as f:
            f.write("=== Starting autotype ===\n")

        # Кодируем скрипт в Base64 для передачи через -EncodedCommand
        # PowerShell ожидает UTF-16 LE
        ps_bytes = ps_content.encode('utf-16-le')
        ps_base64 = base64.b64encode(ps_bytes).decode('ascii')

        # Запускаем PowerShell в фоновом режиме
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0  # SW_HIDE

        subprocess.Popen(
            ["powershell.exe", "-WindowStyle", "Hidden", "-EncodedCommand", ps_base64],
            startupinfo=startupinfo,
            creationflags=CREATE_NO_WINDOW
        )

        return True
    except Exception as e:
        try:
            with codecs.open(debug_file, 'a', 'utf-8') as f:
                f.write("ERROR: {}\n".format(str(e)))
        except:
            pass
        return False


def run_dynamo_script(script_path, category_name):
    """Запустить Dynamo скрипт через Dynamo Player."""
    script_name = os.path.splitext(os.path.basename(script_path))[0]

    # Способ 1: Открыть Dynamo Player через PostableCommand
    try:
        from Autodesk.Revit.UI import PostableCommand, RevitCommandId

        uiapp = revit.HOST_APP.uiapp

        # Команда открытия Dynamo Player
        cmd_id = RevitCommandId.LookupPostableCommandId(PostableCommand.DynamoPlayer)

        if cmd_id:
            # Копируем имя скрипта в буфер обмена (для ручной вставки Ctrl+V если нужно)
            copy_to_clipboard(script_name)

            # Запускаем PowerShell для автоматического ввода (ждёт появления окна)
            create_autotype_script(script_name)

            # Открываем Dynamo Player
            uiapp.PostCommand(cmd_id)
            return True, script_name, category_name

    except Exception as e1:
        output.print_md("PostableCommand error: {}".format(str(e1)))

    # Способ 2: Открыть в Dynamo напрямую
    try:
        os.startfile(script_path)
        return True, script_name, None
    except Exception as e2:
        return False, "Ошибка запуска: {}".format(str(e2)), None


# === ГЛАВНОЕ ОКНО ===

class DynamoLauncherForm(Form):
    """Диалог выбора и запуска Dynamo скриптов."""

    def __init__(self):
        self.scanner = ScriptScanner(SCRIPTS_FOLDER)
        self.config = parse_yaml(CONFIG_FILE)
        self.current_scripts = []
        self.current_page = 0
        self.selected_script = None
        self.recent = []
        self.favorites = []
        self.folders_synced = False

        self.load_config()
        self.setup_form()
        self.load_categories()

    def load_config(self):
        """Загрузить конфигурацию."""
        self.recent = self.config.get('recent', [])
        if not isinstance(self.recent, list):
            self.recent = []

        self.favorites = self.config.get('favorites', [])
        if not isinstance(self.favorites, list):
            self.favorites = []

        # Статистика запусков: {rel_path: count}
        self.run_counts = self.config.get('run_counts', {})
        if not isinstance(self.run_counts, dict):
            self.run_counts = {}

        # Даты последнего запуска: {rel_path: timestamp}
        self.last_runs = self.config.get('last_runs', {})
        if not isinstance(self.last_runs, dict):
            self.last_runs = {}

    def save_config(self):
        """Сохранить конфигурацию."""
        self.config['recent'] = self.recent[:MAX_RECENT]
        self.config['favorites'] = self.favorites
        self.config['run_counts'] = self.run_counts
        self.config['last_runs'] = self.last_runs
        save_yaml(CONFIG_FILE, self.config)

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Запуск Dynamo скриптов - CPSK"
        self.Width = 900
        self.Height = 550
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MinimumSize = Size(700, 400)

        # === ВЕРХНЯЯ ПАНЕЛЬ (Поиск) ===
        top_panel = Panel()
        top_panel.Dock = DockStyle.Top
        top_panel.Height = 40

        lbl_search = Label()
        lbl_search.Text = "Поиск:"
        lbl_search.Location = Point(10, 12)
        lbl_search.AutoSize = True
        top_panel.Controls.Add(lbl_search)

        self.txt_search = TextBox()
        self.txt_search.Location = Point(60, 10)
        self.txt_search.Width = 250
        self.txt_search.TextChanged += self.on_search_changed
        top_panel.Controls.Add(self.txt_search)

        self.lbl_count = Label()
        self.lbl_count.Location = Point(320, 12)
        self.lbl_count.AutoSize = True
        self.lbl_count.Text = ""
        top_panel.Controls.Add(self.lbl_count)

        btn_refresh = Button()
        btn_refresh.Text = "Обновить"
        btn_refresh.Location = Point(450, 8)
        btn_refresh.Width = 80
        btn_refresh.Click += self.on_refresh_click
        top_panel.Controls.Add(btn_refresh)

        btn_open_folder = Button()
        btn_open_folder.Text = "Открыть папку"
        btn_open_folder.Location = Point(540, 8)
        btn_open_folder.Width = 100
        btn_open_folder.Click += self.on_open_folder_click
        top_panel.Controls.Add(btn_open_folder)

        # === НИЖНЯЯ ПАНЕЛЬ (Кнопки) ===
        bottom_panel = Panel()
        bottom_panel.Dock = DockStyle.Bottom
        bottom_panel.Height = 45

        self.btn_sync = Button()
        self.btn_sync.Text = "Синхр. папки"
        self.btn_sync.Location = Point(10, 10)
        self.btn_sync.Width = 100
        self.btn_sync.Height = 30
        self.btn_sync.Click += self.on_sync_click
        self.btn_sync.BackColor = Color.LightYellow
        bottom_panel.Controls.Add(self.btn_sync)

        self.btn_run = Button()
        self.btn_run.Text = "Запустить"
        self.btn_run.Location = Point(120, 10)
        self.btn_run.Width = 100
        self.btn_run.Height = 30
        self.btn_run.Click += self.on_run_click
        self.btn_run.Enabled = False
        bottom_panel.Controls.Add(self.btn_run)

        self.btn_favorite = Button()
        self.btn_favorite.Text = "В избранное"
        self.btn_favorite.Location = Point(230, 10)
        self.btn_favorite.Width = 100
        self.btn_favorite.Height = 30
        self.btn_favorite.Click += self.on_favorite_click
        self.btn_favorite.Enabled = False
        bottom_panel.Controls.Add(self.btn_favorite)

        self.lbl_info = Label()
        self.lbl_info.Location = Point(340, 15)
        self.lbl_info.Width = 420
        self.lbl_info.ForeColor = Color.Gray
        self.lbl_info.Text = "Выберите скрипт для запуска"
        bottom_panel.Controls.Add(self.lbl_info)

        # Проверить статус синхронизации
        self.update_sync_status()

        # === ОСНОВНАЯ ПАНЕЛЬ ===
        main_split = SplitContainer()
        main_split.Dock = DockStyle.Fill
        main_split.Orientation = Orientation.Vertical
        main_split.SplitterDistance = 80

        # === ЛЕВАЯ ПАНЕЛЬ (Категории) ===
        lbl_cat = Label()
        lbl_cat.Text = "Категории"
        lbl_cat.Dock = DockStyle.Top
        lbl_cat.Height = 25
        lbl_cat.Font = Font(lbl_cat.Font, FontStyle.Bold)

        self.tree_categories = TreeView()
        self.tree_categories.Dock = DockStyle.Fill
        self.tree_categories.AfterSelect += self.on_category_selected

        # ВАЖНО: порядок добавления для Dock - сначала Fill, потом Top
        # НО контролы добавляются в обратном порядке визуально
        # Поэтому: добавляем Top ПОСЛЕДНИМ чтобы он был сверху
        main_split.Panel1.Controls.Add(self.tree_categories)
        main_split.Panel1.Controls.Add(lbl_cat)

        # === ПРАВАЯ ПАНЕЛЬ (Скрипты) ===
        lbl_scripts = Label()
        lbl_scripts.Text = "Скрипты"
        lbl_scripts.Dock = DockStyle.Top
        lbl_scripts.Height = 25
        lbl_scripts.Font = Font(lbl_scripts.Font, FontStyle.Bold)

        # Пагинация
        page_panel = Panel()
        page_panel.Dock = DockStyle.Bottom
        page_panel.Height = 25

        self.btn_prev = Button()
        self.btn_prev.Text = "< Назад"
        self.btn_prev.Location = Point(5, 3)
        self.btn_prev.Width = 70
        self.btn_prev.Click += self.on_prev_page
        self.btn_prev.Enabled = False
        page_panel.Controls.Add(self.btn_prev)

        self.lbl_page = Label()
        self.lbl_page.Location = Point(85, 7)
        self.lbl_page.AutoSize = True
        self.lbl_page.Text = ""
        page_panel.Controls.Add(self.lbl_page)

        self.btn_next = Button()
        self.btn_next.Text = "Далее >"
        self.btn_next.Location = Point(180, 3)
        self.btn_next.Width = 70
        self.btn_next.Click += self.on_next_page
        self.btn_next.Enabled = False
        page_panel.Controls.Add(self.btn_next)

        # Список скриптов
        self.list_scripts = ListBox()
        self.list_scripts.Dock = DockStyle.Fill
        self.list_scripts.SelectionMode = SelectionMode.One
        self.list_scripts.SelectedIndexChanged += self.on_script_selected
        self.list_scripts.DoubleClick += self.on_run_click

        # Панель описания скрипта
        desc_panel = Panel()
        desc_panel.Dock = DockStyle.Bottom
        desc_panel.Height = 85
        desc_panel.Padding = Padding(3)

        # Статистика (автор, запуски)
        self.lbl_stats = Label()
        self.lbl_stats.Dock = DockStyle.Top
        self.lbl_stats.Height = 16
        self.lbl_stats.Text = ""
        self.lbl_stats.ForeColor = Color.DimGray

        lbl_desc_title = Label()
        lbl_desc_title.Text = "Описание:"
        lbl_desc_title.Dock = DockStyle.Top
        lbl_desc_title.Height = 16
        lbl_desc_title.Font = Font(lbl_desc_title.Font, FontStyle.Bold)

        self.lbl_description = LinkLabel()
        self.lbl_description.Dock = DockStyle.Fill
        self.lbl_description.Text = ""
        self.lbl_description.ForeColor = Color.DarkGray
        self.lbl_description.LinkClicked += self.on_link_clicked

        desc_panel.Controls.Add(self.lbl_description)
        desc_panel.Controls.Add(lbl_desc_title)
        desc_panel.Controls.Add(self.lbl_stats)

        # ВАЖНО: Порядок добавления контролов в WinForms
        # Контролы добавляются "снизу вверх" в z-order
        # Fill должен быть добавлен ПЕРВЫМ, Top/Bottom - ПОСЛЕ
        main_split.Panel2.Controls.Add(self.list_scripts)  # Fill - первый
        main_split.Panel2.Controls.Add(desc_panel)          # Bottom (описание)
        main_split.Panel2.Controls.Add(page_panel)          # Bottom (пагинация)
        main_split.Panel2.Controls.Add(lbl_scripts)         # Top - последний

        # ВАЖНО: Порядок добавления в главную форму
        # Fill добавляется ПОСЛЕДНИМ
        self.Controls.Add(main_split)   # Fill - последний
        self.Controls.Add(bottom_panel) # Bottom
        self.Controls.Add(top_panel)    # Top

    def load_categories(self):
        """Загрузить дерево категорий."""
        self.tree_categories.Nodes.Clear()

        # Специальные узлы
        all_node = TreeNode("Все скрипты")
        all_node.Tag = "__all__"
        self.tree_categories.Nodes.Add(all_node)

        recent_node = TreeNode("Недавние")
        recent_node.Tag = "__recent__"
        self.tree_categories.Nodes.Add(recent_node)

        fav_node = TreeNode("Избранное")
        fav_node.Tag = "__favorites__"
        self.tree_categories.Nodes.Add(fav_node)

        # Папки-категории с вложенной структурой
        folder_tree = self.scanner.scan_folder_tree(SCRIPTS_FOLDER)
        self._add_folder_nodes(self.tree_categories.Nodes, folder_tree)

        # Развернуть все узлы
        self.tree_categories.ExpandAll()

        # Выбрать "Все скрипты"
        if self.tree_categories.Nodes.Count > 0:
            self.tree_categories.SelectedNode = self.tree_categories.Nodes[0]

    def _add_folder_nodes(self, parent_nodes, folders):
        """Рекурсивно добавить узлы папок в дерево."""
        for folder in folders:
            node = TreeNode(folder['name'])
            node.Tag = folder['rel_path']
            parent_nodes.Add(node)
            if folder['children']:
                self._add_folder_nodes(node.Nodes, folder['children'])

    def on_category_selected(self, sender, args):
        """Выбрана категория."""
        node = self.tree_categories.SelectedNode
        if node is None:
            return

        tag = node.Tag
        if tag is None:
            return

        self.current_page = 0
        tag_str = str(tag)

        if tag_str == "__all__":
            self.current_scripts = self.scanner.get_all_scripts()
        elif tag_str == "__recent__":
            all_scripts = self.scanner.get_all_scripts()
            recent_set = set(self.recent)
            self.current_scripts = [s for s in all_scripts if s['rel_path'] in recent_set]
        elif tag_str == "__favorites__":
            all_scripts = self.scanner.get_all_scripts()
            fav_set = set(self.favorites)
            self.current_scripts = [s for s in all_scripts if s['rel_path'] in fav_set]
        else:
            self.current_scripts = self.scanner.get_scripts_in_category(tag_str)

        # Применить поиск
        query = self.txt_search.Text
        if query:
            self.current_scripts = self.scanner.search_scripts(query, self.current_scripts)

        self.update_scripts_list()

    def on_search_changed(self, sender, args):
        """Изменён поиск."""
        self.on_category_selected(None, None)

    def update_scripts_list(self):
        """Обновить список скриптов."""
        self.list_scripts.Items.Clear()
        self.selected_script = None
        self.btn_run.Enabled = False
        self.btn_favorite.Enabled = False
        self.lbl_description.Text = ""
        self.lbl_description.Links.Clear()
        self.lbl_stats.Text = ""
        if not self.folders_synced:
            self.lbl_info.Text = "Сначала синхронизируйте папки!"
            self.lbl_info.ForeColor = Color.Red
        else:
            self.lbl_info.Text = "Выберите скрипт для запуска"
            self.lbl_info.ForeColor = Color.Gray

        total = len(self.current_scripts)
        start = self.current_page * PAGE_SIZE
        end = min(start + PAGE_SIZE, total)

        page_scripts = self.current_scripts[start:end]

        for s in page_scripts:
            display = "{} ({})".format(s['name'], s['category'])
            self.list_scripts.Items.Add(display)

        # Счётчик
        self.lbl_count.Text = "Найдено: {}".format(total)

        # Пагинация
        total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
        current = self.current_page + 1

        self.lbl_page.Text = "Стр. {} из {}".format(current, total_pages)
        self.btn_prev.Enabled = self.current_page > 0
        self.btn_next.Enabled = end < total

    def on_prev_page(self, sender, args):
        """Предыдущая страница."""
        if self.current_page > 0:
            self.current_page -= 1
            self.update_scripts_list()

    def on_next_page(self, sender, args):
        """Следующая страница."""
        total = len(self.current_scripts)
        if (self.current_page + 1) * PAGE_SIZE < total:
            self.current_page += 1
            self.update_scripts_list()

    def on_script_selected(self, sender, args):
        """Выбран скрипт."""
        idx = self.list_scripts.SelectedIndex
        if idx < 0:
            self.selected_script = None
            self.btn_run.Enabled = False
            self.btn_favorite.Enabled = False
            self.lbl_description.Text = ""
            self.lbl_description.Links.Clear()
            self.lbl_stats.Text = ""
            if not self.folders_synced:
                self.lbl_info.Text = "Сначала синхронизируйте папки!"
                self.lbl_info.ForeColor = Color.Red
            else:
                self.lbl_info.Text = "Выберите скрипт для запуска"
                self.lbl_info.ForeColor = Color.Gray
            return

        actual_idx = self.current_page * PAGE_SIZE + idx
        if actual_idx < len(self.current_scripts):
            self.selected_script = self.current_scripts[actual_idx]
            rel_path = self.selected_script['rel_path']

            # Информация из .dyn файла
            info = self.scanner.get_script_info(self.selected_script['path'])
            desc = info.get('description', '')
            author = info.get('author', '')

            # Статистика запусков
            run_count = int(self.run_counts.get(rel_path, 0))
            last_run = self.last_runs.get(rel_path, '')

            # Формируем строку статистики
            stats_parts = []
            if author:
                stats_parts.append("Автор: {}".format(author))
            stats_parts.append("Запусков: {}".format(run_count))
            if last_run:
                stats_parts.append("Последний: {}".format(last_run))
            self.lbl_stats.Text = "  |  ".join(stats_parts)

            # Показываем описание в панели
            if desc:
                self.lbl_description.Text = desc
                self.lbl_description.ForeColor = Color.Black
                # Настраиваем ссылки
                self.lbl_description.Links.Clear()
                urls = find_urls(desc)
                for start, length, url in urls:
                    link = self.lbl_description.Links.Add(start, length, url)
            else:
                self.lbl_description.Text = "Нет описания"
                self.lbl_description.ForeColor = Color.Gray
                self.lbl_description.Links.Clear()

            # Путь к файлу внизу
            self.lbl_info.Text = rel_path
            self.lbl_info.ForeColor = Color.Gray

            # Кнопки - запуск только если папки синхронизированы
            self.btn_run.Enabled = self.folders_synced
            self.btn_favorite.Enabled = True

            # Предупреждение если не синхронизировано
            if not self.folders_synced:
                self.lbl_info.Text = "Синхронизируйте папки для запуска!"
                self.lbl_info.ForeColor = Color.Orange

            # Текст кнопки избранного
            if rel_path in self.favorites:
                self.btn_favorite.Text = "Из избранного"
            else:
                self.btn_favorite.Text = "В избранное"

    def update_sync_status(self):
        """Обновить статус синхронизации папок."""
        to_sync = get_folders_to_sync()
        if to_sync:
            self.btn_sync.Text = "Синхр. ({})".format(len(to_sync))
            self.btn_sync.BackColor = Color.LightYellow
            self.folders_synced = False
        else:
            self.btn_sync.Text = "Синхронизировано"
            self.btn_sync.BackColor = Color.LightGreen
            self.folders_synced = True

    def on_sync_click(self, sender, args):
        """Синхронизировать папки с Dynamo Player."""
        to_sync = get_folders_to_sync()

        if not to_sync:
            show_info("Синхронизация", "Все папки уже синхронизированы с Dynamo Player")
            return

        # Показать какие папки будут добавлены
        folder_names = [name for name, path in to_sync]
        details = "Будут добавлены папки:\n- " + "\n- ".join(folder_names)

        if not show_confirm("Синхронизация папок", "Добавить папки в Dynamo Player?",
                            details=details):
            return

        added = sync_all_folders()
        self.update_sync_status()
        self.update_scripts_list()  # Обновить UI после синхронизации

        if added:
            show_warning("Синхронизация завершена",
                         "ПЕРЕЗАПУСТИТЕ REVIT для применения изменений!",
                         details="Добавлены папки:\n- " + "\n- ".join(added))
        else:
            show_info("Синхронизация", "Папки уже были добавлены ранее")

    def on_run_click(self, sender, args):
        """Запустить скрипт."""
        if not self.selected_script:
            return

        script_path = self.selected_script['path']
        rel_path = self.selected_script['rel_path']
        category = self.selected_script['category']

        # Добавить в недавние
        if rel_path in self.recent:
            self.recent.remove(rel_path)
        self.recent.insert(0, rel_path)
        self.recent = self.recent[:MAX_RECENT]

        # Увеличить счётчик запусков
        current_count = int(self.run_counts.get(rel_path, 0))
        self.run_counts[rel_path] = str(current_count + 1)

        # Сохранить дату последнего запуска
        self.last_runs[rel_path] = datetime.now().strftime("%d.%m.%Y %H:%M")

        self.save_config()

        # Запустить
        success, script_name, folder_name = run_dynamo_script(script_path, category)

        if success:
            # Закрываем диалог сразу - PowerShell вставит текст в Dynamo Player
            self.DialogResult = DialogResult.OK
            self.Close()
        else:
            show_error("Ошибка", "Ошибка запуска скрипта",
                       details=script_name)

    def on_favorite_click(self, sender, args):
        """Добавить/удалить из избранного."""
        if not self.selected_script:
            return

        rel_path = self.selected_script['rel_path']

        if rel_path in self.favorites:
            self.favorites.remove(rel_path)
            self.btn_favorite.Text = "В избранное"
        else:
            self.favorites.append(rel_path)
            self.btn_favorite.Text = "Из избранного"

        self.save_config()

    def on_open_folder_click(self, sender, args):
        """Открыть папку со скриптами."""
        try:
            if os.path.exists(SCRIPTS_FOLDER):
                os.startfile(SCRIPTS_FOLDER)
            else:
                show_warning("Внимание", "Папка не найдена",
                             details="Путь: {}".format(SCRIPTS_FOLDER))
        except:
            pass

    def on_refresh_click(self, sender, args):
        """Обновить список."""
        self.scanner.clear_cache()
        self.load_categories()

    def on_link_clicked(self, sender, args):
        """Открыть ссылку в браузере."""
        try:
            url = args.Link.LinkData
            if url:
                webbrowser.open(url)
        except Exception as e:
            output.print_md("Ошибка открытия ссылки: {}".format(str(e)))


# === MAIN ===
if __name__ == "__main__":
    # Создать папку если нет
    if not os.path.exists(SCRIPTS_FOLDER):
        try:
            os.makedirs(SCRIPTS_FOLDER)
            os.makedirs(os.path.join(SCRIPTS_FOLDER, "Examples"))
        except Exception as e:
            output.print_md("Ошибка создания папки: {}".format(str(e)))

        show_info("Первый запуск", "Создана папка для Dynamo скриптов",
                  details="Путь: {}\n\nДобавьте .dyn файлы в подпапки-категории.".format(SCRIPTS_FOLDER))

    # ВАЖНО: В pyRevit используем ShowDialog(), НЕ Application.Run()!
    form = DynamoLauncherForm()
    form.ShowDialog()
