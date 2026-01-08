# -*- coding: utf-8 -*-
"""
Загрузка и запуск Dynamo скриптов с сервера CPSK.
"""

__title__ = "Скрипты\nс сервера"
__author__ = "CPSK"

import clr
import os
import codecs
import time
import random
import base64
import subprocess

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, ListBox, ComboBox,
    Panel, DockStyle, FormStartPosition, FormBorderStyle,
    SelectionMode, Padding, DialogResult, Clipboard,
    ComboBoxStyle
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import revit, script

# Добавляем lib в путь
import sys
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))),
    "lib"
)
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from cpsk_notify import show_error, show_info, show_success, show_warning
from cpsk_auth import require_auth
from cpsk_dynamo_api import DynamoApiClient, get_downloaded_scripts_folder

# Проверка авторизации
if not require_auth():
    sys.exit()

output = script.get_output()

# Папка для скачанных скриптов
DOWNLOAD_FOLDER = get_downloaded_scripts_folder()

# Константы для subprocess
CREATE_NO_WINDOW = 0x08000000
STARTF_USESHOWWINDOW = 0x00000001


# === ДОВЕРЕННЫЕ ПАПКИ DYNAMO ===

def get_dynamo_settings_path():
    """Получить путь к настройкам Dynamo."""
    appdata = os.environ.get('APPDATA', '')
    dynamo_base = os.path.join(appdata, 'Dynamo', 'Dynamo Revit')

    if os.path.exists(dynamo_base):
        versions = []
        for d in os.listdir(dynamo_base):
            dir_path = os.path.join(dynamo_base, d)
            if os.path.isdir(dir_path) and d and d[0].isdigit():
                versions.append(d)
        if versions:
            versions.sort(key=lambda x: [int(p) if p.isdigit() else 0 for p in x.split('.')], reverse=True)
            return os.path.join(dynamo_base, versions[0], 'DynamoSettings.xml')
    return None


def add_dynamo_player_folder(folder_path):
    """Добавить папку в Dynamo Player."""
    settings_path = get_dynamo_settings_path()
    if not settings_path or not os.path.exists(settings_path):
        return False

    try:
        with codecs.open(settings_path, 'r', 'utf-8') as f:
            content = f.read()

        display_name = os.path.basename(folder_path)

        if '<DisplayName>{}</DisplayName>'.format(display_name) in content:
            return True  # Уже добавлена

        folder_id = str(int(time.time() * 1000) + random.randint(0, 999))

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

        if '<DynamoPlayerFolderGroups>' in content and '</Folders>' in content:
            content = content.replace('</Folders>', new_folder_entry + '\n      </Folders>', 1)
            with codecs.open(settings_path, 'w', 'utf-8') as f:
                f.write(content)
            return True

        elif '<DynamoPlayerFolderGroups />' in content or '<DynamoPlayerFolderGroups/>' in content:
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
            return True

    except Exception:
        pass

    return False


def add_trusted_location(folder_path):
    """Добавить папку в доверенные локации Dynamo."""
    settings_path = get_dynamo_settings_path()
    if not settings_path or not os.path.exists(settings_path):
        return False

    try:
        with codecs.open(settings_path, 'r', 'utf-8') as f:
            content = f.read()

        if folder_path in content:
            return True

        if '<TrustedLocations>' in content:
            new_location = '    <string>{}</string>\n  </TrustedLocations>'.format(folder_path)
            content = content.replace('</TrustedLocations>', new_location)
        else:
            insert_point = '</PreferenceSettings>'
            new_section = '''  <TrustedLocations>
    <string>{}</string>
  </TrustedLocations>
</PreferenceSettings>'''.format(folder_path)
            content = content.replace(insert_point, new_section)

        with codecs.open(settings_path, 'w', 'utf-8') as f:
            f.write(content)

        return True

    except Exception:
        pass

    return False


def is_folder_in_dynamo_player(folder_path):
    """Проверить, добавлена ли папка в Dynamo Player."""
    settings_path = get_dynamo_settings_path()
    if not settings_path or not os.path.exists(settings_path):
        return False

    try:
        with codecs.open(settings_path, 'r', 'utf-8') as f:
            content = f.read()

        display_name = os.path.basename(folder_path)
        return '<DisplayName>{}</DisplayName>'.format(display_name) in content
    except Exception:
        pass

    return False


def ensure_download_folder_in_dynamo():
    """Убедиться, что папка скачанных скриптов добавлена в Dynamo Player."""
    if not os.path.exists(DOWNLOAD_FOLDER):
        os.makedirs(DOWNLOAD_FOLDER)

    if not is_folder_in_dynamo_player(DOWNLOAD_FOLDER):
        add_dynamo_player_folder(DOWNLOAD_FOLDER)
        add_trusted_location(DOWNLOAD_FOLDER)
        return True  # Нужен перезапуск
    return False


# === ЗАПУСК DYNAMO PLAYER ===

def copy_to_clipboard(text):
    """Скопировать текст в буфер обмена."""
    try:
        Clipboard.SetText(text)
        return True
    except Exception:
        return False


def create_autotype_script(script_name):
    """Запустить PowerShell для автоматического ввода текста в Dynamo Player."""
    debug_file = os.path.join(DOWNLOAD_FOLDER, "_debug.txt")

    ps_content = '''
$debugFile = "{debug_file}"
function Log($msg) {{
    Add-Content -Path $debugFile -Value $msg -Encoding UTF8
}}

Log "=== PowerShell autotype script ==="
Log "Script name: {script_name}"

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

$script:foundHwnd = [IntPtr]::Zero
$script:foundTitle = ""
$script:foundLeft = 0
$script:foundTop = 0
$script:foundWidth = 0
$script:foundHeight = 0

function Find-DynamoWindow {{
    $script:foundHwnd = [IntPtr]::Zero

    $callback = {{
        param($hwnd, $lparam)

        if ([WinAPI]::IsWindowVisible($hwnd)) {{
            $sb = New-Object System.Text.StringBuilder 256
            [void][WinAPI]::GetWindowText($hwnd, $sb, 256)
            $title = $sb.ToString().Trim()

            if ($title -like "*Dynamo*" -and $title -notlike "*Revit*" -and $title -notlike "*.dyn*") {{
                $rect = New-Object WinAPI+RECT
                [WinAPI]::GetWindowRect($hwnd, [ref]$rect)
                $w = $rect.Right - $rect.Left
                $h = $rect.Bottom - $rect.Top

                if ($w -gt 100 -and $h -gt 100) {{
                    $script:foundHwnd = $hwnd
                    $script:foundTitle = $title
                    $script:foundLeft = $rect.Left
                    $script:foundTop = $rect.Top
                    $script:foundWidth = $w
                    $script:foundHeight = $h
                    return $false
                }}
            }}
        }}
        return $true
    }}

    [WinAPI]::EnumWindows($callback, [IntPtr]::Zero)
    return $script:foundHwnd
}}

Start-Sleep -Milliseconds 800

$maxAttempts = 30
$attempt = 0

Log "Searching for Dynamo Player window..."

while ($attempt -lt $maxAttempts) {{
    $hwnd = Find-DynamoWindow
    if ($hwnd -ne [IntPtr]::Zero) {{
        Log "Found window: '$foundTitle' after $($attempt * 300 + 800)ms"
        break
    }}
    Start-Sleep -Milliseconds 300
    $attempt++
}}

if ($foundHwnd -ne [IntPtr]::Zero) {{
    Log "Activating window..."
    [WinAPI]::ShowWindow($foundHwnd, 9)
    [WinAPI]::SetForegroundWindow($foundHwnd)
    Start-Sleep -Milliseconds 300

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
        if not os.path.exists(DOWNLOAD_FOLDER):
            os.makedirs(DOWNLOAD_FOLDER)

        with codecs.open(debug_file, 'w', 'utf-8') as f:
            f.write("=== Starting autotype ===\n")

        ps_bytes = ps_content.encode('utf-16-le')
        ps_base64 = base64.b64encode(ps_bytes).decode('ascii')

        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0

        subprocess.Popen(
            ["powershell.exe", "-WindowStyle", "Hidden", "-EncodedCommand", ps_base64],
            startupinfo=startupinfo,
            creationflags=CREATE_NO_WINDOW
        )

        return True
    except Exception:
        return False


def run_dynamo_script(script_path):
    """Запустить Dynamo скрипт через Dynamo Player."""
    script_name = os.path.splitext(os.path.basename(script_path))[0]

    try:
        from Autodesk.Revit.UI import PostableCommand, RevitCommandId

        uiapp = revit.HOST_APP.uiapp
        cmd_id = RevitCommandId.LookupPostableCommandId(PostableCommand.DynamoPlayer)

        if cmd_id:
            copy_to_clipboard(script_name)
            create_autotype_script(script_name)
            uiapp.PostCommand(cmd_id)
            return True, script_name

    except Exception as e:
        output.print_md("Ошибка: {}".format(str(e)))

    try:
        os.startfile(script_path)
        return True, script_name
    except Exception as e2:
        return False, "Ошибка запуска: {}".format(str(e2))


# === ГЛАВНОЕ ОКНО ===

class ServerScriptsForm(Form):
    """Диалог выбора и запуска Dynamo скриптов с сервера."""

    def __init__(self):
        self.client = DynamoApiClient()
        self.all_scripts = []
        self.filtered_scripts = []
        self.all_sections = []
        self.selected_script = None
        self.needs_restart = False

        self.setup_form()
        self.load_data()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Dynamo скрипты с сервера - CPSK"
        self.Width = 750
        self.Height = 500
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.MinimumSize = Size(600, 400)

        # === ВЕРХНЯЯ ПАНЕЛЬ ===
        top_panel = Panel()
        top_panel.Dock = DockStyle.Top
        top_panel.Height = 70

        # Поиск
        lbl_search = Label()
        lbl_search.Text = "Поиск:"
        lbl_search.Location = Point(10, 12)
        lbl_search.AutoSize = True
        top_panel.Controls.Add(lbl_search)

        self.txt_search = TextBox()
        self.txt_search.Location = Point(60, 10)
        self.txt_search.Width = 250
        self.txt_search.TextChanged += self.on_filter_changed
        top_panel.Controls.Add(self.txt_search)

        # Секция
        lbl_section = Label()
        lbl_section.Text = "Раздел:"
        lbl_section.Location = Point(10, 42)
        lbl_section.AutoSize = True
        top_panel.Controls.Add(lbl_section)

        self.cmb_section = ComboBox()
        self.cmb_section.Location = Point(60, 40)
        self.cmb_section.Width = 250
        self.cmb_section.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmb_section.SelectedIndexChanged += self.on_filter_changed
        top_panel.Controls.Add(self.cmb_section)

        # Счётчик
        self.lbl_count = Label()
        self.lbl_count.Location = Point(320, 12)
        self.lbl_count.AutoSize = True
        self.lbl_count.Text = ""
        top_panel.Controls.Add(self.lbl_count)

        # Кнопка обновить
        btn_refresh = Button()
        btn_refresh.Text = "Обновить"
        btn_refresh.Location = Point(320, 38)
        btn_refresh.Width = 80
        btn_refresh.Click += self.on_refresh_click
        top_panel.Controls.Add(btn_refresh)

        # === НИЖНЯЯ ПАНЕЛЬ ===
        bottom_panel = Panel()
        bottom_panel.Dock = DockStyle.Bottom
        bottom_panel.Height = 45

        self.btn_run = Button()
        self.btn_run.Text = "Скачать и запустить"
        self.btn_run.Location = Point(10, 10)
        self.btn_run.Width = 140
        self.btn_run.Height = 30
        self.btn_run.Click += self.on_run_click
        self.btn_run.Enabled = False
        bottom_panel.Controls.Add(self.btn_run)

        self.lbl_info = Label()
        self.lbl_info.Location = Point(160, 15)
        self.lbl_info.Width = 500
        self.lbl_info.ForeColor = Color.Gray
        self.lbl_info.Text = "Выберите скрипт для запуска"
        bottom_panel.Controls.Add(self.lbl_info)

        # === СПИСОК СКРИПТОВ ===
        self.list_scripts = ListBox()
        self.list_scripts.Dock = DockStyle.Fill
        self.list_scripts.SelectionMode = SelectionMode.One
        self.list_scripts.SelectedIndexChanged += self.on_script_selected
        self.list_scripts.DoubleClick += self.on_run_click

        # === ПАНЕЛЬ ОПИСАНИЯ ===
        desc_panel = Panel()
        desc_panel.Dock = DockStyle.Bottom
        desc_panel.Height = 80
        desc_panel.Padding = Padding(5)

        self.lbl_author = Label()
        self.lbl_author.Dock = DockStyle.Top
        self.lbl_author.Height = 18
        self.lbl_author.Text = ""
        self.lbl_author.ForeColor = Color.DimGray

        lbl_desc_title = Label()
        lbl_desc_title.Text = "Описание:"
        lbl_desc_title.Dock = DockStyle.Top
        lbl_desc_title.Height = 18
        lbl_desc_title.Font = Font(lbl_desc_title.Font, FontStyle.Bold)

        self.lbl_description = Label()
        self.lbl_description.Dock = DockStyle.Fill
        self.lbl_description.Text = ""
        self.lbl_description.ForeColor = Color.DarkGray

        desc_panel.Controls.Add(self.lbl_description)
        desc_panel.Controls.Add(lbl_desc_title)
        desc_panel.Controls.Add(self.lbl_author)

        # Добавляем контролы
        self.Controls.Add(self.list_scripts)
        self.Controls.Add(desc_panel)
        self.Controls.Add(bottom_panel)
        self.Controls.Add(top_panel)

    def load_data(self):
        """Загрузить данные с сервера."""
        # Загружаем секции
        success, sections, error = self.client.get_sections()
        if success and sections:
            self.all_sections = sections
            self.cmb_section.Items.Clear()
            self.cmb_section.Items.Add("Все разделы")
            for section in sections:
                name = section.get("name", "")
                if name:
                    self.cmb_section.Items.Add(name)
            self.cmb_section.SelectedIndex = 0

        # Загружаем скрипты
        success, scripts, error = self.client.get_scripts()
        if success and scripts:
            self.all_scripts = scripts
            self.filter_scripts()
        else:
            show_error(
                "Ошибка загрузки",
                "Не удалось загрузить список скриптов",
                details=error,
                blocking=True
            )

    def filter_scripts(self):
        """Фильтровать скрипты по поиску и секции."""
        search_text = self.txt_search.Text.lower().strip()
        selected_section = ""
        if self.cmb_section.SelectedIndex > 0:
            selected_section = str(self.cmb_section.SelectedItem)

        self.filtered_scripts = []
        for script in self.all_scripts:
            # Фильтр по секции
            if selected_section:
                if script.get("section_name", "") != selected_section:
                    continue

            # Фильтр по поиску
            if search_text:
                filename = script.get("filename", "").lower()
                comment = (script.get("comment") or "").lower()
                if search_text not in filename and search_text not in comment:
                    continue

            self.filtered_scripts.append(script)

        self.update_scripts_list()

    def update_scripts_list(self):
        """Обновить список скриптов."""
        self.list_scripts.Items.Clear()
        self.selected_script = None
        self.btn_run.Enabled = False
        self.lbl_description.Text = ""
        self.lbl_author.Text = ""
        self.lbl_info.Text = "Выберите скрипт для запуска"
        self.lbl_info.ForeColor = Color.Gray

        for script in self.filtered_scripts:
            filename = script.get("filename", "unknown.dyn")
            section = script.get("section_name", "")
            is_approved = script.get("is_approved", False)

            display = filename
            if section:
                display = "{} [{}]".format(filename, section)
            if is_approved:
                display = "[OK] " + display

            self.list_scripts.Items.Add(display)

        self.lbl_count.Text = "Найдено: {}".format(len(self.filtered_scripts))

    def on_filter_changed(self, sender, args):
        """Изменён фильтр."""
        self.filter_scripts()

    def on_script_selected(self, sender, args):
        """Выбран скрипт."""
        idx = self.list_scripts.SelectedIndex
        if idx < 0 or idx >= len(self.filtered_scripts):
            self.selected_script = None
            self.btn_run.Enabled = False
            self.lbl_description.Text = ""
            self.lbl_author.Text = ""
            self.lbl_info.Text = "Выберите скрипт для запуска"
            return

        self.selected_script = self.filtered_scripts[idx]

        # Информация о скрипте
        author = self.selected_script.get("author_username", "")
        date = self.selected_script.get("date_uploaded", "")
        is_approved = self.selected_script.get("is_approved", False)

        author_parts = []
        if author:
            author_parts.append("Автор: {}".format(author))
        if date:
            # Форматируем дату
            if "T" in str(date):
                date = str(date).split("T")[0]
            author_parts.append("Дата: {}".format(date))
        if is_approved:
            author_parts.append("Проверено")

        self.lbl_author.Text = "  |  ".join(author_parts) if author_parts else ""

        comment = self.selected_script.get("comment") or ""
        if comment:
            self.lbl_description.Text = comment
            self.lbl_description.ForeColor = Color.Black
        else:
            self.lbl_description.Text = "Нет описания"
            self.lbl_description.ForeColor = Color.Gray

        self.lbl_info.Text = self.selected_script.get("filename", "")
        self.lbl_info.ForeColor = Color.Gray

        self.btn_run.Enabled = True

    def on_refresh_click(self, sender, args):
        """Обновить список."""
        self.load_data()

    def on_run_click(self, sender, args):
        """Скачать и запустить скрипт."""
        if not self.selected_script:
            return

        # Проверяем/добавляем папку в Dynamo Player
        self.needs_restart = ensure_download_folder_in_dynamo()

        if self.needs_restart:
            show_warning(
                "Требуется перезапуск",
                "Папка для скриптов добавлена в Dynamo Player.\n"
                "ПЕРЕЗАПУСТИТЕ REVIT и запустите скрипт снова!",
                details="Папка: {}".format(DOWNLOAD_FOLDER)
            )
            self.DialogResult = DialogResult.Cancel
            self.Close()
            return

        # Скачиваем скрипт
        self.lbl_info.Text = "Скачивание..."
        self.lbl_info.ForeColor = Color.Blue
        self.Refresh()

        success, file_path, error = self.client.download_script(
            self.selected_script, DOWNLOAD_FOLDER
        )

        if not success:
            show_error(
                "Ошибка скачивания",
                "Не удалось скачать скрипт",
                details=error,
                blocking=True
            )
            self.lbl_info.Text = "Ошибка скачивания"
            self.lbl_info.ForeColor = Color.Red
            return

        # Запускаем скрипт
        success, script_name = run_dynamo_script(file_path)

        if success:
            self.DialogResult = DialogResult.OK
            self.Close()
        else:
            show_error(
                "Ошибка запуска",
                "Не удалось запустить скрипт",
                details=script_name,
                blocking=True
            )


# === MAIN ===
if __name__ == "__main__":
    form = ServerScriptsForm()
    form.ShowDialog()
