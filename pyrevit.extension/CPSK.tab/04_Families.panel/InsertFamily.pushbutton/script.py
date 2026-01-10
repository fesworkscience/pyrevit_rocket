# -*- coding: utf-8 -*-
"""
Вставка семейств из библиотеки CPSK.
Поиск и загрузка семейств с сервера.
"""

__title__ = "Вставка\nсемейств"
__author__ = "CPSK"

import clr
import os
import sys
import json
import codecs
import tempfile
import ssl
import re

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, Panel, ListView, ListViewItem,
    ComboBox, ComboBoxStyle, ProgressBar, Timer,
    DockStyle, FormStartPosition, FormBorderStyle,
    View, ColumnHeader, SortOrder, AnchorStyles,
    DialogResult
)
from System.Drawing import Point, Size, Font, FontStyle, Color

from Autodesk.Revit.DB import Family, FamilySource, IFamilyLoadOptions
from pyrevit import revit

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Добавляем корень проекта для импорта config
PROJECT_ROOT = os.path.dirname(EXTENSION_DIR)
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# Проверка авторизации
from cpsk_auth import require_auth, AuthService, _create_ssl_context
from cpsk_notify import show_error, show_warning, show_success, show_info

if not require_auth():
    sys.exit()

doc = revit.doc

# API настройки (импорт из config.py)
from config import API_BASE_URL
API_SECTIONS_ENDPOINT = "/api/rocketrevit/sections/"
API_FAMILIES_ENDPOINT = "/api/rocketrevit/revit-families/shared_files/"


def url_decode(s):
    """Декодировать URL-encoded строку."""
    if not s:
        return s
    try:
        # Python 2
        import urllib
        return urllib.unquote(s.encode('utf-8')).decode('utf-8')
    except Exception:
        return s


def api_get(endpoint):
    """
    Выполнить GET запрос к API.

    Returns:
        tuple: (success, data_or_error)
    """
    try:
        import urllib2

        url = API_BASE_URL + endpoint
        request = urllib2.Request(url)
        request.add_header('Accept', 'application/json')
        request.add_header('User-Agent', 'CPSK-pyRevit/1.0')

        # Добавляем токен
        token = AuthService.get_token()
        if token:
            request.add_header('Authorization', 'Bearer {}'.format(token))

        ssl_context = _create_ssl_context()
        if ssl_context:
            response = urllib2.urlopen(request, context=ssl_context, timeout=60)
        else:
            response = urllib2.urlopen(request, timeout=60)

        response_body = response.read().decode('utf-8')
        data = json.loads(response_body)
        return (True, data)

    except Exception as e:
        return (False, str(e))


def download_file(url):
    """
    Скачать файл по URL.

    Returns:
        tuple: (success, bytes_or_error)
    """
    try:
        import urllib2

        request = urllib2.Request(url)
        request.add_header('User-Agent', 'CPSK-pyRevit/1.0')

        token = AuthService.get_token()
        if token:
            request.add_header('Authorization', 'Bearer {}'.format(token))

        ssl_context = _create_ssl_context()
        if ssl_context:
            response = urllib2.urlopen(request, context=ssl_context, timeout=300)
        else:
            response = urllib2.urlopen(request, timeout=300)

        return (True, response.read())

    except Exception as e:
        return (False, str(e))


class FamilyLoadOptions(IFamilyLoadOptions):
    """Опции загрузки семейства - всегда перезаписывать."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        """
        Семейство найдено в проекте.
        В IronPython out параметры возвращаются через tuple.
        """
        # Return: (bool result, bool overwriteParameterValues)
        return True, True  # Загрузить и перезаписать параметры

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        """
        Вложенное семейство найдено.
        В IronPython out параметры возвращаются через tuple.
        """
        # Return: (bool result, FamilySource source, bool overwriteParameterValues)
        return True, FamilySource.Family, True


class FamilySearchForm(Form):
    """Форма поиска и вставки семейств."""

    def __init__(self):
        self.selected_family = None
        self.all_families = []
        self.sections = []
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Поиск и вставка семейств - CPSK"
        self.Width = 800
        self.Height = 600
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # Поиск
        lbl_search = Label()
        lbl_search.Text = "Поиск:"
        lbl_search.Location = Point(12, y + 3)
        lbl_search.AutoSize = True
        self.Controls.Add(lbl_search)

        self.txt_search = TextBox()
        self.txt_search.Location = Point(70, y)
        self.txt_search.Size = Size(300, 23)
        self.txt_search.TextChanged += self.on_filter_changed
        self.Controls.Add(self.txt_search)

        # Раздел
        lbl_section = Label()
        lbl_section.Text = "Раздел:"
        lbl_section.Location = Point(390, y + 3)
        lbl_section.AutoSize = True
        self.Controls.Add(lbl_section)

        self.cmb_section = ComboBox()
        self.cmb_section.Location = Point(450, y)
        self.cmb_section.Size = Size(200, 23)
        self.cmb_section.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmb_section.SelectedIndexChanged += self.on_filter_changed
        self.Controls.Add(self.cmb_section)

        # Кнопка обновления
        btn_refresh = Button()
        btn_refresh.Text = "Обновить"
        btn_refresh.Location = Point(670, y - 2)
        btn_refresh.Size = Size(80, 27)
        btn_refresh.Click += self.on_refresh
        self.Controls.Add(btn_refresh)

        y += 40

        # Список семейств
        self.list_families = ListView()
        self.list_families.Location = Point(12, y)
        self.list_families.Size = Size(760, 440)
        self.list_families.View = View.Details
        self.list_families.FullRowSelect = True
        self.list_families.GridLines = True
        self.list_families.MultiSelect = False

        self.list_families.Columns.Add("Название", 250)
        self.list_families.Columns.Add("Раздел", 150)
        self.list_families.Columns.Add("Автор", 120)
        self.list_families.Columns.Add("Дата", 100)
        self.list_families.Columns.Add("Статус", 60)
        self.list_families.Columns.Add("Комментарий", 70)

        self.list_families.DoubleClick += self.on_family_double_click
        self.list_families.SelectedIndexChanged += self.on_selection_changed
        self.Controls.Add(self.list_families)

        y += 450

        # Статус
        self.lbl_status = Label()
        self.lbl_status.Text = "Загрузка данных..."
        self.lbl_status.Location = Point(12, y)
        self.lbl_status.Size = Size(400, 23)
        self.Controls.Add(self.lbl_status)

        # Прогресс
        self.progress = ProgressBar()
        self.progress.Location = Point(12, y + 25)
        self.progress.Size = Size(760, 10)
        self.progress.Visible = False
        self.Controls.Add(self.progress)

        y += 45

        # Кнопки
        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(610, y)
        btn_cancel.Size = Size(75, 27)
        btn_cancel.Click += self.on_cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

        self.btn_insert = Button()
        self.btn_insert.Text = "Вставить"
        self.btn_insert.Location = Point(695, y)
        self.btn_insert.Size = Size(75, 27)
        self.btn_insert.Enabled = False
        self.btn_insert.Click += self.on_insert
        self.Controls.Add(self.btn_insert)
        self.AcceptButton = self.btn_insert

        # Загрузка данных при показе формы
        self.Load += self.on_form_load

    def on_form_load(self, sender, args):
        """Загрузка данных при открытии формы."""
        self.load_data()

    def load_data(self):
        """Загрузить разделы и семейства."""
        self.lbl_status.Text = "Загрузка данных..."
        self.progress.Visible = True
        self.progress.Style = System.Windows.Forms.ProgressBarStyle.Marquee

        # Загружаем разделы
        success, data = api_get(API_SECTIONS_ENDPOINT)
        if success and isinstance(data, list):
            self.sections = data
        else:
            self.sections = []

        # Заполняем комбобокс разделов
        self.cmb_section.Items.Clear()
        self.cmb_section.Items.Add("Все разделы")
        for section in self.sections:
            name = section.get("name", "") if isinstance(section, dict) else str(section)
            self.cmb_section.Items.Add(name)
        self.cmb_section.SelectedIndex = 0

        # Загружаем семейства
        success, data = api_get(API_FAMILIES_ENDPOINT)
        if success and isinstance(data, list):
            # Фильтруем только одобренные
            self.all_families = [f for f in data if f.get("is_approved", False)]
        else:
            self.all_families = []
            if not success:
                show_error("Ошибка загрузки", "Не удалось загрузить список семейств", details=str(data))

        self.progress.Visible = False
        self.update_family_list()

    def update_family_list(self):
        """Обновить список семейств с учётом фильтров."""
        self.list_families.Items.Clear()

        search_text = self.txt_search.Text.lower() if self.txt_search.Text else ""
        selected_section = ""
        if self.cmb_section.SelectedIndex > 0:
            selected_section = str(self.cmb_section.SelectedItem)

        count = 0
        for family in self.all_families:
            # Получаем имя файла (защита от None)
            file_input = family.get("file_input") or ""
            file_name = os.path.basename(file_input) if file_input else "Unknown"
            file_name = url_decode(file_name) or "Unknown"
            if file_name.lower().endswith(".rfa"):
                file_name = file_name[:-4]

            section_name = family.get("section_name") or ""
            author = family.get("author_username") or ""
            comment = family.get("comment") or ""

            # Применяем фильтр поиска
            if search_text:
                searchable = "{}{}{}{}".format(file_name, section_name, author, comment).lower()
                if search_text not in searchable:
                    continue

            # Применяем фильтр раздела
            if selected_section and section_name != selected_section:
                continue

            # Добавляем в список
            item = ListViewItem(file_name)
            item.SubItems.Add(section_name)
            item.SubItems.Add(author)

            # Дата
            date_str = family.get("date_uploaded") or ""
            if date_str:
                try:
                    # Парсим дату ISO формата
                    date_str = date_str.split("T")[0]  # Берём только дату
                except Exception:
                    pass
            item.SubItems.Add(date_str)

            # Статус
            is_approved = family.get("is_approved", False)
            item.SubItems.Add("OK" if is_approved else "-")

            item.SubItems.Add(comment[:50] if comment else "")
            item.Tag = family  # Сохраняем полные данные

            self.list_families.Items.Add(item)
            count += 1

            # Ограничиваем количество
            if count >= 500:
                break

        self.lbl_status.Text = "Показано {} из {} семейств".format(count, len(self.all_families))

    def on_filter_changed(self, sender, args):
        """Изменился фильтр."""
        self.update_family_list()

    def on_selection_changed(self, sender, args):
        """Изменился выбор в списке."""
        self.btn_insert.Enabled = self.list_families.SelectedItems.Count > 0

    def on_family_double_click(self, sender, args):
        """Двойной клик по семейству."""
        if self.list_families.SelectedItems.Count > 0:
            self.selected_family = self.list_families.SelectedItems[0].Tag
            self.DialogResult = DialogResult.OK
            self.Close()

    def on_insert(self, sender, args):
        """Кнопка Вставить."""
        if self.list_families.SelectedItems.Count > 0:
            self.selected_family = self.list_families.SelectedItems[0].Tag
            self.DialogResult = DialogResult.OK
            self.Close()

    def on_cancel(self, sender, args):
        """Кнопка Отмена."""
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def on_refresh(self, sender, args):
        """Кнопка Обновить."""
        self.load_data()


def get_safe_filename(filename):
    """Получить безопасное имя файла."""
    if not filename:
        return "family.rfa"

    # Декодируем URL
    filename = url_decode(filename)

    # Убираем недопустимые символы
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')

    # Ограничиваем длину
    if len(filename) > 100:
        filename = filename[:100]

    # Добавляем .rfa если нет
    if not filename.lower().endswith('.rfa'):
        filename += '.rfa'

    return filename


def load_family_into_revit(family_data):
    """
    Скачать и загрузить семейство в Revit.

    Args:
        family_data: dict с данными семейства из API

    Returns:
        tuple: (success, message)
    """
    try:
        # Получаем URL для скачивания
        download_url = family_data.get("download_link") or family_data.get("file_input")
        if not download_url:
            return (False, "URL для скачивания не найден")

        # Если относительный URL, добавляем базовый
        if download_url.startswith('/'):
            download_url = API_BASE_URL + download_url

        # Получаем имя файла
        file_input = family_data.get("file_input") or ""
        filename = os.path.basename(file_input) if file_input else "family.rfa"
        filename = get_safe_filename(filename)

        # Создаём временную папку
        temp_dir = os.path.join(tempfile.gettempdir(), "CPSK_Families")
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

        local_path = os.path.join(temp_dir, filename)

        # Скачиваем файл
        success, data = download_file(download_url)
        if not success:
            return (False, "Ошибка скачивания: {}".format(data))

        # Сохраняем файл
        with open(local_path, 'wb') as f:
            f.write(data)

        # Проверяем что файл создался
        if not os.path.exists(local_path):
            return (False, "Не удалось сохранить файл")

        # Загружаем в Revit
        if doc.IsReadOnly:
            return (False, "Документ открыт только для чтения")

        load_options = FamilyLoadOptions()
        loaded_family = clr.Reference[Family]()

        try:
            with revit.Transaction("Загрузка семейства CPSK"):
                success = doc.LoadFamily(local_path, load_options, loaded_family)

                if not success:
                    raise Exception("LoadFamily вернул False - возможно семейство уже загружено или повреждено")

            # Транзакция успешно завершена
            family_name = loaded_family.Value.Name if loaded_family.Value else filename

            # Удаляем временный файл
            try:
                os.remove(local_path)
            except Exception:
                pass

            return (True, family_name)

        except Exception as e:
            return (False, "Ошибка Revit API: {}".format(str(e)))

    except Exception as e:
        return (False, "Ошибка: {}".format(str(e)))


# === MAIN ===
if __name__ == "__main__":
    form = FamilySearchForm()
    result = form.ShowDialog()

    if result == DialogResult.OK and form.selected_family:
        # Показываем что загружаем
        file_input = form.selected_family.get("file_input") or ""
        filename = os.path.basename(file_input) if file_input else "family"
        filename = url_decode(filename) or "family"

        # Загружаем семейство
        success, result_msg = load_family_into_revit(form.selected_family)

        if success:
            show_success(
                "Семейство загружено",
                "Семейство '{}' успешно загружено в проект".format(result_msg)
            )
        else:
            show_error(
                "Ошибка загрузки",
                "Не удалось загрузить семейство",
                details=result_msg
            )
