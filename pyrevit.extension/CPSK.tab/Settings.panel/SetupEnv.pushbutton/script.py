# -*- coding: utf-8 -*-
"""
SetupEnv - Установка Python окружения и зависимостей.
Создаёт .venv и устанавливает пакеты из requirements.txt.
Пути настраиваются в cpsk_settings.yaml.
"""

__title__ = "Установить\nокружение"
__author__ = "CPSK"

import clr
import os
import sys
import subprocess
import codecs

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, Panel, CheckBox, ProgressBar,
    DockStyle, FormStartPosition, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    DialogResult, GroupBox, ProgressBarStyle, OpenFileDialog,
    FolderBrowserDialog
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import script, forms

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Импортируем модуль конфигурации
from cpsk_config import (
    load_settings, save_settings, get_setting, set_setting,
    get_venv_path, get_requirements_path, get_venv_python, get_venv_pip,
    find_system_python, get_python_version, check_environment,
    get_absolute_path, get_relative_path, EXTENSION_DIR
)

output = script.get_output()


class SetupEnvForm(Form):
    """Диалог установки окружения."""

    def __init__(self):
        self.settings = load_settings()
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "CPSK - Установка окружения"
        self.Width = 600
        self.Height = 520
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # Заголовок
        lbl_title = Label()
        lbl_title.Text = "Настройка Python окружения для CPSK Tools"
        lbl_title.Location = Point(20, y)
        lbl_title.Size = Size(540, 25)
        lbl_title.Font = Font(lbl_title.Font.FontFamily, 11, FontStyle.Bold)
        self.Controls.Add(lbl_title)

        y += 35

        # === Группа: Системный Python ===
        grp_python = GroupBox()
        grp_python.Text = "Системный Python (для создания venv)"
        grp_python.Location = Point(15, y)
        grp_python.Size = Size(555, 75)

        self.txt_python = TextBox()
        self.txt_python.Location = Point(15, 22)
        self.txt_python.Width = 430
        grp_python.Controls.Add(self.txt_python)

        btn_browse_py = Button()
        btn_browse_py.Text = "Обзор..."
        btn_browse_py.Location = Point(455, 20)
        btn_browse_py.Width = 85
        btn_browse_py.Click += self.on_browse_python
        grp_python.Controls.Add(btn_browse_py)

        self.lbl_python_ver = Label()
        self.lbl_python_ver.Text = ""
        self.lbl_python_ver.Location = Point(15, 50)
        self.lbl_python_ver.Size = Size(525, 18)
        grp_python.Controls.Add(self.lbl_python_ver)

        self.Controls.Add(grp_python)

        y += 90

        # === Группа: Виртуальное окружение ===
        grp_venv = GroupBox()
        grp_venv.Text = "Виртуальное окружение (.venv)"
        grp_venv.Location = Point(15, y)
        grp_venv.Size = Size(555, 75)

        lbl_venv = Label()
        lbl_venv.Text = "Путь:"
        lbl_venv.Location = Point(15, 25)
        lbl_venv.AutoSize = True
        grp_venv.Controls.Add(lbl_venv)

        self.txt_venv = TextBox()
        self.txt_venv.Location = Point(55, 22)
        self.txt_venv.Width = 390
        grp_venv.Controls.Add(self.txt_venv)

        btn_browse_venv = Button()
        btn_browse_venv.Text = "Обзор..."
        btn_browse_venv.Location = Point(455, 20)
        btn_browse_venv.Width = 85
        btn_browse_venv.Click += self.on_browse_venv
        grp_venv.Controls.Add(btn_browse_venv)

        self.lbl_venv_status = Label()
        self.lbl_venv_status.Text = ""
        self.lbl_venv_status.Location = Point(55, 50)
        self.lbl_venv_status.Size = Size(485, 18)
        grp_venv.Controls.Add(self.lbl_venv_status)

        self.Controls.Add(grp_venv)

        y += 90

        # === Группа: Requirements ===
        grp_req = GroupBox()
        grp_req.Text = "Файл зависимостей (requirements.txt)"
        grp_req.Location = Point(15, y)
        grp_req.Size = Size(555, 75)

        lbl_req = Label()
        lbl_req.Text = "Путь:"
        lbl_req.Location = Point(15, 25)
        lbl_req.AutoSize = True
        grp_req.Controls.Add(lbl_req)

        self.txt_req = TextBox()
        self.txt_req.Location = Point(55, 22)
        self.txt_req.Width = 390
        grp_req.Controls.Add(self.txt_req)

        btn_browse_req = Button()
        btn_browse_req.Text = "Обзор..."
        btn_browse_req.Location = Point(455, 20)
        btn_browse_req.Width = 85
        btn_browse_req.Click += self.on_browse_req
        grp_req.Controls.Add(btn_browse_req)

        self.lbl_req_status = Label()
        self.lbl_req_status.Text = ""
        self.lbl_req_status.Location = Point(55, 50)
        self.lbl_req_status.Size = Size(485, 18)
        grp_req.Controls.Add(self.lbl_req_status)

        self.Controls.Add(grp_req)

        y += 90

        # === Прогресс и статус ===
        self.progress = ProgressBar()
        self.progress.Location = Point(15, y)
        self.progress.Size = Size(555, 22)
        self.progress.Visible = False
        self.Controls.Add(self.progress)

        y += 28

        self.lbl_status = Label()
        self.lbl_status.Text = ""
        self.lbl_status.Location = Point(15, y)
        self.lbl_status.Size = Size(555, 40)
        self.Controls.Add(self.lbl_status)

        y += 48

        # === Кнопки ===
        self.btn_create_venv = Button()
        self.btn_create_venv.Text = "1. Создать venv"
        self.btn_create_venv.Location = Point(15, y)
        self.btn_create_venv.Size = Size(130, 32)
        self.btn_create_venv.Click += self.on_create_venv
        self.Controls.Add(self.btn_create_venv)

        self.btn_install_deps = Button()
        self.btn_install_deps.Text = "2. Установить пакеты"
        self.btn_install_deps.Location = Point(155, y)
        self.btn_install_deps.Size = Size(145, 32)
        self.btn_install_deps.Click += self.on_install_deps
        self.Controls.Add(self.btn_install_deps)

        self.btn_full_install = Button()
        self.btn_full_install.Text = "Полная установка"
        self.btn_full_install.Location = Point(310, y)
        self.btn_full_install.Size = Size(130, 32)
        self.btn_full_install.Click += self.on_full_install
        self.Controls.Add(self.btn_full_install)

        btn_save = Button()
        btn_save.Text = "Сохранить"
        btn_save.Location = Point(450, y)
        btn_save.Size = Size(90, 32)
        btn_save.Click += self.on_save_click
        self.Controls.Add(btn_save)

        y += 42

        btn_check = Button()
        btn_check.Text = "Проверить окружение"
        btn_check.Location = Point(15, y)
        btn_check.Size = Size(150, 28)
        btn_check.Click += self.on_check_click
        self.Controls.Add(btn_check)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(480, y)
        btn_close.Size = Size(90, 28)
        btn_close.Click += self.on_close_click
        self.Controls.Add(btn_close)

        # Загрузить данные
        self.load_data()

    def load_data(self):
        """Загрузить данные из настроек."""
        # Python
        python_path = get_setting("environment.python_path", "")
        if not python_path:
            python_path = find_system_python() or ""
        self.txt_python.Text = python_path
        self.update_python_info()

        # Venv (показываем абсолютный путь для наглядности)
        venv_rel = get_setting("environment.venv_path", "lib/.venv")
        self.txt_venv.Text = get_absolute_path(venv_rel)

        # Requirements
        req_rel = get_setting("environment.requirements_path", "lib/requirements.txt")
        self.txt_req.Text = get_absolute_path(req_rel)

        self.check_status()

    def update_python_info(self):
        """Обновить информацию о Python."""
        path = self.txt_python.Text.strip()
        if path and os.path.exists(path):
            version = get_python_version(path)
            self.lbl_python_ver.Text = version if version else "Неизвестная версия"
            self.lbl_python_ver.ForeColor = Color.Green
        else:
            self.lbl_python_ver.Text = "Python не найден"
            self.lbl_python_ver.ForeColor = Color.Red

    def check_status(self):
        """Проверить статус установки."""
        # Venv
        venv_path = self.txt_venv.Text.strip()
        venv_python = os.path.join(venv_path, "Scripts", "python.exe")
        if os.path.exists(venv_python):
            version = get_python_version(venv_python)
            self.lbl_venv_status.Text = "Установлено: {}".format(version)
            self.lbl_venv_status.ForeColor = Color.Green
        else:
            self.lbl_venv_status.Text = "Не установлено"
            self.lbl_venv_status.ForeColor = Color.Orange

        # Requirements
        req_path = self.txt_req.Text.strip()
        if os.path.exists(req_path):
            with codecs.open(req_path, 'r', 'utf-8') as f:
                lines = [l.strip() for l in f if l.strip() and not l.startswith('#')]
            self.lbl_req_status.Text = "Найден: {} пакетов".format(len(lines))
            self.lbl_req_status.ForeColor = Color.Green
        else:
            self.lbl_req_status.Text = "Файл не найден"
            self.lbl_req_status.ForeColor = Color.Red

    def on_browse_python(self, sender, args):
        """Выбор Python."""
        selected = forms.pick_file(file_ext='exe', title="Выберите python.exe")
        if selected:
            self.txt_python.Text = selected
            self.update_python_info()

    def on_browse_venv(self, sender, args):
        """Выбор папки venv."""
        dialog = FolderBrowserDialog()
        dialog.Description = "Выберите папку для виртуального окружения"
        if dialog.ShowDialog() == DialogResult.OK:
            self.txt_venv.Text = dialog.SelectedPath
            self.check_status()

    def on_browse_req(self, sender, args):
        """Выбор requirements.txt."""
        selected = forms.pick_file(file_ext='txt', title="Выберите requirements.txt")
        if selected:
            self.txt_req.Text = selected
            self.check_status()

    def on_save_click(self, sender, args):
        """Сохранить настройки."""
        # Python - сохраняем как есть (абсолютный путь)
        set_setting("environment.python_path", self.txt_python.Text.strip())

        # Venv и requirements - преобразуем в относительные если возможно
        venv_path = self.txt_venv.Text.strip()
        set_setting("environment.venv_path", get_relative_path(venv_path))

        req_path = self.txt_req.Text.strip()
        set_setting("environment.requirements_path", get_relative_path(req_path))

        self.lbl_status.Text = "Настройки сохранены"
        self.lbl_status.ForeColor = Color.Green

    def on_check_click(self, sender, args):
        """Проверить окружение."""
        self.check_status()

        # Детальная проверка
        result = check_environment()

        if result["is_ready"]:
            self.lbl_status.Text = "Окружение готово к работе!"
            self.lbl_status.ForeColor = Color.Green
        else:
            errors = "\n".join(result["errors"][:3])
            self.lbl_status.Text = "Проблемы: {}".format(errors)
            self.lbl_status.ForeColor = Color.Red

    def on_create_venv(self, sender, args):
        """Создать виртуальное окружение."""
        python_path = self.txt_python.Text.strip()
        venv_path = self.txt_venv.Text.strip()

        if not python_path or not os.path.exists(python_path):
            MessageBox.Show("Укажите корректный путь к Python!", "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        if not venv_path:
            MessageBox.Show("Укажите путь для venv!", "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        self.start_progress("Создание виртуального окружения...")

        try:
            # Удалить старый venv если есть
            if os.path.exists(venv_path):
                import shutil
                shutil.rmtree(venv_path)

            # Создать venv
            result = subprocess.call([python_path, "-m", "venv", venv_path])

            if result != 0:
                raise Exception("Ошибка создания venv (код {})".format(result))

            self.stop_progress()
            self.check_status()
            self.lbl_status.Text = "Виртуальное окружение создано!"
            self.lbl_status.ForeColor = Color.Green

        except Exception as e:
            self.stop_progress()
            self.lbl_status.Text = "Ошибка: {}".format(str(e))
            self.lbl_status.ForeColor = Color.Red

    def on_install_deps(self, sender, args):
        """Установить зависимости."""
        venv_path = self.txt_venv.Text.strip()
        req_path = self.txt_req.Text.strip()

        pip_path = os.path.join(venv_path, "Scripts", "pip.exe")
        if not os.path.exists(pip_path):
            MessageBox.Show("Сначала создайте виртуальное окружение!", "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        if not os.path.exists(req_path):
            MessageBox.Show("Файл requirements.txt не найден!", "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        self.start_progress("Установка зависимостей (это может занять несколько минут)...")

        try:
            # Обновить pip
            subprocess.call([pip_path, "install", "--upgrade", "pip"])

            # Установить зависимости
            result = subprocess.call([pip_path, "install", "-r", req_path])

            if result != 0:
                raise Exception("Ошибка установки (код {})".format(result))

            self.stop_progress()
            self.check_status()
            self.lbl_status.Text = "Зависимости установлены!"
            self.lbl_status.ForeColor = Color.Green

            # Обновить статус в настройках
            set_setting("environment.installed", True)

        except Exception as e:
            self.stop_progress()
            self.lbl_status.Text = "Ошибка: {}".format(str(e))
            self.lbl_status.ForeColor = Color.Red

    def on_full_install(self, sender, args):
        """Полная установка: venv + зависимости."""
        python_path = self.txt_python.Text.strip()
        venv_path = self.txt_venv.Text.strip()
        req_path = self.txt_req.Text.strip()

        if not python_path or not os.path.exists(python_path):
            MessageBox.Show("Укажите корректный путь к Python!", "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        if not venv_path:
            MessageBox.Show("Укажите путь для venv!", "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        if not os.path.exists(req_path):
            MessageBox.Show("Файл requirements.txt не найден!", "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        # Сохраняем настройки перед установкой
        self.on_save_click(None, None)

        self.start_progress("Шаг 1/3: Создание виртуального окружения...")

        try:
            # Удалить старый venv
            if os.path.exists(venv_path):
                import shutil
                shutil.rmtree(venv_path)

            # Создать venv
            result = subprocess.call([python_path, "-m", "venv", venv_path])
            if result != 0:
                raise Exception("Ошибка создания venv")

            self.lbl_status.Text = "Шаг 2/3: Обновление pip..."
            System.Windows.Forms.Application.DoEvents()

            pip_path = os.path.join(venv_path, "Scripts", "pip.exe")
            subprocess.call([pip_path, "install", "--upgrade", "pip"])

            self.lbl_status.Text = "Шаг 3/3: Установка зависимостей..."
            System.Windows.Forms.Application.DoEvents()

            result = subprocess.call([pip_path, "install", "-r", req_path])
            if result != 0:
                raise Exception("Ошибка установки зависимостей")

            self.stop_progress()
            self.check_status()

            set_setting("environment.installed", True)

            self.lbl_status.Text = "Установка завершена успешно!"
            self.lbl_status.ForeColor = Color.Green

            MessageBox.Show(
                "Окружение успешно установлено!\n\n"
                "Теперь инструменты CPSK будут использовать это окружение.",
                "Готово",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

        except Exception as e:
            self.stop_progress()
            self.lbl_status.Text = "Ошибка: {}".format(str(e))
            self.lbl_status.ForeColor = Color.Red
            MessageBox.Show("Ошибка: {}".format(str(e)), "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)

    def start_progress(self, message):
        """Показать прогресс."""
        self.progress.Visible = True
        self.progress.Style = ProgressBarStyle.Marquee
        self.lbl_status.Text = message
        self.lbl_status.ForeColor = Color.Black
        self.btn_create_venv.Enabled = False
        self.btn_install_deps.Enabled = False
        self.btn_full_install.Enabled = False
        System.Windows.Forms.Application.DoEvents()

    def stop_progress(self):
        """Скрыть прогресс."""
        self.progress.Visible = False
        self.btn_create_venv.Enabled = True
        self.btn_install_deps.Enabled = True
        self.btn_full_install.Enabled = True

    def on_close_click(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    form = SetupEnvForm()
    form.ShowDialog()
