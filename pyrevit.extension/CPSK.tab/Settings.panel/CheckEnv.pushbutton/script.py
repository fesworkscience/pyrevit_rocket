# -*- coding: utf-8 -*-
"""
CheckEnv - Проверка Python окружения.
Показывает статус установки и готовности окружения.
"""

__title__ = "Проверить\nокружение"
__author__ = "CPSK"

import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, GroupBox,
    FormStartPosition, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import script

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from cpsk_config import (
    check_environment, get_venv_path, get_requirements_path,
    get_setting, EXTENSION_DIR
)

output = script.get_output()


class CheckEnvForm(Form):
    """Диалог проверки окружения."""

    def __init__(self):
        self.setup_form()
        self.run_check()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "CPSK - Проверка окружения"
        self.Width = 500
        self.Height = 400
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # Заголовок
        lbl_title = Label()
        lbl_title.Text = "Статус Python окружения"
        lbl_title.Location = Point(20, y)
        lbl_title.Size = Size(440, 28)
        lbl_title.Font = Font(lbl_title.Font.FontFamily, 12, FontStyle.Bold)
        self.Controls.Add(lbl_title)

        y += 40

        # Общий статус
        self.lbl_overall = Label()
        self.lbl_overall.Location = Point(20, y)
        self.lbl_overall.Size = Size(440, 25)
        self.lbl_overall.Font = Font(self.lbl_overall.Font.FontFamily, 10, FontStyle.Bold)
        self.Controls.Add(self.lbl_overall)

        y += 35

        # Группа деталей
        grp_details = GroupBox()
        grp_details.Text = "Детали"
        grp_details.Location = Point(15, y)
        grp_details.Size = Size(455, 200)

        detail_y = 25

        # Venv
        lbl_venv_title = Label()
        lbl_venv_title.Text = "Виртуальное окружение:"
        lbl_venv_title.Location = Point(15, detail_y)
        lbl_venv_title.Size = Size(180, 18)
        lbl_venv_title.Font = Font(lbl_venv_title.Font.FontFamily, 9, FontStyle.Bold)
        grp_details.Controls.Add(lbl_venv_title)

        self.lbl_venv = Label()
        self.lbl_venv.Location = Point(200, detail_y)
        self.lbl_venv.Size = Size(240, 18)
        grp_details.Controls.Add(self.lbl_venv)

        detail_y += 22

        self.lbl_venv_path = Label()
        self.lbl_venv_path.Location = Point(15, detail_y)
        self.lbl_venv_path.Size = Size(425, 18)
        self.lbl_venv_path.ForeColor = Color.Gray
        grp_details.Controls.Add(self.lbl_venv_path)

        detail_y += 30

        # Python версия
        lbl_py_title = Label()
        lbl_py_title.Text = "Python версия:"
        lbl_py_title.Location = Point(15, detail_y)
        lbl_py_title.Size = Size(180, 18)
        lbl_py_title.Font = Font(lbl_py_title.Font.FontFamily, 9, FontStyle.Bold)
        grp_details.Controls.Add(lbl_py_title)

        self.lbl_python = Label()
        self.lbl_python.Location = Point(200, detail_y)
        self.lbl_python.Size = Size(240, 18)
        grp_details.Controls.Add(self.lbl_python)

        detail_y += 30

        # Requirements
        lbl_req_title = Label()
        lbl_req_title.Text = "Requirements.txt:"
        lbl_req_title.Location = Point(15, detail_y)
        lbl_req_title.Size = Size(180, 18)
        lbl_req_title.Font = Font(lbl_req_title.Font.FontFamily, 9, FontStyle.Bold)
        grp_details.Controls.Add(lbl_req_title)

        self.lbl_req = Label()
        self.lbl_req.Location = Point(200, detail_y)
        self.lbl_req.Size = Size(240, 18)
        grp_details.Controls.Add(self.lbl_req)

        detail_y += 30

        # Пакеты
        lbl_pkg_title = Label()
        lbl_pkg_title.Text = "Ключевые пакеты:"
        lbl_pkg_title.Location = Point(15, detail_y)
        lbl_pkg_title.Size = Size(180, 18)
        lbl_pkg_title.Font = Font(lbl_pkg_title.Font.FontFamily, 9, FontStyle.Bold)
        grp_details.Controls.Add(lbl_pkg_title)

        self.lbl_packages = Label()
        self.lbl_packages.Location = Point(200, detail_y)
        self.lbl_packages.Size = Size(240, 18)
        grp_details.Controls.Add(self.lbl_packages)

        detail_y += 30

        # Ошибки
        self.lbl_errors = Label()
        self.lbl_errors.Location = Point(15, detail_y)
        self.lbl_errors.Size = Size(425, 35)
        self.lbl_errors.ForeColor = Color.Red
        grp_details.Controls.Add(self.lbl_errors)

        self.Controls.Add(grp_details)

        y += 215

        # Кнопки
        btn_refresh = Button()
        btn_refresh.Text = "Обновить"
        btn_refresh.Location = Point(15, y)
        btn_refresh.Size = Size(100, 28)
        btn_refresh.Click += self.on_refresh_click
        self.Controls.Add(btn_refresh)

        btn_setup = Button()
        btn_setup.Text = "Установить окружение"
        btn_setup.Location = Point(125, y)
        btn_setup.Size = Size(150, 28)
        btn_setup.Click += self.on_setup_click
        self.Controls.Add(btn_setup)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(390, y)
        btn_close.Size = Size(90, 28)
        btn_close.Click += self.on_close_click
        self.Controls.Add(btn_close)

    def run_check(self):
        """Выполнить проверку."""
        result = check_environment()

        # Общий статус
        if result["is_ready"]:
            self.lbl_overall.Text = "Окружение готово к работе"
            self.lbl_overall.ForeColor = Color.Green
        else:
            self.lbl_overall.Text = "Окружение требует настройки"
            self.lbl_overall.ForeColor = Color.Red

        # Venv
        if result["venv_exists"]:
            self.lbl_venv.Text = "Установлено"
            self.lbl_venv.ForeColor = Color.Green
        else:
            self.lbl_venv.Text = "Не найдено"
            self.lbl_venv.ForeColor = Color.Red

        self.lbl_venv_path.Text = get_venv_path()

        # Python
        if result["venv_version"]:
            self.lbl_python.Text = result["venv_version"]
            self.lbl_python.ForeColor = Color.Green
        else:
            self.lbl_python.Text = "Не определена"
            self.lbl_python.ForeColor = Color.Gray

        # Requirements
        if result["requirements_exists"]:
            self.lbl_req.Text = "Найден ({} пакетов)".format(result["requirements_count"])
            self.lbl_req.ForeColor = Color.Green
        else:
            self.lbl_req.Text = "Не найден"
            self.lbl_req.ForeColor = Color.Red

        # Пакеты
        if result["packages_installed"]:
            self.lbl_packages.Text = "Установлены (ifcopenshell, ifctester)"
            self.lbl_packages.ForeColor = Color.Green
        else:
            self.lbl_packages.Text = "Не установлены"
            self.lbl_packages.ForeColor = Color.Orange

        # Ошибки
        if result["errors"]:
            self.lbl_errors.Text = "\n".join(result["errors"][:2])
        else:
            self.lbl_errors.Text = ""

    def on_refresh_click(self, sender, args):
        """Обновить проверку."""
        self.run_check()

    def on_setup_click(self, sender, args):
        """Открыть окно установки."""
        self.Close()
        # Импортируем и запускаем SetupEnv
        try:
            setup_path = os.path.join(
                EXTENSION_DIR, "CPSK.tab", "Settings.panel",
                "SetupEnv.pushbutton", "script.py"
            )
            if os.path.exists(setup_path):
                exec(compile(open(setup_path).read(), setup_path, 'exec'))
        except Exception as e:
            MessageBox.Show(
                "Ошибка открытия настроек:\n{}".format(str(e)),
                "Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

    def on_close_click(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    form = CheckEnvForm()
    form.ShowDialog()
