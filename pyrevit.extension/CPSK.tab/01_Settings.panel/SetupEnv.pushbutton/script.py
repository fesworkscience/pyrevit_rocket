# -*- coding: utf-8 -*-
"""
SetupEnv - Проверка и установка Python окружения.
Показывает статус, пути и позволяет установить/переустановить окружение.

Venv создаётся в C:\cpsk_envs\pyrevit_rocket (вне OneDrive).
"""

__title__ = "Окружение"
__author__ = "CPSK"

import clr
import os
import sys
import subprocess
import shutil

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, ProgressBar,
    FormStartPosition, FormBorderStyle,
    GroupBox, ProgressBarStyle
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import script, forms

from cpsk_notify import show_error, show_info, show_success, show_confirm, show_toast

# === Проверка движка pyRevit ===
# CPSK Tools требует IronPython 2.7
import platform
_engine_name = platform.python_implementation()
_engine_version = sys.version.split()[0]

if _engine_name != "IronPython" or not _engine_version.startswith("2.7"):
    show_toast(
        "Неверный движок pyRevit",
        "CPSK Tools требует IronPython 2.7!\n"
        "Текущий: {} {}\n\n"
        "Откройте pyRevit Settings -> Engines\n"
        "и выберите IronPython 2 Engine.".format(_engine_name, _engine_version),
        notification_type="warning",
        auto_close=15
    )

# Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR))))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Проверка авторизации
from cpsk_auth import require_auth
if not require_auth():
    import sys as _sys
    _sys.exit()

from cpsk_config import (
    get_setting, set_setting,
    get_venv_path, get_requirements_path, get_venv_python, get_venv_pip,
    find_system_python, get_python_version, check_environment,
    VENV_BASE_DIR, reset_environment_cache
)

output = script.get_output()


class SetupEnvForm(Form):
    """Диалог проверки и установки окружения."""

    def __init__(self):
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "CPSK - Окружение"
        self.Width = 520
        self.Height = 530
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # === Общий статус ===
        self.lbl_overall = Label()
        self.lbl_overall.Location = Point(20, y)
        self.lbl_overall.Size = Size(470, 28)
        self.lbl_overall.Font = Font(self.lbl_overall.Font.FontFamily, 12, FontStyle.Bold)
        self.Controls.Add(self.lbl_overall)

        y += 40

        # === Группа: Статус проверки ===
        grp_status = GroupBox()
        grp_status.Text = "Статус окружения"
        grp_status.Location = Point(15, y)
        grp_status.Size = Size(475, 155)

        detail_y = 22

        # Venv
        lbl_venv_title = Label()
        lbl_venv_title.Text = "Виртуальное окружение:"
        lbl_venv_title.Location = Point(15, detail_y)
        lbl_venv_title.Size = Size(160, 18)
        grp_status.Controls.Add(lbl_venv_title)

        self.lbl_venv_status = Label()
        self.lbl_venv_status.Location = Point(180, detail_y)
        self.lbl_venv_status.Size = Size(280, 18)
        grp_status.Controls.Add(self.lbl_venv_status)

        detail_y += 24

        # Python версия
        lbl_py_title = Label()
        lbl_py_title.Text = "Python:"
        lbl_py_title.Location = Point(15, detail_y)
        lbl_py_title.Size = Size(160, 18)
        grp_status.Controls.Add(lbl_py_title)

        self.lbl_python_status = Label()
        self.lbl_python_status.Location = Point(180, detail_y)
        self.lbl_python_status.Size = Size(280, 18)
        grp_status.Controls.Add(self.lbl_python_status)

        detail_y += 24

        # Пакеты
        lbl_pkg_title = Label()
        lbl_pkg_title.Text = "Ключевые пакеты:"
        lbl_pkg_title.Location = Point(15, detail_y)
        lbl_pkg_title.Size = Size(160, 18)
        grp_status.Controls.Add(lbl_pkg_title)

        self.lbl_packages_status = Label()
        self.lbl_packages_status.Location = Point(180, detail_y)
        self.lbl_packages_status.Size = Size(280, 18)
        grp_status.Controls.Add(self.lbl_packages_status)

        detail_y += 24

        # Requirements
        lbl_req_title = Label()
        lbl_req_title.Text = "Requirements.txt:"
        lbl_req_title.Location = Point(15, detail_y)
        lbl_req_title.Size = Size(160, 18)
        grp_status.Controls.Add(lbl_req_title)

        self.lbl_req_status = Label()
        self.lbl_req_status.Location = Point(180, detail_y)
        self.lbl_req_status.Size = Size(280, 18)
        grp_status.Controls.Add(self.lbl_req_status)

        detail_y += 24

        # Соответствие пакетов
        lbl_match_title = Label()
        lbl_match_title.Text = "Соответствие пакетов:"
        lbl_match_title.Location = Point(15, detail_y)
        lbl_match_title.Size = Size(160, 18)
        grp_status.Controls.Add(lbl_match_title)

        self.lbl_match_status = Label()
        self.lbl_match_status.Location = Point(180, detail_y)
        self.lbl_match_status.Size = Size(280, 18)
        grp_status.Controls.Add(self.lbl_match_status)

        self.Controls.Add(grp_status)

        y += 170

        # === Группа: Пути ===
        grp_paths = GroupBox()
        grp_paths.Text = "Пути"
        grp_paths.Location = Point(15, y)
        grp_paths.Size = Size(475, 70)

        self.lbl_venv_path = Label()
        self.lbl_venv_path.Location = Point(15, 20)
        self.lbl_venv_path.Size = Size(445, 18)
        self.lbl_venv_path.ForeColor = Color.FromArgb(80, 80, 80)
        grp_paths.Controls.Add(self.lbl_venv_path)

        self.lbl_req_path = Label()
        self.lbl_req_path.Location = Point(15, 42)
        self.lbl_req_path.Size = Size(445, 18)
        self.lbl_req_path.ForeColor = Color.FromArgb(80, 80, 80)
        grp_paths.Controls.Add(self.lbl_req_path)

        self.Controls.Add(grp_paths)

        y += 85

        # === Группа: Системный Python ===
        grp_python = GroupBox()
        grp_python.Text = "Системный Python (для создания venv)"
        grp_python.Location = Point(15, y)
        grp_python.Size = Size(475, 55)

        self.txt_python = TextBox()
        self.txt_python.Location = Point(15, 22)
        self.txt_python.Width = 365
        grp_python.Controls.Add(self.txt_python)

        btn_browse = Button()
        btn_browse.Text = "Обзор..."
        btn_browse.Location = Point(390, 20)
        btn_browse.Width = 70
        btn_browse.Click += self.on_browse_python
        grp_python.Controls.Add(btn_browse)

        self.Controls.Add(grp_python)

        y += 70

        # === Прогресс ===
        self.progress = ProgressBar()
        self.progress.Location = Point(15, y)
        self.progress.Size = Size(475, 20)
        self.progress.Visible = False
        self.Controls.Add(self.progress)

        y += 25

        self.lbl_progress = Label()
        self.lbl_progress.Location = Point(15, y)
        self.lbl_progress.Size = Size(475, 20)
        self.Controls.Add(self.lbl_progress)

        y += 30

        # === Кнопки ===
        self.btn_install = Button()
        self.btn_install.Text = "Установить окружение"
        self.btn_install.Location = Point(15, y)
        self.btn_install.Size = Size(155, 32)
        self.btn_install.Click += self.on_install
        self.Controls.Add(self.btn_install)

        btn_refresh = Button()
        btn_refresh.Text = "Обновить"
        btn_refresh.Location = Point(180, y)
        btn_refresh.Size = Size(80, 32)
        btn_refresh.Click += self.on_refresh
        self.Controls.Add(btn_refresh)

        self.btn_delete = Button()
        self.btn_delete.Text = "Удалить"
        self.btn_delete.Location = Point(270, y)
        self.btn_delete.Size = Size(80, 32)
        self.btn_delete.Click += self.on_delete
        self.Controls.Add(self.btn_delete)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(400, y)
        btn_close.Size = Size(90, 32)
        btn_close.Click += self.on_close
        self.Controls.Add(btn_close)

        # Загрузка данных
        self.load_data()

    def load_data(self):
        """Загрузить данные."""
        # Системный Python
        python_path = get_setting("environment.python_path", "")
        if not python_path:
            python_path = find_system_python() or ""
        self.txt_python.Text = python_path

        # Пути
        self.lbl_venv_path.Text = "venv: {}".format(get_venv_path())
        self.lbl_req_path.Text = "req: {}".format(get_requirements_path())

        # Проверка
        self.run_check()

    def run_check(self):
        """Выполнить проверку окружения."""
        result = check_environment()

        # Общий статус
        if result["is_ready"]:
            self.lbl_overall.Text = "Окружение готово к работе"
            self.lbl_overall.ForeColor = Color.Green
            self.btn_install.Text = "Переустановить"
        elif result["needs_update"]:
            self.lbl_overall.Text = "! Требуется обновление окружения"
            self.lbl_overall.ForeColor = Color.FromArgb(255, 140, 0)  # Orange
            self.btn_install.Text = "Обновить окружение"
        else:
            self.lbl_overall.Text = "Окружение требует настройки"
            self.lbl_overall.ForeColor = Color.Red
            self.btn_install.Text = "Установить окружение"

        # Venv
        if result["venv_exists"]:
            self.lbl_venv_status.Text = "Установлено"
            self.lbl_venv_status.ForeColor = Color.Green
        else:
            self.lbl_venv_status.Text = "Не найдено"
            self.lbl_venv_status.ForeColor = Color.Red

        # Python
        if result["venv_version"]:
            self.lbl_python_status.Text = result["venv_version"]
            self.lbl_python_status.ForeColor = Color.Green
        else:
            self.lbl_python_status.Text = "Не определена"
            self.lbl_python_status.ForeColor = Color.Gray

        # Пакеты
        if result["packages_installed"]:
            self.lbl_packages_status.Text = "Установлены (ifcopenshell, ifctester)"
            self.lbl_packages_status.ForeColor = Color.Green
        else:
            self.lbl_packages_status.Text = "Не установлены"
            self.lbl_packages_status.ForeColor = Color.FromArgb(255, 140, 0)

        # Requirements
        if result["requirements_exists"]:
            self.lbl_req_status.Text = "Найден ({} пакетов)".format(result["requirements_count"])
            self.lbl_req_status.ForeColor = Color.Green
        else:
            self.lbl_req_status.Text = "Не найден"
            self.lbl_req_status.ForeColor = Color.Red

        # Соответствие пакетов
        if not result["venv_exists"]:
            self.lbl_match_status.Text = "- (venv не установлен)"
            self.lbl_match_status.ForeColor = Color.Gray
        elif result["packages_match"]:
            self.lbl_match_status.Text = "Все пакеты актуальны"
            self.lbl_match_status.ForeColor = Color.Green
        else:
            # Формируем сообщение о несоответствии
            issues = []
            if result["missing_packages"]:
                issues.append("нет: {}".format(", ".join(result["missing_packages"])))
            if result["outdated_packages"]:
                # Показываем версии: pkg (0.7.0 -> >=0.8.0)
                outdated_info = []
                for p in result["outdated_packages"]:
                    outdated_info.append("{}: {}->{}".format(
                        p["package"], p["installed"], p["required"]))
                issues.append("версии: {}".format(", ".join(outdated_info)))

            if issues:
                self.lbl_match_status.Text = "! " + "; ".join(issues)
            else:
                self.lbl_match_status.Text = "! Требуется обновление"
            self.lbl_match_status.ForeColor = Color.FromArgb(255, 140, 0)

    def on_browse_python(self, sender, args):
        """Выбор Python."""
        selected = forms.pick_file(file_ext='exe', title="Выберите python.exe")
        if selected:
            self.txt_python.Text = selected

    def on_refresh(self, sender, args):
        """Обновить проверку."""
        self.run_check()
        self.lbl_progress.Text = "Проверка выполнена"
        self.lbl_progress.ForeColor = Color.Green

    def on_delete(self, sender, args):
        """Удалить окружение."""
        venv_path = get_venv_path()

        if not os.path.exists(venv_path):
            show_info("Информация", "Окружение не найдено",
                      details="Путь: {}".format(venv_path))
            return

        # Подтверждение
        if not show_confirm("Подтверждение", "Удалить виртуальное окружение?",
                            details="Путь: {}".format(venv_path)):
            return

        try:
            self.lbl_progress.Text = "Удаление окружения..."
            self.lbl_progress.ForeColor = Color.Black
            System.Windows.Forms.Application.DoEvents()

            shutil.rmtree(venv_path)

            set_setting("environment.installed", False)
            self.run_check()

            self.lbl_progress.Text = "Окружение удалено"
            self.lbl_progress.ForeColor = Color.Green

            show_success("Готово", "Окружение успешно удалено")

        except Exception as e:
            self.lbl_progress.Text = "Ошибка удаления"
            self.lbl_progress.ForeColor = Color.Red
            show_error("Ошибка", "Ошибка удаления окружения",
                       details=str(e))

    def run_command(self, cmd, description, log_lines):
        """Выполнить команду с логированием."""
        log_lines.append("=" * 50)
        log_lines.append("Команда: {}".format(description))
        log_lines.append("Запуск: {}".format(" ".join(cmd)))
        log_lines.append("-" * 50)

        try:
            # Очищаем переменные окружения IronPython, которые мешают CPython
            # Без этого CPython не находит модуль encodings
            env = os.environ.copy()
            env.pop('PYTHONHOME', None)
            env.pop('PYTHONPATH', None)
            env.pop('IRONPYTHONPATH', None)

            # CREATE_NO_WINDOW = 0x08000000 - скрывает окно CMD
            CREATE_NO_WINDOW = 0x08000000
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                shell=False,
                creationflags=CREATE_NO_WINDOW,
                env=env
            )
            stdout, stderr = process.communicate()

            stdout_text = stdout.decode('utf-8', errors='replace') if stdout else ""
            stderr_text = stderr.decode('utf-8', errors='replace') if stderr else ""

            if stdout_text:
                log_lines.append("STDOUT:")
                log_lines.append(stdout_text)
            if stderr_text:
                log_lines.append("STDERR:")
                log_lines.append(stderr_text)

            log_lines.append("Код возврата: {}".format(process.returncode))
            log_lines.append("")

            return process.returncode, stdout_text, stderr_text

        except Exception as e:
            log_lines.append("ИСКЛЮЧЕНИЕ: {}".format(str(e)))
            log_lines.append("")
            return -1, "", str(e)

    def save_log(self, log_lines, success=False):
        """Сохранить лог в файл."""
        from datetime import datetime
        import codecs

        # Лог в корне extension
        log_path = os.path.join(EXTENSION_DIR, "setup_env.log")

        with codecs.open(log_path, 'w', 'utf-8') as f:
            f.write("CPSK Setup Environment Log\n")
            f.write("Дата: {}\n".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            f.write("Статус: {}\n\n".format("УСПЕХ" if success else "ОШИБКА"))
            f.write("\n".join(log_lines))

        return log_path

    def on_install(self, sender, args):
        """Установить окружение."""
        python_path = self.txt_python.Text.strip()
        venv_path = get_venv_path()
        req_path = get_requirements_path()

        log_lines = []
        log_lines.append("Python: {}".format(python_path))
        log_lines.append("Venv: {}".format(venv_path))
        log_lines.append("Requirements: {}".format(req_path))
        log_lines.append("Base dir: {}".format(VENV_BASE_DIR))
        log_lines.append("")

        if not python_path or not os.path.exists(python_path):
            show_error("Ошибка", "Укажите корректный путь к Python!",
                       details="Указанный путь: {}".format(python_path))
            return

        if not os.path.exists(req_path):
            show_error("Ошибка", "Файл requirements.txt не найден",
                       details="Путь: {}".format(req_path))
            return

        # Сохраняем путь к Python
        set_setting("environment.python_path", python_path)

        pip_path = os.path.join(venv_path, "Scripts", "pip.exe")
        venv_python_exe = os.path.join(venv_path, "Scripts", "python.exe")

        log_lines.append("=" * 50)
        log_lines.append("НАЧАЛО УСТАНОВКИ")
        log_lines.append("=" * 50)
        log_lines.append("")

        try:
            # Шаг 0: Создание базовой директории
            self.start_progress("Подготовка...")

            if not os.path.exists(VENV_BASE_DIR):
                log_lines.append(">>> ШАГ 0: СОЗДАНИЕ БАЗОВОЙ ДИРЕКТОРИИ")
                log_lines.append("Путь: {}".format(VENV_BASE_DIR))
                os.makedirs(VENV_BASE_DIR)
                log_lines.append("Создана успешно")
                log_lines.append("")

            # Шаг 1: Удаление старого venv
            self.lbl_progress.Text = "Шаг 1/3: Удаление старого venv..."
            System.Windows.Forms.Application.DoEvents()

            log_lines.append(">>> ШАГ 1: УДАЛЕНИЕ СТАРОГО VENV")
            log_lines.append("Путь: {}".format(venv_path))

            if os.path.exists(venv_path):
                log_lines.append("Папка существует, удаляем...")
                shutil.rmtree(venv_path)
                log_lines.append("Папка удалена")
            else:
                log_lines.append("Папка не существует, пропускаем")

            log_lines.append("")

            # Шаг 2: Создание venv
            self.lbl_progress.Text = "Шаг 2/3: Создание venv..."
            System.Windows.Forms.Application.DoEvents()

            log_lines.append(">>> ШАГ 2: СОЗДАНИЕ VENV")

            code, stdout, stderr = self.run_command(
                [python_path, "-m", "venv", venv_path],
                "Создание venv",
                log_lines
            )

            if code != 0:
                log_path = self.save_log(log_lines, success=False)
                raise Exception("Ошибка создания venv (код {})\nЛог: {}".format(code, log_path))

            # Проверяем что pip создан
            if not os.path.exists(pip_path):
                log_path = self.save_log(log_lines, success=False)
                raise Exception("pip.exe не найден после создания venv\nЛог: {}".format(log_path))

            log_lines.append("Venv создан успешно")
            log_lines.append("")

            # Шаг 3: Установка пакетов
            self.lbl_progress.Text = "Шаг 3/3: Установка пакетов..."
            System.Windows.Forms.Application.DoEvents()

            log_lines.append(">>> ШАГ 3: УСТАНОВКА ПАКЕТОВ")

            # Обновление pip (--no-cache-dir чтобы избежать проблем с OneDrive)
            log_lines.append("Обновление pip...")
            code, stdout, stderr = self.run_command(
                [venv_python_exe, "-m", "pip", "install", "--no-cache-dir", "--upgrade", "pip"],
                "Обновление pip",
                log_lines
            )

            # Установка пакетов (--no-cache-dir чтобы избежать проблем с OneDrive)
            log_lines.append("Установка из requirements.txt...")
            code, stdout, stderr = self.run_command(
                [pip_path, "install", "--no-cache-dir", "-r", req_path],
                "Установка пакетов",
                log_lines
            )

            if code != 0:
                log_path = self.save_log(log_lines, success=False)
                raise Exception("Ошибка установки пакетов (код {})\nЛог: {}".format(code, log_path))

            log_lines.append("")
            log_lines.append("=" * 50)
            log_lines.append("УСТАНОВКА ЗАВЕРШЕНА УСПЕШНО")
            log_lines.append("=" * 50)

            self.stop_progress()
            self.run_check()

            set_setting("environment.installed", True)
            reset_environment_cache()  # Сбросить кэш для новых проверок

            # Сохраняем успешный лог
            log_path = self.save_log(log_lines, success=True)

            self.lbl_progress.Text = "Установка завершена!"
            self.lbl_progress.ForeColor = Color.Green

            show_success("Готово", "Окружение успешно установлено!",
                         details="Путь: {}\nЛог: {}".format(venv_path, log_path))

        except Exception as e:
            self.stop_progress()
            error_msg = str(e)
            self.lbl_progress.Text = "Ошибка! См. лог"
            self.lbl_progress.ForeColor = Color.Red

            # Сохраняем лог если ещё не сохранён
            if "Лог:" not in error_msg:
                log_lines.append("ИСКЛЮЧЕНИЕ: {}".format(error_msg))
                log_path = self.save_log(log_lines, success=False)
                error_msg = "{}\n\nЛог: {}".format(error_msg, log_path)

            show_error("Ошибка", "Ошибка установки окружения",
                       details=error_msg)

    def start_progress(self, message):
        """Показать прогресс."""
        self.progress.Visible = True
        self.progress.Style = ProgressBarStyle.Marquee
        self.lbl_progress.Text = message
        self.lbl_progress.ForeColor = Color.Black
        self.btn_install.Enabled = False
        System.Windows.Forms.Application.DoEvents()

    def stop_progress(self):
        """Скрыть прогресс."""
        self.progress.Visible = False
        self.btn_install.Enabled = True

    def on_close(self, sender, args):
        """Закрыть."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    form = SetupEnvForm()
    form.ShowDialog()
