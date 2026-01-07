# -*- coding: utf-8 -*-
"""
CPSK Tools - Startup Script
Проверяет окружение при запуске и показывает всплывающее уведомление.
"""

import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, Button, Timer,
    FormStartPosition, FormBorderStyle, FormWindowState
)
from System.Drawing import Point, Size, Color, Font, FontStyle, ContentAlignment

# Добавляем lib в путь
EXTENSION_DIR = os.path.dirname(os.path.abspath(__file__))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

try:
    from cpsk_config import check_environment, get_setting
except ImportError:
    # Модуль ещё не установлен
    check_environment = None
    get_setting = None


class ToastNotification(Form):
    """Всплывающее уведомление (Toast) в правом нижнем углу."""

    def __init__(self, title, message, is_success=True, auto_close_seconds=5):
        self.auto_close_seconds = auto_close_seconds
        self.is_success = is_success
        self._title = title
        self._message = message
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "CPSK"
        self.Width = 320
        self.Height = 100
        self.FormBorderStyle = FormBorderStyle.None
        self.ShowInTaskbar = False
        self.TopMost = True
        self.StartPosition = FormStartPosition.Manual

        # Позиция в правом нижнем углу
        screen = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea
        self.Location = Point(screen.Right - self.Width - 10, screen.Bottom - self.Height - 10)

        # Цвет фона
        if self.is_success:
            self.BackColor = Color.FromArgb(240, 255, 240)  # Светло-зелёный
            border_color = Color.FromArgb(34, 139, 34)
        else:
            self.BackColor = Color.FromArgb(255, 250, 240)  # Светло-оранжевый
            border_color = Color.FromArgb(255, 140, 0)

        # Заголовок
        lbl_title = Label()
        lbl_title.Text = self._title
        lbl_title.Location = Point(10, 8)
        lbl_title.Size = Size(250, 20)
        lbl_title.Font = Font(lbl_title.Font.FontFamily, 10, FontStyle.Bold)
        lbl_title.ForeColor = border_color
        self.Controls.Add(lbl_title)

        # Сообщение
        lbl_msg = Label()
        lbl_msg.Text = self._message
        lbl_msg.Location = Point(10, 32)
        lbl_msg.Size = Size(280, 40)
        lbl_msg.ForeColor = Color.FromArgb(60, 60, 60)
        self.Controls.Add(lbl_msg)

        # Кнопка закрытия
        btn_close = Button()
        btn_close.Text = "X"
        btn_close.Location = Point(self.Width - 25, 5)
        btn_close.Size = Size(20, 20)
        btn_close.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        btn_close.ForeColor = Color.Gray
        btn_close.Click += self.on_close
        self.Controls.Add(btn_close)

        # Таймер для автозакрытия
        if self.auto_close_seconds > 0:
            self.timer = Timer()
            self.timer.Interval = self.auto_close_seconds * 1000
            self.timer.Tick += self.on_timer_tick
            self.timer.Start()

    def on_close(self, sender, args):
        """Закрыть уведомление."""
        if hasattr(self, 'timer'):
            self.timer.Stop()
        self.Close()

    def on_timer_tick(self, sender, args):
        """Автозакрытие по таймеру."""
        self.timer.Stop()
        self.Close()


def show_toast(title, message, is_success=True, auto_close_seconds=5):
    """Показать всплывающее уведомление (неблокирующее)."""
    try:
        toast = ToastNotification(title, message, is_success, auto_close_seconds)
        toast.Show()  # Show вместо ShowDialog - неблокирующее
    except Exception:
        pass  # Игнорируем ошибки в startup


def check_and_notify():
    """Проверить окружение и показать уведомление."""
    # Проверяем настройки
    if get_setting is None:
        return

    # Проверяем нужно ли показывать уведомление
    show_notification = get_setting("notifications.show_startup_check", True)
    if not show_notification:
        return

    # Проверяем окружение
    if check_environment is None:
        return

    try:
        result = check_environment()

        if result["is_ready"]:
            show_toast(
                "CPSK Tools",
                "Окружение готово к работе.\n{}".format(result.get("venv_version", "")),
                is_success=True,
                auto_close_seconds=3
            )
        else:
            show_env_warning = get_setting("notifications.show_env_warnings", True)
            if show_env_warning:
                errors = result.get("errors", [])
                error_msg = errors[0] if errors else "Требуется настройка"
                show_toast(
                    "CPSK Tools - Внимание",
                    "{}.\nНастройки > Установить окружение".format(error_msg),
                    is_success=False,
                    auto_close_seconds=7
                )
    except Exception:
        pass  # Игнорируем ошибки


# === MAIN - выполняется при загрузке extension ===
if __name__ == "__main__" or True:  # True чтобы выполнялось при импорте
    # Запускаем проверку асинхронно через небольшую задержку
    try:
        startup_timer = Timer()
        startup_timer.Interval = 2000  # 2 секунды после загрузки
        startup_timer.Tick += lambda s, e: (startup_timer.Stop(), check_and_notify())
        startup_timer.Start()
    except Exception:
        pass
