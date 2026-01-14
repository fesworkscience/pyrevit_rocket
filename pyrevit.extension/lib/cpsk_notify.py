# -*- coding: utf-8 -*-
"""
CPSK Notification Module
Единый модуль уведомлений в стандартном стиле WinForms.

Использование:
    from cpsk_notify import show_error, show_warning, show_info, show_success

    # Блокирующее уведомление (по умолчанию)
    show_error("Заголовок", "Сообщение об ошибке")

    # С деталями (разворачиваемое)
    show_error("Ошибка API", "Не удалось подключиться", details="HTTP 405\\nURL: https://...")

    # Неблокирующий тост (автозакрытие 7 сек)
    show_success("Готово", "Операция завершена", blocking=False, auto_close=7)

    # Другие типы
    show_warning("Внимание", "Это предупреждение")
    show_info("Информация", "Это информация")
    show_success("Успех", "Операция завершена")
"""

import clr
import ctypes

clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, Panel, Timer, LinkLabel,
    DockStyle, FormStartPosition, FormBorderStyle,
    ScrollBars, AnchorStyles, Clipboard, DialogResult
)
from System.Drawing import Point, Size, Font, FontStyle, Rectangle, Color
from System.Diagnostics import Process, ProcessStartInfo


def get_revit_window_bounds():
    """
    Получить границы окна Revit.
    Возвращает (left, top, right, bottom) или None если не удалось.
    """
    try:
        # Попробуем получить через pyrevit
        from pyrevit import HOST_APP
        hwnd = HOST_APP.proc_window
        if hwnd:
            # Используем Windows API для получения границ окна
            class RECT(ctypes.Structure):
                _fields_ = [
                    ('left', ctypes.c_long),
                    ('top', ctypes.c_long),
                    ('right', ctypes.c_long),
                    ('bottom', ctypes.c_long)
                ]
            rect = RECT()
            ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
            return (rect.left, rect.top, rect.right, rect.bottom)
    except:
        pass

    # Fallback - используем рабочую область экрана
    screen = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea
    return (screen.Left, screen.Top, screen.Right, screen.Bottom)


class NotificationType:
    """Типы уведомлений."""
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"
    SUCCESS = "success"


# Иконки и заголовки для типов уведомлений
NOTIFICATION_ICONS = {
    NotificationType.ERROR: "[X]",
    NotificationType.WARNING: "[!]",
    NotificationType.INFO: "[i]",
    NotificationType.SUCCESS: "[OK]"
}


class NotificationForm(Form):
    """Уведомление в стандартном стиле WinForms с возможностью развернуть детали."""

    def __init__(self, title, message, details=None, notification_type=NotificationType.INFO,
                 blocking=True, auto_close=0, link_url=None, link_text=None):
        """
        Args:
            title: Заголовок уведомления
            message: Краткое сообщение
            details: Детальная информация (опционально, разворачивается)
            notification_type: Тип уведомления (error/warning/info/success)
            blocking: True - модальное окно, False - неблокирующий тост
            auto_close: Секунды до автозакрытия (0 = не закрывать)
            link_url: URL для кликабельной ссылки (опционально)
            link_text: Текст ссылки (по умолчанию = link_url)
        """
        self.title_text = title
        self.message_text = message
        self.details_text = details or ""
        self.notification_type = notification_type
        self.blocking = blocking
        self.auto_close = auto_close
        self.link_url = link_url
        self.link_text = link_text or link_url
        self.expanded = False
        self.collapsed_height = 130
        self.expanded_height = 300
        self.timer = None
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        icon = NOTIFICATION_ICONS.get(self.notification_type, "[i]")

        # Базовые настройки формы - стандартный стиль WinForms
        self.Text = "{} {}".format(icon, self.title_text)
        self.Width = 400
        self.Height = self.collapsed_height
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.TopMost = True

        if self.blocking:
            # Блокирующее - по центру экрана
            self.StartPosition = FormStartPosition.CenterScreen
            self.ShowInTaskbar = True
        else:
            # Неблокирующее - справа внизу окна Revit
            self.StartPosition = FormStartPosition.Manual
            self.ShowInTaskbar = False
            self._update_position()

        y = 15

        # Сообщение
        self.lbl_message = Label()
        self.lbl_message.Text = self.message_text
        self.lbl_message.Location = Point(15, y)
        self.lbl_message.Size = Size(360, 40)
        self.Controls.Add(self.lbl_message)

        y += 45

        # Кликабельная ссылка (если задана)
        if self.link_url:
            self.lbl_link = LinkLabel()
            self.lbl_link.Text = self.link_text
            self.lbl_link.Location = Point(15, y)
            self.lbl_link.Size = Size(360, 20)
            self.lbl_link.LinkColor = Color.Blue
            self.lbl_link.LinkClicked += self.on_link_clicked
            self.Controls.Add(self.lbl_link)
            y += 25
            # Увеличиваем высоту формы
            self.collapsed_height += 25
            self.Height = self.collapsed_height

        # Кнопки
        btn_x = 15

        if self.details_text:
            self.btn_details = Button()
            self.btn_details.Text = "Детали..."
            self.btn_details.Location = Point(btn_x, y)
            self.btn_details.Width = 80
            self.btn_details.Click += self.on_toggle_details
            self.Controls.Add(self.btn_details)
            btn_x += 90

            self.btn_copy = Button()
            self.btn_copy.Text = "Копировать"
            self.btn_copy.Location = Point(btn_x, y)
            self.btn_copy.Width = 85
            self.btn_copy.Click += self.on_copy
            self.Controls.Add(self.btn_copy)

        btn_ok = Button()
        btn_ok.Text = "OK"
        btn_ok.Location = Point(self.Width - 90, y)
        btn_ok.Width = 70
        btn_ok.Click += self.on_close
        self.Controls.Add(btn_ok)
        self.AcceptButton = btn_ok
        self.CancelButton = btn_ok

        y += 35

        # Детали (скрыты по умолчанию)
        self.txt_details = TextBox()
        self.txt_details.Location = Point(15, y)
        self.txt_details.Size = Size(360, 120)
        self.txt_details.Multiline = True
        self.txt_details.WordWrap = True
        self.txt_details.ScrollBars = ScrollBars.Vertical
        self.txt_details.ReadOnly = True
        self.txt_details.Font = Font("Segoe UI", 9)
        # Нормализуем переносы строк для Windows
        details_normalized = self.details_text.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
        self.txt_details.Text = details_normalized
        self.txt_details.Visible = False
        self.Controls.Add(self.txt_details)

        # Таймер для автозакрытия
        if self.auto_close > 0:
            self.timer = Timer()
            self.timer.Interval = self.auto_close * 1000
            self.timer.Tick += self.on_timer_tick
            self.timer.Start()

    def _update_position(self):
        """Обновить позицию окна (справа внизу окна Revit)."""
        bounds = get_revit_window_bounds()
        left, top, right, bottom = bounds
        # Позиционируем справа внизу с отступом 20px
        x = right - self.Width - 20
        y = bottom - self.Height - 20
        self.Location = Point(x, y)

    def on_toggle_details(self, sender, args):
        """Показать/скрыть детали."""
        self.expanded = not self.expanded

        if self.expanded:
            self.Height = self.expanded_height
            self.txt_details.Visible = True
            self.btn_details.Text = "Скрыть"
        else:
            self.Height = self.collapsed_height
            self.txt_details.Visible = False
            self.btn_details.Text = "Детали..."

        # Обновляем позицию для неблокирующего окна
        if not self.blocking:
            self._update_position()

    def on_copy(self, sender, args):
        """Копировать детали в буфер обмена."""
        try:
            full_text = "=== {} ===\n{}\n\n=== Детали ===\n{}".format(
                self.title_text,
                self.message_text,
                self.details_text
            )
            Clipboard.SetText(full_text)

            # Визуальная обратная связь
            old_text = self.btn_copy.Text
            self.btn_copy.Text = "Скопировано!"
            self.btn_copy.Enabled = False

            # Вернуть через таймер
            restore_timer = Timer()
            restore_timer.Interval = 1500

            def on_tick(s, e):
                self.btn_copy.Text = old_text
                self.btn_copy.Enabled = True
                restore_timer.Stop()
                restore_timer.Dispose()

            restore_timer.Tick += on_tick
            restore_timer.Start()

        except Exception:
            pass

    def on_link_clicked(self, sender, args):
        """Открыть ссылку в браузере."""
        try:
            psi = ProcessStartInfo(self.link_url)
            psi.UseShellExecute = True
            Process.Start(psi)
        except Exception:
            # Если не удалось открыть - копируем в буфер
            try:
                Clipboard.SetText(self.link_url)
            except:
                pass

    def on_close(self, sender, args):
        """Закрыть форму."""
        if self.timer:
            self.timer.Stop()
            self.timer.Dispose()
        self.Close()

    def on_timer_tick(self, sender, args):
        """Автозакрытие по таймеру."""
        if self.timer:
            self.timer.Stop()
        self.Close()


class ConfirmationForm(Form):
    """Диалог подтверждения в стандартном стиле WinForms."""

    def __init__(self, title, message, details=None):
        self.title_text = title
        self.message_text = message
        self.details_text = details or ""
        self.expanded = False
        self.collapsed_height = 130
        self.expanded_height = 300
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "[?] {}".format(self.title_text)
        self.Width = 400
        self.Height = self.collapsed_height
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ShowInTaskbar = True
        self.TopMost = True
        self.StartPosition = FormStartPosition.CenterScreen

        y = 15

        # Сообщение
        lbl_message = Label()
        lbl_message.Text = self.message_text
        lbl_message.Location = Point(15, y)
        lbl_message.Size = Size(360, 40)
        self.Controls.Add(lbl_message)

        y += 45

        # Кнопки
        btn_x = 15

        if self.details_text:
            self.btn_details = Button()
            self.btn_details.Text = "Детали..."
            self.btn_details.Location = Point(btn_x, y)
            self.btn_details.Width = 80
            self.btn_details.Click += self.on_toggle_details
            self.Controls.Add(self.btn_details)

        # Кнопки Да/Нет справа
        btn_no = Button()
        btn_no.Text = "Нет"
        btn_no.Location = Point(self.Width - 170, y)
        btn_no.Width = 70
        btn_no.Click += self.on_no
        self.Controls.Add(btn_no)
        self.CancelButton = btn_no

        btn_yes = Button()
        btn_yes.Text = "Да"
        btn_yes.Location = Point(self.Width - 90, y)
        btn_yes.Width = 70
        btn_yes.Click += self.on_yes
        self.Controls.Add(btn_yes)
        self.AcceptButton = btn_yes

        y += 35

        # Детали (скрыты)
        self.txt_details = TextBox()
        self.txt_details.Location = Point(15, y)
        self.txt_details.Size = Size(360, 120)
        self.txt_details.Multiline = True
        self.txt_details.ScrollBars = ScrollBars.Both
        self.txt_details.ReadOnly = True
        self.txt_details.Font = Font("Consolas", 9)
        self.txt_details.Text = self.details_text
        self.txt_details.Visible = False
        self.Controls.Add(self.txt_details)

    def on_toggle_details(self, sender, args):
        """Показать/скрыть детали."""
        self.expanded = not self.expanded
        if self.expanded:
            self.Height = self.expanded_height
            self.txt_details.Visible = True
            self.btn_details.Text = "Скрыть"
        else:
            self.Height = self.collapsed_height
            self.txt_details.Visible = False
            self.btn_details.Text = "Детали..."

    def on_yes(self, sender, args):
        """Да."""
        self.DialogResult = DialogResult.Yes
        self.Close()

    def on_no(self, sender, args):
        """Нет."""
        self.DialogResult = DialogResult.No
        self.Close()


def _show_notification(title, message, details=None, notification_type=NotificationType.INFO,
                       blocking=True, auto_close=0, link_url=None, link_text=None):
    """Внутренняя функция показа уведомления."""
    form = NotificationForm(
        title, message, details,
        notification_type=notification_type,
        blocking=blocking,
        auto_close=auto_close,
        link_url=link_url,
        link_text=link_text
    )

    if blocking:
        form.ShowDialog()
    else:
        form.Show()


def show_error(title, message, details=None, blocking=True, auto_close=0, link_url=None, link_text=None):
    """
    Показать уведомление об ошибке.

    Args:
        title: Заголовок ошибки
        message: Краткое сообщение
        details: Детальная информация (опционально, разворачивается)
        blocking: True - модальное окно, False - неблокирующий тост
        auto_close: Секунды до автозакрытия (0 = не закрывать, только для blocking=False)
        link_url: URL для кликабельной ссылки (опционально)
        link_text: Текст ссылки (по умолчанию = link_url)
    """
    _show_notification(title, message, details, NotificationType.ERROR, blocking, auto_close, link_url, link_text)


def show_warning(title, message, details=None, blocking=True, auto_close=0, link_url=None, link_text=None):
    """
    Показать предупреждение.

    Args:
        title: Заголовок
        message: Краткое сообщение
        details: Детальная информация (опционально)
        blocking: True - модальное окно, False - неблокирующий тост
        auto_close: Секунды до автозакрытия
        link_url: URL для кликабельной ссылки (опционально)
        link_text: Текст ссылки (по умолчанию = link_url)
    """
    _show_notification(title, message, details, NotificationType.WARNING, blocking, auto_close, link_url, link_text)


def show_info(title, message, details=None, blocking=True, auto_close=0, link_url=None, link_text=None):
    """
    Показать информационное уведомление.

    Args:
        title: Заголовок
        message: Краткое сообщение
        details: Детальная информация (опционально)
        blocking: True - модальное окно, False - неблокирующий тост
        auto_close: Секунды до автозакрытия
        link_url: URL для кликабельной ссылки (опционально)
        link_text: Текст ссылки (по умолчанию = link_url)
    """
    _show_notification(title, message, details, NotificationType.INFO, blocking, auto_close, link_url, link_text)


def show_success(title, message, details=None, blocking=True, auto_close=0, link_url=None, link_text=None):
    """
    Показать уведомление об успехе.

    Args:
        title: Заголовок
        message: Краткое сообщение
        details: Детальная информация (опционально)
        blocking: True - модальное окно, False - неблокирующий тост
        auto_close: Секунды до автозакрытия
        link_url: URL для кликабельной ссылки (опционально)
        link_text: Текст ссылки (по умолчанию = link_url)
    """
    _show_notification(title, message, details, NotificationType.SUCCESS, blocking, auto_close, link_url, link_text)


def show_toast(title, message, details=None, notification_type="success", auto_close=7):
    """
    Показать неблокирующий тост (удобная обёртка).

    Args:
        title: Заголовок
        message: Краткое сообщение
        details: Детальная информация (опционально)
        notification_type: Тип (success/warning/error/info)
        auto_close: Секунды до автозакрытия (по умолчанию 7)
    """
    type_map = {
        "success": NotificationType.SUCCESS,
        "warning": NotificationType.WARNING,
        "error": NotificationType.ERROR,
        "info": NotificationType.INFO
    }
    ntype = type_map.get(notification_type, NotificationType.SUCCESS)
    _show_notification(title, message, details, ntype, blocking=False, auto_close=auto_close)


def show_confirm(title, message, details=None):
    """
    Показать диалог подтверждения (Да/Нет).

    Args:
        title: Заголовок
        message: Текст вопроса
        details: Детальная информация (опционально)

    Returns:
        bool: True если пользователь нажал "Да", False если "Нет" или закрыл окно
    """
    form = ConfirmationForm(title, message, details)
    result = form.ShowDialog()
    return result == DialogResult.Yes
