#! python3
# -*- coding: utf-8 -*-
"""
Login - Авторизация пользователя.
Вход/выход из системы CPSK.
"""

__title__ = "Логин"
__author__ = "CPSK"

import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, Panel, CheckBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    Application, Cursor, Cursors
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import script

# Добавляем lib в путь если нужно
lib_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))), "lib")
if lib_path not in sys.path:
    sys.path.insert(0, lib_path)

from cpsk_auth import AuthService
from cpsk_notify import show_error, show_success, show_info, show_confirm

output = script.get_output()


class LoginForm(Form):
    """Диалог авторизации."""

    def __init__(self):
        self.result = None
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "CPSK - Авторизация"
        self.Width = 380
        self.Height = 280
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 20

        # Заголовок
        lbl_title = Label()
        lbl_title.Text = "Вход в систему CPSK"
        lbl_title.Location = Point(20, y)
        lbl_title.Size = Size(320, 25)
        lbl_title.Font = Font(lbl_title.Font.FontFamily, 12, FontStyle.Bold)
        self.Controls.Add(lbl_title)

        y += 40

        # Email / Логин
        lbl_email = Label()
        lbl_email.Text = "Email или логин:"
        lbl_email.Location = Point(20, y)
        lbl_email.AutoSize = True
        self.Controls.Add(lbl_email)

        y += 22
        self.txt_email = TextBox()
        self.txt_email.Location = Point(20, y)
        self.txt_email.Width = 320
        self.Controls.Add(self.txt_email)

        y += 35

        # Пароль
        lbl_password = Label()
        lbl_password.Text = "Пароль:"
        lbl_password.Location = Point(20, y)
        lbl_password.AutoSize = True
        self.Controls.Add(lbl_password)

        y += 22
        self.txt_password = TextBox()
        self.txt_password.Location = Point(20, y)
        self.txt_password.Width = 320
        self.txt_password.PasswordChar = '*'
        self.txt_password.KeyDown += self.on_key_down
        self.Controls.Add(self.txt_password)

        y += 40

        # Кнопки
        self.btn_login = Button()
        self.btn_login.Text = "Войти"
        self.btn_login.Location = Point(20, y)
        self.btn_login.Size = Size(100, 30)
        self.btn_login.Click += self.on_login_click
        self.Controls.Add(self.btn_login)

        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(240, y)
        btn_cancel.Size = Size(100, 30)
        btn_cancel.Click += self.on_cancel_click
        self.Controls.Add(btn_cancel)

        # Статус
        y += 40
        self.lbl_status = Label()
        self.lbl_status.Text = ""
        self.lbl_status.Location = Point(20, y)
        self.lbl_status.Size = Size(320, 40)
        self.lbl_status.ForeColor = Color.Gray
        self.Controls.Add(self.lbl_status)

    def on_key_down(self, sender, args):
        """Обработка Enter в поле пароля."""
        if args.KeyCode == System.Windows.Forms.Keys.Enter:
            self.on_login_click(sender, args)

    def on_login_click(self, sender, args):
        """Обработчик кнопки Войти."""
        email = self.txt_email.Text.strip()
        password = self.txt_password.Text

        if not email:
            self.lbl_status.Text = "Введите email или логин"
            self.lbl_status.ForeColor = Color.Red
            return

        if not password:
            self.lbl_status.Text = "Введите пароль"
            self.lbl_status.ForeColor = Color.Red
            return

        # Показываем статус загрузки
        self.lbl_status.Text = "Подключение к серверу..."
        self.lbl_status.ForeColor = Color.Gray
        self.btn_login.Enabled = False
        self.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        try:
            # Выполняем авторизацию (возвращает 3 значения: success, error, details)
            result = AuthService.login(email, password)
            success = result[0]
            error = result[1] if len(result) > 1 else None
            details = result[2] if len(result) > 2 else None

            if success:
                self.lbl_status.Text = "Успешный вход!"
                self.lbl_status.ForeColor = Color.Green
                Application.DoEvents()

                self.result = True
                self.Close()

                # Показываем уведомление после закрытия формы
                show_success(
                    "Авторизация успешна",
                    "Добро пожаловать, {}!".format(email)
                )
            else:
                self.lbl_status.Text = error or "Ошибка авторизации"
                self.lbl_status.ForeColor = Color.Red

                # Показываем детальную ошибку
                show_error(
                    "Ошибка авторизации",
                    error or "Не удалось войти в систему",
                    details
                )

        except Exception as e:
            self.lbl_status.Text = "Ошибка: {}".format(str(e))
            self.lbl_status.ForeColor = Color.Red

            show_error(
                "Ошибка",
                "Произошла непредвиденная ошибка",
                str(e)
            )

        finally:
            self.btn_login.Enabled = True
            self.Cursor = Cursors.Default

    def on_cancel_click(self, sender, args):
        """Закрыть форму."""
        self.Close()


class LoggedInForm(Form):
    """Диалог для авторизованного пользователя."""

    def __init__(self, username):
        self.username = username
        self.setup_form()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "CPSK - Профиль"
        self.Width = 350
        self.Height = 180
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 20

        # Заголовок
        lbl_title = Label()
        lbl_title.Text = "Вы авторизованы"
        lbl_title.Location = Point(20, y)
        lbl_title.Size = Size(290, 25)
        lbl_title.Font = Font(lbl_title.Font.FontFamily, 12, FontStyle.Bold)
        self.Controls.Add(lbl_title)

        y += 35

        # Имя пользователя
        lbl_user = Label()
        lbl_user.Text = "Пользователь: {}".format(self.username)
        lbl_user.Location = Point(20, y)
        lbl_user.Size = Size(290, 20)
        self.Controls.Add(lbl_user)

        y += 40

        # Кнопки
        btn_logout = Button()
        btn_logout.Text = "Выйти"
        btn_logout.Location = Point(20, y)
        btn_logout.Size = Size(100, 30)
        btn_logout.Click += self.on_logout_click
        self.Controls.Add(btn_logout)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(210, y)
        btn_close.Size = Size(100, 30)
        btn_close.Click += self.on_close_click
        self.Controls.Add(btn_close)

    def on_logout_click(self, sender, args):
        """Выход из системы."""
        if not show_confirm("Подтверждение", "Вы уверены, что хотите выйти?"):
            return

        AuthService.logout()
        show_info("Выход", "Вы вышли из системы")
        self.Close()

    def on_close_click(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    # Проверяем статус авторизации
    if AuthService.is_authenticated():
        # Уже авторизован - показываем профиль
        username = AuthService.get_username()
        form = LoggedInForm(username)
        form.ShowDialog()
    else:
        # Не авторизован - показываем форму входа
        form = LoginForm()
        form.ShowDialog()
