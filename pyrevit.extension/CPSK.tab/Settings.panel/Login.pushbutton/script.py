# -*- coding: utf-8 -*-
"""
Login - Авторизация пользователя.
Заглушка для будущей реализации.
"""

__title__ = "Логин"
__author__ = "CPSK"

import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows.Forms import (
    Form, Label, TextBox, Button, Panel, CheckBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    DialogResult
)
from System.Drawing import Point, Size, Color, Font, FontStyle

from pyrevit import script

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
        self.Controls.Add(self.txt_password)

        y += 35

        # Запомнить меня
        self.chk_remember = CheckBox()
        self.chk_remember.Text = "Запомнить меня"
        self.chk_remember.Location = Point(20, y)
        self.chk_remember.AutoSize = True
        self.Controls.Add(self.chk_remember)

        y += 35

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
        self.lbl_status.Size = Size(320, 20)
        self.lbl_status.ForeColor = Color.Gray
        self.Controls.Add(self.lbl_status)

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

        # TODO: Реальная авторизация
        self.lbl_status.Text = "Функция в разработке..."
        self.lbl_status.ForeColor = Color.Orange

        MessageBox.Show(
            "Авторизация пока не реализована.\n\n"
            "Email: {}\n"
            "Запомнить: {}".format(email, "Да" if self.chk_remember.Checked else "Нет"),
            "В разработке",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )

    def on_cancel_click(self, sender, args):
        """Закрыть форму."""
        self.Close()


# === MAIN ===
if __name__ == "__main__":
    form = LoginForm()
    form.ShowDialog()
