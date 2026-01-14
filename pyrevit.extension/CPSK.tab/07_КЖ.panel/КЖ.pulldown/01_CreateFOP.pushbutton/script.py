# -*- coding: utf-8 -*-
"""
Создать параметр ФОП - создание общего параметра для Smart Openings.

Создаёт параметр CPSK_RebarCutData и привязывает его
к категории Structural Rebar как Instance параметр.
"""

__title__ = "Создать\nпараметр ФОП"
__author__ = "CPSK"

# 1. Стандартные импорты
import clr
import os
import sys

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

# 2. System и WinForms импорты
import System
from System.Windows.Forms import (
    Form, Label, Button, Panel, TextBox, GroupBox,
    DockStyle, FormStartPosition, FormBorderStyle,
    DialogResult, AnchorStyles
)
from System.Drawing import Point, Size, Font, FontStyle, Color

# 3. Добавляем lib в путь
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# 4. Проверка авторизации
from cpsk_auth import require_auth
from cpsk_notify import show_error, show_success, show_warning, show_info

if not require_auth():
    sys.exit()

# 5. pyrevit и Revit API
from pyrevit import revit

from Autodesk.Revit.DB import Transaction

# 6. Импорт модулей для работы с параметрами
from cpsk_shared_params import (
    REBAR_CUT_DATA_PARAM,
    get_shared_param_file_path,
    check_shared_param_exists,
    check_param_visibility,
    ensure_rebar_cut_param_with_info
)

# 7. Настройки
doc = revit.doc
uidoc = revit.uidoc
app = doc.Application


# === ФОРМА ===

class CreateFOPForm(Form):
    """Форма для создания параметра ФОП."""

    def __init__(self):
        self.setup_form()
        self.update_status()

    def setup_form(self):
        """Настройка формы."""
        self.Text = "Создание параметра ФОП"
        self.Width = 500
        self.Height = 320
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False

        y = 15

        # Заголовок
        lbl_title = Label()
        lbl_title.Text = "Параметр для Smart Openings"
        lbl_title.Location = Point(15, y)
        lbl_title.Size = Size(460, 25)
        lbl_title.Font = Font(lbl_title.Font.FontFamily, 11, FontStyle.Bold)
        self.Controls.Add(lbl_title)
        y += 35

        # Группа информации
        grp_info = GroupBox()
        grp_info.Text = "Информация о параметре"
        grp_info.Location = Point(15, y)
        grp_info.Size = Size(460, 100)
        self.Controls.Add(grp_info)

        # Имя параметра
        lbl_name = Label()
        lbl_name.Text = "Имя: {}".format(REBAR_CUT_DATA_PARAM)
        lbl_name.Location = Point(15, 25)
        lbl_name.Size = Size(430, 20)
        grp_info.Controls.Add(lbl_name)

        # Категория
        lbl_cat = Label()
        lbl_cat.Text = "Категория: Structural Rebar (Несущая арматура)"
        lbl_cat.Location = Point(15, 45)
        lbl_cat.Size = Size(430, 20)
        grp_info.Controls.Add(lbl_cat)

        # Тип
        lbl_type = Label()
        lbl_type.Text = "Тип: Instance (по экземпляру), Text, Группа: Data"
        lbl_type.Location = Point(15, 65)
        lbl_type.Size = Size(430, 20)
        grp_info.Controls.Add(lbl_type)

        y += 110

        # Группа статуса
        grp_status = GroupBox()
        grp_status.Text = "Текущий статус"
        grp_status.Location = Point(15, y)
        grp_status.Size = Size(460, 60)
        self.Controls.Add(grp_status)

        self.lbl_status = Label()
        self.lbl_status.Text = "Проверка..."
        self.lbl_status.Location = Point(15, 22)
        self.lbl_status.Size = Size(430, 30)
        grp_status.Controls.Add(self.lbl_status)

        y += 70

        # Кнопки
        self.btn_create = Button()
        self.btn_create.Text = "Создать параметр"
        self.btn_create.Location = Point(120, y)
        self.btn_create.Size = Size(150, 35)
        self.btn_create.Click += self.on_create
        self.Controls.Add(self.btn_create)

        btn_close = Button()
        btn_close.Text = "Закрыть"
        btn_close.Location = Point(280, y)
        btn_close.Size = Size(100, 35)
        btn_close.Click += self.on_close
        self.Controls.Add(btn_close)

    def update_status(self):
        """Обновить статус параметра."""
        exists, is_bound = check_shared_param_exists(doc)
        exists_in_file, is_visible = check_param_visibility(app)

        if is_bound and is_visible:
            self.lbl_status.Text = "Параметр создан и привязан. Видимость: ДА"
            self.lbl_status.ForeColor = Color.Green
            self.btn_create.Enabled = False
            self.btn_create.Text = "Уже создан"
        elif is_bound and not is_visible:
            self.lbl_status.Text = "Параметр привязан, но СКРЫТ! Нажмите для исправления."
            self.lbl_status.ForeColor = Color.Red
            self.btn_create.Enabled = True
            self.btn_create.Text = "Исправить"
        elif exists_in_file and not is_visible:
            self.lbl_status.Text = "Параметр в файле ФОП скрыт. Нажмите для исправления."
            self.lbl_status.ForeColor = Color.Red
            self.btn_create.Enabled = True
            self.btn_create.Text = "Исправить"
        elif exists_in_file:
            self.lbl_status.Text = "Параметр есть в файле, но не привязан к проекту."
            self.lbl_status.ForeColor = Color.FromArgb(255, 140, 0)
            self.btn_create.Enabled = True
            self.btn_create.Text = "Привязать"
        else:
            self.lbl_status.Text = "Параметр не существует. Нажмите для создания."
            self.lbl_status.ForeColor = Color.Black
            self.btn_create.Enabled = True
            self.btn_create.Text = "Создать параметр"

    def on_create(self, sender, args):
        """Обработчик создания параметра."""
        try:
            with Transaction(doc, "Создать параметр ФОП") as t:
                t.Start()
                success, message, was_created = ensure_rebar_cut_param_with_info(doc, app, force_rebind=True)
                t.Commit()

            if success:
                show_success("Параметр создан", message)
                self.update_status()
            else:
                show_error("Ошибка", "Не удалось создать параметр", details=message)

        except Exception as e:
            show_error("Ошибка", "Исключение при создании параметра", details=str(e))

    def on_close(self, sender, args):
        """Обработчик закрытия."""
        self.DialogResult = DialogResult.Cancel
        self.Close()


# === MAIN ===

def main():
    """Главная функция."""
    form = CreateFOPForm()
    form.ShowDialog()


if __name__ == "__main__":
    main()
