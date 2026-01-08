#! python3
# -*- coding: utf-8 -*-
"""Create grid axes for building layout."""

__title__ = "Create\nAxes"
__author__ = "CPSK"

import os
import sys

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

import System.Windows.Forms as WinForms
import System.Drawing as Drawing

# Добавляем lib в путь для импорта cpsk_auth
SCRIPT_DIR = os.path.dirname(__file__)
LIB_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))), "lib")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# Проверка авторизации
from cpsk_auth import require_auth
from cpsk_notify import show_error
if not require_auth():
    sys.exit()

doc = revit.doc
output = script.get_output()


class AxesForm(WinForms.Form):
    """Form for grid axes creation parameters."""

    def __init__(self):
        self.result = None
        self.setup_form()

    def setup_form(self):
        self.Text = "Create Axes - CPSK"
        self.Width = 400
        self.Height = 400
        self.StartPosition = WinForms.FormStartPosition.CenterScreen
        self.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False

        y = 20

        # === X AXES (vertical lines) ===
        self.add_label("X AXES (Vertical - 1, 2, 3...)", 20, y, bold=True)
        y += 30

        self.add_label("Count:", 20, y)
        self.txt_x_count = self.add_textbox("9", 180, y)
        y += 30

        self.add_label("Spacing (mm):", 20, y)
        self.txt_x_spacing = self.add_textbox("12000", 180, y)
        y += 30

        self.add_label("Line length (mm):", 20, y)
        self.txt_x_length = self.add_textbox("30000", 180, y)
        y += 40

        # === Y AXES (horizontal lines) ===
        self.add_label("Y AXES (Horizontal - A, B, C...)", 20, y, bold=True)
        y += 30

        self.add_label("Count:", 20, y)
        self.txt_y_count = self.add_textbox("2", 180, y)
        y += 30

        self.add_label("Spacing (mm):", 20, y)
        self.txt_y_spacing = self.add_textbox("24000", 180, y)
        y += 30

        self.add_label("Line length (mm):", 20, y)
        self.txt_y_length = self.add_textbox("110000", 180, y)
        y += 50

        # === BUTTONS ===
        btn_create = WinForms.Button()
        btn_create.Text = "Create Axes"
        btn_create.Location = Drawing.Point(80, y)
        btn_create.Width = 100
        btn_create.Click += self.on_create
        self.Controls.Add(btn_create)

        btn_cancel = WinForms.Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Location = Drawing.Point(200, y)
        btn_cancel.Width = 100
        btn_cancel.Click += self.on_cancel
        self.Controls.Add(btn_cancel)

    def add_label(self, text, x, y, bold=False):
        label = WinForms.Label()
        label.Text = text
        label.Location = Drawing.Point(x, y)
        label.AutoSize = True
        if bold:
            label.Font = Drawing.Font(label.Font, Drawing.FontStyle.Bold)
        self.Controls.Add(label)
        return label

    def add_textbox(self, default, x, y):
        textbox = WinForms.TextBox()
        textbox.Text = default
        textbox.Location = Drawing.Point(x, y)
        textbox.Width = 100
        self.Controls.Add(textbox)
        return textbox

    def on_create(self, sender, args):
        try:
            self.result = {
                'x_count': int(self.txt_x_count.Text),
                'x_spacing': float(self.txt_x_spacing.Text),
                'x_length': float(self.txt_x_length.Text),
                'y_count': int(self.txt_y_count.Text),
                'y_spacing': float(self.txt_y_spacing.Text),
                'y_length': float(self.txt_y_length.Text),
            }
            self.DialogResult = WinForms.DialogResult.OK
            self.Close()
        except ValueError as e:
            show_error("Ошибка", "Неверный ввод", details=str(e))

    def on_cancel(self, sender, args):
        self.DialogResult = WinForms.DialogResult.Cancel
        self.Close()


def mm_to_feet(mm):
    """Convert millimeters to feet."""
    return mm / 304.8


def get_axis_name_x(index):
    """Get numeric axis name (1, 2, 3...)."""
    return str(index + 1)


def get_axis_name_y(index):
    """Get alphabetic axis name (A, B, C... AA, AB...)."""
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if index < 26:
        return letters[index]
    else:
        return letters[index // 26 - 1] + letters[index % 26]


def create_axes(params):
    """Create grid axes based on parameters."""

    x_count = params['x_count']
    x_spacing = mm_to_feet(params['x_spacing'])
    x_length = mm_to_feet(params['x_length'])

    y_count = params['y_count']
    y_spacing = mm_to_feet(params['y_spacing'])
    y_length = mm_to_feet(params['y_length'])

    # Calculate extents
    total_x = (x_count - 1) * x_spacing
    total_y = (y_count - 1) * y_spacing

    created_x = 0
    created_y = 0

    with revit.Transaction("Create Grid Axes"):

        # Create X axes (vertical lines - numbered)
        y_start = -x_length * 0.1  # 10% extension below
        y_end = total_y + x_length * 0.1  # 10% extension above

        for i in range(x_count):
            x = i * x_spacing

            start = XYZ(x, y_start, 0)
            end = XYZ(x, y_end, 0)
            line = Line.CreateBound(start, end)

            try:
                grid = Grid.Create(doc, line)
                grid.Name = get_axis_name_x(i)
                created_x += 1
            except Exception as e:
                output.print_md("**Error** creating X axis {}: {}".format(i + 1, str(e)))

        # Create Y axes (horizontal lines - lettered)
        x_start = -y_length * 0.1  # 10% extension left
        x_end = total_x + y_length * 0.1  # 10% extension right

        for j in range(y_count):
            y = j * y_spacing

            start = XYZ(x_start, y, 0)
            end = XYZ(x_end, y, 0)
            line = Line.CreateBound(start, end)

            try:
                grid = Grid.Create(doc, line)
                grid.Name = get_axis_name_y(j)
                created_y += 1
            except Exception as e:
                output.print_md("**Error** creating Y axis {}: {}".format(
                    get_axis_name_y(j), str(e)))

    return created_x, created_y


# === MAIN ===
if __name__ == "__main__":
    form = AxesForm()
    result = form.ShowDialog()

    if result == WinForms.DialogResult.OK and form.result:
        count_x, count_y = create_axes(form.result)

        output.print_md("## Axes Created")
        output.print_md("- **X axes** (1, 2, 3...): {}".format(count_x))
        output.print_md("- **Y axes** (A, B, C...): {}".format(count_y))
        output.print_md("- **Total**: {} axes".format(count_x + count_y))
