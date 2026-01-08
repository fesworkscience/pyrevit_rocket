# -*- coding: utf-8 -*-
"""Create building levels (elevations)."""

__title__ = "Create\nLevels"
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
from cpsk_notify import show_error, show_warning
if not require_auth():
    sys.exit()

doc = revit.doc
output = script.get_output()


class LevelsForm(WinForms.Form):
    """Form for levels creation parameters."""

    def __init__(self):
        self.result = None
        self.setup_form()

    def setup_form(self):
        self.Text = "Create Levels - CPSK"
        self.Width = 450
        self.Height = 480
        self.StartPosition = WinForms.FormStartPosition.CenterScreen
        self.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False

        y = 20

        # === STANDARD LEVELS ===
        self.add_label("STANDARD LEVELS", 20, y, bold=True)
        y += 30

        self.chk_foundation = WinForms.CheckBox()
        self.chk_foundation.Text = "Foundation bottom"
        self.chk_foundation.Location = Drawing.Point(20, y)
        self.chk_foundation.AutoSize = True
        self.chk_foundation.Checked = True
        self.Controls.Add(self.chk_foundation)

        self.add_label("Elevation (mm):", 250, y)
        self.txt_fnd_elev = self.add_textbox("-1500", 360, y)
        y += 30

        self.chk_ground = WinForms.CheckBox()
        self.chk_ground.Text = "Ground level (0.000)"
        self.chk_ground.Location = Drawing.Point(20, y)
        self.chk_ground.AutoSize = True
        self.chk_ground.Checked = True
        self.Controls.Add(self.chk_ground)
        y += 30

        self.chk_column_top = WinForms.CheckBox()
        self.chk_column_top.Text = "Column top"
        self.chk_column_top.Location = Drawing.Point(20, y)
        self.chk_column_top.AutoSize = True
        self.chk_column_top.Checked = True
        self.Controls.Add(self.chk_column_top)

        self.add_label("Elevation (mm):", 250, y)
        self.txt_col_top = self.add_textbox("7772", 360, y)
        y += 30

        self.chk_roof = WinForms.CheckBox()
        self.chk_roof.Text = "Roof level"
        self.chk_roof.Location = Drawing.Point(20, y)
        self.chk_roof.AutoSize = True
        self.chk_roof.Checked = True
        self.Controls.Add(self.chk_roof)

        self.add_label("Elevation (mm):", 250, y)
        self.txt_roof = self.add_textbox("10000", 360, y)
        y += 50

        # === CUSTOM LEVELS ===
        self.add_label("CUSTOM LEVELS", 20, y, bold=True)
        y += 25

        self.add_label("Enter elevations separated by comma (mm):", 20, y)
        y += 25

        self.txt_custom = WinForms.TextBox()
        self.txt_custom.Location = Drawing.Point(20, y)
        self.txt_custom.Width = 390
        self.txt_custom.Text = ""
        self.Controls.Add(self.txt_custom)
        y += 30

        self.add_label("Example: 3000, 6000, 9000", 20, y)
        y += 40

        # === NAMING ===
        self.add_label("NAMING OPTIONS", 20, y, bold=True)
        y += 30

        self.add_label("Name prefix:", 20, y)
        self.txt_prefix = self.add_textbox("Level", 150, y)
        y += 50

        # === BUTTONS ===
        btn_create = WinForms.Button()
        btn_create.Text = "Create Levels"
        btn_create.Location = Drawing.Point(100, y)
        btn_create.Width = 120
        btn_create.Click += self.on_create
        self.Controls.Add(btn_create)

        btn_cancel = WinForms.Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Location = Drawing.Point(240, y)
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
        textbox.Width = 70
        self.Controls.Add(textbox)
        return textbox

    def on_create(self, sender, args):
        try:
            levels = []

            # Standard levels
            if self.chk_foundation.Checked:
                levels.append({
                    'name': 'Foundation Bottom',
                    'elevation': float(self.txt_fnd_elev.Text)
                })

            if self.chk_ground.Checked:
                levels.append({
                    'name': 'Ground Level',
                    'elevation': 0.0
                })

            if self.chk_column_top.Checked:
                levels.append({
                    'name': 'Column Top',
                    'elevation': float(self.txt_col_top.Text)
                })

            if self.chk_roof.Checked:
                levels.append({
                    'name': 'Roof Level',
                    'elevation': float(self.txt_roof.Text)
                })

            # Custom levels
            if self.txt_custom.Text.strip():
                prefix = self.txt_prefix.Text.strip() or "Level"
                custom_elevs = [
                    float(x.strip())
                    for x in self.txt_custom.Text.split(',')
                    if x.strip()
                ]
                for i, elev in enumerate(custom_elevs):
                    levels.append({
                        'name': "{} {}".format(prefix, i + 1),
                        'elevation': elev
                    })

            if not levels:
                show_warning("Уровни", "Выберите хотя бы один уровень для создания")
                return

            self.result = {'levels': levels}
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


def create_levels(params):
    """Create levels based on parameters."""

    levels = params['levels']
    created_count = 0

    # Sort by elevation
    levels.sort(key=lambda x: x['elevation'])

    with revit.Transaction("Create Levels"):
        for level_data in levels:
            name = level_data['name']
            elevation_mm = level_data['elevation']
            elevation_ft = mm_to_feet(elevation_mm)

            try:
                level = Level.Create(doc, elevation_ft)

                # Try to set name (may fail if name exists)
                try:
                    level.Name = name
                except:
                    # If name exists, add elevation to name
                    level.Name = "{} ({})".format(name, int(elevation_mm))

                created_count += 1
                output.print_md("- **{}**: {:.3f} m ({:.0f} mm)".format(
                    level.Name,
                    elevation_mm / 1000,
                    elevation_mm
                ))

            except Exception as e:
                output.print_md("**Error** creating level '{}': {}".format(
                    name, str(e)))

    return created_count


# === MAIN ===
if __name__ == "__main__":
    form = LevelsForm()
    result = form.ShowDialog()

    if result == WinForms.DialogResult.OK and form.result:
        output.print_md("## Levels Created")

        count = create_levels(form.result)

        output.print_md("---")
        output.print_md("**Total**: {} levels created".format(count))
