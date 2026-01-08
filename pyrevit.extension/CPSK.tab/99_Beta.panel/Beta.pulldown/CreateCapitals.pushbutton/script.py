#! python3
# -*- coding: utf-8 -*-
"""Create capitals (bearing plates) on top of columns."""

__title__ = "Create\nCapitals"
__author__ = "CPSK"

import os
import sys

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
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
uidoc = revit.uidoc
output = script.get_output()


class CapitalsForm(WinForms.Form):
    """Form for capitals creation parameters."""

    def __init__(self):
        self.result = None
        self.setup_form()
        self.load_family_types()
        self.load_levels()

    def setup_form(self):
        self.Text = "Create Capitals - CPSK"
        self.Width = 450
        self.Height = 520
        self.StartPosition = WinForms.FormStartPosition.CenterScreen
        self.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False

        y = 15

        # === GRID SETTINGS ===
        self.add_label("GRID SETTINGS", 20, y, bold=True)
        y += 28

        self.add_label("X axis count:", 20, y)
        self.txt_x_count = self.add_textbox("9", 220, y)
        y += 28

        self.add_label("Y axis count:", 20, y)
        self.txt_y_count = self.add_textbox("2", 220, y)
        y += 28

        self.add_label("X spacing (mm):", 20, y)
        self.txt_x_spacing = self.add_textbox("12000", 220, y)
        y += 28

        self.add_label("Y spacing (mm):", 20, y)
        self.txt_y_spacing = self.add_textbox("24000", 220, y)
        y += 35

        # === CAPITAL DIMENSIONS ===
        self.add_label("CAPITAL DIMENSIONS", 20, y, bold=True)
        y += 28

        self.add_label("Capital width (mm):", 20, y)
        self.txt_width = self.add_textbox("400", 220, y)
        y += 28

        self.add_label("Capital depth (mm):", 20, y)
        self.txt_depth = self.add_textbox("600", 220, y)
        y += 28

        self.add_label("Capital height (mm):", 20, y)
        self.txt_height = self.add_textbox("200", 220, y)
        y += 28

        self.add_label("Column top elevation (mm):", 20, y)
        self.txt_col_height = self.add_textbox("7772", 220, y)
        y += 35

        # === REVIT SETTINGS ===
        self.add_label("REVIT SETTINGS", 20, y, bold=True)
        y += 28

        self.add_label("Family type:", 20, y)
        self.cmb_family = WinForms.ComboBox()
        self.cmb_family.Location = Drawing.Point(150, y)
        self.cmb_family.Width = 250
        self.cmb_family.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.Controls.Add(self.cmb_family)
        y += 28

        self.add_label("Level:", 20, y)
        self.cmb_level = WinForms.ComboBox()
        self.cmb_level.Location = Drawing.Point(150, y)
        self.cmb_level.Width = 250
        self.cmb_level.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.Controls.Add(self.cmb_level)
        y += 45

        # === BUTTONS ===
        btn_create = WinForms.Button()
        btn_create.Text = "Create Capitals"
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
        textbox.Width = 100
        self.Controls.Add(textbox)
        return textbox

    def load_family_types(self):
        """Load available generic model or structural connection types."""
        # Try structural columns first (capitals can be column type)
        categories = [
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_StructuralFraming,
        ]

        self.family_types = {}

        for cat in categories:
            collector = FilteredElementCollector(doc)\
                .OfCategory(cat)\
                .WhereElementIsElementType()

            for fam_type in collector:
                name = Element.Name.GetValue(fam_type)
                cat_name = cat.ToString().replace("OST_", "")
                full_name = "[{}] {}".format(cat_name, name)
                self.family_types[full_name] = fam_type.Id
                self.cmb_family.Items.Add(full_name)

        if self.cmb_family.Items.Count > 0:
            self.cmb_family.SelectedIndex = 0

    def load_levels(self):
        """Load available levels."""
        collector = FilteredElementCollector(doc)\
            .OfClass(Level)\
            .WhereElementIsNotElementType()

        self.levels = {}
        sorted_levels = sorted(collector, key=lambda x: x.Elevation)

        for level in sorted_levels:
            name = level.Name
            self.levels[name] = level.Id
            self.cmb_level.Items.Add(name)

        if self.cmb_level.Items.Count > 0:
            self.cmb_level.SelectedIndex = 0

    def on_create(self, sender, args):
        try:
            self.result = {
                'x_count': int(self.txt_x_count.Text),
                'y_count': int(self.txt_y_count.Text),
                'x_spacing': float(self.txt_x_spacing.Text),
                'y_spacing': float(self.txt_y_spacing.Text),
                'width': float(self.txt_width.Text),
                'depth': float(self.txt_depth.Text),
                'height': float(self.txt_height.Text),
                'col_height': float(self.txt_col_height.Text),
                'family_type_id': self.family_types.get(
                    str(self.cmb_family.SelectedItem)),
                'level_id': self.levels.get(
                    str(self.cmb_level.SelectedItem)),
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


def create_capitals(params):
    """Create capitals on column grid."""

    x_count = params['x_count']
    y_count = params['y_count']
    x_spacing = mm_to_feet(params['x_spacing'])
    y_spacing = mm_to_feet(params['y_spacing'])

    cap_width = mm_to_feet(params['width'])
    cap_depth = mm_to_feet(params['depth'])
    cap_height = mm_to_feet(params['height'])
    col_height = mm_to_feet(params['col_height'])

    family_type_id = params['family_type_id']
    level_id = params['level_id']

    if not family_type_id:
        show_warning("Капители", "Не выбран тип семейства!")
        return 0

    family_type = doc.GetElement(family_type_id)
    level = doc.GetElement(level_id)

    # Get family category to determine creation method
    family = family_type.Family
    category_id = family.FamilyCategoryId

    created_count = 0

    with revit.Transaction("Create Capitals"):

        for ix in range(x_count):
            for iy in range(y_count):
                x = ix * x_spacing
                y = iy * y_spacing
                z = col_height  # Place on top of column

                point = XYZ(x, y, z)

                try:
                    # Create family instance
                    if category_id == ElementId(BuiltInCategory.OST_StructuralColumns):
                        instance = doc.Create.NewFamilyInstance(
                            point,
                            family_type,
                            level,
                            StructuralType.Column
                        )
                    else:
                        instance = doc.Create.NewFamilyInstance(
                            point,
                            family_type,
                            level,
                            StructuralType.NonStructural
                        )

                    created_count += 1

                except Exception as e:
                    output.print_md("**Error** at ({}, {}): {}".format(ix, iy, str(e)))

    return created_count


# === MAIN ===
if __name__ == "__main__":
    form = CapitalsForm()
    result = form.ShowDialog()

    if result == WinForms.DialogResult.OK and form.result:
        count = create_capitals(form.result)

        output.print_md("## Capitals Created")
        output.print_md("- **Total**: {} capitals".format(count))
        output.print_md("- **Grid**: {}x{}".format(
            form.result['x_count'],
            form.result['y_count']
        ))
        output.print_md("- **Position**: top of columns at {} mm".format(
            form.result['col_height']
        ))
        output.print_md("")
        output.print_md("### Dimensions")
        output.print_md("- Width: {} mm".format(form.result['width']))
        output.print_md("- Depth: {} mm".format(form.result['depth']))
        output.print_md("- Height: {} mm".format(form.result['height']))
