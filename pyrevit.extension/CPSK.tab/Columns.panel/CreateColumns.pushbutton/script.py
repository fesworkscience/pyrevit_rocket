# -*- coding: utf-8 -*-
"""Create structural columns on grid intersections."""

__title__ = "Create\nColumns"
__author__ = "CPSK"

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI import *
from pyrevit import revit, forms, script

import System.Windows.Forms as WinForms
import System.Drawing as Drawing

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


class ColumnForm(WinForms.Form):
    """Form for column creation parameters."""

    def __init__(self):
        self.result = None
        self.setup_form()
        self.load_column_types()
        self.load_levels()

    def setup_form(self):
        self.Text = "Create Columns - CPSK"
        self.Width = 450
        self.Height = 550
        self.StartPosition = WinForms.FormStartPosition.CenterScreen
        self.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False

        y = 20

        # === GRID SECTION ===
        self.add_label("GRID SETTINGS", 20, y, bold=True)
        y += 30

        self.add_label("Columns X count:", 20, y)
        self.txt_x_count = self.add_textbox("9", 200, y)
        y += 30

        self.add_label("Columns Y count:", 20, y)
        self.txt_y_count = self.add_textbox("2", 200, y)
        y += 30

        self.add_label("X spacing (mm):", 20, y)
        self.txt_x_spacing = self.add_textbox("12000", 200, y)
        y += 30

        self.add_label("Y spacing (mm):", 20, y)
        self.txt_y_spacing = self.add_textbox("24000", 200, y)
        y += 40

        # === COLUMN PARAMETERS ===
        self.add_label("COLUMN PARAMETERS", 20, y, bold=True)
        y += 30

        self.add_label("Width (mm):", 20, y)
        self.txt_width = self.add_textbox("400", 200, y)
        y += 30

        self.add_label("Depth (mm):", 20, y)
        self.txt_depth = self.add_textbox("600", 200, y)
        y += 40

        # === REVIT SETTINGS ===
        self.add_label("REVIT SETTINGS", 20, y, bold=True)
        y += 30

        self.add_label("Column Type:", 20, y)
        self.cmb_column_type = WinForms.ComboBox()
        self.cmb_column_type.Location = Drawing.Point(200, y)
        self.cmb_column_type.Width = 200
        self.cmb_column_type.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.Controls.Add(self.cmb_column_type)
        y += 30

        self.add_label("Base Level:", 20, y)
        self.cmb_base_level = WinForms.ComboBox()
        self.cmb_base_level.Location = Drawing.Point(200, y)
        self.cmb_base_level.Width = 200
        self.cmb_base_level.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.Controls.Add(self.cmb_base_level)
        y += 30

        self.add_label("Top Level:", 20, y)
        self.cmb_top_level = WinForms.ComboBox()
        self.cmb_top_level.Location = Drawing.Point(200, y)
        self.cmb_top_level.Width = 200
        self.cmb_top_level.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.Controls.Add(self.cmb_top_level)
        y += 50

        # === BUTTONS ===
        btn_create = WinForms.Button()
        btn_create.Text = "Create Columns"
        btn_create.Location = Drawing.Point(100, y)
        btn_create.Width = 120
        btn_create.Click += self.on_create
        self.Controls.Add(btn_create)

        btn_cancel = WinForms.Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Location = Drawing.Point(230, y)
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

    def load_column_types(self):
        """Load available column types from project."""
        collector = FilteredElementCollector(doc)\
            .OfCategory(BuiltInCategory.OST_StructuralColumns)\
            .WhereElementIsElementType()

        self.column_types = {}
        for col_type in collector:
            name = Element.Name.GetValue(col_type)
            self.column_types[name] = col_type.Id
            self.cmb_column_type.Items.Add(name)

        if self.cmb_column_type.Items.Count > 0:
            self.cmb_column_type.SelectedIndex = 0

    def load_levels(self):
        """Load available levels from project."""
        collector = FilteredElementCollector(doc)\
            .OfClass(Level)\
            .WhereElementIsNotElementType()

        self.levels = {}
        sorted_levels = sorted(collector, key=lambda x: x.Elevation)

        for level in sorted_levels:
            name = level.Name
            self.levels[name] = level.Id
            self.cmb_base_level.Items.Add(name)
            self.cmb_top_level.Items.Add(name)

        if self.cmb_base_level.Items.Count > 0:
            self.cmb_base_level.SelectedIndex = 0
        if self.cmb_top_level.Items.Count > 1:
            self.cmb_top_level.SelectedIndex = 1
        elif self.cmb_top_level.Items.Count > 0:
            self.cmb_top_level.SelectedIndex = 0

    def on_create(self, sender, args):
        try:
            self.result = {
                'x_count': int(self.txt_x_count.Text),
                'y_count': int(self.txt_y_count.Text),
                'x_spacing': float(self.txt_x_spacing.Text),
                'y_spacing': float(self.txt_y_spacing.Text),
                'width': float(self.txt_width.Text),
                'depth': float(self.txt_depth.Text),
                'column_type_id': self.column_types.get(
                    str(self.cmb_column_type.SelectedItem)),
                'base_level_id': self.levels.get(
                    str(self.cmb_base_level.SelectedItem)),
                'top_level_id': self.levels.get(
                    str(self.cmb_top_level.SelectedItem)),
            }
            self.DialogResult = WinForms.DialogResult.OK
            self.Close()
        except ValueError as e:
            WinForms.MessageBox.Show(
                "Invalid input: " + str(e),
                "Error",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Error
            )

    def on_cancel(self, sender, args):
        self.DialogResult = WinForms.DialogResult.Cancel
        self.Close()


def mm_to_feet(mm):
    """Convert millimeters to feet."""
    return mm / 304.8


def create_columns(params):
    """Create columns based on parameters."""

    x_count = params['x_count']
    y_count = params['y_count']
    x_spacing = mm_to_feet(params['x_spacing'])
    y_spacing = mm_to_feet(params['y_spacing'])

    column_type_id = params['column_type_id']
    base_level_id = params['base_level_id']
    top_level_id = params['top_level_id']

    if not column_type_id:
        forms.alert("No column type selected!")
        return 0

    base_level = doc.GetElement(base_level_id)
    top_level = doc.GetElement(top_level_id)
    column_type = doc.GetElement(column_type_id)

    created_count = 0

    with revit.Transaction("Create Columns"):
        for i in range(x_count):
            for j in range(y_count):
                x = i * x_spacing
                y = j * y_spacing

                point = XYZ(x, y, base_level.Elevation)

                try:
                    column = doc.Create.NewFamilyInstance(
                        point,
                        column_type,
                        base_level,
                        StructuralType.Column
                    )

                    # Set top level constraint
                    top_constraint = column.get_Parameter(
                        BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
                    if top_constraint:
                        top_constraint.Set(top_level_id)

                    # Set top offset to 0
                    top_offset = column.get_Parameter(
                        BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
                    if top_offset:
                        top_offset.Set(0.0)

                    created_count += 1

                except Exception as e:
                    output.print_md("**Error** creating column at ({}, {}): {}".format(
                        i, j, str(e)))

    return created_count


# === MAIN ===
if __name__ == "__main__":
    # Check for column types
    collector = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_StructuralColumns)\
        .WhereElementIsElementType()

    if not list(collector):
        forms.alert(
            "No structural column families loaded in project!\n\n"
            "Please load a column family first.",
            title="No Column Types"
        )
    else:
        form = ColumnForm()
        result = form.ShowDialog()

        if result == WinForms.DialogResult.OK and form.result:
            count = create_columns(form.result)

            if count > 0:
                output.print_md("## Columns Created")
                output.print_md("- **Total**: {} columns".format(count))
                output.print_md("- **Grid**: {}x{}".format(
                    form.result['x_count'],
                    form.result['y_count']
                ))
                output.print_md("- **Spacing**: {} x {} mm".format(
                    form.result['x_spacing'],
                    form.result['y_spacing']
                ))
