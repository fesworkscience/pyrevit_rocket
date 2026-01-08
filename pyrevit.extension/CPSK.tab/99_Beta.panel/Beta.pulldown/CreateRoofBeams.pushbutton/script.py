# -*- coding: utf-8 -*-
"""Create roof beams (end gable beams) following truss profile."""

__title__ = "Roof\nBeams"
__author__ = "CPSK"

import os
import sys
import math

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


class RoofBeamsForm(WinForms.Form):
    """Form for roof beams creation parameters."""

    def __init__(self):
        self.result = None
        self.setup_form()
        self.load_beam_types()
        self.load_levels()

    def setup_form(self):
        self.Text = "Create Roof Beams - CPSK"
        self.Width = 480
        self.Height = 550
        self.StartPosition = WinForms.FormStartPosition.CenterScreen
        self.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False

        y = 15

        # === BUILDING GEOMETRY ===
        self.add_label("BUILDING GEOMETRY", 20, y, bold=True)
        y += 28

        self.add_label("Building length (mm):", 20, y)
        self.txt_length = self.add_textbox("96000", 250, y)
        y += 28

        self.add_label("Building width / span (mm):", 20, y)
        self.txt_span = self.add_textbox("24000", 250, y)
        y += 35

        # === TRUSS PROFILE ===
        self.add_label("TRUSS PROFILE (for beam placement)", 20, y, bold=True)
        y += 28

        self.add_label("Column top height (mm):", 20, y)
        self.txt_col_height = self.add_textbox("7772", 250, y)
        y += 28

        self.add_label("Truss height at center (mm):", 20, y)
        self.txt_truss_height = self.add_textbox("2736", 250, y)
        y += 28

        self.add_label("Slope angle (degrees):", 20, y)
        self.txt_slope = self.add_textbox("8", 250, y)
        y += 28

        self.add_label("Panel count (per half span):", 20, y)
        self.txt_panel_count = self.add_textbox("4", 250, y)
        y += 35

        # === BEAM DIMENSIONS ===
        self.add_label("BEAM DIMENSIONS", 20, y, bold=True)
        y += 28

        self.add_label("Roof beam width (mm):", 20, y)
        self.txt_beam_width = self.add_textbox("271", 250, y)
        y += 28

        self.add_label("Roof beam height (mm):", 20, y)
        self.txt_beam_height = self.add_textbox("279", 250, y)
        y += 35

        # === REVIT SETTINGS ===
        self.add_label("REVIT SETTINGS", 20, y, bold=True)
        y += 28

        self.add_label("Beam family type:", 20, y)
        self.cmb_beam_type = WinForms.ComboBox()
        self.cmb_beam_type.Location = Drawing.Point(180, y)
        self.cmb_beam_type.Width = 250
        self.cmb_beam_type.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.Controls.Add(self.cmb_beam_type)
        y += 28

        self.add_label("Level:", 20, y)
        self.cmb_level = WinForms.ComboBox()
        self.cmb_level.Location = Drawing.Point(180, y)
        self.cmb_level.Width = 250
        self.cmb_level.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.Controls.Add(self.cmb_level)
        y += 45

        # === BUTTONS ===
        btn_create = WinForms.Button()
        btn_create.Text = "Create Roof Beams"
        btn_create.Location = Drawing.Point(110, y)
        btn_create.Width = 130
        btn_create.Click += self.on_create
        self.Controls.Add(btn_create)

        btn_cancel = WinForms.Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Location = Drawing.Point(260, y)
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

    def load_beam_types(self):
        """Load available beam/framing types."""
        collector = FilteredElementCollector(doc)\
            .OfCategory(BuiltInCategory.OST_StructuralFraming)\
            .WhereElementIsElementType()

        self.beam_types = {}
        for beam_type in collector:
            name = Element.Name.GetValue(beam_type)
            self.beam_types[name] = beam_type.Id
            self.cmb_beam_type.Items.Add(name)

        if self.cmb_beam_type.Items.Count > 0:
            self.cmb_beam_type.SelectedIndex = 0

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
                'length': float(self.txt_length.Text),
                'span': float(self.txt_span.Text),
                'col_height': float(self.txt_col_height.Text),
                'truss_height': float(self.txt_truss_height.Text),
                'slope': float(self.txt_slope.Text),
                'panel_count': int(self.txt_panel_count.Text),
                'beam_width': float(self.txt_beam_width.Text),
                'beam_height': float(self.txt_beam_height.Text),
                'beam_type_id': self.beam_types.get(
                    str(self.cmb_beam_type.SelectedItem)),
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


def calculate_roof_profile(params):
    """Calculate roof beam profile points following truss top chord."""
    span = params['span']
    truss_height = params['truss_height']
    slope = params['slope']
    panel_count = params['panel_count']
    col_height = params['col_height']

    total_panels = panel_count * 2
    panel_length = span / total_panels

    z_bottom = col_height
    z_top_center = z_bottom + truss_height

    slope_rad = math.radians(slope)
    rise = (span / 2.0) * math.tan(slope_rad)
    z_top_support = z_top_center - rise

    # Calculate top chord nodes (roof profile)
    profile_nodes = []
    for idx in range(total_panels + 1):
        y_pos = idx * panel_length
        if y_pos <= span / 2.0:
            t = y_pos / (span / 2.0) if span > 0 else 0
            z_pos = z_top_support + t * rise
        else:
            t = (y_pos - span / 2.0) / (span / 2.0) if span > 0 else 0
            z_pos = z_top_center - t * rise
        profile_nodes.append((y_pos, z_pos))

    return profile_nodes


def create_roof_beams(params):
    """Create roof beams at building ends following truss profile."""

    length = mm_to_feet(params['length'])
    span = params['span']

    beam_type_id = params['beam_type_id']
    level_id = params['level_id']

    if not beam_type_id:
        show_warning("Балки крыши", "Не выбран тип балки!")
        return 0

    beam_type = doc.GetElement(beam_type_id)
    level = doc.GetElement(level_id)

    # Calculate roof profile
    profile_nodes = calculate_roof_profile(params)

    # Convert to feet
    profile_ft = [(mm_to_feet(y), mm_to_feet(z)) for y, z in profile_nodes]

    # X positions for end beams (start and end of building)
    x_positions = [0, length]

    created_count = 0

    with revit.Transaction("Create Roof Beams"):

        for beam_x in x_positions:

            # Create beams along the roof profile
            for idx in range(len(profile_ft) - 1):
                y1, z1 = profile_ft[idx]
                y2, z2 = profile_ft[idx + 1]

                start_pt = XYZ(beam_x, y1, z1)
                end_pt = XYZ(beam_x, y2, z2)

                try:
                    line = Line.CreateBound(start_pt, end_pt)
                    beam = doc.Create.NewFamilyInstance(
                        line,
                        beam_type,
                        level,
                        StructuralType.Beam
                    )
                    created_count += 1

                except Exception as e:
                    output.print_md("**Error** creating roof beam: {}".format(str(e)))

    return created_count


# === MAIN ===
if __name__ == "__main__":
    # Check for beam types
    collector = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_StructuralFraming)\
        .WhereElementIsElementType()

    if not list(collector):
        show_warning("Балки крыши", "В проекте нет семейств балок!\nЗагрузите семейство балок.")
    else:
        form = RoofBeamsForm()
        result = form.ShowDialog()

        if result == WinForms.DialogResult.OK and form.result:
            count = create_roof_beams(form.result)

            total_panels = form.result['panel_count'] * 2
            beams_per_end = total_panels

            output.print_md("## Roof Beams Created")
            output.print_md("- **Total beams**: {}".format(count))
            output.print_md("- **Beams per end**: {}".format(beams_per_end))
            output.print_md("- **End count**: 2 (start and end)")
            output.print_md("")
            output.print_md("### Profile")
            output.print_md("- Span: {} mm".format(form.result['span']))
            output.print_md("- Truss height: {} mm".format(form.result['truss_height']))
            output.print_md("- Slope: {}°".format(form.result['slope']))
