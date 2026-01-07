# -*- coding: utf-8 -*-
"""Create Warren-type trusses on column grid."""

__title__ = "Create\nTruss"
__author__ = "CPSK"

import clr
import math

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import revit, forms, script

import System.Windows.Forms as WinForms
import System.Drawing as Drawing

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


class TrussForm(WinForms.Form):
    """Form for truss creation parameters."""

    def __init__(self):
        self.result = None
        self.setup_form()
        self.load_beam_types()
        self.load_levels()

    def setup_form(self):
        self.Text = "Create Trusses - CPSK"
        self.Width = 500
        self.Height = 650
        self.StartPosition = WinForms.FormStartPosition.CenterScreen
        self.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        self.MaximizeBox = False

        y = 15

        # === GRID SETTINGS ===
        self.add_label("GRID SETTINGS (from existing grids)", 20, y, bold=True)
        y += 28

        self.add_label("X axis count (along building):", 20, y)
        self.txt_x_count = self.add_textbox("9", 280, y)
        y += 28

        self.add_label("X spacing (mm):", 20, y)
        self.txt_x_spacing = self.add_textbox("12000", 280, y)
        y += 28

        self.add_label("Building width / span (mm):", 20, y)
        self.txt_span = self.add_textbox("24000", 280, y)
        y += 35

        # === TRUSS GEOMETRY ===
        self.add_label("TRUSS GEOMETRY", 20, y, bold=True)
        y += 28

        self.add_label("Truss height at center (mm):", 20, y)
        self.txt_truss_height = self.add_textbox("2736", 280, y)
        y += 28

        self.add_label("Slope angle (degrees):", 20, y)
        self.txt_slope = self.add_textbox("8", 280, y)
        y += 28

        self.add_label("Panel count (per half span):", 20, y)
        self.txt_panel_count = self.add_textbox("4", 280, y)
        y += 28

        self.add_label("Column top height (mm):", 20, y)
        self.txt_col_height = self.add_textbox("7772", 280, y)
        y += 35

        # === MEMBER SIZES ===
        self.add_label("MEMBER SIZES", 20, y, bold=True)
        y += 28

        self.add_label("Top chord size (mm):", 20, y)
        self.txt_top_size = self.add_textbox("115", 280, y)
        y += 28

        self.add_label("Bottom chord size (mm):", 20, y)
        self.txt_bottom_size = self.add_textbox("142", 280, y)
        y += 28

        self.add_label("Diagonal size (mm):", 20, y)
        self.txt_diag_size = self.add_textbox("87", 280, y)
        y += 35

        # === REVIT SETTINGS ===
        self.add_label("REVIT SETTINGS", 20, y, bold=True)
        y += 28

        self.add_label("Beam family type:", 20, y)
        self.cmb_beam_type = WinForms.ComboBox()
        self.cmb_beam_type.Location = Drawing.Point(200, y)
        self.cmb_beam_type.Width = 250
        self.cmb_beam_type.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.Controls.Add(self.cmb_beam_type)
        y += 28

        self.add_label("Base level:", 20, y)
        self.cmb_level = WinForms.ComboBox()
        self.cmb_level.Location = Drawing.Point(200, y)
        self.cmb_level.Width = 250
        self.cmb_level.DropDownStyle = WinForms.ComboBoxStyle.DropDownList
        self.Controls.Add(self.cmb_level)
        y += 45

        # === BUTTONS ===
        btn_create = WinForms.Button()
        btn_create.Text = "Create Trusses"
        btn_create.Location = Drawing.Point(120, y)
        btn_create.Width = 120
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
                'x_count': int(self.txt_x_count.Text),
                'x_spacing': float(self.txt_x_spacing.Text),
                'span': float(self.txt_span.Text),
                'truss_height': float(self.txt_truss_height.Text),
                'slope': float(self.txt_slope.Text),
                'panel_count': int(self.txt_panel_count.Text),
                'col_height': float(self.txt_col_height.Text),
                'top_size': float(self.txt_top_size.Text),
                'bottom_size': float(self.txt_bottom_size.Text),
                'diag_size': float(self.txt_diag_size.Text),
                'beam_type_id': self.beam_types.get(
                    str(self.cmb_beam_type.SelectedItem)),
                'level_id': self.levels.get(
                    str(self.cmb_level.SelectedItem)),
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


def calculate_truss_nodes(params):
    """Calculate top and bottom chord node positions."""
    span = params['span']
    truss_height = params['truss_height']
    slope = params['slope']
    panel_count = params['panel_count']
    col_height = params['col_height']

    total_panels = panel_count * 2
    panel_length = span / total_panels

    z_bottom = col_height  # Bottom chord at column top
    z_top_center = z_bottom + truss_height

    slope_rad = math.radians(slope)
    rise = (span / 2.0) * math.tan(slope_rad)
    z_top_support = z_top_center - rise

    # Top chord nodes (follow roof slope)
    top_nodes = []
    for idx in range(total_panels + 1):
        y_pos = idx * panel_length
        if y_pos <= span / 2.0:
            t = y_pos / (span / 2.0) if span > 0 else 0
            z_pos = z_top_support + t * rise
        else:
            t = (y_pos - span / 2.0) / (span / 2.0) if span > 0 else 0
            z_pos = z_top_center - t * rise
        top_nodes.append((y_pos, z_pos))

    # Bottom chord nodes (horizontal at z_bottom)
    bottom_nodes = []
    for idx in range(total_panels - 1):
        y_pos = (idx + 1) * panel_length
        bottom_nodes.append((y_pos, z_bottom))

    return top_nodes, bottom_nodes, z_bottom, z_top_center


def create_beam(doc, start_pt, end_pt, beam_type, level):
    """Create a structural beam between two points."""
    line = Line.CreateBound(start_pt, end_pt)
    beam = doc.Create.NewFamilyInstance(
        line,
        beam_type,
        level,
        StructuralType.Beam
    )
    return beam


def create_trusses(params):
    """Create Warren-type trusses."""
    x_count = params['x_count']
    x_spacing = mm_to_feet(params['x_spacing'])
    span = params['span']

    beam_type_id = params['beam_type_id']
    level_id = params['level_id']

    if not beam_type_id:
        forms.alert("No beam type selected!")
        return 0, 0, 0

    beam_type = doc.GetElement(beam_type_id)
    level = doc.GetElement(level_id)

    # Calculate node positions (in mm)
    top_nodes, bottom_nodes, z_bottom, z_top_center = calculate_truss_nodes(params)

    # Convert to feet
    top_nodes_ft = [(mm_to_feet(y), mm_to_feet(z)) for y, z in top_nodes]
    bottom_nodes_ft = [(mm_to_feet(y), mm_to_feet(z)) for y, z in bottom_nodes]
    z_bottom_ft = mm_to_feet(z_bottom)

    # Truss positions (skip first and last column - end trusses)
    truss_x_positions = []
    for idx in range(1, x_count - 1):
        truss_x_positions.append(idx * x_spacing)

    top_count = 0
    bottom_count = 0
    diag_count = 0

    with revit.Transaction("Create Trusses"):

        for truss_x in truss_x_positions:

            # === TOP CHORDS ===
            for idx in range(len(top_nodes_ft) - 1):
                y1, z1 = top_nodes_ft[idx]
                y2, z2 = top_nodes_ft[idx + 1]

                start_pt = XYZ(truss_x, y1, z1)
                end_pt = XYZ(truss_x, y2, z2)

                try:
                    create_beam(doc, start_pt, end_pt, beam_type, level)
                    top_count += 1
                except Exception as e:
                    output.print_md("Error creating top chord: {}".format(str(e)))

            # === BOTTOM CHORDS ===
            # Add start and end points for bottom chord
            all_bottom = [(0, z_bottom_ft)] + bottom_nodes_ft + [(mm_to_feet(span), z_bottom_ft)]

            for idx in range(len(all_bottom) - 1):
                y1, z1 = all_bottom[idx]
                y2, z2 = all_bottom[idx + 1]

                start_pt = XYZ(truss_x, y1, z1)
                end_pt = XYZ(truss_x, y2, z2)

                try:
                    create_beam(doc, start_pt, end_pt, beam_type, level)
                    bottom_count += 1
                except Exception as e:
                    output.print_md("Error creating bottom chord: {}".format(str(e)))

            # === DIAGONALS (Warren pattern) ===
            total_panels = params['panel_count'] * 2
            panel_length_ft = mm_to_feet(span) / total_panels

            for idx in range(total_panels):
                # Bottom node for this panel (midpoint)
                if idx < len(bottom_nodes_ft):
                    bottom_y = bottom_nodes_ft[idx][0]
                else:
                    bottom_y = (idx + 0.5) * panel_length_ft

                # Top nodes at panel edges
                top_left_y, top_left_z = top_nodes_ft[idx]
                top_right_y, top_right_z = top_nodes_ft[idx + 1]

                # Diagonal 1: bottom-left to top-right
                pt_bottom = XYZ(truss_x, bottom_y, z_bottom_ft)
                pt_top_left = XYZ(truss_x, top_left_y, top_left_z)
                pt_top_right = XYZ(truss_x, top_right_y, top_right_z)

                try:
                    create_beam(doc, pt_bottom, pt_top_left, beam_type, level)
                    diag_count += 1
                except:
                    pass

                try:
                    create_beam(doc, pt_bottom, pt_top_right, beam_type, level)
                    diag_count += 1
                except:
                    pass

    return top_count, bottom_count, diag_count


# === MAIN ===
if __name__ == "__main__":
    # Check for beam types
    collector = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_StructuralFraming)\
        .WhereElementIsElementType()

    if not list(collector):
        forms.alert(
            "No structural framing families loaded!\n\n"
            "Please load a beam/framing family first.",
            title="No Beam Types"
        )
    else:
        form = TrussForm()
        result = form.ShowDialog()

        if result == WinForms.DialogResult.OK and form.result:
            top, bottom, diag = create_trusses(form.result)

            output.print_md("## Trusses Created")
            output.print_md("- **Top chords**: {}".format(top))
            output.print_md("- **Bottom chords**: {}".format(bottom))
            output.print_md("- **Diagonals**: {}".format(diag))
            output.print_md("- **Total members**: {}".format(top + bottom + diag))
            output.print_md("")
            output.print_md("### Parameters")
            output.print_md("- Truss count: {}".format(form.result['x_count'] - 2))
            output.print_md("- Span: {} mm".format(form.result['span']))
            output.print_md("- Height: {} mm".format(form.result['truss_height']))
            output.print_md("- Slope: {}°".format(form.result['slope']))
