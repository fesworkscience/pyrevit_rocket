# -*- coding: utf-8 -*-
"""Export IFC from the active 3D view for GIP Vision stage 1."""

__title__ = "Экспорт модели\nв GIP Vision"
__author__ = "CPSK"

import clr
import os
import sys
import json
import codecs
import math
from datetime import datetime

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System import Enum
from System.Collections.Generic import List
from System.Drawing import Font, FontStyle, Point, Size
from System.Threading import Thread
from System.Windows.Forms import Application, Button, Clipboard, Form, FormBorderStyle, FormStartPosition, Label
from pyrevit import forms
from pyrevit import revit
from Autodesk.Revit.DB import (
    ElementId,
    Family,
    FamilyInstance,
    FamilyPlacementType,
    FamilySource,
    FamilySymbol,
    FilteredElementCollector,
    IFamilyLoadOptions,
    View3D,
    XYZ,
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException


SCRIPT_DIR = os.path.dirname(__file__)
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
LIB_DIR = os.path.join(EXTENSION_DIR, "lib")
GIP_PANEL_DIR = os.path.dirname(SCRIPT_DIR)
DOOR_FAMILY_PATH = os.path.join(GIP_PANEL_DIR, "GIP_Vision_Door.rfa")
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)


from cpsk_auth import require_auth
from cpsk_config import get_setting, set_setting
from cpsk_logger import Logger
from cpsk_notify import show_confirm, show_error, show_success, show_warning
from gipvision.api_client import (
    create_session_by_plane,
    create_session_by_refimage,
    create_session_by_scan,
    resolve_onetime_code,
)
from gipvision.config import load_settings, save_settings
from gipvision.ifc_export import describe_view3d_state, export_current_3d_view_to_ifc


if not require_auth():
    sys.exit()


SCRIPT_NAME = "GIPVisionExportViewIFC"
FEET_TO_METERS = 0.3048
RESULT_XAML_PATH = os.path.join(GIP_PANEL_DIR, "Create AR Session.pushbutton", "result_ui.xaml")
Logger.init(SCRIPT_NAME)
Logger.info(SCRIPT_NAME, "Script started")

doc = revit.doc
uidoc = revit.uidoc


class FamilyLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        return True, True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        return True, FamilySource.Family, True


class ScanPlacementForm(Form):
    def __init__(self):
        Form.__init__(self)
        self.result = "cancel"

        self.Text = "GIP Vision Door"
        self.Width = 430
        self.Height = 205
        self.TopMost = True
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False
        self.MinimizeBox = False

        lbl_title = Label()
        lbl_title.Text = "Экземпляр двери размещен"
        lbl_title.Font = Font("Segoe UI", 11.0, FontStyle.Bold)
        lbl_title.Location = Point(16, 16)
        lbl_title.Size = Size(380, 24)
        self.Controls.Add(lbl_title)

        lbl_body = Label()
        lbl_body.Text = (
            "Измените параметры 'Ширина' и 'Высота' у выделенного семейства "
            "в палитре свойств Revit, затем нажмите 'Подтвердить'."
        )
        lbl_body.Location = Point(16, 48)
        lbl_body.Size = Size(380, 56)
        self.Controls.Add(lbl_body)

        lbl_hint = Label()
        lbl_hint.Text = (
            "Кнопка 'Переустановить' удалит текущий экземпляр и позволит "
            "выбрать новую грань стены."
        )
        lbl_hint.Location = Point(16, 104)
        lbl_hint.Size = Size(380, 36)
        self.Controls.Add(lbl_hint)

        btn_cancel = Button()
        btn_cancel.Text = "Отмена"
        btn_cancel.Location = Point(16, 144)
        btn_cancel.Size = Size(110, 32)
        btn_cancel.Click += self._on_cancel
        self.Controls.Add(btn_cancel)

        btn_retry = Button()
        btn_retry.Text = "Переустановить"
        btn_retry.Location = Point(136, 144)
        btn_retry.Size = Size(120, 32)
        btn_retry.Click += self._on_retry
        self.Controls.Add(btn_retry)

        btn_ok = Button()
        btn_ok.Text = "Подтвердить"
        btn_ok.Location = Point(266, 144)
        btn_ok.Size = Size(130, 32)
        btn_ok.Click += self._on_confirm
        self.Controls.Add(btn_ok)

    def _on_confirm(self, sender, args):
        self.result = "confirm"
        self.Close()

    def _on_retry(self, sender, args):
        self.result = "retry"
        self.Close()

    def _on_cancel(self, sender, args):
        self.result = "cancel"
        self.Close()


class ResultWindow(forms.WPFWindow):
    def __init__(self, xaml_path, data):
        forms.WPFWindow.__init__(self, xaml_path)
        self.txtPin.Text = _as_text(data.get("onetime_code"))
        self.txtDeeplink.Text = _as_text(data.get("deeplink"))
        self.txtExpires.Text = _as_text(data.get("expires_at"))
        self.txtIfcPath.Text = _as_text(data.get("ifc_path"))
        self.txtResolve.Text = _as_text(data.get("resolve_text"))

    def _copy_text(self, value):
        text = (value or "").strip()
        if not text:
            return
        Clipboard.SetText(text)
        show_success("GIP Vision", "Скопировано в буфер обмена.")

    def btnCopyPin_Click(self, sender, args):
        self._copy_text(self.txtPin.Text)

    def btnCopyDeeplink_Click(self, sender, args):
        self._copy_text(self.txtDeeplink.Text)

    def btnClose_Click(self, sender, args):
        self.Close()


def _is_exportable_3d_view(view):
    return view is not None and isinstance(view, View3D) and not view.IsTemplate


def _get_default_export_folder():
    settings = load_settings()
    saved_folder = (settings.get("export_folder", "") or "").strip()
    if saved_folder:
        return saved_folder

    saved_folder = get_setting("gip_vision.export_folder", "")
    if saved_folder:
        return saved_folder

    doc_path = getattr(doc, "PathName", "")
    if doc_path:
        return os.path.dirname(doc_path)

    return ""


def _get_saved_scenario():
    scenario = get_setting("gip_vision.scenario", "by_plane")
    if scenario not in ("by_plane", "by_refimage", "by_scan"):
        scenario = "by_plane"
    return scenario


def _build_view_info_text(view_state):
    section_box_text = "включен" if view_state.get("is_section_box_active") else "выключен"
    return "ID вида: {0}\nSection Box: {1}".format(
        view_state.get("view_id", ""),
        section_box_text
    )


def _build_success_details(ifc_path, view_state):
    return _build_success_details_with_stub(ifc_path, None, view_state, "by_plane", None)


def _build_success_details_with_stub(ifc_path, request_stub_path, view_state, scenario, selection_data):
    details = [
        "Сценарий: {0}".format(scenario),
        "Вид: {0}".format(view_state.get("view_name", "")),
        "ID вида: {0}".format(view_state.get("view_id", "")),
        "Section Box: {0}".format("включен" if view_state.get("is_section_box_active") else "выключен"),
        "IFC: {0}".format(ifc_path),
        "Лог: {0}".format(Logger.get_log_path())
    ]
    if scenario == "by_refimage" and selection_data:
        details.append(
            "Точка: ({0}, {1}, {2})".format(
                selection_data.get("x"),
                selection_data.get("y"),
                selection_data.get("z")
            )
        )
    elif scenario == "by_scan" and selection_data:
        details.append("Door instance ID: {0}".format(selection_data.get("instance_id")))
        details.append(
            "Размеры: width={0}, height={1}".format(
                selection_data.get("width"),
                selection_data.get("height")
            )
        )
    if request_stub_path:
        details.append("JSON запроса: {0}".format(request_stub_path))
    try:
        details.append("Размер IFC: {0} байт".format(os.path.getsize(ifc_path)))
    except Exception:
        Logger.warning(SCRIPT_NAME, "Unable to read IFC file size for success notification.")
    return "\n".join(details)


def _build_request_stub_path(ifc_path, scenario):
    return os.path.splitext(ifc_path)[0] + ".{0}.request.json".format(scenario)


def _as_text(value):
    if value is None:
        return ""
    return str(value)


def _feet_to_meters(value):
    return _round_value(value * FEET_TO_METERS)


def _serialize_xyz_meters(xyz):
    return {
        "x": _feet_to_meters(xyz.X),
        "y": _feet_to_meters(xyz.Y),
        "z": _feet_to_meters(xyz.Z)
    }


def _serialize_xyz_array_meters(xyz):
    return [
        _feet_to_meters(xyz.X),
        _feet_to_meters(xyz.Y),
        _feet_to_meters(xyz.Z)
    ]


def _normalize_xyz(xyz):
    length = math.sqrt(xyz.X * xyz.X + xyz.Y * xyz.Y + xyz.Z * xyz.Z)
    if length <= 0.0000001:
        raise Exception("Cannot normalize zero-length vector.")
    return {
        "x": xyz.X / length,
        "y": xyz.Y / length,
        "z": xyz.Z / length
    }


def _normalize_revit_xyz(xyz):
    length = math.sqrt(xyz.X * xyz.X + xyz.Y * xyz.Y + xyz.Z * xyz.Z)
    if length <= 0.0000001:
        raise Exception("Cannot normalize zero-length vector.")
    return XYZ(xyz.X / length, xyz.Y / length, xyz.Z / length)


def _round_value(value):
    return round(float(value), 6)


def _compute_angles_from_normal(normal_xyz):
    nx = normal_xyz["x"]
    ny = normal_xyz["y"]
    nz = normal_xyz["z"]

    ry = math.degrees(math.atan2(nx, nz))
    horizontal = math.sqrt((nx * nx) + (nz * nz))
    rx = -math.degrees(math.atan2(ny, horizontal))
    rz = 0.0

    return {
        "rx": _round_value(rx),
        "ry": _round_value(ry),
        "rz": _round_value(rz)
    }


def _serialize_xyz(xyz):
    return {
        "x": _round_value(xyz.X),
        "y": _round_value(xyz.Y),
        "z": _round_value(xyz.Z)
    }


def _serialize_xyz_array(xyz):
    return [
        _round_value(xyz.X),
        _round_value(xyz.Y),
        _round_value(xyz.Z)
    ]


def _get_door_family_name():
    return os.path.splitext(os.path.basename(DOOR_FAMILY_PATH))[0]


def _find_loaded_family_by_name(family_name):
    for family in FilteredElementCollector(doc).OfClass(Family):
        if getattr(family, "Name", "") == family_name:
            return family
    return None


def _get_first_symbol_from_family(family):
    if family is None:
        return None

    symbol_ids = family.GetFamilySymbolIds()
    for symbol_id in symbol_ids:
        symbol = doc.GetElement(symbol_id)
        if isinstance(symbol, FamilySymbol):
            return symbol
    return None


def _ensure_door_family_symbol():
    family_name = _get_door_family_name()
    family = _find_loaded_family_by_name(family_name)

    if family is None:
        if not os.path.exists(DOOR_FAMILY_PATH):
            raise Exception("Door family file not found: {0}".format(DOOR_FAMILY_PATH))

        Logger.info(SCRIPT_NAME, "Loading door family: {0}".format(DOOR_FAMILY_PATH))
        loaded_family = clr.Reference[Family]()
        with revit.Transaction("Load GIP Vision door family"):
            ok = doc.LoadFamily(DOOR_FAMILY_PATH, FamilyLoadOptions(), loaded_family)
            if not ok and loaded_family.Value is None:
                raise Exception("Revit failed to load family: {0}".format(DOOR_FAMILY_PATH))

        family = loaded_family.Value or _find_loaded_family_by_name(family_name)
        if family is None:
            raise Exception("Loaded door family was not found in the document: {0}".format(family_name))

    symbol = _get_first_symbol_from_family(family)
    if symbol is None:
        raise Exception("No family symbol found for door family: {0}".format(family_name))

    if not symbol.IsActive:
        Logger.info(SCRIPT_NAME, "Activating door family symbol: {0}".format(symbol.Id.IntegerValue))
        with revit.Transaction("Activate GIP Vision door symbol"):
            symbol.Activate()
            doc.Regenerate()

    try:
        placement_type = str(getattr(family, "FamilyPlacementType", ""))
        Logger.info(SCRIPT_NAME, "Door family placement type: {0}".format(placement_type))
    except Exception:
        Logger.warning(SCRIPT_NAME, "Unable to read door family placement type.")

    return family, symbol


def _get_parameter_double_value(element, parameter_names):
    if element is None:
        return None

    for name in parameter_names:
        param = element.LookupParameter(name)
        if param and param.HasValue:
            try:
                return param.AsDouble()
            except Exception:
                continue
    return None


def _get_family_instance_origin(instance):
    location = getattr(instance, "Location", None)
    if location is not None and hasattr(location, "Point") and location.Point is not None:
        return location.Point

    transform = instance.GetTransform()
    return transform.Origin


def _build_scan_selection(instance):
    origin = _get_family_instance_origin(instance)
    width = _get_parameter_double_value(instance, ["Ширина", "Width"])
    if width is None:
        width = _get_parameter_double_value(getattr(instance, "Symbol", None), ["Ширина", "Width"])
    if width is None:
        raise Exception("Door family parameter 'Ширина' / 'Width' was not found.")

    height = _get_parameter_double_value(instance, ["Высота", "Height"])
    if height is None:
        height = _get_parameter_double_value(getattr(instance, "Symbol", None), ["Высота", "Height"])
    if height is None:
        raise Exception("Door family parameter 'Высота' / 'Height' was not found.")

    width_direction = None
    if hasattr(instance, "HandOrientation") and instance.HandOrientation is not None:
        width_direction = _normalize_revit_xyz(instance.HandOrientation)
    else:
        width_direction = _normalize_revit_xyz(instance.GetTransform().BasisX)

    wall_normal = None
    if hasattr(instance, "FacingOrientation") and instance.FacingOrientation is not None:
        wall_normal = _normalize_revit_xyz(instance.FacingOrientation)
    else:
        wall_normal = _normalize_revit_xyz(instance.GetTransform().BasisY)

    up_direction = XYZ.BasisZ

    corners = [
        origin,
        origin + (width_direction.Multiply(width)),
        origin + (width_direction.Multiply(width)) + (up_direction.Multiply(height)),
        origin + (up_direction.Multiply(height))
    ]

    scan_selection = {
        "units": "revit_internal_feet",
        "family_path": DOOR_FAMILY_PATH,
        "family_name": getattr(getattr(instance, "Symbol", None), "FamilyName", "") or _get_door_family_name(),
        "instance_id": instance.Id.IntegerValue,
        "symbol_id": getattr(getattr(instance, "Symbol", None), "Id", None).IntegerValue if getattr(instance, "Symbol", None) else "",
        "origin": _serialize_xyz(origin),
        "width": _round_value(width),
        "height": _round_value(height),
        "width_direction": _serialize_xyz(width_direction),
        "wall_normal": _serialize_xyz(wall_normal),
        "corners": [_serialize_xyz_array(corner) for corner in corners]
    }
    Logger.data(SCRIPT_NAME, "by_scan selected door family", scan_selection)
    return scan_selection


def _build_request_payload(scenario, point_selection, scan_selection):
    endpoint_path = "/session/{0}/".format(scenario)
    fields = {}
    notes = []

    if scenario == "by_refimage":
        point = point_selection["point"]
        fields = {
            "point_of_view.x": _feet_to_meters(point["x"]),
            "point_of_view.y": _feet_to_meters(point["y"]),
            "point_of_view.z": _feet_to_meters(point["z"]),
            "point_of_view.rx": point["rx"],
            "point_of_view.ry": point["ry"],
            "point_of_view.rz": point["rz"],
            "point_of_view.coordinate_system": "z-up"
        }
        notes.append("point_of_view converted from Revit internal feet to meters before sending.")
        notes.append("point_of_view.rz is currently fixed to 0 and may require calibration.")
    elif scenario == "by_scan":
        corners_meters = []
        for corner in scan_selection["corners"]:
            corners_meters.append([
                _feet_to_meters(corner[0]),
                _feet_to_meters(corner[1]),
                _feet_to_meters(corner[2])
            ])

        wall_normal = [
            scan_selection["wall_normal"]["x"],
            scan_selection["wall_normal"]["y"],
            scan_selection["wall_normal"]["z"]
        ]
        fields = {
            "alignment.corners": json.dumps(corners_meters, ensure_ascii=False),
            "alignment.wall_normal": json.dumps(wall_normal, ensure_ascii=False),
            "alignment.coordinate_system": "z-up"
        }
        notes.append("alignment.corners converted from Revit internal feet to meters before sending.")
        notes.append("by_scan uses helper family GIP_Vision_Door.rfa placed after IFC export.")
        notes.append("Door insertion point is assumed to be the left bottom corner of the opening.")
        notes.append("Door width uses HandOrientation, height uses global Z-up, and wall normal uses FacingOrientation.")

    return {
        "endpoint_path": endpoint_path,
        "fields": fields,
        "notes": notes
    }


def _pick_scan_face_point():
    if not uidoc:
        raise Exception("Active Revit UI document is not available.")

    Logger.info(SCRIPT_NAME, "Prompting user to pick a wall face point for by_scan.")
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.PointOnElement,
            "Укажите точку на грани стены для размещения GIP Vision Door"
        )
    except OperationCanceledException:
        Logger.info(SCRIPT_NAME, "by_scan face point selection canceled by user.")
        return None

    point = ref.GlobalPoint
    element = doc.GetElement(ref.ElementId)
    if element is None:
        raise Exception("Selected element for by_scan was not found.")

    face = element.GetGeometryObjectFromReference(ref)
    if face is None or not hasattr(face, "Project") or not hasattr(face, "ComputeNormal"):
        raise Exception("Selected point for by_scan is not on a supported face.")

    projected = face.Project(point)
    if projected is None:
        raise Exception("Could not project by_scan point to face UV coordinates.")

    face_normal = _normalize_revit_xyz(face.ComputeNormal(projected.UVPoint))
    if abs(face_normal.Z) > 0.95:
        raise Exception("Для сценария по двери нужно выбирать вертикальную грань стены, а не горизонтальную поверхность.")

    # For this helper family, local +X must go to the user's "right" along the wall.
    # The opposite cross-product mirrored the instance from the left-bottom insertion point.
    width_direction = XYZ.BasisZ.CrossProduct(face_normal)
    try:
        width_direction = _normalize_revit_xyz(width_direction)
    except Exception:
        width_direction = _normalize_revit_xyz(XYZ.BasisX.CrossProduct(face_normal))

    face_pick = {
        "reference": ref,
        "element_id": ref.ElementId.IntegerValue,
        "point": point,
        "face_normal": face_normal,
        "width_direction": width_direction
    }
    Logger.data(SCRIPT_NAME, "by_scan face pick", {
        "element_id": face_pick["element_id"],
        "point": _serialize_xyz(point),
        "face_normal": _serialize_xyz(face_normal),
        "width_direction": _serialize_xyz(width_direction)
    })
    return face_pick


def _select_revit_element(element_id):
    if not uidoc or element_id is None:
        return
    id_list = List[ElementId]([element_id])
    uidoc.Selection.SetElementIds(id_list)
    try:
        uidoc.ShowElements(element_id)
    except Exception as ex:
        Logger.warning(SCRIPT_NAME, "Failed to focus selected element in Revit UI: {0}".format(str(ex)))
        show_warning(
            "GIP Vision",
            "Не удалось показать выбранный элемент в активном окне Revit.",
            details=str(ex),
            blocking=False
        )


def _place_single_scan_instance(symbol, face_pick):
    Logger.info(SCRIPT_NAME, "Creating single by_scan family instance on selected wall face.")
    with revit.Transaction("Place GIP Vision Door"):
        instance = doc.Create.NewFamilyInstance(
            face_pick["reference"],
            face_pick["point"],
            face_pick["width_direction"],
            symbol
        )
        doc.Regenerate()

    if instance is None:
        raise Exception("Revit did not return a family instance for by_scan placement.")

    Logger.info(SCRIPT_NAME, "Created by_scan family instance: {0}".format(instance.Id.IntegerValue))
    _select_revit_element(instance.Id)
    return instance


def _delete_scan_instance(instance_id):
    if instance_id is None:
        return

    existing = doc.GetElement(instance_id)
    if existing is None:
        return

    Logger.info(SCRIPT_NAME, "Deleting by_scan helper family instance: {0}".format(instance_id.IntegerValue))
    with revit.Transaction("Delete GIP Vision Door"):
        doc.Delete(instance_id)


def _wait_for_scan_confirmation(instance_id):
    _select_revit_element(instance_id)
    form = ScanPlacementForm()
    form.Show()

    while form.Visible:
        Application.DoEvents()
        Thread.Sleep(100)

    Logger.info(SCRIPT_NAME, "by_scan confirmation form result: {0}".format(form.result))
    return form.result


def _place_scan_door_family():
    if not uidoc:
        raise Exception("Active Revit UI document is not available.")

    family, symbol = _ensure_door_family_symbol()
    Logger.info(SCRIPT_NAME, "Door family ready for by_scan: {0}".format(getattr(family, "Name", "")))

    while True:
        show_warning(
            "GIP Vision",
            "Сейчас нужно указать грань стены и точку вставки двери.",
            details=(
                "Команда разместит один экземпляр GIP Vision Door на выбранной грани, "
                "после чего можно будет изменить параметры 'Ширина' и 'Высота' "
                "в палитре свойств и подтвердить размещение в отдельном окне."
            )
        )

        face_pick = _pick_scan_face_point()
        if face_pick is None:
            return None

        instance = _place_single_scan_instance(symbol, face_pick)
        action = _wait_for_scan_confirmation(instance.Id)

        if action == "confirm":
            current_instance = doc.GetElement(instance.Id)
            if current_instance is None:
                raise Exception("Placed by_scan family instance was deleted before confirmation.")
            return _build_scan_selection(current_instance)

        _delete_scan_instance(instance.Id)
        if action == "retry":
            Logger.info(SCRIPT_NAME, "by_scan instance deleted for repositioning.")
            continue

        Logger.info(SCRIPT_NAME, "by_scan canceled on confirmation step.")
        return None


def _pick_marker_point():
    if not uidoc:
        raise Exception("Active Revit UI document is not available.")

    Logger.info(SCRIPT_NAME, "Prompting user to pick a point on element surface for by_refimage.")
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.PointOnElement,
            "Укажите точку на поверхности для сценария по маркеру"
        )
    except OperationCanceledException:
        Logger.info(SCRIPT_NAME, "Point selection canceled by user.")
        return None

    point = ref.GlobalPoint
    element = doc.GetElement(ref.ElementId)
    if element is None:
        raise Exception("Selected element was not found.")

    geom_obj = element.GetGeometryObjectFromReference(ref)
    if geom_obj is None or not hasattr(geom_obj, "Project") or not hasattr(geom_obj, "ComputeNormal"):
        raise Exception("Selected point is not on a supported face.")

    projected = geom_obj.Project(point)
    if projected is None:
        raise Exception("Could not project selected point to face UV coordinates.")

    normal_xyz = _normalize_xyz(geom_obj.ComputeNormal(projected.UVPoint))
    rotation = _compute_angles_from_normal(normal_xyz)
    point_of_view = _serialize_xyz(point)
    point_of_view["rx"] = rotation["rx"]
    point_of_view["ry"] = rotation["ry"]
    point_of_view["rz"] = rotation["rz"]
    point_of_view["coordinate_system"] = "z-up"

    selection_debug = {
        "units": "revit_internal_feet",
        "element_id": ref.ElementId.IntegerValue,
        "point": point_of_view,
        "normal": {
            "x": _round_value(normal_xyz["x"]),
            "y": _round_value(normal_xyz["y"]),
            "z": _round_value(normal_xyz["z"])
        }
    }
    Logger.data(SCRIPT_NAME, "by_refimage picked point", selection_debug)
    return selection_debug


def _show_result_dialog(session_data, resolve_data, ifc_path):
    resolve_text = "Проверка resolve не выполнялась."
    if resolve_data:
        if resolve_data.get("ok"):
            resolved = resolve_data.get("data") or {}
            resolve_text = (
                "Проверка resolve: OK\n"
                "Resolve deeplink: {0}\n"
                "Resolve spatial_project_id: {1}"
            ).format(
                _as_text(resolved.get("deeplink")) or "-",
                _as_text(resolved.get("spatial_project_id")) or "-"
            )
        else:
            resolve_text = (
                "Проверка resolve: ошибка HTTP {0}\n{1}"
            ).format(
                resolve_data.get("status"),
                _as_text((resolve_data.get("data") or {}).get("detail"))
                or _as_text(resolve_data.get("raw"))
                or _as_text(resolve_data.get("error"))
            )

    data = {
        "onetime_code": _as_text(session_data.get("onetime_code")),
        "deeplink": _as_text(session_data.get("deeplink")),
        "expires_at": _as_text(session_data.get("expires_at")),
        "ifc_path": ifc_path,
        "resolve_text": resolve_text
    }
    if not os.path.exists(RESULT_XAML_PATH):
        show_warning(
            "GIP Vision",
            "AR-сессия создана, но result_ui.xaml не найден.",
            details="\n".join([
                "PIN: {0}".format(data["onetime_code"]),
                "Deeplink: {0}".format(data["deeplink"]),
                "Expires: {0}".format(data["expires_at"]),
                "IFC: {0}".format(data["ifc_path"]),
                data["resolve_text"]
            ])
        )
        return
    wnd = ResultWindow(RESULT_XAML_PATH, data)
    wnd.show_dialog()


def _show_api_error(title, response_obj):
    detail = ""
    data = response_obj.get("data") if response_obj else None
    if isinstance(data, dict):
        detail = data.get("detail") or data.get("message") or data.get("error") or ""
    raw = response_obj.get("raw") if response_obj else ""
    status = response_obj.get("status") if response_obj else ""
    err = response_obj.get("error") if response_obj else ""

    text = "{0}\nHTTP: {1}".format(title, status)
    if detail:
        text += "\n{0}".format(detail)
    elif raw:
        text += "\n{0}".format(raw)
    elif err:
        text += "\n{0}".format(err)

    show_error("GIP Vision", text)


def _send_session_request(ifc_path, scenario, api_key, point_selection, scan_selection):
    if scenario == "by_plane":
        Logger.data(SCRIPT_NAME, "GIP Vision API request", {
            "scenario": scenario,
            "endpoint_path": "/session/by_plane/",
            "ifc_path": ifc_path,
            "has_api_key": bool(api_key),
            "fields": {}
        })
        return create_session_by_plane(ifc_path, api_key)

    request_payload = _build_request_payload(scenario, point_selection, scan_selection)
    Logger.data(SCRIPT_NAME, "GIP Vision API request", {
        "scenario": scenario,
        "endpoint_path": request_payload["endpoint_path"],
        "ifc_path": ifc_path,
        "has_api_key": bool(api_key),
        "fields": request_payload["fields"]
    })
    if scenario == "by_refimage":
        return create_session_by_refimage(ifc_path, api_key, request_payload["fields"])
    return create_session_by_scan(ifc_path, api_key, request_payload["fields"])


def _save_request_stub(ifc_path, request_stub_path, view_state, scenario, point_selection, scan_selection):
    request_payload = _build_request_payload(scenario, point_selection, scan_selection)
    endpoint_path = request_payload["endpoint_path"]

    model_file_size = 0
    if os.path.exists(ifc_path):
        model_file_size = os.path.getsize(ifc_path)

    payload = {
        "stub_kind": "gip_vision_request_stub",
        "created_at_utc": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "scenario": scenario,
        "request": {
            "method": "POST",
            "url": "https://api-cpsk-superapp.gip.su/api/gip-vision/v1" + endpoint_path,
            "content_type": "multipart/form-data",
            "file_field": "model_file"
        },
        "model_file": {
            "path": ifc_path,
            "name": os.path.basename(ifc_path),
            "size_bytes": model_file_size
        },
        "revit": {
            "document": {
                "title": getattr(doc, "Title", ""),
                "path": getattr(doc, "PathName", "")
            },
            "view": view_state
        },
        "notes": [
            "This JSON mirrors the multipart request fields sent by Export View IFC."
        ]
    }

    if request_payload["fields"]:
        payload["request"]["fields"] = request_payload["fields"]

    if scenario == "by_refimage":
        payload["selection_debug"] = point_selection
    elif scenario == "by_scan":
        payload["selection_debug"] = scan_selection

    payload["notes"].extend(request_payload["notes"])

    with codecs.open(request_stub_path, "w", "utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)

    Logger.file_saved(SCRIPT_NAME, request_stub_path, "GIP Vision request stub")
    return request_stub_path


class ExportWindow(forms.WPFWindow):
    def __init__(self, xaml_path, active_view):
        forms.WPFWindow.__init__(self, xaml_path)
        self.active_view = active_view
        self.result = None
        self.settings = load_settings()

        view_state = describe_view3d_state(active_view)
        self.txtViewName.Text = view_state.get("view_name", "")
        self.txtViewInfo.Text = _build_view_info_text(view_state)
        self.txtExportFolder.Text = _get_default_export_folder()
        self.txtFilePrefix.Text = (
            self.settings.get("file_prefix")
            or get_setting("gip_vision.file_prefix", "gipvision_export_current_view")
        )
        self.txtApiKey.Text = self.settings.get("api_key", "")
        self.chkResolve.IsChecked = bool(self.settings.get("auto_resolve", True))
        self._set_selected_scenario(_get_saved_scenario())

    def _set_selected_scenario(self, scenario):
        for item in self.cmbScenario.Items:
            if getattr(item, "Tag", None) == scenario:
                self.cmbScenario.SelectedItem = item
                return
        self.cmbScenario.SelectedIndex = 0

    def _get_selected_scenario(self):
        item = self.cmbScenario.SelectedItem
        if item is None:
            return "by_plane"
        tag = getattr(item, "Tag", None)
        return tag or "by_plane"

    def btnBrowse_Click(self, sender, args):
        folder = forms.pick_folder()
        if folder:
            self.txtExportFolder.Text = folder

    def btnCancel_Click(self, sender, args):
        Logger.info(SCRIPT_NAME, "Export dialog canceled by user.")
        self.Close()

    def btnRun_Click(self, sender, args):
        export_folder = (self.txtExportFolder.Text or "").strip()
        file_prefix = (self.txtFilePrefix.Text or "").strip()
        api_key = (self.txtApiKey.Text or "").strip()
        scenario = self._get_selected_scenario()

        if not export_folder:
            Logger.warning(SCRIPT_NAME, "Validation failed: export folder is empty.")
            show_warning("GIP Vision", "Укажите папку для IFC экспорта.")
            return
        if not file_prefix:
            Logger.warning(SCRIPT_NAME, "Validation failed: file prefix is empty.")
            show_warning("GIP Vision", "Укажите имя IFC файла.")
            return

        self.result = {
            "api_key": api_key,
            "export_folder": export_folder,
            "file_prefix": file_prefix,
            "scenario": scenario,
            "auto_resolve": bool(self.chkResolve.IsChecked)
        }
        Logger.data(SCRIPT_NAME, "Export dialog result", self.result)
        self.Close()


def main():
    if not doc:
        Logger.error(SCRIPT_NAME, "Active Revit document is not available.")
        show_error("GIP Vision", "Нет активного документа Revit.")
        return

    active_view = doc.ActiveView
    if not _is_exportable_3d_view(active_view):
        Logger.warning(SCRIPT_NAME, "Active view is not an exportable 3D view.")
        show_warning(
            "GIP Vision",
            "Для этого прототипа нужно открыть непустой 3D-вид.",
            details="Активный вид должен быть 3D и не быть шаблоном."
        )
        return

    view_state = describe_view3d_state(active_view)
    Logger.data(SCRIPT_NAME, "Active view before export", view_state)

    xaml_file = os.path.join(SCRIPT_DIR, "ui.xaml")
    wnd = ExportWindow(xaml_file, active_view)
    wnd.show_dialog()

    if not wnd.result:
        Logger.info(SCRIPT_NAME, "No export parameters received. Script finished without export.")
        return

    save_settings({
        "api_key": wnd.result["api_key"],
        "export_folder": wnd.result["export_folder"],
        "file_prefix": wnd.result["file_prefix"],
        "auto_resolve": wnd.result["auto_resolve"]
    })
    set_setting("gip_vision.export_folder", wnd.result["export_folder"])
    set_setting("gip_vision.file_prefix", wnd.result["file_prefix"])
    set_setting("gip_vision.scenario", wnd.result["scenario"])
    Logger.info(SCRIPT_NAME, "Export settings saved to gipvision.config and cpsk_config.")

    point_selection = None
    scan_selection = None
    if wnd.result["scenario"] == "by_refimage":
        point_selection = _pick_marker_point()
        if point_selection is None:
            Logger.info(SCRIPT_NAME, "Script finished after by_refimage point selection cancel.")
            return

    confirm_details = "\n".join([
        "Документ: {0}".format(getattr(doc, "Title", "")),
        "Сценарий: {0}".format(wnd.result["scenario"]),
        "Вид: {0}".format(view_state.get("view_name", "")),
        "Section Box: {0}".format("включен" if view_state.get("is_section_box_active") else "выключен"),
        "Папка: {0}".format(wnd.result["export_folder"]),
        "Файл: {0}".format(wnd.result["file_prefix"])
    ])
    if point_selection:
        confirm_details += "\nТочка: ({0}, {1}, {2})".format(
            point_selection["point"]["x"],
            point_selection["point"]["y"],
            point_selection["point"]["z"]
        )
    if wnd.result["scenario"] == "by_scan":
        confirm_details += "\nПосле экспорта IFC Revit предложит разместить семейство GIP_Vision_Door.rfa."

    if not show_confirm("GIP Vision", "Начать экспорт IFC из активного 3D-вида?", details=confirm_details):
        Logger.info(SCRIPT_NAME, "Export canceled on confirmation step.")
        return

    try:
        Logger.info(SCRIPT_NAME, "Opening Revit transaction for IFC export.")
        with revit.Transaction("GIP Vision IFC Export Current View"):
            ifc_path = export_current_3d_view_to_ifc(
                doc,
                active_view,
                wnd.result["export_folder"],
                wnd.result["file_prefix"],
                SCRIPT_NAME
            )
        Logger.info(SCRIPT_NAME, "Revit transaction finished.")
        if wnd.result["scenario"] == "by_scan":
            scan_selection = _place_scan_door_family()
            if scan_selection is None:
                show_warning(
                    "GIP Vision",
                    "IFC уже экспортирован, но размещение семейства двери отменено.",
                    details="JSON-заглушка по сценарию by_scan не была сохранена."
                )
                return
        request_stub_path = _build_request_stub_path(ifc_path, wnd.result["scenario"])
        Logger.info(SCRIPT_NAME, "Saving request trace: {0}".format(request_stub_path))
        _save_request_stub(
            ifc_path,
            request_stub_path,
            view_state,
            wnd.result["scenario"],
            point_selection,
            scan_selection
        )
        Logger.info(SCRIPT_NAME, "Sending GIP Vision API request for scenario: {0}".format(wnd.result["scenario"]))
        response = _send_session_request(
            ifc_path,
            wnd.result["scenario"],
            wnd.result["api_key"],
            point_selection,
            scan_selection
        )
        if not response.get("ok"):
            Logger.data(SCRIPT_NAME, "GIP Vision API error response", {
                "status": response.get("status"),
                "data": response.get("data"),
                "raw": response.get("raw"),
                "error": response.get("error")
            })
            Logger.result(SCRIPT_NAME, False, "GIP Vision API request failed.", details=[
                ifc_path,
                request_stub_path
            ])
            _show_api_error("Не удалось создать AR-сессию", response)
            return

        payload = response.get("data") or {}
        code = payload.get("onetime_code")
        if not code:
            show_error(
                "GIP Vision",
                "Сессия создана, но сервер не вернул onetime_code.\nОтвет:\n{0}".format(response.get("raw") or payload)
            )
            return

        resolve_resp = None
        if wnd.result.get("auto_resolve"):
            Logger.info(SCRIPT_NAME, "Sending resolve-by-onetime-code request.")
            resolve_resp = resolve_onetime_code(code, wnd.result["api_key"])

        Logger.result(SCRIPT_NAME, True, "IFC export and GIP Vision API request completed successfully.", details=[
            ifc_path,
            request_stub_path
        ])
        _show_result_dialog(payload, resolve_resp, ifc_path)
    except Exception as ex:
        Logger.exception(SCRIPT_NAME, "IFC export failed.")
        show_error(
            "GIP Vision",
            "Не удалось экспортировать IFC, отправить API-запрос или сохранить JSON трассировки.",
            details=str(ex)
        )


if __name__ == "__main__":
    main()
