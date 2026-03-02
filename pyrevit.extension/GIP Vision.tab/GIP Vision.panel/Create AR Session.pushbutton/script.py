# coding: utf-8
import os
import traceback

import clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Clipboard
from pyrevit import forms
from pyrevit import revit
from pyrevit import script

from gipvision.config import load_settings, save_settings
from gipvision.ifc_export import export_active_document_to_ifc
from gipvision.api_client import create_session_by_plane, resolve_onetime_code


logger = script.get_logger()
output = script.get_output()

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document if uidoc else None


class SessionWindow(forms.WPFWindow):
    def __init__(self, xaml_path, settings):
        forms.WPFWindow.__init__(self, xaml_path)
        self.settings = settings
        self.result = None

        self.txtApiKey.Text = settings.get("api_key", "")
        self.txtExportFolder.Text = settings.get("export_folder", "")
        self.txtFilePrefix.Text = settings.get("file_prefix", "gipvision_export")
        self.chkResolve.IsChecked = bool(settings.get("auto_resolve", True))

    def btnBrowse_Click(self, sender, args):
        folder = forms.pick_folder()
        if folder:
            self.txtExportFolder.Text = folder

    def btnCancel_Click(self, sender, args):
        self.Close()

    def btnRun_Click(self, sender, args):
        export_folder = (self.txtExportFolder.Text or "").strip()
        file_prefix = (self.txtFilePrefix.Text or "").strip()
        api_key = (self.txtApiKey.Text or "").strip()

        if not export_folder:
            forms.alert("Укажите папку для IFC", exitscript=False)
            return
        if not file_prefix:
            forms.alert("Укажите имя IFC файла", exitscript=False)
            return

        self.result = {
            "api_key": api_key,
            "export_folder": export_folder,
            "file_prefix": file_prefix,
            "auto_resolve": bool(self.chkResolve.IsChecked)
        }
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
        forms.alert("Скопировано в буфер обмена.", title="GIP Vision", exitscript=False)

    def btnCopyPin_Click(self, sender, args):
        self._copy_text(self.txtPin.Text)

    def btnCopyDeeplink_Click(self, sender, args):
        self._copy_text(self.txtDeeplink.Text)

    def btnClose_Click(self, sender, args):
        self.Close()


def _as_text(value):
    if value is None:
        return ""
    return str(value)


def _show_result_dialog(session_data, resolve_data, ifc_path):
    resolve_text = "Проверка resolve не выполнялась."
    if resolve_data:
        if resolve_data.get("ok"):
            r = resolve_data.get("data") or {}
            resolve_text = (
                "Проверка resolve: OK\n"
                "Resolve deeplink: {0}\n"
                "Resolve spatial_project_id: {1}"
            ).format(_as_text(r.get("deeplink")) or "-", _as_text(r.get("spatial_project_id")) or "-")
        else:
            resolve_text = (
                "Проверка resolve: ошибка HTTP {0}\n{1}"
            ).format(
                resolve_data.get("status"),
                _as_text((resolve_data.get("data") or {}).get("detail")) or _as_text(resolve_data.get("raw")) or _as_text(resolve_data.get("error"))
            )

    data = {
        "onetime_code": _as_text(session_data.get("onetime_code")),
        "deeplink": _as_text(session_data.get("deeplink")),
        "expires_at": _as_text(session_data.get("expires_at")),
        "ifc_path": ifc_path,
        "resolve_text": resolve_text
    }
    xaml_file = os.path.join(os.path.dirname(__file__), "result_ui.xaml")
    wnd = ResultWindow(xaml_file, data)
    wnd.show_dialog()


def _show_error(title, response_obj):
    detail = ""
    data = response_obj.get("data") if response_obj else None
    if isinstance(data, dict):
        detail = data.get("detail") or data.get("message") or ""
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

    forms.alert(text, title="GIP Vision", exitscript=False)


def main():
    if not doc:
        forms.alert("Нет активного документа Revit.", title="GIP Vision", exitscript=False)
        return

    settings = load_settings()
    xaml_file = os.path.join(os.path.dirname(__file__), "ui.xaml")
    wnd = SessionWindow(xaml_file, settings)
    wnd.show_dialog()

    if not wnd.result:
        return

    save_settings(wnd.result)

    try:
        proceed = forms.alert(
            "Начать экспорт в IFC?",
            title="GIP Vision",
            yes=True,
            no=True,
            exitscript=False
        )
        if not proceed:
            return

        with revit.Transaction("GIP Vision IFC Export"):
            ifc_path = export_active_document_to_ifc(
                doc,
                wnd.result["export_folder"],
                wnd.result["file_prefix"]
            )

        response = create_session_by_plane(ifc_path, wnd.result["api_key"])
        if not response.get("ok"):
            _show_error("Не удалось создать AR-сессию", response)
            return

        payload = response.get("data") or {}
        code = payload.get("onetime_code")
        if not code:
            forms.alert(
                "Сессия создана, но сервер не вернул onetime_code.\nОтвет:\n{0}".format(response.get("raw") or payload),
                title="GIP Vision",
                exitscript=False
            )
            return

        resolve_resp = None
        if wnd.result.get("auto_resolve"):
            resolve_resp = resolve_onetime_code(code, wnd.result["api_key"])

        _show_result_dialog(payload, resolve_resp, ifc_path)

        output.print_md("### GIP Vision session")
        output.print_md("- IFC: `{0}`".format(ifc_path))
        output.print_md("- One-time code: `{0}`".format(code))
        output.print_md("- Deeplink: `{0}`".format(payload.get("deeplink")))
        output.print_md("- Expires at (UTC): `{0}`".format(payload.get("expires_at")))

    except Exception as ex:
        logger.error(str(ex))
        logger.debug(traceback.format_exc())
        forms.alert("Ошибка выполнения:\n{0}".format(ex), title="GIP Vision", exitscript=False)


if __name__ == "__main__":
    main()
