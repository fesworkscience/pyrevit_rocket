# coding: utf-8
import io
import json
import os

try:
    from System import Environment
except Exception:
    Environment = None

_SETTINGS_DIR_NAME = "GIPVision"
_SETTINGS_FILE_NAME = "settings.json"


def _appdata_path():
    if Environment:
        return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    return os.path.expanduser("~")


def _settings_path():
    base = _appdata_path()
    return os.path.join(base, "pyRevit", _SETTINGS_DIR_NAME, _SETTINGS_FILE_NAME)


def get_default_export_folder():
    return os.path.join(_appdata_path(), "pyRevit", _SETTINGS_DIR_NAME, "exports")


def load_settings():
    path = _settings_path()
    defaults = {
        "api_key": "",
        "export_folder": get_default_export_folder(),
        "file_prefix": "gipvision_export",
        "auto_resolve": True
    }
    if not os.path.exists(path):
        return defaults

    try:
        with io.open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        defaults.update(data)
        return defaults
    except Exception:
        return defaults


def save_settings(settings_dict):
    path = _settings_path()
    folder = os.path.dirname(path)
    if not os.path.exists(folder):
        os.makedirs(folder)

    with io.open(path, "w", encoding="utf-8") as f:
        json.dump(settings_dict, f, ensure_ascii=False, indent=2)
