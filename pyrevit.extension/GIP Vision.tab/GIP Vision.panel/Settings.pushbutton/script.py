# coding: utf-8
from pyrevit import forms

from gipvision.config import load_settings, save_settings


def main():
    settings = load_settings()

    api_key = forms.ask_for_string(
        default=settings.get("api_key", ""),
        prompt="Введите X-API-Key (можно оставить пустым для публичного by_plane):",
        title="GIP Vision Settings"
    )
    if api_key is None:
        return

    export_folder = forms.pick_folder()
    if not export_folder:
        export_folder = settings.get("export_folder", "")

    file_prefix = forms.ask_for_string(
        default=settings.get("file_prefix", "gipvision_export"),
        prompt="Префикс IFC-файла:",
        title="GIP Vision Settings"
    )
    if file_prefix is None:
        return

    settings["api_key"] = api_key.strip()
    settings["export_folder"] = export_folder
    settings["file_prefix"] = file_prefix.strip() or "gipvision_export"

    save_settings(settings)
    forms.alert("Настройки сохранены.", title="GIP Vision", exitscript=False)


if __name__ == "__main__":
    main()
