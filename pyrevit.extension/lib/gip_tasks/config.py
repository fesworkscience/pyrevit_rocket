# -*- coding: utf-8 -*-
import json
import os


DEFAULTS = {
    "API_BASE_URL": "",
    "AUTH_TOKEN": "",
    "COMPANY_ID": "demo-company",
    "CURRENT_PROJECT_ID": "",
    "CURRENT_PROJECT_NAME": "",
    "CURRENT_PROJECT_CODE": "",
    "CURRENT_DISCIPLINE": "КЖ",
    "CURRENT_USER": "",
    "CURRENT_USER_ROLE": "project_coordinator",
    "LOCAL_MODE": True,
    "LOCAL_CACHE_PATH": "",
    "API_TIMEOUT_SECONDS": 3,
    "POLLING_INTERVAL_SECONDS": 30,
    "MARKER_FAMILY_NAME": "CPSK_Маркер задания",
    "MARKER_FAMILY_TYPE": "Новое",
    "MARKER_FAMILY_PATH": r"C:\Users\saukouma\Documents\BIM библиотека\2-Семейства\Задания\CPSK_Маркер задания.rfa",
}


def _ensure_dir(path):
    if path and not os.path.isdir(path):
        os.makedirs(path)
    return path


def appdata_dir():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return _ensure_dir(os.path.join(base, "GIPTasks"))


def extension_root():
    return os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir, os.pardir))


def config_path():
    return os.path.join(appdata_dir(), "config.json")


def _normalize(data):
    data["MARKER_FAMILY_NAME"] = DEFAULTS["MARKER_FAMILY_NAME"]
    data["MARKER_FAMILY_TYPE"] = DEFAULTS["MARKER_FAMILY_TYPE"]
    if not data.get("CURRENT_DISCIPLINE"):
        data["CURRENT_DISCIPLINE"] = DEFAULTS["CURRENT_DISCIPLINE"]
    if not data.get("LOCAL_CACHE_PATH"):
        data["LOCAL_CACHE_PATH"] = os.path.join(appdata_dir(), "gip_tasks_cache.sqlite")
    data["LOCAL_MODE"] = not bool((data.get("API_BASE_URL") or "").strip())
    return data


def load():
    data = DEFAULTS.copy()
    path = config_path()
    if os.path.exists(path):
        try:
            with open(path, "rb") as fp:
                data.update(json.loads(fp.read().decode("utf-8")))
        except Exception:
            pass
    return _normalize(data)


def save(values):
    data = DEFAULTS.copy()
    data.update(values or {})
    data = _normalize(data)
    _ensure_dir(os.path.dirname(config_path()))
    raw = json.dumps(data, indent=2, ensure_ascii=False, sort_keys=True)
    with open(config_path(), "wb") as fp:
        fp.write(raw.encode("utf-8"))
    return data


def set_current_project(project):
    data = load()
    data["CURRENT_PROJECT_ID"] = project.get("id") or project.get("project_id") or ""
    data["CURRENT_PROJECT_NAME"] = project.get("name") or project.get("project_name") or ""
    data["CURRENT_PROJECT_CODE"] = project.get("code") or project.get("project_code") or ""
    if project.get("user_role"):
        data["CURRENT_USER_ROLE"] = project.get("user_role")
    return save(data)


def can_create_project(settings=None):
    role = (settings or load()).get("CURRENT_USER_ROLE") or ""
    return role in ("admin", "project_coordinator", "lead_specialist")


def is_local_mode(settings=None):
    settings = settings or load()
    return bool(settings.get("LOCAL_MODE")) or not bool((settings.get("API_BASE_URL") or "").strip())
